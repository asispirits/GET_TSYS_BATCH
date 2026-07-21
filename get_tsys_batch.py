#!/usr/bin/env python3
"""Generate the store-level TSYS PAX batch-not-closed report.

The report flags an active TSYS store when it has approved authorization
activity in the configured authorization window and no accepted batch for the
store's account in the report window. It deliberately does not claim that a
specific termID or physical device failed to batch. Optional configuration
rows provide display overrides only; MXConnect's active merchant roster is the
authoritative store list.
"""

import csv
import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from argparse import ArgumentParser
from datetime import datetime, time, timedelta
from pathlib import Path
from sys import stdout
from tkinter import filedialog, messagebox, ttk
from urllib.parse import quote

import requests


BASE_URL = "https://api.mxconnect.com"
AUTH_PATH = "/security/v1/apiKey/authenticate"
BATCH_PATH = "/report/v1/tsys/batch/export"
AUTHORIZATION_PATH = "/report/v1/tsys/authorization/export"
UAR_PATH = "/boarding/v1/uar"
TSYS_PRODUCT_ID = "3"
EMAIL_KEYS = ["URL", "STORENAME", "DEVICE", "AMOUNT", "TERMID"]
REVIEW_KEYS = [
    "reason",
    "URL",
    "STORENAME",
    "DEVICE",
    "accountNumber",
    "termID",
    "terminalNumber",
    "details",
]
DETAIL_KEYS = [
    "accountNumber",
    "STORENAME",
    "terminalNumber",
    "approvedAmount",
    "approvedCount",
    "batchStatus",
]
RAW_BATCH_KEYS = [
    "created", "rejected", "batchNumber", "accountNumber", "termID",
    "salesAmount", "salesCount", "refundAmount", "refundCount", "netAmount",
    "netCount", "PPSNotFundedTotal", "PPSFundedTotal", "bankNumber", "batchDate",
    "fileId", "filePath", "fileName", "id", "accountId", "entityId", "acl",
    "labels", "locationId", "name", "uar", "domain",
]


def text(value):
    if value is None:
        return ""
    return str(value).strip()


def parse_amount(value):
    if value is None or value == "":
        return 0.0
    try:
        return float(str(value).replace(",", "").replace("$", "").strip())
    except (TypeError, ValueError):
        return 0.0


def is_approved(record):
    return text(record.get("authorizationResponseStatus")).casefold() == "approved"


def is_accepted_batch(record):
    return text(record.get("rejected")).casefold() in {"no", "false", "0", "n"}


def clean_api_key(value):
    value = text(value)
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        value = value[1:-1].strip()
    return value


def application_entrypoint():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(__file__).resolve()


def application_cli_entrypoint():
    entrypoint = application_entrypoint()
    if getattr(sys, "frozen", False):
        cli_entrypoint = entrypoint.with_name("TSYS_PAX_BATCH_REPORT_CLI.exe")
        if cli_entrypoint.exists():
            return cli_entrypoint
    return entrypoint


def load_config(path):
    with path.open("r", encoding="utf-8") as input_file:
        config = json.load(input_file)
    if not isinstance(config, dict):
        raise ValueError(f"Config must be a JSON object: {path}")
    devices = config.get("devices")
    if not isinstance(devices, list):
        raise ValueError(f"Config must contain a devices array: {path}")
    return config, devices


def resolve_output_directory(config_path, config):
    raw_output = text(config.get("outputDirectory")) or "./tsys-auditdata"
    output_path = Path(raw_output).expanduser()
    if not output_path.is_absolute():
        output_path = config_path.parent / output_path
    return output_path.resolve()


def portable_output_directory(config_path, output_path):
    selected = Path(text(output_path)).expanduser()
    if not selected.is_absolute():
        return text(output_path) or "./tsys-auditdata"
    relative = os.path.relpath(str(selected), str(config_path.parent))
    return relative.replace(os.sep, "/") or "."


def save_config(path, config, devices, output_directory=None):
    updated = dict(config)
    if output_directory is not None:
        updated["outputDirectory"] = portable_output_directory(path, output_directory)
    updated["devices"] = devices
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    with temporary_path.open("w", encoding="utf-8", newline="\n") as output_file:
        json.dump(updated, output_file, indent=2, ensure_ascii=False)
        output_file.write("\n")
    temporary_path.replace(path)


def write_csv(rows, path, keys):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(path.suffix + ".tmp")
    with temporary_path.open("w", encoding="utf-8", newline="") as output_file:
        writer = csv.DictWriter(output_file, fieldnames=keys, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    temporary_path.replace(path)


class MxConnectClient:
    def __init__(self, base_url, api_key):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})
        self.token = None

    def authenticate(self):
        response = self.session.post(
            f"{self.base_url}{AUTH_PATH}",
            json={"value": self.api_key},
            headers={"Content-Type": "application/json"},
            timeout=90,
        )
        if response.status_code != 200:
            body = response.text[:500].replace("\n", " ")
            raise RuntimeError(
                f"MXConnect authentication failed with {response.status_code} "
                f"at {self.base_url}{AUTH_PATH}: {body}"
            )
        data = response.json()
        self.token = text(data.get("token"))
        if not self.token:
            raise RuntimeError("MXConnect authentication response did not contain a token.")

    def request(self, method, url, payload=None, retry=True):
        if not self.token:
            self.authenticate()

        headers = {"Authorization": f"Bearer {self.token}"}
        if method == "POST":
            headers["Content-Type"] = "application/json"
            response = self.session.post(url, json=payload or {}, headers=headers, timeout=120)
        else:
            response = self.session.get(url, headers=headers, timeout=120)

        if response.status_code == 401 and retry:
            self.token = None
            self.authenticate()
            return self.request(method, url, payload, retry=False)

        if response.status_code >= 400:
            body = response.text[:500].replace("\n", " ")
            raise RuntimeError(f"MXConnect request failed with {response.status_code}: {body}")
        return response.json()


def response_page(data):
    if isinstance(data, list):
        return data, None
    if not isinstance(data, dict):
        raise ValueError("Unexpected MXConnect response shape.")
    page = data.get("records")
    if page is None:
        page = data.get("data")
    if page is None:
        page = []
    if not isinstance(page, list):
        raise ValueError("Unexpected MXConnect records response shape.")
    total = data.get("totalRecords")
    try:
        total = int(total) if total is not None else None
    except (TypeError, ValueError):
        total = None
    return page, total


def fetch_all(client, url):
    records = []
    payload = {}
    data = client.request("POST", url, payload)
    expected_total = None

    while True:
        page, total = response_page(data)
        if total is not None:
            expected_total = total
        records.extend(page)
        if not page:
            if expected_total is not None and len(records) != expected_total:
                raise RuntimeError(
                    f"MXConnect returned an incomplete export: received {len(records)} "
                    f"of {expected_total} records."
                )
            break
        if expected_total is not None and len(records) >= expected_total:
            if len(records) != expected_total:
                raise RuntimeError(
                    f"MXConnect returned {len(records)} of {expected_total} records."
                )
            break

        scroll_id = data.get("_scroll_id") or data.get("scrollId")
        if not scroll_id:
            raise RuntimeError("MXConnect returned more records but no scroll ID.")
        payload = {"scrollId": scroll_id}
        data = client.request("POST", url, payload)
        print(f"Fetched {len(records)} records" + (f" of {expected_total}" if expected_total else ""))
        stdout.flush()

    return records


def date_window(days):
    end_time = datetime.combine(datetime.today(), time.min) + timedelta(hours=4)
    return end_time - timedelta(days=days), end_time


def batch_url(base_url, start_date, end_date):
    filter_object = {
        "must": [
            {"bool": {"should": [], "minimum_should_match": 1}},
            {
                "bool": {
                    "should": [
                        {
                            "range": {
                                "batchDate": {
                                    "gte": f"{start_date}T00:00:00.000Z",
                                    "lte": f"{end_date}T23:59:59.999Z",
                                }
                            }
                        }
                    ],
                    "minimum_should_match": 1,
                }
            },
        ]
    }
    encoded_filter = quote(json.dumps(filter_object, separators=(",", ":")), safe="")
    return (
        f"{base_url}{BATCH_PATH}?filter={encoded_filter}&size=100&from=0"
    )


def authorization_url(base_url, auth_window, account_numbers):
    fields = [
        "accountNumber",
        "terminalNumber",
        "authorizedAmount",
        "authorizationResponseStatus",
    ]
    encoded_accounts = quote(json.dumps(sorted(account_numbers)), safe="")
    encoded_fields = quote(json.dumps(fields, separators=(",", ":")), safe="")
    return (
        f"{base_url}{AUTHORIZATION_PATH}?dr_type=q&dr_quick={quote(auth_window)}"
        f"&accountNumbers={encoded_accounts}&fields={encoded_fields}&size=100&from=0"
    )


def uar_url(base_url, offset):
    return f"{base_url}{UAR_PATH}?size=100&from={offset}"


def fetch_uar(client):
    records = []
    offset = 0
    while True:
        data = client.request("GET", uar_url(client.base_url, offset))
        if isinstance(data, list):
            page = data
            total = None
        else:
            page, total = response_page(data)
        if not page:
            break
        records.extend(page)
        offset += len(page)
        if total is not None and offset >= total:
            break
        if len(page) < 100:
            break
        print(f"Fetched {offset} UAR records" + (f" of {total}" if total else ""))
        stdout.flush()
    return records


def active_tsys_accounts(uar_records):
    accounts = set()
    for record in uar_records:
        account = record.get("account") or {}
        if (
            text(account.get("number"))
            and text((account.get("product") or {}).get("id")) == TSYS_PRODUCT_ID
            and account.get("active") is True
            and text(account.get("status")) not in {"Closed", "Terminated", "Suspended"}
        ):
            accounts.add(text(account.get("number")))
    return accounts


def active_tsys_roster(uar_records):
    roster = {}
    for record in uar_records:
        account = record.get("account") or {}
        account_number = text(account.get("number"))
        if not account_number:
            continue
        if (
            text((account.get("product") or {}).get("id")) != TSYS_PRODUCT_ID
            or account.get("active") is not True
            or text(account.get("status")) in {"Closed", "Terminated", "Suspended"}
        ):
            continue
        roster[account_number] = {
            "storeName": text((record.get("location") or {}).get("name")),
            "status": text(account.get("status")),
        }
    return roster


def store_display_overrides(raw_devices):
    """Use configured account-level values only as display overrides.

    Account numbers are optional here. They are not used to validate a device;
    they only allow a manually maintained URL/store-name override to follow a
    known merchant account.
    """
    overrides = {}
    conflicts = set()
    for raw in raw_devices:
        if raw.get("enabled", True) is False:
            continue
        account = text(raw.get("accountNumber"))
        if not account:
            continue
        candidate = {
            "URL": text(raw.get("url")),
            "STORENAME": text(raw.get("storeName")),
            "DEVICE": text(raw.get("device")) or "PAX",
        }
        existing = overrides.get(account)
        if existing is not None and existing != candidate:
            conflicts.add(account)
        else:
            overrides[account] = candidate
    return overrides, conflicts


def create_store_report(uar_records, raw_devices, auth_summary, current_batch_records):
    roster = active_tsys_roster(uar_records)
    overrides, override_conflicts = store_display_overrides(raw_devices)
    batched_accounts = {
        text(record.get("accountNumber"))
        for record in current_batch_records
        if text(record.get("accountNumber"))
    }

    auth_by_account = {}
    for (account, terminal), detail in auth_summary.items():
        account_summary = auth_by_account.setdefault(
            account,
            {"amount": 0.0, "count": 0, "terminals": []},
        )
        account_summary["amount"] += detail["amount"]
        account_summary["count"] += detail["count"]
        account_summary["terminals"].append((terminal, detail))

    email_rows = []
    detail_rows = []
    review_rows = []
    for account in sorted(auth_by_account):
        if account in batched_accounts:
            continue
        merchant = roster.get(account)
        if merchant is None:
            review_rows.append(
                {
                    "reason": "ACCOUNT_NOT_FOUND_IN_ACTIVE_TSYS_ROSTER",
                    "URL": "",
                    "STORENAME": "",
                    "DEVICE": "PAX",
                    "accountNumber": account,
                    "termID": "",
                    "terminalNumber": "",
                    "details": "",
                }
            )
            continue

        override = overrides.get(account, {})
        if account in override_conflicts:
            review_rows.append(
                {
                    "reason": "CONFLICTING_STORE_DISPLAY_OVERRIDES",
                    "URL": "",
                    "STORENAME": merchant["storeName"],
                    "DEVICE": "PAX",
                    "accountNumber": account,
                    "termID": "",
                    "terminalNumber": "",
                    "details": "More than one configured display row exists for this account.",
                }
            )
            continue

        store_name = override.get("STORENAME") or merchant["storeName"]
        if not store_name:
            review_rows.append(
                {
                    "reason": "MISSING_STORE_NAME",
                    "URL": override.get("URL", ""),
                    "STORENAME": "",
                    "DEVICE": override.get("DEVICE", "PAX"),
                    "accountNumber": account,
                    "termID": "",
                    "terminalNumber": "",
                    "details": "Active TSYS account has no location name.",
                }
            )
            continue

        summary = auth_by_account[account]
        email_rows.append(
            {
                "URL": override.get("URL", ""),
                "STORENAME": store_name,
                "DEVICE": override.get("DEVICE", "PAX"),
                "AMOUNT": f"{summary['amount']:.2f}",
                "TERMID": "",
            }
        )
        for terminal, detail in summary["terminals"]:
            detail_rows.append(
                {
                    "accountNumber": account,
                    "STORENAME": store_name,
                    "terminalNumber": terminal,
                    "approvedAmount": f"{detail['amount']:.2f}",
                    "approvedCount": detail["count"],
                    "batchStatus": "No accepted batch for account in report window",
                }
            )

    return email_rows, detail_rows, review_rows


def build_auth_summary(records, require_approved):
    summary = {}
    for record in records:
        if require_approved and not is_approved(record):
            continue
        account = text(record.get("accountNumber"))
        if not account:
            continue
        terminal = text(record.get("terminalNumber"))
        entry = summary.setdefault(
            (account, terminal), {"amount": 0.0, "count": 0}
        )
        entry["amount"] += parse_amount(record.get("authorizedAmount"))
        entry["count"] += 1
    return summary


def batch_index(records):
    by_pair = {}
    term_accounts = {}
    for record in records:
        account = text(record.get("accountNumber"))
        term = text(record.get("termID"))
        if not account or not term:
            continue
        pair = (account, term)
        batch_date = text(record.get("batchDate"))
        if pair not in by_pair or batch_date > by_pair[pair]["lastBatchDate"]:
            by_pair[pair] = {"lastBatchDate": batch_date}
        term_accounts.setdefault(term, set()).add(account)
    return by_pair, term_accounts


def write_raw_batch_history(records, path):
    keys = list(RAW_BATCH_KEYS)
    for record in records:
        for key in record:
            if key not in keys:
                keys.append(key)
    write_csv(records, path, keys)


def write_term_history(records, path):
    by_pair, _ = batch_index(records)
    rows = []
    for (account, term), detail in sorted(by_pair.items()):
        rows.append(
            {
                "accountNumber": account,
                "termID": term,
                "lastBatchDate": detail["lastBatchDate"],
            }
        )
    write_csv(rows, path, ["accountNumber", "termID", "lastBatchDate"])


def write_active_roster(records, path):
    rows = []
    for record in records:
        account = record.get("account") or {}
        location = record.get("location") or {}
        product = account.get("product") or {}
        rows.append(
            {
                "accountNumber": text(account.get("number")),
                "storeName": text(location.get("name")),
                "active": account.get("active"),
                "status": text(account.get("status")),
                "productId": text(product.get("id")),
                "processor": text((account.get("processor") or {}).get("name")),
            }
        )
    write_csv(
        rows,
        path,
        ["accountNumber", "storeName", "active", "status", "productId", "processor"],
    )


def make_client(config):
    base_url = text(config.get("apiBaseUrl")) or BASE_URL
    if base_url.endswith(AUTH_PATH):
        base_url = base_url[: -len(AUTH_PATH)].rstrip("/")
    api_key_name = text(config.get("apiKeyEnvironmentVariable")) or "MXCONNECT_API_KEY"
    api_key = clean_api_key(os.environ.get(api_key_name))
    if not api_key:
        raise RuntimeError(f"Set {api_key_name} before running the API report.")
    return MxConnectClient(base_url, api_key)


def run_historical(args):
    config_path = Path(args.config or "config.json").expanduser().resolve()
    config, _ = load_config(config_path)
    historical_days = args.historical_days or int(config.get("historicalLookbackDays", 90))
    output_directory = resolve_output_directory(config_path, config) / "historical"
    client = make_client(config)
    client.authenticate()

    start_time, end_time = date_window(historical_days)
    print(
        f"Historical batch window: {start_time:%Y-%m-%d} - {end_time:%Y-%m-%d}"
    )
    base_url = text(config.get("apiBaseUrl")) or BASE_URL
    if base_url.endswith(AUTH_PATH):
        base_url = base_url[: -len(AUTH_PATH)].rstrip("/")
    batch_records = fetch_all(
        client,
        batch_url(base_url, start_time.strftime("%Y-%m-%d"), end_time.strftime("%Y-%m-%d")),
    )
    accepted_batch_records = [record for record in batch_records if is_accepted_batch(record)]
    print(f"Accepted batch records: {len(accepted_batch_records)} of {len(batch_records)}")
    write_raw_batch_history(accepted_batch_records, output_directory / "batch_history.csv")
    write_term_history(accepted_batch_records, output_directory / "termid_account_history.csv")

    roster_records = fetch_uar(client)
    write_active_roster(roster_records, output_directory / "active_tsys_roster.csv")
    print(f"Wrote historical data to {output_directory}")


def parse_args():
    parser = ArgumentParser(description="Generate the store-level TSYS PAX batch report.")
    parser.add_argument(
        "--ui", action="store_true",
        help="Open the configuration UI (default when no command is supplied)",
    )
    parser.add_argument(
        "--run-report", action="store_true",
        help="Run the report without opening the UI; intended for Task Scheduler",
    )
    parser.add_argument(
        "--refresh-historical", action="store_true",
        help="Refresh historical batch and active-roster CSV files without opening the UI",
    )
    parser.add_argument(
        "-c", "--config", default=None,
        help="JSON configuration file (default: ./config.json)",
    )
    parser.add_argument(
        "--batch-days", type=int, default=None,
        help="Batch lookback in days (default: config batchLookbackDays or 3)",
    )
    parser.add_argument(
        "--historical-days", type=int, default=None,
        help="Historical batch export lookback (default: config historicalLookbackDays or 90)",
    )
    parser.add_argument(
        "--auth-window", default=None,
        help="MXConnect authorization quick window (default: config authorizationWindow or last_24_h)",
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="Email CSV output path",
    )
    parser.add_argument(
        "--review-output", default=None,
        help="Store review CSV output path",
    )
    parser.add_argument(
        "-tf", "--tsys_filename", dest="raw_batch_output", default=None,
        help="Optional raw batch CSV output path retained for compatibility",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    if args.refresh_historical:
        run_historical(args)
        return 0

    config_path = Path(args.config or "config.json").expanduser().resolve()
    config, raw_devices = load_config(config_path)

    batch_days = args.batch_days or int(config.get("batchLookbackDays", 3))
    auth_window = args.auth_window or text(config.get("authorizationWindow")) or "last_24_h"
    require_approved = True
    base_url = text(config.get("apiBaseUrl")) or BASE_URL
    if base_url.endswith(AUTH_PATH):
        base_url = base_url[: -len(AUTH_PATH)].rstrip("/")

    output_directory = resolve_output_directory(config_path, config)

    email_path = Path(args.output).expanduser() if args.output else output_directory / f"pinpad_batch_not_closed_{batch_days}_days.csv"
    review_path = Path(args.review_output).expanduser() if args.review_output else output_directory / "needs_mapping_or_review.csv"
    detail_path = output_directory / "in_use_not_batched_detail.csv"
    raw_path = Path(args.raw_batch_output).expanduser() if args.raw_batch_output else None

    print(f"Loaded {len(raw_devices)} optional store display overrides.")
    print(f"MXConnect base URL: {base_url}")
    print(f"Batch window: {batch_days} days; authorization window: {auth_window}")
    stdout.flush()

    # Clear any previous email data before starting a new API run. A failed
    # run must not leave yesterday's CSV looking like the current result.
    write_csv([], email_path, EMAIL_KEYS)

    client = make_client(config)
    client.authenticate()

    report_start, report_end = date_window(batch_days)
    print("Fetching TSYS batch records for the report window...")
    batch_records = fetch_all(
        client,
        batch_url(base_url, report_start.strftime("%Y-%m-%d"), report_end.strftime("%Y-%m-%d")),
    )
    accepted_batch_records = [record for record in batch_records if is_accepted_batch(record)]
    print(f"Accepted batch records: {len(accepted_batch_records)} of {len(batch_records)}")
    print("Fetching active TSYS merchant roster...")
    roster_records = fetch_uar(client)
    active_accounts = active_tsys_accounts(roster_records)
    print(f"Active TSYS merchants: {len(active_accounts)}")

    current_records = accepted_batch_records
    batched_accounts = {
        text(record.get("accountNumber"))
        for record in current_records
        if text(record.get("accountNumber"))
    }
    print(f"Accounts with an accepted batch in the report window: {len(batched_accounts)}")

    candidate_accounts = active_accounts - batched_accounts
    print(f"Active TSYS stores with no accepted batch: {len(candidate_accounts)}")
    print("Fetching authorization detail for candidate stores...")
    authorization_records = fetch_all(
        client,
        authorization_url(base_url, auth_window, candidate_accounts),
    ) if candidate_accounts else []
    auth_summary = build_auth_summary(authorization_records, require_approved)
    in_use_accounts = {account for account, _ in auth_summary}
    print(f"Active TSYS stores with approved authorization activity: {len(in_use_accounts)}")

    email_rows, detail_rows, review_rows = create_store_report(
        roster_records,
        raw_devices,
        auth_summary,
        current_records,
    )

    write_csv(email_rows, email_path, EMAIL_KEYS)
    write_csv(review_rows, review_path, REVIEW_KEYS)
    write_csv(detail_rows, detail_path, DETAIL_KEYS)
    write_raw_batch_history(accepted_batch_records, output_directory / "batch_history.csv")
    write_term_history(accepted_batch_records, output_directory / "termid_account_history.csv")
    if raw_path:
        write_raw_batch_history(batch_records, raw_path)

    print(f"{len(email_rows)} store(s) would be included in the email export.")
    print(f"{len(detail_rows)} authorization terminal detail row(s) were written.")
    print(f"{len(review_rows)} store review row(s) were excluded.")
    print(f"Wrote {email_path}")
    print(f"Wrote {review_path}")
    print(f"Wrote {detail_path}")
    print(f"Wrote {output_directory / 'batch_history.csv'}")
    print(f"Wrote {output_directory / 'termid_account_history.csv'}")
    if raw_path:
        print(f"Wrote {raw_path}")
    return 0


UI_COLUMNS = [
    ("enabled", "Enabled", 70),
    ("storeName", "Store Name", 220),
    ("url", "URL", 220),
    ("device", "Device", 85),
    ("accountNumber", "Account Number", 160),
]


def default_config():
    return {
        "apiBaseUrl": BASE_URL,
        "apiKeyEnvironmentVariable": "MXCONNECT_API_KEY",
        "batchLookbackDays": 3,
        "historicalLookbackDays": 90,
        "authorizationWindow": "last_24_h",
        "requireAuthorizationActivity": True,
        "outputDirectory": "./tsys-auditdata",
        "devices": [],
    }


class BatchReportUi(tk.Tk):
    def __init__(self, initial_config=None):
        super().__init__()
        self.title("TSYS_PAX_BATCH_REPORT")
        self.geometry("1500x900")
        self.minsize(1100, 650)
        self.configure(bg="#101216")

        self.script_path = application_entrypoint()
        self.command_path = application_cli_entrypoint()
        self.config_path_var = tk.StringVar(
            value=str(Path(initial_config).expanduser().resolve())
            if initial_config else str(self.script_path.with_name("config.json"))
        )
        self.output_directory_var = tk.StringVar(value="./tsys-auditdata")
        self.batch_days_var = tk.IntVar(value=3)
        self.filter_var = tk.StringVar()
        self.devices = []
        self.config = default_config()
        self.loaded_config_path = None
        self.process = None
        self.edit_control = None

        self.configure_styles()
        self.build_controls()
        self.filter_var.trace_add("write", lambda *_: self.refresh_table())
        self.load_config_path(self.config_path_var.get(), log=False)

    def configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#101216")
        style.configure("TLabel", background="#101216", foreground="#e5e7eb")
        style.configure("TButton", padding=(8, 5))
        style.configure("Treeview", background="#050607", foreground="#e5e7eb", fieldbackground="#050607", rowheight=24)
        style.configure("Treeview.Heading", background="#263241", foreground="#ffffff")
        style.map("Treeview", background=[("selected", "#174a66")])

    def build_controls(self):
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Config file").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(top, textvariable=self.config_path_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(top, text="Open config", command=self.open_config).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(top, text="Output folder").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(top, textvariable=self.output_directory_var).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(top, text="Browse folder", command=self.browse_output).grid(row=1, column=2, padx=(8, 0), pady=4)

        options = ttk.Frame(top)
        options.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 4))
        ttk.Label(options, text="Report timeframe (days)").pack(side="left", padx=(0, 5))
        tk.Spinbox(options, from_=1, to=365, width=5, textvariable=self.batch_days_var).pack(side="left", padx=(0, 14))

        actions = ttk.Frame(top)
        actions.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 8))
        ttk.Button(actions, text="Save config", command=self.save_config_from_ui).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Run report", command=self.run_report).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Refresh historical data", command=self.refresh_historical).pack(side="left")

        filter_frame = ttk.Frame(outer)
        filter_frame.pack(fill="x", pady=(4, 8))
        filter_frame.columnconfigure(1, weight=1)
        ttk.Label(filter_frame, text="Filter devices").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(filter_frame, textvariable=self.filter_var).grid(row=0, column=1, sticky="ew")
        self.device_count_label = ttk.Label(filter_frame, text="0 shown / 0 total")
        self.device_count_label.grid(row=0, column=2, sticky="e", padx=(8, 0))

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_frame, columns=[item[0] for item in UI_COLUMNS], show="headings", selectmode="extended")
        for key, heading, width in UI_COLUMNS:
            self.tree.heading(key, text=heading)
            self.tree.column(key, width=width, minwidth=60, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.bind("<Double-1>", self.begin_cell_edit)

        self.log_box = tk.Text(outer, height=6, bg="#050607", fg="#e5e7eb", insertbackground="#ffffff", relief="flat", wrap="word")
        self.log_box.pack(fill="x", pady=(10, 0))

    def log(self, message):
        self.log_box.insert("end", f"[{datetime.now():%H:%M:%S}] {message}\n")
        self.log_box.see("end")

    def load_config_path(self, raw_path, log=True):
        path = Path(raw_path).expanduser().resolve()
        self.config_path_var.set(str(path))
        if not path.exists():
            self.config = default_config()
            self.devices = []
            self.loaded_config_path = str(path)
            self.output_directory_var.set(str((path.parent / "tsys-auditdata").resolve()))
            self.refresh_table()
            if log:
                self.log(f"Config does not exist yet: {path}")
            return
        try:
            config, devices = load_config(path)
            self.config = config
            self.devices = [dict(device) for device in devices]
            self.loaded_config_path = str(path)
            self.output_directory_var.set(str(resolve_output_directory(path, config)))
            self.batch_days_var.set(int(config.get("batchLookbackDays", 3)))
            self.refresh_table()
            if log:
                self.log(f"Loaded {len(self.devices)} devices.")
        except Exception as error:
            self.log(f"ERROR: {error}")

    def open_config(self):
        path = filedialog.askopenfilename(
            title="Open config",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path:
            self.load_config_path(path)

    def browse_output(self):
        path = filedialog.askdirectory(
            title="Select output folder",
            initialdir=self.output_directory_var.get() or str(Path.cwd()),
        )
        if path:
            self.output_directory_var.set(path)

    def display_value(self, device, key):
        if key == "enabled":
            return "Yes" if device.get(key, True) else "No"
        return text(device.get(key))

    def refresh_table(self):
        filter_value = self.filter_var.get().strip().casefold()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for index, device in enumerate(self.devices):
            searchable = "|".join(text(device.get(key)) for key, _, _ in UI_COLUMNS).casefold()
            if filter_value and filter_value not in searchable:
                continue
            values = [self.display_value(device, key) for key, _, _ in UI_COLUMNS]
            self.tree.insert("", "end", iid=str(index), values=values)
            shown += 1
        self.device_count_label.configure(text=f"{shown:,} shown / {len(self.devices):,} total")

    def begin_cell_edit(self, event):
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        column_index = int(column_id[1:]) - 1
        if column_index < 0 or column_index >= len(UI_COLUMNS):
            return
        device_index = int(row_id)
        key = UI_COLUMNS[column_index][0]
        if key == "enabled":
            self.devices[device_index][key] = not bool(self.devices[device_index].get(key, True))
            self.refresh_table()
            return
        bounds = self.tree.bbox(row_id, column_id)
        if not bounds:
            return
        x, y, width, height = bounds
        self.commit_cell_edit()
        self.edit_control = tk.Entry(self.tree, bg="#1f2937", fg="#ffffff", insertbackground="#ffffff", relief="solid")
        self.edit_control.insert(0, text(self.devices[device_index].get(key)))
        self.edit_control.place(x=x, y=y, width=width, height=height)
        self.edit_control.focus_set()
        self.edit_control.bind("<Return>", lambda _: self.commit_cell_edit(device_index, key))
        self.edit_control.bind("<Escape>", lambda _: self.commit_cell_edit())
        self.edit_control.bind("<FocusOut>", lambda _: self.commit_cell_edit(device_index, key))

    def commit_cell_edit(self, device_index=None, key=None):
        if self.edit_control is None:
            return
        if device_index is not None and key is not None:
            self.devices[device_index][key] = self.edit_control.get().strip()
        self.edit_control.destroy()
        self.edit_control = None
        self.refresh_table()

    def save_config_from_ui(self, announce=True):
        try:
            self.commit_cell_edit()
            path = Path(self.config_path_var.get()).expanduser().resolve()
            config = dict(self.config or default_config())
            config["batchLookbackDays"] = max(1, int(self.batch_days_var.get()))
            config["historicalLookbackDays"] = max(
                int(config.get("historicalLookbackDays", 90)),
                config["batchLookbackDays"],
            )
            config["requireAuthorizationActivity"] = True
            save_config(path, config, self.devices, self.output_directory_var.get())
            self.config = config
            self.loaded_config_path = str(path)
            self.config_path_var.set(str(path))
            if announce:
                self.log(f"Saved {path}.")
            return path
        except Exception as error:
            self.log(f"ERROR: {error}")
            return None

    def start_command(self, command):
        if self.process is not None and self.process.poll() is None:
            self.log("A report is already running.")
            return
        config_path = self.save_config_from_ui()
        if config_path is None:
            return
        command = (
            [str(self.command_path), *command, "--config", str(config_path)]
            if getattr(sys, "frozen", False)
            else [sys.executable, str(self.command_path), *command, "--config", str(config_path)]
        )
        try:
            self.process = subprocess.Popen(
                command,
                cwd=str(self.script_path.parent),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                env=os.environ.copy(),
            )
            threading.Thread(target=self.read_process_output, daemon=True).start()
        except Exception as error:
            self.log(f"ERROR: {error}")

    def read_process_output(self):
        process = self.process
        if process is None or process.stdout is None:
            return
        for line in process.stdout:
            self.after(0, self.log, line.rstrip())
        exit_code = process.wait()
        self.after(0, self.report_finished, exit_code)

    def report_finished(self, exit_code):
        self.log("Finished." if exit_code == 0 else f"ERROR: report exited with code {exit_code}.")
        self.process = None

    def run_report(self):
        self.start_command(["--run-report"])

    def refresh_historical(self):
        self.start_command(["--refresh-historical"])


def launch_ui(initial_config=None):
    BatchReportUi(initial_config).mainloop()


if __name__ == "__main__":
    try:
        if len(sys.argv) == 1 or "--ui" in sys.argv:
            initial_config = None
            for flag in ("--config", "-c"):
                if flag in sys.argv:
                    position = sys.argv.index(flag)
                    if position + 1 < len(sys.argv):
                        initial_config = sys.argv[position + 1]
                    break
            launch_ui(initial_config)
        else:
            raise SystemExit(main())
    except (requests.RequestException, RuntimeError, ValueError, OSError, json.JSONDecodeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
