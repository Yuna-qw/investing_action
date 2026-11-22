"""Standalone v1.5 Investing downloader.

SOURCE_SELECTOR = 1 -> international primary, Chinese fallback
SOURCE_SELECTOR = 2 -> Chinese primary, international fallback

This file is self-contained and does not import the old 01 downloader or the
three earlier crawler files. It follows aSetting, RawData/FilePath, and the
failure-log convention used by the v1.5 project.
"""
import csv
import argparse
import json
import re
import socket
import time
from datetime import datetime
from http.client import HTTPException
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

import pandas as pd

from util import aSetting


# ===================== User selector =====================
SOURCE_SELECTOR = 1          # 1=international, 2=chinese
ENABLE_FALLBACK = True       # if primary fails, try the other site
START_DATE = "1980-01-01"
TIMEOUT_SECONDS = 20
REQUEST_ATTEMPTS = 3         # total attempts per URL, including the first request
RETRY_DELAY_SECONDS = 2      # exponential backoff: 2s, 4s
SLEEP_SECONDS = 0.5
MAX_CANDIDATES = 3
LOG_SAVE_INTERVAL = 100

PROJECT_ROOT = Path(__file__).resolve().parent
MASTER = PROJECT_ROOT / aSetting.MasterSheet
RAW_DATA_PATH = PROJECT_ROOT / aSetting.RawDataPath
LOG_PATH = PROJECT_ROOT / aSetting.LogPath
END_DATE = aSetting.EndDate_Str_2
UPDATE = aSetting.Update
CSV_HEADER = ["Date", "Price", "Open", "High", "Low", "Vol.", "Change %"]
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/142 Safari/537.36"
}


def normalize_date(value):
    if value is None or str(value).strip() in ("", "nan", "NaT"):
        return ""
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%b %d, %Y", "%B %d, %Y", "%b %Y", "%B %Y"):
        try:
            parsed = datetime.strptime(text, fmt)
            return f"{parsed.year}/{parsed.month}/{parsed.day}"
        except ValueError:
            pass
    match = re.match(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$", text)
    if match:
        return f"{match.group(1)}/{int(match.group(2))}/{int(match.group(3))}"
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return f"{parsed.year}/{parsed.month}/{parsed.day}"
    except ValueError:
        return text


def localized_url(url, site):
    path = urlparse(str(url)).path
    host = "cn.investing.com" if site == "chinese" else "www.investing.com"
    return f"https://{host}{path}"


def request_bytes(url, headers=None):
    request = Request(url, headers=headers or HEADERS)
    for attempt in range(1, REQUEST_ATTEMPTS + 1):
        try:
            with urlopen(request, timeout=TIMEOUT_SECONDS) as response:
                # Only return a complete response; never parse IncompleteRead.partial.
                return response.status, response.read().decode("utf-8", errors="replace")
        except HTTPError as exc:
            retryable = exc.code in (408, 429, 500, 502, 503, 504)
            exc.close()
            if not retryable or attempt == REQUEST_ATTEMPTS:
                raise
            reason = f"HTTP {exc.code}"
        except (HTTPException, URLError, OSError) as exc:
            if attempt == REQUEST_ATTEMPTS:
                raise
            reason = type(exc).__name__
        delay = RETRY_DELAY_SECONDS * (2 ** (attempt - 1))
        print(f"    Request interrupted ({reason}); retry {attempt + 1}/{REQUEST_ATTEMPTS} in {delay}s: {url}", flush=True)
        time.sleep(delay)


def extract_ids(html):
    script = re.search(r'<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', html, re.I | re.S)
    if script:
        try:
            value = json.loads(script.group(1))
            found = []

            def walk(node):
                if isinstance(node, dict):
                    for key, child in node.items():
                        if key in ("instrument_id", "instrumentId", "pair_ID", "pairId") and str(child).isdigit() and str(child) != "0":
                            found.append(str(child))
                        walk(child)
                elif isinstance(node, list):
                    for child in node:
                        walk(child)

            walk(value)
            if found:
                return list(dict.fromkeys(found))
        except (ValueError, TypeError):
            pass
    found = re.findall(r'"(?:instrument_id|instrumentId|pair_ID|pairId)"\s*:\s*"?(\d+)', html, re.I)
    return [item for item in dict.fromkeys(found) if item != "0"]


def fetch_history(site, page_url, instrument_id):
    domain = "cn" if site == "chinese" else "www"
    api_url = (
        f"https://api.investing.com/api/financialdata/historical/{instrument_id}"
        f"?start-date={START_DATE}&end-date={END_DATE}"
        "&time-frame=Monthly&add-missing-rows=false"
    )
    headers = dict(HEADERS, Accept="application/json, text/plain, */*", **{
        "domain-id": domain,
        "Referer": page_url,
        "Origin": f"https://{domain}.investing.com",
    })
    status, text = request_bytes(api_url, headers)
    if status != 200:
        raise RuntimeError(f"history HTTP {status}: {text[:160]}")
    return json.loads(text), api_url


def payload_rows(payload):
    data = payload.get("data", []) if isinstance(payload, dict) else []
    if isinstance(data, dict):
        data = data.get("data", data.get("rows", []))
    rows = []
    for item in data if isinstance(data, list) else []:
        if not isinstance(item, dict):
            continue
        row = [
            normalize_date(item.get("rowDate") or item.get("date") or item.get("rowDateTimestamp")),
            item.get("last_close", item.get("close", "")),
            item.get("last_open", item.get("open", "")),
            item.get("last_max", item.get("high", "")),
            item.get("last_min", item.get("low", "")),
            item.get("volume", item.get("last_volume", "")),
            item.get("change_precent", item.get("change_percent", "")),
        ]
        if row[0]:
            rows.append(row)
    return rows


def write_csv(path, rows):
    # 有效数据写入前，调用方会先确保目录和表头文件已经存在。
    if not rows:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(CSV_HEADER)
        writer.writerows(rows)
    return True


def ensure_csv_table(path):
    """先创建目录和带表头的 CSV，确保下载失败时表结构仍然保留。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        with path.open("w", encoding="utf-8", newline="") as handle:
            csv.writer(handle).writerow(CSV_HEADER)


def selected_sources():
    if SOURCE_SELECTOR == 1:
        primary, backup = "international", "chinese"
    elif SOURCE_SELECTOR == 2:
        primary, backup = "chinese", "international"
    else:
        raise ValueError("SOURCE_SELECTOR must be 1 or 2")
    return [primary, backup] if ENABLE_FALLBACK else [primary]


def load_master_rows():
    master = pd.read_excel(MASTER, sheet_name="RawData", engine="openpyxl")
    mask = master["Choose_to_Download"].astype(str).str.lower().eq("yes")
    if "Update_time" in master.columns:
        mask &= pd.to_numeric(master["Update_time"], errors="coerce").eq(int(UPDATE))
    return master.loc[mask].copy()


def initialize_output_tables(rows):
    """在下载开始前，为所有目标项创建目录和带表头的 CSV。"""
    for _, row in rows.iterrows():
        output = PROJECT_ROOT / RAW_DATA_PATH / (str(row["FilePath"]) + ".csv")
        ensure_csv_table(output)


def download_one(row, sources):
    name = str(row.get("Subcategory", row.get("Name_investing", row["NoSecurity"])))
    page_url = str(row["Link_original"])
    output = PROJECT_ROOT / RAW_DATA_PATH / (str(row["FilePath"]) + ".csv")

    # 先创建目标目录和表头；即使后续没有数据，CSV 文件也不会缺失。
    ensure_csv_table(output)
    if output.exists() and output.stat().st_size > 50:
        return {"Subcategory": name, "FilePath": row["FilePath"], "Result": "skipped", "Source": "existing"}

    errors = []
    for site in sources:
        try:
            localized = localized_url(page_url, site)
            status, html = request_bytes(localized)
            if status != 200:
                raise RuntimeError(f"page HTTP {status}")
            ids = extract_ids(html)
            if not ids:
                raise RuntimeError("instrument_id not found")
            for instrument_id in ids[:MAX_CANDIDATES]:
                try:
                    payload, api_url = fetch_history(site, localized, instrument_id)
                    rows = payload_rows(payload)
                    if rows and write_csv(output, rows):
                        return {"Subcategory": name, "FilePath": row["FilePath"], "Result": "successful", "Source": site, "Instrument_ID": instrument_id, "Rows": len(rows), "API_URL": api_url}
                except (HTTPException, HTTPError, URLError, socket.timeout, TimeoutError, OSError, RuntimeError, ValueError) as exc:
                    errors.append(f"{site}/{instrument_id}:{type(exc).__name__}:{str(exc)[:100]}")
            errors.append(f"{site}:empty history")
        except (HTTPException, HTTPError, URLError, socket.timeout, TimeoutError, OSError, RuntimeError, ValueError) as exc:
            errors.append(f"{site}:{type(exc).__name__}:{str(exc)[:100]}")
    return {"Subcategory": name, "Link_original": page_url, "FilePath": row["FilePath"], "Result": "fail", "Source": "none", "Reason": "; ".join(errors)}


def save_logs(report):
    LOG_PATH.mkdir(parents=True, exist_ok=True)
    json_path = LOG_PATH / f"selector_download_{UPDATE}.json"
    json_path.write_text(json.dumps({"selector": SOURCE_SELECTOR, "sources": selected_sources(), "items": report}, ensure_ascii=False, indent=2), encoding="utf-8")
    failed = [item for item in report if item.get("Result") == "fail"]
    failed_path = LOG_PATH / f"download_failed_{UPDATE}.xlsx"
    pd.DataFrame(failed).to_excel(failed_path, index=False)
    return json_path, failed_path, len(failed)


def main():
    parser = argparse.ArgumentParser(description="Download Investing.com history for master securities.")
    parser.add_argument("--limit", type=int, default=None,
                        help="Only download the first N eligible master rows; omit for full download.")
    parser.add_argument("--fallback-full", action="store_true",
                        help="If the limited run has fewer successful downloads than its limit, retry the full eligible list.")
    args = parser.parse_args()
    rows = load_master_rows()
    sources = selected_sources()
    full_rows = rows
    if args.limit is not None:
        if args.limit < 1:
            parser.error("--limit must be a positive integer")
        rows = rows.head(args.limit)
    report = []
    print(f"Master={MASTER}")
    print(f"Selector={SOURCE_SELECTOR}, sources={sources}, target_count={len(rows)}, full_count={len(full_rows)}")
    initialize_output_tables(rows)
    print(f"Initialized {len(rows)} output tables", flush=True)
    for index, (_, row) in enumerate(rows.iterrows(), 1):
        name = str(row.get("Subcategory", row.get("Name_investing", row["NoSecurity"])))
        print(f"[{index}/{len(rows)}] {name}", flush=True)
        result = download_one(row, sources)
        report.append(result)
        print(f"    -> {result['Result']} {result.get('Source', '')}", flush=True)

        # 已存在的本地文件无需等待；只有实际发起下载时才限速。
        if result.get("Result") != "skipped":
            time.sleep(SLEEP_SECONDS)

        # 避免每个文件都重写一次 JSON 和 Excel；定期保存以保留中途进度。
        if index % LOG_SAVE_INTERVAL == 0:
            save_logs(report)
    successful_count = sum(item.get("Result") == "successful" for item in report)
    if args.fallback_full and args.limit is not None and successful_count < args.limit and len(rows) < len(full_rows):
        print(f"Only {successful_count}/{args.limit} downloads succeeded; falling back to all {len(full_rows)} eligible rows.", flush=True)
        report = []
        rows = full_rows
        initialize_output_tables(rows)
        for index, (_, row) in enumerate(rows.iterrows(), 1):
            name = str(row.get("Subcategory", row.get("Name_investing", row["NoSecurity"])))
            print(f"[full {index}/{len(rows)}] {name}", flush=True)
            result = download_one(row, sources)
            report.append(result)
            print(f"    -> {result['Result']} {result.get('Source', '')}", flush=True)
            if result.get("Result") != "skipped":
                time.sleep(SLEEP_SECONDS)
            if index % LOG_SAVE_INTERVAL == 0:
                save_logs(report)
    json_path, failed_path, failed_count = save_logs(report)
    print(f"Finished: successful={len(rows) - failed_count}, failed={failed_count}")
    print(f"JSON log: {json_path}")
    print(f"Failed log: {failed_path}")


if __name__ == "__main__":
    main()
