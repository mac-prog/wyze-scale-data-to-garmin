import os
import json
import hashlib
from pathlib import Path
from datetime import datetime, timedelta
from dotenv import load_dotenv
from wyze_sdk import Client
import pandas as pd


CACHE_DIR = Path("wyze_scale_cache")
OUTPUT_XLSX = "Wyze_Scale_Raw_All.xlsx"
OUTPUT_JSON = "Wyze_Scale_Raw_All.json"

# Start very early so we do not miss older records.
# Wyze Scale launched years after this, but this is safe.
START_DATE = datetime(2010, 1, 1)

# Monthly chunks reduce API risk and allow resume.
CHUNK_DAYS = 31


def object_to_dict(obj):
    """
    Convert Wyze SDK objects into dictionaries as safely as possible.
    Keeps everything we can see.
    """
    if obj is None:
        return None

    if isinstance(obj, dict):
        return {k: object_to_dict(v) for k, v in obj.items()}

    if isinstance(obj, list):
        return [object_to_dict(v) for v in obj]

    if isinstance(obj, tuple):
        return [object_to_dict(v) for v in obj]

    if isinstance(obj, (str, int, float, bool)):
        return obj

    if isinstance(obj, datetime):
        return obj.isoformat()

    # Try common conversion methods
    for method_name in ["to_dict", "dict", "as_dict"]:
        method = getattr(obj, method_name, None)
        if callable(method):
            try:
                return object_to_dict(method())
            except Exception:
                pass

    # Try object attributes
    try:
        data = vars(obj)
        if data:
            return {k: object_to_dict(v) for k, v in data.items()}
    except Exception:
        pass

    # Final fallback
    return repr(obj)


def flatten_record(record_dict):
    """
    Flatten nested dictionaries for Excel.
    """
    try:
        flat = pd.json_normalize(record_dict, sep=".").to_dict(orient="records")
        return flat[0] if flat else record_dict
    except Exception:
        return record_dict


def stable_hash(data):
    raw = json.dumps(data, sort_keys=True, default=str, ensure_ascii=False)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def fetch_chunk(client, start_time, end_time):
    records = client.scales.get_records(
        start_time=start_time,
        end_time=end_time,
    )
    return [object_to_dict(r) for r in records]


def main():
    load_dotenv()

    email = os.getenv("WYZE_EMAIL")
    password = os.getenv("WYZE_PASSWORD")
    key_id = os.getenv("WYZE_KEY_ID")
    api_key = os.getenv("WYZE_API_KEY")

    missing = [
        name for name, value in {
            "WYZE_EMAIL": email,
            "WYZE_PASSWORD": password,
            "WYZE_KEY_ID": key_id,
            "WYZE_API_KEY": api_key,
        }.items()
        if not value
    ]

    if missing:
        raise RuntimeError(f"Missing required .env values: {', '.join(missing)}")

    CACHE_DIR.mkdir(exist_ok=True)

    print("Logging into Wyze...")
    client = Client(
        email=email,
        password=password,
        key_id=key_id,
        api_key=api_key,
    )
    print("Wyze login successful.")

    export_started_at = datetime.now()
    end_date = datetime.now() + timedelta(days=1)

    print()
    print("Checking scale client metadata...")

    scale_metadata = {}

    for name in ["list", "info", "get_goal_weight"]:
        method = getattr(client.scales, name, None)
        if callable(method):
            try:
                result = method()
                scale_metadata[name] = object_to_dict(result)
                print(f"  {name}: OK")
            except Exception as e:
                scale_metadata[name] = {
                    "error_type": type(e).__name__,
                    "error": str(e),
                }
                print(f"  {name}: {type(e).__name__} - {e}")

    print()
    print(f"Exporting all Wyze scale records from {START_DATE.date()} to {end_date.date()}")
    print(f"Cache folder: {CACHE_DIR}")
    print("Resume mode: ON")
    print()

    all_records = []
    fetch_log = []

    current_start = START_DATE

    while current_start < end_date:
        current_end = min(current_start + timedelta(days=CHUNK_DAYS), end_date)
        cache_file = CACHE_DIR / f"scale_records_{current_start.date()}_to_{current_end.date()}.json"

        print(f"Fetching {current_start.date()} to {current_end.date()}...")

        if cache_file.exists():
            try:
                records = json.loads(cache_file.read_text(encoding="utf-8"))
                print(f"  Loaded from cache: {len(records)} record(s)")
                status = "cached"
                error = ""
            except Exception as e:
                print(f"  Cache read failed. Refetching. Error: {e}")
                records = fetch_chunk(client, current_start, current_end)
                cache_file.write_text(
                    json.dumps(records, indent=2, ensure_ascii=False, default=str),
                    encoding="utf-8",
                )
                print(f"  Fetched: {len(records)} record(s)")
                status = "fetched_after_cache_error"
                error = str(e)
        else:
            try:
                records = fetch_chunk(client, current_start, current_end)
                cache_file.write_text(
                    json.dumps(records, indent=2, ensure_ascii=False, default=str),
                    encoding="utf-8",
                )
                print(f"  Fetched: {len(records)} record(s)")
                status = "fetched"
                error = ""
            except Exception as e:
                print(f"  ERROR: {type(e).__name__} - {e}")
                records = []
                status = "error"
                error = f"{type(e).__name__}: {e}"

        all_records.extend(records)

        fetch_log.append({
            "chunk_start": current_start.isoformat(),
            "chunk_end": current_end.isoformat(),
            "records_found": len(records),
            "status": status,
            "error": error,
        })

        current_start = current_end

    print()
    print("Deduplicating records...")

    deduped = {}
    for record in all_records:
        h = stable_hash(record)
        deduped[h] = record

    unique_records = list(deduped.values())

    print(f"Total raw records collected: {len(all_records)}")
    print(f"Unique records after dedupe: {len(unique_records)}")

    print()
    print("Building Excel output...")

    raw_rows = []
    flat_rows = []

    for i, record in enumerate(unique_records, start=1):
        raw_rows.append({
            "row_number": i,
            "record_hash": stable_hash(record),
            "raw_json": json.dumps(record, ensure_ascii=False, default=str),
        })

        flat = flatten_record(record)
        flat["row_number"] = i
        flat["record_hash"] = stable_hash(record)
        flat_rows.append(flat)

    df_raw = pd.DataFrame(raw_rows)
    df_flat = pd.DataFrame(flat_rows)
    df_log = pd.DataFrame(fetch_log)

    metadata_rows = []
    for key, value in scale_metadata.items():
        metadata_rows.append({
            "metadata_type": key,
            "raw_json": json.dumps(value, ensure_ascii=False, default=str),
        })

    df_metadata = pd.DataFrame(metadata_rows)

    summary = {
        "export_started_at": export_started_at.isoformat(),
        "export_finished_at": datetime.now().isoformat(),
        "start_date_used": START_DATE.isoformat(),
        "end_date_used": end_date.isoformat(),
        "total_raw_records_collected": len(all_records),
        "unique_records_after_dedupe": len(unique_records),
        "output_xlsx": OUTPUT_XLSX,
        "output_json": OUTPUT_JSON,
    }

    full_json = {
        "summary": summary,
        "scale_metadata": scale_metadata,
        "records": unique_records,
        "fetch_log": fetch_log,
    }

    Path(OUTPUT_JSON).write_text(
        json.dumps(full_json, indent=2, ensure_ascii=False, default=str),
        encoding="utf-8",
    )

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name="Summary", index=False)
        df_flat.to_excel(writer, sheet_name="ScaleRecordsFlattened", index=False)
        df_raw.to_excel(writer, sheet_name="ScaleRecordsRawJSON", index=False)
        df_metadata.to_excel(writer, sheet_name="ScaleMetadata", index=False)
        df_log.to_excel(writer, sheet_name="FetchLog", index=False)

    print()
    print("DONE.")
    print(f"Created Excel file: {OUTPUT_XLSX}")
    print(f"Created JSON backup: {OUTPUT_JSON}")
    print(f"Cache folder: {CACHE_DIR}")


if __name__ == "__main__":
    main()