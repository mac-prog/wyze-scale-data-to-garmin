import os
import json
import math
import shutil
from pathlib import Path
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import pandas as pd


BASE_DIR = Path(__file__).resolve().parents[1]
RAW_JSON = BASE_DIR / "raw" / "Wyze_Scale_Raw_All.json"
LOCAL_CLEAN_DIR = BASE_DIR / "clean"
PROJECT_DIR = Path(__file__).resolve().parents[1]
PRIVATE_CLEAN_DIR = Path(os.getenv("WYZE_OUTPUT_DIR", PROJECT_DIR / "output")) / "clean"

OUTPUT_XLSX = LOCAL_CLEAN_DIR / "Wyze_Scale_Clean_Master.xlsx"
OUTPUT_CSV_ALL = LOCAL_CLEAN_DIR / "Wyze_Scale_Clean_Master_All_Records.csv"
OUTPUT_CSV_ANALYSIS = LOCAL_CLEAN_DIR / "Wyze_Scale_Analysis_Ready.csv"
OUTPUT_CSV_DAILY = LOCAL_CLEAN_DIR / "Wyze_Scale_Daily_Analysis.csv"

PRIVATE_OUTPUT_XLSX = PRIVATE_CLEAN_DIR / "Wyze_Scale_Clean_Master.xlsx"
PRIVATE_OUTPUT_CSV_ALL = PRIVATE_CLEAN_DIR / "Wyze_Scale_Clean_Master_All_Records.csv"
PRIVATE_OUTPUT_CSV_ANALYSIS = PRIVATE_CLEAN_DIR / "Wyze_Scale_Analysis_Ready.csv"
PRIVATE_OUTPUT_CSV_DAILY = PRIVATE_CLEAN_DIR / "Wyze_Scale_Daily_Analysis.csv"

BODY_METRIC_COLUMNS = [
    "body_fat",
    "body_water",
    "bone_mineral",
    "muscle",
    "protein",
    "bmr",
    "body_vfr",
    "metabolic_age",
]

PREFERRED_COLUMNS = [
    "local_date",
    "local_datetime",
    "utc_datetime",
    "weight",
    "weight_kg",
    "bmi",
    "body_fat",
    "muscle",
    "body_water",
    "protein",
    "bone_mineral",
    "body_vfr",
    "bmr",
    "metabolic_age",
    "age",
    "height",
    "measure_type",
    "body_type",
    "is_valid_weight",
    "has_body_composition",
    "is_analysis_ready",
    "quality_notes",
    "id",
    "measure_ts",
    "timezone",
    "user_id",
    "family_member_id",
    "device_id",
    "mac",
    "impedance",
]


def safe_local_datetime(measure_ts, timezone_name):
    try:
        tz = ZoneInfo(timezone_name or "America/Los_Angeles")
    except Exception:
        tz = ZoneInfo("America/Los_Angeles")

    utc_dt = datetime.fromtimestamp(int(measure_ts) / 1000, tz=timezone.utc)
    local_dt = utc_dt.astimezone(tz)
    return utc_dt, local_dt


def clean_number(value):
    if value is None:
        return None
    try:
        value = float(value)
    except Exception:
        return value

    if math.isnan(value):
        return None

    return value


def main():
    LOCAL_CLEAN_DIR.mkdir(exist_ok=True)
    PRIVATE_CLEAN_DIR.mkdir(parents=True, exist_ok=True)

    if not RAW_JSON.exists():
        raise FileNotFoundError(f"Cannot find raw JSON file: {RAW_JSON}")

    with RAW_JSON.open("r", encoding="utf-8") as f:
        data = json.load(f)

    records = data.get("records", [])
    summary = data.get("summary", {})

    if not records:
        raise ValueError("No records found in JSON file.")

    df = pd.DataFrame(records)

    raw_count = len(df)

    if "id" in df.columns:
        df = df.drop_duplicates(subset=["id"], keep="last")
    else:
        dedupe_cols = [c for c in ["user_id", "measure_ts", "weight", "bmi"] if c in df.columns]
        df = df.drop_duplicates(subset=dedupe_cols, keep="last")

    deduped_count = len(df)

    # Convert timestamps.
    utc_values = []
    local_values = []
    local_dates = []

    for _, row in df.iterrows():
        utc_dt, local_dt = safe_local_datetime(row.get("measure_ts"), row.get("timezone"))
        utc_values.append(utc_dt.replace(tzinfo=None))
        local_values.append(local_dt.replace(tzinfo=None))
        local_dates.append(local_dt.date())

    df["utc_datetime"] = utc_values
    df["local_datetime"] = local_values
    df["local_date"] = local_dates

    # Normalize numeric columns.
    numeric_columns = [
        "age",
        "bmi",
        "bmr",
        "body_fat",
        "body_vfr",
        "body_water",
        "bone_mineral",
        "height",
        "metabolic_age",
        "muscle",
        "protein",
        "weight",
    ]

    for col in numeric_columns:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)

    # Replace Wyze invalid body-composition values with blanks.
    for col in BODY_METRIC_COLUMNS:
        if col in df.columns:
            df.loc[df[col] == -1, col] = None

    # Weight is in pounds based on the Wyze output.
    df["weight_kg"] = df["weight"] * 0.45359237

    # Quality flags.
    df["is_valid_weight"] = df["weight"].between(80, 400, inclusive="both")
    df["has_body_composition"] = df["body_fat"].notna() & df["muscle"].notna() & df["body_water"].notna()
    df["is_analysis_ready"] = df["is_valid_weight"]

    def build_quality_notes(row):
        notes = []

        if not row["is_valid_weight"]:
            notes.append("Invalid or implausible weight")

        if not row["has_body_composition"]:
            notes.append("Missing body composition metrics")

        if pd.isna(row.get("body_fat")):
            notes.append("Body fat missing")

        return "; ".join(notes)

    df["quality_notes"] = df.apply(build_quality_notes, axis=1)

    # Sort oldest to newest.
    df = df.sort_values(["local_datetime", "id"], ascending=True).reset_index(drop=True)

    # Reorder columns.
    ordered = [c for c in PREFERRED_COLUMNS if c in df.columns]
    remaining = [c for c in df.columns if c not in ordered]
    df = df[ordered + remaining]

    analysis_df = df[df["is_analysis_ready"]].copy()

    # One analysis row per local day: use the last valid weigh-in of each day.
    daily_df = (
        analysis_df
        .sort_values("local_datetime")
        .groupby("local_date", as_index=False)
        .tail(1)
        .reset_index(drop=True)
    )

    # Summary sheet.
    summary_rows = [
        ["Source JSON", str(RAW_JSON)],
        ["Export started at", summary.get("export_started_at", "")],
        ["Export finished at", summary.get("export_finished_at", "")],
        ["Raw records in uploaded JSON", raw_count],
        ["Records after script dedupe", deduped_count],
        ["Analysis-ready records", len(analysis_df)],
        ["Daily analysis records", len(daily_df)],
        ["Unique user IDs", df["user_id"].nunique() if "user_id" in df.columns else ""],
        ["First local date", str(df["local_date"].min())],
        ["Last local date", str(df["local_date"].max())],
        ["Lowest valid weight", analysis_df["weight"].min()],
        ["Highest valid weight", analysis_df["weight"].max()],
        ["Latest valid weight", analysis_df.iloc[-1]["weight"] if len(analysis_df) else ""],
        ["Notes", "Invalid body-composition values of -1 were converted to blank. Implausible weights outside 80-400 lb were flagged and excluded from analysis-ready rows."],
    ]

    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    # Save CSV outputs.
    df.to_csv(OUTPUT_CSV_ALL, index=False)
    analysis_df.to_csv(OUTPUT_CSV_ANALYSIS, index=False)
    daily_df.to_csv(OUTPUT_CSV_DAILY, index=False)

    # Save Excel output.
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        df.to_excel(writer, sheet_name="Clean_All", index=False)
        analysis_df.to_excel(writer, sheet_name="Analysis_Ready", index=False)
        daily_df.to_excel(writer, sheet_name="Daily_Analysis", index=False)

        workbook = writer.book

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            ws.freeze_panes = "A2"

            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)

            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    value = cell.value
                    if value is None:
                        continue
                    max_length = max(max_length, len(str(value)))

                adjusted_width = min(max(max_length + 2, 10), 35)
                ws.column_dimensions[column_letter].width = adjusted_width

        # Apply number/date formatting.
        for ws_name in ["Clean_All", "Analysis_Ready", "Daily_Analysis"]:
            ws = workbook[ws_name]
            headers = [cell.value for cell in ws[1]]

            for idx, header in enumerate(headers, start=1):
                if header in ["local_datetime", "utc_datetime"]:
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row=row, column=idx).number_format = "yyyy-mm-dd hh:mm"
                elif header == "local_date":
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row=row, column=idx).number_format = "yyyy-mm-dd"
                elif header in ["weight", "weight_kg", "bmi", "body_fat", "muscle", "body_water", "protein", "bone_mineral"]:
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row=row, column=idx).number_format = "0.00"

    print("Done.")
    print(f"Raw JSON records read: {raw_count}")
    print(f"Records after script dedupe: {deduped_count}")
    print(f"Analysis-ready records: {len(analysis_df)}")
    print(f"Daily analysis records: {len(daily_df)}")
    print("")

    shutil.copy2(OUTPUT_XLSX, PRIVATE_OUTPUT_XLSX)
    shutil.copy2(OUTPUT_CSV_ALL, PRIVATE_OUTPUT_CSV_ALL)
    shutil.copy2(OUTPUT_CSV_ANALYSIS, PRIVATE_OUTPUT_CSV_ANALYSIS)
    shutil.copy2(OUTPUT_CSV_DAILY, PRIVATE_OUTPUT_CSV_DAILY)

    print(f"Created local Excel file: {OUTPUT_XLSX}")
    print(f"Created local CSV file: {OUTPUT_CSV_ALL}")
    print(f"Created local CSV file: {OUTPUT_CSV_ANALYSIS}")
    print(f"Created local CSV file: {OUTPUT_CSV_DAILY}")
    print("")
    print(f"Copied private Excel file: {PRIVATE_OUTPUT_XLSX}")
    print(f"Copied private CSV file: {PRIVATE_OUTPUT_CSV_ALL}")
    print(f"Copied private CSV file: {PRIVATE_OUTPUT_CSV_ANALYSIS}")
    print(f"Copied private CSV file: {PRIVATE_OUTPUT_CSV_DAILY}")


if __name__ == "__main__":
    main()
