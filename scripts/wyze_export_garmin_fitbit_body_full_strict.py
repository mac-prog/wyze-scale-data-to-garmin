import os
from pathlib import Path
from datetime import datetime

import pandas as pd

PROJECT_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = Path(os.getenv("WYZE_OUTPUT_DIR", PROJECT_DIR / "output"))

input_file = OUTPUT_DIR / "clean" / "Wyze_Scale_Clean_Master.xlsx"
output_dir = OUTPUT_DIR / "garmin"

timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
output_file = output_dir / f"Wyze_To_Garmin_FITBIT_BODY_{timestamp}.csv"

sheet_name = "Daily_Analysis"

print("Reading clean workbook...")
print(input_file)
print("Sheet:", sheet_name)

df = pd.read_excel(input_file, sheet_name=sheet_name)

required_columns = ["local_datetime", "weight", "bmi", "body_fat"]
missing = [col for col in required_columns if col not in df.columns]
if missing:
    raise ValueError(f"Missing required columns: {missing}")

df["local_datetime"] = pd.to_datetime(df["local_datetime"], errors="coerce")
df["weight"] = pd.to_numeric(df["weight"], errors="coerce")
df["bmi"] = pd.to_numeric(df["bmi"], errors="coerce")
df["body_fat"] = pd.to_numeric(df["body_fat"], errors="coerce")

before = len(df)

df = df.dropna(subset=["local_datetime", "weight", "bmi", "body_fat"])
df = df.sort_values("local_datetime")

after = len(df)

rows = []
rows.append("Body")
rows.append("Date,Weight,BMI,Fat")

for _, row in df.iterrows():
    date_text = row["local_datetime"].strftime("%d-%m-%Y")
    weight_text = f"{float(row['weight']):.2f}"
    bmi_text = f"{float(row['bmi']):.1f}"
    fat_text = f"{float(row['body_fat']):.1f}"
    rows.append(f'"{date_text}","{weight_text}","{bmi_text}","{fat_text}"')

content = "\n".join(rows) + "\n"

output_dir.mkdir(parents=True, exist_ok=True)
output_file.write_text(content, encoding="utf-8")

print()
print("Garmin Fitbit Body CSV created:")
print(output_file)

print()
print("Rows before strict filtering:", before)
print("Rows after strict filtering:", after)

print()
print("First 10 lines:")
for line in rows[:10]:
    print(line)

print()
print("Last 10 lines:")
for line in rows[-10:]:
    print(line)

print()
print("Done.")
