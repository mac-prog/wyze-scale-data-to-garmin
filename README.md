# Wyze Scale Garmin Import

Unofficial local-only workflow for exporting Wyze Scale data and creating Garmin Connect-compatible Fitbit Body CSV files.

## What this project does

This workflow helps you:

1. Export Wyze Scale records.
2. Save raw JSON and Excel files locally.
3. Clean and deduplicate scale records.
4. Build a clean master workbook.
5. Build a trend analysis workbook.
6. Create a Garmin Connect import CSV using Fitbit Body format.

The Garmin upload is manual. This project creates the CSV file. You import it into Garmin Connect yourself.

## Confirmed Garmin import fields

The following fields have been confirmed to import into Garmin Connect:

- Weight
- BMI
- Body Fat

Other Wyze body composition fields may be available in local exports, but they are not guaranteed to import into Garmin Connect.

## Privacy model

This tool runs locally on your own computer.

Your Wyze credentials should stay in your local `.env` file. Your exported JSON, CSV, and Excel files should stay on your machine.

Do not share your Wyze password, Wyze API keys, `.env` file, raw JSON exports, CSV files, Excel files, or health data.

## Disclaimer

This is an unofficial community tool. It is not affiliated with, endorsed by, or supported by Wyze or Garmin.

This project may stop working if Wyze changes authentication, API behavior, rate limits, or device support.

Garmin import compatibility is not guaranteed. The currently confirmed imported fields are Weight, BMI, and Body Fat.

Use this tool only with your own account and your own data.

## Requirements

- Windows
- Python 3.10 or newer
- PowerShell
- Wyze account
- Wyze API key credentials
- Garmin Connect account

## Setup

Create a virtual environment:

```powershell
python -m venv .venv