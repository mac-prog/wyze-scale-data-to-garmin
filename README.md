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

## Configuration

This project uses a local `.env` file for configuration.

Create your `.env` file by copying the example file:

```powershell
Copy-Item .env.example .env
```
Then open `.env`:

```powershell
notepad .env
```

Fill in your own Wyze credentials and Wyze API credentials:

```env
WYZE_EMAIL=your_email@example.com
WYZE_PASSWORD=your_wyze_password
WYZE_KEY_ID=your_wyze_key_id
WYZE_API_KEY=your_wyze_api_key
```

Your `.env` file stays on your own computer. Do not upload it to GitHub. Do not share it with anyone.

## Output folder

By default, generated files are saved inside this project folder under:

```text
.\output
```

You can choose a different output folder by setting `WYZE_OUTPUT_DIR` in your `.env` file.

Example:

```env
WYZE_OUTPUT_DIR=C:\Users\YourName\Documents\Wyze Scale Output
```

If `WYZE_OUTPUT_DIR` is blank, the workflow uses:

```text
.\output
```

## Working files and final outputs

The workflow creates local working files and final output files.

Local working files may include:

```text
raw\
clean\
wyze_scale_cache\
```

Final output folders may include:

```text
output\clean
output\analysis
output\garmin
```

The Garmin import CSV is created in:

```text
output\garmin
```

or in your custom `WYZE_OUTPUT_DIR\garmin` folder if you set `WYZE_OUTPUT_DIR`.

Do not commit generated output files to GitHub. These files may contain personal health data, timestamps, account data, user IDs, or device IDs.
## Setup

Create a virtual environment:

```powershell
python -m venv .venv