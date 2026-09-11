# Automated Report Processor

A Python-based automation tool for processing production reports and generating daily reject rate reports.

*Tailored for Bosch engineers working with automated insertion machines.*

## Features

- Scans production report folders by date
- Extracts part number, pickup count, throw count, and generated time from TXT reports
- Calculates component reject rates
- Generates a structured Excel report
- Automatically sends the report via Gmail

## Output

The generated Excel report contains:

| Project No. | Part No. | Pickup Count | Throw Count | Reject Rate | Generated Time |
|-------------|----------|---:|---:|---:|---|
| xxx         | xxx      | 168 | 1 | 0.60% | 19:54 |

## Requirements

*The requirements below can be ignored when using the packaged `.exe` version.*

- Python 3
- openpyxl

```bash
pip install openpyxl
```

## Build

The application can be modified and then packaged as a Windows executable using PyInstaller:

```bash
pyinstaller --onefile main.py
```

## Email Configuration

Gmail credentials are stored using environment variables:

```text
GMAIL_ADDRESS
GMAIL_APP_PASSWORD
```

On Windows, use the following commands to set the environment variables:

```bash
setx GMAIL_ADDRESS "your_email@gmail.com"
setx GMAIL_APP_PASSWORD "your_app_password"
```

**Restart the terminal or application after setting the variables.**

**Note:** Do not store email passwords directly in the source code.

## Automatic Scheduling on Windows

Windows Task Scheduler to be used to run the application automatically every day.

### Setup

1. Open **Task Scheduler**.
2. Select **Create Basic Task**.
3. Enter a name, for example:
   ```text
   Daily Report Processor
   ```
4. Select **Daily** as the trigger.
5. Set the execution time, for example **06:00**.
6. Select **Start a program**.
7. Select the packaged executable:
   ```text
   DailyReportProcessor.exe
   ```
8. Complete the setup and save the task.

The application will then run automatically at the scheduled time and generate the daily report from **yesterday's** data.

*Make sure the computer is not fully shut down and is connected to the internet at the scheduled time.*