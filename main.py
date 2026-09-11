import smtplib
import os
import sys
from datetime import datetime, timedelta
from openpyxl import Workbook
from email.message import EmailMessage

# Set up the email
sender_email = os.getenv("GMAIL_ADDRESS")
app_password = os.getenv("GMAIL_APP_PASSWORD")

# Specify the folder name and date to be extracted
if getattr(sys, "frozen", False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

report_folder_name = os.path.join(base_path, "Report")
date = (datetime.now() - timedelta(days=1)).strftime("%d_%m_%Y")

if not os.path.isdir(report_folder_name):
    raise FileNotFoundError(
        f"Report folder not found: {report_folder_name}"
    )
# Store all valid results
results = []

# Loop through the report folder and filter out the report files for the specified date
projects = os.listdir(report_folder_name)

for project in projects:
    project_path = os.path.join(report_folder_name, project)

    if os.path.isdir(project_path):
        date_path = os.path.join(project_path, date)

        if os.path.isdir(date_path):
            report_files = os.listdir(date_path)

            for filename in report_files:
                if filename.endswith(".txt"):
                    file_path = os.path.join(date_path, filename)

                    # Extract generated time from filename
                    try:
                        time_part = os.path.splitext(filename)[0]
                        hour, minute, second = time_part.split("_")
                        generated_time = f"{hour}:{minute}"
                    except ValueError:
                        print(f"Error: Invalid filename format: {filename}")
                        continue

                    # Extract data from the report files and calculate the reject rate
                    try:
                        with open(file_path, "r", encoding="utf-8") as f:
                            lines = f.readlines()

                        file_results = []

                        for line in lines[2:]:
                            parts = line.split()

                            if len(parts) != 5:
                                raise ValueError("Invalid report format")

                            part_name = parts[0]
                            pickup_count = int(parts[1])
                            throw_count = int(parts[4])

                            if pickup_count == 0:
                                raise ValueError("Pickup count cannot be 0")

                            reject_rate = throw_count / pickup_count

                            result = {
                                "Project No.": project,
                                "Part No.": part_name,
                                "Pickup Count": pickup_count,
                                "Throw Count": throw_count,
                                "Reject Rate": reject_rate,
                                "Generated Time": generated_time
                            }

                            file_results.append(result)

                        # Only save the file if the whole file is valid
                        results.extend(file_results)

                    except (ValueError, UnicodeDecodeError, OSError) as e:
                        print(f"Error: Failed to process {filename}: {e}")
                        continue


# Write the results to an Excel file
output_file = f"daily_report_{date}.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Reject Rate Report"

headers = [
    "Project No.",
    "Part No.",
    "Pickup Count",
    "Throw Count",
    "Reject Rate",
    "Generated Time"
]

ws.append(headers)

for result in results:
    ws.append([
        result["Project No."],
        result["Part No."],
        result["Pickup Count"],
        result["Throw Count"],
        result["Reject Rate"],
        result["Generated Time"]
    ])

for cell in ws["E"][1:]:
    cell.number_format = "0.00%"

# Set column widths
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 25
ws.column_dimensions["C"].width = 15
ws.column_dimensions["D"].width = 15
ws.column_dimensions["E"].width = 15
ws.column_dimensions["F"].width = 15

wb.save(output_file)
print(f"Report generated: {output_file}")

# Send the report via email
msg = EmailMessage()

msg["Subject"] = f"Daily Reject Rate Report - {date}"
# Input your email address and the receiver's email address here
msg["From"] = sender_email
msg["To"] = sender_email

msg.set_content(f"Please find attached the daily reject rate report for {date}.")

# Attach the Excel file to the email
with open(output_file, "rb") as f:
    file_data = f.read()

msg.add_attachment(
    file_data,
    maintype="application",
    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    filename=output_file
)

# Send the email using Gmail's SMTP server
try:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(msg)
    print("Email sent successfully!")

except smtplib.SMTPException as e:
    print(f"Error: Failed to send email: {e}")