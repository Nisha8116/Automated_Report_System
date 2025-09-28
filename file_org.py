import os, shutil

# Path to Downloads folder
source_folder = r"C:\Users\NISHANTHINI S\Downloads"
dest_folder = r"C:\Users\NISHANTHINI S\OrganizedFiles"

file_types = {
    "Images": [".jpg", ".png", ".jpeg"],
    "PDFs": [".pdf"],
    "Docs": [".docx", ".txt", ".pptx"],
    "Excel": [".xlsx", ".csv"]
}

if not os.path.exists(dest_folder):
    os.makedirs(dest_folder)

count = {}
for filename in os.listdir(source_folder):
    file_path = os.path.join(source_folder, filename)
    if os.path.isfile(file_path):
        for folder, extensions in file_types.items():
            if filename.lower().endswith(tuple(extensions)):
                target = os.path.join(dest_folder, folder)
                os.makedirs(target, exist_ok=True)
                shutil.move(file_path, os.path.join(target, filename))
                count[folder] = count.get(folder, 0) + 1

print("Files organized:", count)
# ---- After moving all files ----

import pandas as pd
from datetime import datetime

report_file = "daily_report.xlsx"
data = {"Category": list(count.keys()), "Files Moved": list(count.values())}
df = pd.DataFrame(data)

with pd.ExcelWriter(report_file, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, index=False, sheet_name="Report")

print("Report generated:", report_file)




import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

sender = "nishanthinironaldo@gmail.com"
password = "ljkm llcz enbq duvb"
receiver = "krishnakumariradhakrishnan29@gmail.com"

# Create the email
msg = MIMEMultipart()
msg["From"] = sender
msg["To"] = receiver
msg["Subject"] = "Daily File Organizer Report"

body = "Attached is today's report."
msg.attach(MIMEText(body, "plain"))

# Attach the Excel report
filename = "daily_report.xlsx"
with open(filename, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")
    msg.attach(part)

# Send via Gmail SMTP
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender, password)
server.sendmail(sender, receiver, msg.as_string())
server.quit()

print("Email sent successfully ✅")