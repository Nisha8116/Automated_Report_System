import os
import shutil
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

source_folder = r"C:\Users\NISHANTHINI S\Downloads"
dest_folder = r"C:\Users\NISHANTHINI S\Desktop\OrganizedFiles"

file_types = {
    "Images": [".jpg", ".png", ".jpeg"],
    "PDFs": [".pdf"],
    "Docs": [".docx", ".txt", ".pptx"],
    "Excel": [".xlsx", ".csv"]
}
try:
    script_dir = os.path.dirname(os.path.abspath(_file_))
except NameError:
    script_dir = os.getcwd()
report_file_name = "daily_report.xlsx"
report_file_path = os.path.join(script_dir, report_file_name) 
if not os.path.exists(dest_folder):
    os.makedirs(dest_folder)
count = {}
files_moved_count = 0
for filename in os.listdir(source_folder):
    file_path = os.path.join(source_folder, filename)
    if os.path.isfile(file_path):
        for folder, extensions in file_types.items():
            if filename.lower().endswith(tuple(extensions)):
                target = os.path.join(dest_folder, folder)
                os.makedirs(target, exist_ok=True)
                destination_path = os.path.join(target, filename)
                if not os.path.exists(destination_path):
                    shutil.move(file_path, destination_path)
                    count[folder] = count.get(folder, 0) + 1
                    files_moved_count += 1
print(f"Files organized: {files_moved_count} total files moved.")
print("Breakdown by folder:", count)
data = {"Category": list(count.keys()), "Files Moved": list(count.values())}
df = pd.DataFrame(data)
with pd.ExcelWriter(report_file_path, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, index=False, sheet_name="Report")
print("Report generated:", report_file_path)
sender = "nishanthinironaldo@gmail.com"
password = "ljkm llcz enbq duvb" 
receiver = "krishnakumariradhakrishnan29@gmail.com"
msg = MIMEMultipart()
msg["From"] = sender
msg["To"] = receiver
msg["Subject"] = f"Daily File Organizer Report - {datetime.now().strftime('%Y-%m-%d')}"
body = f"Hello,\n\nThe file organization script finished today.\n\nTotal files moved: {files_moved_count}\n\nAttached is the detailed report."
msg.attach(MIMEText(body, "plain"))
with open(report_file_path, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {report_file_name}")
    msg.attach(part)
try:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender, password)
    server.sendmail(sender, receiver, msg.as_string())
    server.quit()
    print("Email sent successfully ✅")
except smtplib.SMTPAuthenticationError:
    print("Email Error ❌: Authentication failed. Did you use an *App Password*?")
except Exception as e:
    print(f"Email Error ❌: An error occurred while sending the email: {e}")
