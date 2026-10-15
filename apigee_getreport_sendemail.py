
#!/usr/bin/env python2.7
# -*- coding: utf-8 -*-

import os
import sys
import calendar
import datetime
import subprocess
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Konfiguráció
PYTHON = "/usr/bin/python2.7"
SCRIPT = "/opt/apigee/apigee_excel_export.py"
OUTPUT_DIR = "/tmp"

SMTP_HOST = "smtp.example.com"
SMTP_PORT = 25

MAIL_FROM = "report@example.com"
MAIL_TO = ["recipient@example.com"]


def previous_month():
    today = datetime.date.today()
    first = today.replace(day=1)
    last = first - datetime.timedelta(days=1)

    start = last.replace(day=1)
    return start, last


def export_report(start, end, output):
    cmd = [
        PYTHON,
        SCRIPT,
        "--table", 'analytics."org.env.fact"',
        "--start-date", str(start),
        "--end-date", str(end),
        "--granularity", "monthly",
        "--output", output
    ]

    subprocess.check_call(
        cmd,
        cwd=os.path.dirname(SCRIPT)
    )

    if not os.path.isfile(output):
        raise RuntimeError("Export file missing")

    if os.path.getsize(output) == 0:
        raise RuntimeError("Export file is empty")


def send_email(output, month):

    msg = MIMEMultipart()
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(MAIL_TO)
    msg["Subject"] = "Apigee havi riport - " + month

    body = (
        "Csatolva az Apigee havi riport.\n"
        "Idoszak: %s\n" % month
    )

    msg.attach(MIMEText(body, "plain", "utf-8"))

    with open(output, "rb") as f:
        part = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
        part.set_payload(f.read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        "attachment",
        filename=os.path.basename(output)
    )

    msg.attach(part)

    smtp = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)

    try:
        smtp.sendmail(MAIL_FROM, MAIL_TO, msg.as_string())
    finally:
        try:
            smtp.quit()
        except smtplib.SMTPServerDisconnected:
            pass

def main():
    start, end = previous_month()
    month = start.strftime("%Y-%m")

    output = os.path.join(
        OUTPUT_DIR,
        "apigee_report_%s.xlsx" % month
    )

    # Régi fájl eltávolítása a téves küldés ellen
    if os.path.exists(output):
        os.remove(output)

    export_report(start, end, output)
    send_email(output, month)

    print("Report sent: %s" % output)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        sys.stderr.write("ERROR: %s\n" % exc)
        sys.exit(1)
