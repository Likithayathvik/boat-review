import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

EMAIL_USERNAME = os.environ['EMAIL_USERNAME']
EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
NOTIFY_EMAIL   = 'likithamdl06@gmail.com'
XLSX_FILE      = 'Android_HearablesApp_Review.xlsx'

def send():
    msg = MIMEMultipart()
    msg['From']    = EMAIL_USERNAME
    msg['To']      = NOTIFY_EMAIL
    msg['Subject'] = f"✅ boAt HApp Weekly Review Report — {datetime.now().strftime('%d %b %Y')}"

    body = """Hi Likitha,

The weekly boAt Hearables App review analysis is complete.

Please find the Excel report attached.

— boAt R&D QA Automation"""

    msg.attach(MIMEText(body, 'plain'))

    # Attach Excel file
    with open(XLSX_FILE, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='xlsx')
        attachment.add_header(
            'Content-Disposition', 'attachment',
            filename=f"Android_HearablesApp_Review_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        )
        msg.attach(attachment)

    # Send via Gmail SMTP
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        server.send_message(msg)

    print(f"✅ Email sent to {NOTIFY_EMAIL}")

if __name__ == '__main__':
    send()