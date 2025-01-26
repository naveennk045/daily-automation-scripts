import os
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import pandas as pd

load_dotenv()


# Email account credentials
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")

# Load recipient details from a CSV file
recipients = pd.read_csv(r'V:\Projects\daily-automation-scripts\email\recipients.csv')

#Email Content
subject = "Welcome to My GitHub Repository!"
body_template = """
Hi {name},

I hope this email finds you well! I'm excited to share my GitHub repository with you: daily-automation-scripts. 
Here, you'll find a collection of Python scripts designed to simplify day-to-day tasks and enhance productivity.

Check it out here: https://github.com/naveennk045?tab=repositories

Feel free to explore, use the scripts, and share your feedback. If you have any suggestions or contributions, they’re always welcome!

Best regards,  
NK
"""

# Set up the SMTP server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(EMAIL, PASSWORD)

for _, row in recipients.iterrows():
    name, to_email = row['Name'], row['Email']

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = to_email
    msg['Subject'] = subject
    body = body_template.format(name=name)
    msg.attach(MIMEText(body, 'plain'))

    # Add an attachment
    filename = r'V:\Projects\daily-automation-scripts\email\recipients_list.pdf' 
    with open(filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={filename}')
    msg.attach(part)

    # Send the email
    server.sendmail(EMAIL, to_email, msg.as_string())
    print(f"Email sent to {name} ({to_email})")

server.quit()