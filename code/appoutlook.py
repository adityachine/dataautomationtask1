import pandas as pd
import logging
import smtplib
import ssl
from email.message import EmailMessage
import os
import schedule
import time
from datetime import datetime

# ========================
# CONFIGURATION & LOGGING
# ========================

INPUT_CSV = 'online_shoppers_intention.csv'
OUTPUT_CSV = 'pivot_output.csv'
SENDER_EMAIL = 'adityachineaasplintership@outlook.com'
APP_PASSWORD = 'ylgaqytqhzsmslxb'
RECIPIENT_EMAIL = 'adityahemantchine41@gmail.com'
EMAIL_SUBJECT = "Daily Pivot Report - Online Shoppers Intention"
EMAIL_BODY = "Please find attached the pivot table report generated today."

logging.basicConfig(
    filename='automation_log.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# ========================
# FUNCTION DEFINITIONS
# ========================

def load_data(filepath):
    try:
        df = pd.read_csv(filepath)
        logging.info("✅ CSV loaded successfully.")
        return df
    except FileNotFoundError:
        logging.error(f"❌ File not found: {filepath}")
        raise
    except Exception as e:
        logging.error(f"❌ Error loading CSV: {e}")
        raise

def clean_data(df):
    try:
        df['Weekend'] = df['Weekend'].astype(str)
        df['Revenue'] = df['Revenue'].astype(str)

        numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
        df[numeric_cols] = df[numeric_cols].fillna(0)

        logging.info("✅ Data cleaned successfully.")
        return df
    except Exception as e:
        logging.error(f"❌ Error cleaning data: {e}")
        raise

def create_pivot_table(df):
    try:
        pivot = pd.pivot_table(
            df,
            values='PageValues',
            index='VisitorType',
            columns='Weekend',
            aggfunc='mean',
            fill_value=0
        )
        logging.info("✅ Pivot table created.")
        return pivot
    except Exception as e:
        logging.error(f"❌ Pivot creation failed: {e}")
        raise

def save_csv(df, output_path):
    try:
        df.to_csv(output_path)
        logging.info(f"✅ Pivot saved to {output_path}.")
    except Exception as e:
        logging.error(f"❌ Error saving CSV: {e}")
        raise

def send_email(sender, password, recipient, subject, body, attachment_path):
    try:
        msg = EmailMessage()
        msg['From'] = sender
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.set_content(body)

        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            filename = os.path.basename(attachment_path)
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=filename)

        context = ssl.create_default_context()
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.starttls(context=context)
            smtp.login(sender, password)
            smtp.send_message(msg)

        logging.info("✅ Email sent successfully.")
        print("📧 Email sent successfully.")
    except Exception as e:
        logging.error(f"❌ Failed to send email: {e}")
        print("❌ Failed to send email:", e)

# ========================
# MAIN AUTOMATION WORKFLOW
# ========================

def run_automation():
    print(f"🔄 Running automation at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}...")

    try:
        df = load_data(INPUT_CSV)
        cleaned_df = clean_data(df)
        pivot_df = create_pivot_table(cleaned_df)
        save_csv(pivot_df, OUTPUT_CSV)
        send_email(SENDER_EMAIL, APP_PASSWORD, RECIPIENT_EMAIL, EMAIL_SUBJECT, EMAIL_BODY, OUTPUT_CSV)

        print("✅ Automation complete.")
    except Exception as err:
        print(f"❌ Automation failed: {err}")

# ========================
# SCHEDULING
# ========================

# Set time to run the automation daily (24-hour format, e.g. "09:00" for 9 AM)
schedule.every().day.at("14:57").do(run_automation)

print("⏳ Scheduler started. Waiting for scheduled time...")

# Infinite loop to keep the script running
while True:
    schedule.run_pending()
    time.sleep(60)
