import pandas as pd
import logging
import smtplib
import ssl
from email.message import EmailMessage
import os

# === SETUP LOGGING ===
logging.basicConfig(
    filename='automation_log.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# === STEP 1: Load CSV Data ===
def load_data(filepath):
    try:
        df = pd.read_csv(filepath)
        logging.info("CSV file loaded successfully.")
        return df
    except FileNotFoundError:
        logging.error(f"File not found: {filepath}")
        raise
    except Exception as e:
        logging.error(f"Error loading data: {e}")
        raise

# === STEP 2: Clean and Preprocess ===
def clean_data(df):
    try:
        df['Weekend'] = df['Weekend'].astype(str)
        df['Revenue'] = df['Revenue'].astype(str)

        num_cols = df.select_dtypes(include=['float64', 'int64']).columns
        df[num_cols] = df[num_cols].fillna(0)

        logging.info("Data cleaned and preprocessed successfully.")
        return df
    except Exception as e:
        logging.error(f"Error during cleaning: {e}")
        raise

# === STEP 3: Create Pivot Table ===
def create_pivot_table(df):
    try:
        pivot_df = pd.pivot_table(
            df,
            values='PageValues',
            index='VisitorType',
            columns='Weekend',
            aggfunc='mean',
            fill_value=0
        )
        logging.info("Pivot table created successfully.")
        return pivot_df
    except Exception as e:
        logging.error(f"Pivot table creation failed: {e}")
        raise

# === STEP 4: Save to CSV ===
def save_output(df, output_path):
    try:
        df.to_csv(output_path)
        logging.info(f"Pivot table saved to {output_path}")
    except Exception as e:
        logging.error(f"Saving output failed: {e}")
        raise

# === STEP 5: Send Email ===
def send_email_report(sender_email, app_password, recipient_email, subject, body, attachment_path):
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg.set_content(body)

        with open(attachment_path, 'rb') as file:
            data = file.read()
            filename = os.path.basename(attachment_path)
            msg.add_attachment(data, maintype='application', subtype='octet-stream', filename=filename)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(msg)

        logging.info("Email report sent successfully.")
        print("📧 Email sent successfully.")
    except Exception as e:
        logging.error(f"Email sending failed: {e}")
        print("❌ Failed to send email.")

# === MAIN WORKFLOW ===
def automate_workflow():
    input_file = 'online_shoppers_intention.csv'
    output_file = 'pivot_output.csv'

    print("🔄 Starting automation...")
    df = load_data(input_file)
    df = clean_data(df)
    pivot = create_pivot_table(df)
    save_output(pivot, output_file)

    # EMAIL CONFIGURATION
    sender_email = "adityachineaasplinternship@gmail.com"
    app_password = "fgeo zebn vkul njwx"  # Use Gmail App Password
    recipient_email = "adityahemantchine41@gmail.com"
    subject = "Daily Pivot Report - Online Shoppers Intention"
    body = "Please find attached the pivot table report generated today."

    send_email_report(sender_email, app_password, recipient_email, subject, body, output_file)
    print("✅ Automation complete.")

# === Run Script ===
if __name__ == "__main__":
    automate_workflow()
