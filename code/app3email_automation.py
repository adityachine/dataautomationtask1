import pandas as pd
import logging
import smtplib
import ssl
from email.message import EmailMessage
import os
from openpyxl.styles import Font
from openpyxl import load_workbook

# === LOGGING SETUP ===
logging.basicConfig(
    filename='automation_log.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# === LOAD CSV ===
def load_data(filepath):
    try:
        df = pd.read_csv(filepath)
        logging.info("CSV file loaded successfully.")
        return df
    except Exception as e:
        logging.error(f"Failed to load CSV: {e}")
        raise

# === CLEAN DATA ===
def clean_data(df):
    try:
        df['Weekend'] = df['Weekend'].astype(str)
        df['Revenue'] = df['Revenue'].astype(str)
        numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
        df[numeric_cols] = df[numeric_cols].fillna(0)
        logging.info("Data cleaned successfully.")
        return df
    except Exception as e:
        logging.error(f"Error during data cleaning: {e}")
        raise

# === CREATE PIVOT TABLE ===
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
        pivot['Total_Mean_PageValues'] = pivot.sum(axis=1)
        pivot = pivot.sort_values(by='Total_Mean_PageValues', ascending=False)
        logging.info("Pivot table created successfully.")
        return pivot
    except Exception as e:
        logging.error(f"Error creating pivot table: {e}")
        raise

# === SAVE EXCEL REPORT ===
def save_to_excel(clean_df, pivot_df, file_path):
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            clean_df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
            pivot_df.to_excel(writer, sheet_name='Pivot_Table')

            workbook = writer.book
            ws1 = workbook['Cleaned_Data']
            ws2 = workbook['Pivot_Table']

            for cell in ws1[1]:
                cell.font = Font(bold=True)
            for cell in ws2[1]:
                cell.font = Font(bold=True)

        logging.info(f"Excel report saved at: {file_path}")
    except Exception as e:
        logging.error(f"Error saving Excel: {e}")
        raise

# === GET EMAIL LIST ===
def get_email_list(filepath):
    try:
        df = pd.read_excel(filepath)
        df.columns = [col.strip().lower() for col in df.columns]
        email_col = next((col for col in df.columns if 'email' in col), None)

        if not email_col:
            logging.error("Email column not found in Excel file.")
            return []

        emails = df[email_col].dropna().astype(str)
        valid_emails = [email for email in emails if '@' in email and '.' in email]
        logging.info(f"Loaded {len(valid_emails)} valid emails.")
        return valid_emails
    except Exception as e:
        logging.error(f"Error reading email list: {e}")
        return []

# === SEND EMAIL ===
def send_email_report(sender, password, recipients, subject, body, attachment):
    try:
        if not recipients:
            logging.error("No valid recipients to send email.")
            print("❌ No valid recipients found.")
            return

        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = ', '.join(recipients)
        msg.set_content(body)

        with open(attachment, 'rb') as file:
            msg.add_attachment(
                file.read(),
                maintype='application',
                subtype='octet-stream',
                filename=os.path.basename(attachment)
            )

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(sender, password)
            smtp.send_message(msg)

        logging.info("Email report sent successfully.")
        print("📧 Email sent.")
    except Exception as e:
        logging.error(f"Email sending failed: {e}")
        print("❌ Failed to send email.")

# === MAIN AUTOMATION WORKFLOW ===
def automate_workflow():
    input_csv = 'online_shoppers_intention.csv'
    email_excel = 'email_list.xlsx'
    output_excel = 'pivot_report.xlsx'

    print("🔄 Starting automation...")

    try:
        df = load_data(input_csv)
        clean_df = clean_data(df)
        pivot_df = create_pivot_table(clean_df)
        save_to_excel(clean_df, pivot_df, output_excel)

        sender = "adityachineaasplinternship@gmail.com"
        password = "fgeo zebn vkul njwx"  # Use Gmail App Password here
        recipients = get_email_list(email_excel)

        subject = "📊 Daily Pivot Report - Online Shoppers Intention"
        body = "Hi,\n\nPlease find today's attached report.\n\nBest regards,\nAutomation Bot"
        send_email_report(sender, password, recipients, subject, body, output_excel)

        print("✅ Automation complete.")
    except Exception as e:
        logging.error(f"Automation workflow failed: {e}")
        print("❌ Automation failed.")

# === RUN SCRIPT ===
if __name__ == "__main__":
    automate_workflow()
