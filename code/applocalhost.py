import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import logging
import smtplib
import ssl
import os
import schedule
import time
from email.message import EmailMessage

# === SETUP LOGGING ===
logging.basicConfig(
    filename='automation_log.txt',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

# === Load Data ===
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

# === Clean Data ===
def clean_data(df):
    try:
        df['Weekend'] = df['Weekend'].astype(str)
        df['Revenue'] = df['Revenue'].astype(str)
        num_cols = df.select_dtypes(include=['float64', 'int64']).columns
        df[num_cols] = df[num_cols].fillna(0)
        logging.info("Data cleaned successfully.")
        return df
    except Exception as e:
        logging.error(f"Error during cleaning: {e}")
        raise

# === Create Advanced Pivot Table ===
def create_pivot_table(df):
    try:
        pivot = pd.pivot_table(
            df,
            values='PageValues',
            index=['VisitorType', 'Revenue'],
            columns='Weekend',
            aggfunc=['mean', 'sum', 'count'],
            fill_value=0
        )
        logging.info("Advanced pivot table created.")
        return pivot
    except Exception as e:
        logging.error(f"Pivot table error: {e}")
        raise

# === Visualizations ===
def create_visualizations(df):
    try:
        sns.set(style="whitegrid")

        plt.figure(figsize=(10, 6))
        sns.boxplot(x="Revenue", y="PageValues", data=df)
        plt.title("Page Values by Revenue")
        plt.savefig("boxplot_pagevalues_by_revenue.png")
        plt.close()

        plt.figure(figsize=(10, 6))
        sns.countplot(x="VisitorType", hue="Weekend", data=df)
        plt.title("Visitor Type vs Weekend")
        plt.savefig("visitor_weekend_count.png")
        plt.close()

        logging.info("Visualizations created.")
    except Exception as e:
        logging.error(f"Visualization error: {e}")
        raise

# === Save Files ===
def save_output(df, output_path):
    try:
        df.to_csv(output_path)
        logging.info(f"Pivot table saved to {output_path}")
    except Exception as e:
        logging.error(f"Saving pivot table failed: {e}")
        raise

# === Extract Recipient Email from Data ===
def get_recipient_email(df):
    try:
        email_col = [col for col in df.columns if 'email' in col.lower()]
        if email_col:
            return df[email_col[0]].dropna().unique()[0]  # take first non-null unique email
        else:
            raise Exception("No email column found.")
    except Exception as e:
        logging.error(f"Email extraction failed: {e}")
        raise

# === Send Email ===
def send_email_report(sender_email, app_password, recipient_email, subject, body, attachments):
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg.set_content(body)

        for path in attachments:
            with open(path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(path)
                msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(sender_email, app_password)
            smtp.send_message(msg)

        logging.info("Email sent successfully.")
        print("📧 Email sent.")
    except Exception as e:
        logging.error(f"Email failed: {e}")
        print("❌ Failed to send email.")

# === Main Workflow ===
def automate_workflow():
    print("🔄 Starting automation...")
    input_file = 'online_shoppers_intention.csv'
    output_file = 'pivot_output.csv'

    df = load_data(input_file)
    recipient_email = get_recipient_email(df)
    df = clean_data(df)
    pivot = create_pivot_table(df)
    save_output(pivot, output_file)
    create_visualizations(df)

    sender_email = "adityachineaasplinternship@gmail.com"
    app_password = "fgeo zebn vkul njwx"
    subject = "Automated Pivot & Analysis Report"
    body = "Attached are today's pivot report and visual analytics."

    attachments = [output_file, "boxplot_pagevalues_by_revenue.png", "visitor_weekend_count.png"]
    send_email_report(sender_email, app_password, recipient_email, subject, body, attachments)
    print("✅ Workflow complete.")

# === Schedule Automation ===
schedule.every().day.at("15:05").do(automate_workflow)  # Runs daily at 8 AM

if __name__ == "__main__":
    print("⏳ Waiting for scheduled job...")
    while True:
        schedule.run_pending()
        time.sleep(60)
