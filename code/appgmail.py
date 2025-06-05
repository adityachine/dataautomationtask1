import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import smtplib
from email.message import EmailMessage
from pathlib import Path
import schedule
import time
import os

# === CONFIGURATION ===
CSV_FILE = 'online_shoppers_intention.csv'
SENDER_EMAIL = 'adityachineaasplinternship@gmail.com'
APP_PASSWORD = 'wxpt weul cgre cdoc'
SUBJECT = "Automated Pivot & Analysis Report"
BODY = "Attached are today's pivot report and visual analytics."
RECIPIENTS = [
    "adityabhartichine4141@gmail.com",
    "adityahemantchine41@gmail.com",
    "aditya.chine@alignedautomation.com"
]

def run_analysis_and_send_email():
    try:
        sns.set(style="whitegrid")

        print("🔄 Loading CSV file...")
        df = pd.read_csv(CSV_FILE)
        print(f"✅ Loaded: {df.shape[0]} rows, {df.shape[1]} columns.")

        print("🚨 Checking missing/negative values...")
        print(df.isnull().sum().to_frame("Missing Values"))
        print((df.select_dtypes(include='number') < 0).sum().to_frame("Negative Values"))

        print("📌 Sorting by 'BounceRates'...")
        sorted_df = df.sort_values(by='BounceRates', ascending=False) if 'BounceRates' in df.columns else df.copy()

        print("📊 Creating pivot table...")
        pivot = pd.pivot_table(
            df,
            index=['VisitorType', 'Weekend'],
            values=['PageValues', 'BounceRates', 'ExitRates', 'Administrative_Duration'],
            aggfunc=['mean', 'sum', 'count'],
            margins=True,
            margins_name='Grand Total'
        )

        print("📈 Creating charts...")
        charts = {
            "visitor_pie_chart.png": lambda: df['VisitorType'].value_counts().plot.pie(
                autopct='%1.1f%%', title='Visitor Type Distribution'),
            "pagevalues_bar_chart.png": lambda: df.groupby('VisitorType')['PageValues'].mean().plot(
                kind='bar', title='Avg PageValues by Visitor Type', color='skyblue'),
            "exit_vs_bounce.png": lambda: sns.lineplot(data=df, x='ExitRates', y='BounceRates'),
            "correlation_heatmap.png": lambda: sns.heatmap(
                df.corr(numeric_only=True), annot=True, cmap='coolwarm', fmt=".2f")
        }

        for filename, plot_func in charts.items():
            plt.figure(figsize=(8, 6))
            plot_func()
            plt.tight_layout()
            plt.savefig(filename)
            plt.close()

        print("💾 Saving Excel file...")
        with pd.ExcelWriter("outputexcelfile.xlsx", engine='openpyxl') as writer:
            sorted_df.to_excel(writer, index=False, sheet_name='Sorted_Data')
            pivot.to_excel(writer, sheet_name='Pivot_Insights')

        print("📨 Composing email...")
        msg = EmailMessage()
        msg['From'] = SENDER_EMAIL
        msg['To'] = ', '.join(RECIPIENTS)
        msg['Subject'] = SUBJECT
        msg.set_content(BODY)

        attachments = [
            "outputexcelfile.xlsx",
            "visitor_pie_chart.png",
            "pagevalues_bar_chart.png",
            "exit_vs_bounce.png",
            "correlation_heatmap.png"
        ]

        for file_path in attachments:
            with open(file_path, 'rb') as f:
                file_data = f.read()
                file_name = Path(file_path).name
                maintype = "application" if file_path.endswith('.xlsx') else "image"
                subtype = "octet-stream" if file_path.endswith('.xlsx') else "png"
                msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

        print("📤 Sending email...")
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(SENDER_EMAIL, APP_PASSWORD)
                smtp.send_message(msg)
            print("✅ Email sent successfully.")
        except Exception as email_err:
            print("❌ Email failed to send:", email_err)

    except Exception as e:
        print("❌ Error occurred during analysis or report creation:", e)


# ========================
# ⏰ SCHEDULE TASK DAILY
# ========================
schedule.every().day.at("15:55").do(run_analysis_and_send_email)

print("⏳ Scheduler started. Waiting for scheduled time...")

while True:
    schedule.run_pending()
    time.sleep(60)
