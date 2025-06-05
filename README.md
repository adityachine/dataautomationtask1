# 🧠 Data Automation Task

This project automates the process of analyzing **online shopper behavior** using a dataset, generating visualizations and reports, and sending those via email on a schedule using Python.

## 📁 Project Structure
dataautomationtask1:
  - app1.py
  - app2email.py
  - app3email_automation.py
  - appgmail.py
  - appoutlook.py
  - applocalhost.py
  - automation_log.txt
  - client_emails.xlsx          # 📊 Excel file with client emails
  - recipients.csv              # 📊 Static CSV for recipients
  - online_shoppers_intention.csv  # 📊 Source data file
  - correlation_heatmap.png     # 📈 Heatmap visualization
  - exit_vs_bounce.png          # 📈 Bounce vs Exit chart
  - pagevalues_bar_chart.png    # 📈 Page values visualization
  - visitor_pie_chart.png       # 📈 Visitor distribution pie chart
  - outputexcelfile.xlsx        # 📊 Final Excel report
  - pivot_output.csv            # 📜 Processed pivot CSV
  - README.md                   # 📄 Documentation

## ⚙️ Key Features

- ✅ Loads and validates **online shopper intention data**
- ✅ Checks for missing or negative values
- ✅ Creates pivot tables for summary analysis
- ✅ Generates charts (bar, pie, heatmaps)
- ✅ Saves output in Excel and image formats
- ✅ Sends automated emails with attachments using:
  - Gmail (via `appgmail.py`)
  - Outlook SMTP (via `appoutlook.py`)
- ✅ Supports scheduling with time-based automation

---

## 📊 Sample Outputs

- `outputexcelfile.xlsx`: Cleaned data with pivot table
- `correlation_heatmap.png`: Correlation among numerical features
- `visitor_pie_chart.png`: Pie chart of visitor types
- `pagevalues_bar_chart.png`: Bar chart of page values
- `exit_vs_bounce.png`: Bounce vs Exit rates

---

## ✉️ Email Setup

- `recipients.csv`: List of email recipients
- `client_emails.xlsx`: Used for dynamic client contact info
- Configure your email credentials in the respective Python script (`appgmail.py`, `appoutlook.py`)
  > **Note**: Use app passwords if 2FA is enabled. Avoid plain passwords in code.

---

## 🧪 Scripts Overview

| Script Name           | Purpose                                    |
|-----------------------|--------------------------------------------|
| `app1.py`             | Base automation logic                      |
| `app2email.py`        | Processes and sends report via email       |
| `app3email_automation.py` | Extended automation with error logging |
| `appgmail.py`         | Uses Gmail SMTP for email automation       |
| `appoutlook.py`       | Uses Outlook SMTP                          |
| `applocalhost.py`     | Sends email via local SMTP server          |

---

## 🛠️ Setup Instructions

1. Clone the repository:
   ```bash
   git clone https://github.com/adityachine/dataautomationtask1.git
   cd dataautomationtask1
Create a virtual environment:

bash
Copy
Edit
python -m venv myenv
myenv\Scripts\activate   # On Windows
Install required packages:

bash
Copy
Edit
pip install -r requirements.txt
Run a script:

bash
Copy
Edit
python code/appgmail.py
📅 Scheduling
Each script includes scheduling logic (based on time). You can customize run frequency using Python’s schedule or time.sleep() approach.

🔐 Security Notice
Avoid committing files with real credentials.

Add .env or use environment variables to securely manage secrets.

Add .gitignore to prevent data files from being tracked.

📬 Contact
Maintained by Aditya Chine
GitHub: @adityachine



