import logging
import pandas as pd
import teradatasql
import win32com.client as win32
import os

# --- Configurations ---
TERADATA_HOST = "your_teradata_host"
TERADATA_USER = "your_username"
TERADATA_PASS = "your_password"
SQL_QUERY = "SELECT * FROM your_table"
EXCEL_OUTPUT_PATH = "output/report.xlsx"
EMAIL_TO = ["recipient1@example.com", "recipient2@example.com"]
EMAIL_SUBJECT = "Automated Data Report"
EMAIL_BODY = "Attached is the latest data report."

# --- Logging setup ---
logging.basicConfig(filename='automation.log', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

def run_query_and_export():
    try:
        conn = teradatasql.connect(
            host=TERADATA_HOST,
            user=TERADATA_USER,
            password=TERADATA_PASS
        )
        df = pd.read_sql(SQL_QUERY, conn)
        conn.close()

        os.makedirs(os.path.dirname(EXCEL_OUTPUT_PATH), exist_ok=True)
        df.to_excel(EXCEL_OUTPUT_PATH, index=False)
        logging.info("Data exported to Excel successfully.")
        return True
    except Exception as e:
        logging.error(f"Failed to export data: {e}")
        return False

def send_email_with_outlook(to_list, subject, body, attachment_path):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = "; ".join(to_list)
        mail.Subject = subject
        mail.Body = body

        if attachment_path and os.path.isfile(attachment_path):
            mail.Attachments.Add(os.path.abspath(attachment_path))

        mail.Send()
        logging.info("Email sent successfully via Outlook.")
        return True
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return False

if __name__ == "__main__":
    if run_query_and_export():
        send_email_with_outlook(EMAIL_TO, EMAIL_SUBJECT, EMAIL_BODY, EXCEL_OUTPUT_PATH)
