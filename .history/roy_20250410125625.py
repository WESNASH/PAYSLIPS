import sys
print(sys.executable)
import pandas as pd
from fpdf import FPDF
import yagmail
from datetime import datetime


def load_data_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df

def generate_payslip(row):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    

def send_email(subject, body, recipient_email, payslip_file):
    # Use an App Password instead of your real password!
    yag = yagmail.SMTP('roymakanjira@gmail.com', 'your_app_password_here')
    yag.send(
        to=recipient_email,
        subject=subject,
        contents=body,
        attachments=payslip_file
    )

def main():
    file_path = 'employeesDetails.xlsx'  # Excel file should be in the same directory
    data = load_data_from_excel(file_path)

    for index, row in data.iterrows():
        payslip_file = generate_payslip(row)

        subject = f"Payslip for {row['Name']}"
        body = f"""
Hello {row['Name']},

Please find attached your payslip for the current month.

Best regards,  
Your Company
        """

        send_email(subject, body, row['Email Address'], payslip_file)

        print(f"Payslip for {row['Name']} has been sent to {row['Email Address']}!")
        print(f"Generated file: {payslip_file}")

# Ensure the script runs only when executed directly
if __name__ == "__main__":
    main()