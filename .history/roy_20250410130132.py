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
    pdf.add_page()a

    # Title
    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(200, 10, txt="Payslip", ln=True, align="C")
    pdf.ln(10)

    # Employee Information
    pdf.set_font("Arial", size=12)
    pdf.cell(100, 10, f"Employee ID: {row['Employee ID']}", ln=True)
    pdf.cell(100, 10, f"Name: {row['Name']}", ln=True)
    pdf.cell(100, 10, f"Email: {row['Email Address']}", ln=True)
    pdf.cell(100, 10, f"Basic Salary: {row['Basic Salary']:.2f}", ln=True)
    pdf.cell(100, 10, f"Allowance: {row['Allowance']:.2f}", ln=True)
    pdf.cell(100, 10, f"Deductions: {row['Deductions']:.2f}", ln=True)

    # Calculate Net Salary
    net_pay = row['Basic Salary'] + row['Allowance'] - row['Deductions']
    pdf.cell(100, 10, f"Net Pay: {net_pay:.2f}", ln=True)
    pdf.ln(10)

    # Generate filename
    current_month = datetime.now().strftime('%B_%Y')
    file_name = f"Payslip_{row['Name'].replace(' ', '')}{current_month}.pdf"

    # Save PDF
    pdf.output(file_name)
    return file_name

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