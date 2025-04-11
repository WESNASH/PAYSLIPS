# Import necessary libraries
import pandas as pd
import os
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Retrieve credentials
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASS = os.getenv('EMAIL_PASS')

# Read employee data
data = pd.read_excel('employees 1.xlsx', sheet_name='Sheet1')

# Calculate net salary
data["Net Salary"] = data["Basic salary"] + data["Allowances"] - data["Deductions"]

# Ensure the 'payslips' directory exists
os.makedirs('payslips', exist_ok=True)

# Function to generate PDF payslips
def generate_pdf(employee):
 try:
        #
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt=f"Payslip for {employee['Name']}", ln=True, align='C')
    pdf.cell(200, 10, txt=f"Employee ID: {employee['Employee ID']}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: {employee['Basic salary']}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: {employee['Allowances']}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: {employee['Deductions']}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: {employee['Net Salary']}", ln=True)

    # Save PDF in the 'payslips' directory
    pdf_file_path = f'payslips/{employee["Name"].replace(" ", "_")}_payslip.pdf'
    pdf.output(pdf_file_path)
    print(f"Generated payslip at: {pdf_file_path}")
 except Exception as e:
    print(f"Error: {employee['Name']} does not have a valid email address. Payslip not generated.")

# Generate payslips for each employee
for index, employee in data.iterrows():
    generate_pdf(employee)

# Function to send email with payslip
def send_email(recipient, employee):
    t
    attachment = f'payslips/{employee["Name"].replace(" ", "_")}_payslip.pdf'
    if os.path.exists(attachment):
        yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASS)
        subject = f"Your Payslip for {employee['Name']}"
        body = "Please find attached your payslip."
        yag.send(to=recipient, subject=subject, contents=body, attachments=attachment)
        print(f"Email sent to {recipient}.")
    else:
        print(f"Error: Payslip {attachment} does not exist.")

# Send emails with payslips
for index, employee in data.iterrows():
    send_email(employee['Email'], employee)