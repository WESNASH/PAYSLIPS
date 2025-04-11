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
def generate_pdf(employee):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Title
        pdf.cell(200, 10, txt=f"Payslip for {employee['Name']}", ln=True, align='C')
        pdf.ln(10)  # Add a line break

        # Table Header
        pdf.set_font("Arial", 'B', size=12)
        pdf.cell(80, 10, "Description", border=1)
        pdf.cell(40, 10, "Amount", border=1)
        pdf.ln()  # Move to the next line

        # Table Rows
        pdf.set_font("Arial", size=12)
        pdf.cell(80, 10, "Employee ID", border=1)
        pdf.cell(40, 10, str(employee["Employee ID"]), border=1)
        pdf.ln()

        pdf.cell(80, 10, "Basic Salary", border=1)
        pdf.cell(40, 10, str(employee["Basic salary"]), border=1)
        pdf.ln()

        pdf.cell(80, 10, "Allowances", border=1)
        pdf.cell(40, 10, str(employee["Allowances"]), border=1)
        pdf.ln()

        pdf.cell(80, 10, "Deductions", border=1)
        pdf.cell(40, 10, str(employee["Deductions"]), border=1)
        pdf.ln()

        pdf.cell(80, 10, "Net Salary", border=1)
        pdf.cell(40, 10, str(employee["Net Salary"]), border=1)
        pdf.ln()

        # Save PDF in the 'payslips' directory
        pdf_file_path = f'payslips/{employee["Name"].replace(" ", "_")}_payslip.pdf'
        pdf.output(pdf_file_path)
        print(f"Generated payslip at: {pdf_file_path}")
    
    except Exception as e:
        print(f"Error generating payslip for {employee['Name']}: {e}")

# Send emails with payslips
for index, employee in data.iterrows():
    send_email(employee['Email'], employee)