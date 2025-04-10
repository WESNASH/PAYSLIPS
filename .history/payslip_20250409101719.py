####Payslip generator 
# Import necessary libraries
import pandas as pd

data = pd.read_excel('employees 1.xlsx', sheet_name='Sheet1')

print(data.head())

# calculating net salary 
data["Net Salary"] = data["Basic salary"] + data["Allowances"] - data["Deductions"]
print(data)

#generate pdf payslips
from fpdf import FPDF

def generate_pdf(employee):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt = f"Payslip for {employee['Name']}", ln = True, align = 'C')
    pdf.cell(200, 10, txt = f"Employee ID: {employee['Employee ID']}", ln = True, align = 'C')
    pdf.cell(200, 10, txt = f"Basic Salary: {employee['Basic salary']}", ln = True, align = 'C')
    pdf.cell(200, 10, txt = f"Allowances: {employee['Allowances']}", ln = True, align = 'C')
    pdf.cell(200, 10, txt = f"Deductions: {employee['Deductions']}", ln = True, align = 'C')
###save pdf 
    pdf_file_path = f"{employee['Name']}_payslip.pdf"
    pdf.output(pdf_file_path)

#### Generate payslips for each employee
for index, employee in data.iterrows():
    generate_pdf(employee)
    print(f"Payslip generated for {employee['Name']}")

# The above code will generate a PDF payslip for each employee in the Excel file and save it with the employee's name.

##Emailing Payslips
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(employee_email, employee_name, payslip_path):
    msg = MIMEMultipart()
    msg['From'] = 'nyabun'
    msg['To'] = employee_email
    msg['Subject'] = "Your Payslip for This Month"

    body = f"Dear {employee_name},\n\nPlease find attached your payslip for this month."
    msg.attach(MIMEText(body, 'plain'))

    with open(payslip_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={payslip_path}')
        msg.attach(part)

    with smtplib.SMTP('smtp.example.com', 587) as server:
        server.starttls()
        server.login('your_email@example.com', 'your_password')
        server.send_message(msg)
    
    
    


