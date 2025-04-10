####Payslip generator 
# Import necessary libraries
import pandas as pd

data = pd.read_excel('employees 1.xlsx', sheet_name='Sheet1')

print(data)

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
    pdf.cell(200, 10, txt = f"Employee ID: {employee['Employee ID']}", ln = True)
    pdf.cell(200, 10, txt = f"Basic Salary: {employee['Basic salary']}", ln = True)
    pdf.cell(200, 10, txt = f"Allowances: {employee['Allowances']}", ln = True, )
    pdf.cell(200, 10, txt = f"Deductions: {employee['Deductions']}", ln = True, )
    pdf.cell(200, 10, txt = f"Net Salary: {employee['Net Salary']}", ln = True, )
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
from email.mime.base import MIMEBase
from email import encoders

def send_email(employee_email, employee_name, payslip_path):
    msg = MIMEMultipart()
    msg['From'] = 'nyabungawesley@gmail.com'  # Your email address
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

    # SMTP configuration
    smtp_server = 'smtp.gmail.com'  # Gmail's SMTP server
    smtp_port = 587  # Commonly used port for TLS
    email_user = 'nyabungawesley@gmail.com'  # Your email address
    email_password = 'Tadiwanash21'  # Replace with your email password or App Password

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Upgrade the connection to secure
            server.login(email_user, email_password)
            server.send_message(msg)
        print(f"Email sent to {employee_email} successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Example usage
send_email('wesnashtadiwa.com', 'employee Name', 'payslip.pdf')

    
    
    


