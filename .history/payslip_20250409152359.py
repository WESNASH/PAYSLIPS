####Payslip generator 
# Import necessary libraries
import pandas as pd

data = pd.read_excel('employees 1.xlsx', sheet_name='Sheet1')

print(data)

# calculating net salary 
data["Net Salary"] = data["Basic salary"] + data["Allowances"] - data["Deductions"]
print(data)

df = pd.DataFrame(data)

class PayslipPDF(FPDF):
    def header(self):
        self.set_font("Arial", 'B', 16)
        self.cell(0, 10, 'Uncommon.org', ln=True, align='C')
        self.ln(10)  # Line break

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def generate_pdf(employee):
    pdf = PayslipPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Add content with better alignment and spacing
    pdf.cell(0, 10, '', ln=True)  # Empty line for spacing
    pdf.cell(0, 10, txt=f"Payslip for: {employee['Name']}", ln=True, align='C')
    pdf.cell(0, 10, txt=f"Employee ID: {employee['Employee ID']}", ln=True, align='C')
    pdf.cell(0, 10, '', ln=True)  # Empty line for spacing

    pdf.cell(100, 10, txt="Basic Salary:", align='R')
    pdf.cell(0, 10, txt=f"${employee['Basic salary']:.2f}", ln=True)

    pdf.cell(100, 10, txt="Allowances:", align='R')
    pdf.cell(0, 10, txt=f"${employee['Allowances']:.2f}", ln=True)

    pdf.cell(100, 10, txt="Deductions:", align='R')
    pdf.cell(0, 10, txt=f"${employee['Deductions']:.2f}", ln=True)

    pdf.cell(100, 10, txt="Net Salary:", align='R')
    pdf.cell(0, 10, txt=f"${employee['Net Salary']:.2f}", ln=True)

    # Save the PDF
    pdf_file_path = f"payslips/{employee['Name']}_payslip.pdf"
    os.makedirs('payslips', exist_ok=True)  # Ensure the directory exists
    pdf.output(pdf_file_path)

# Generate payslips for each employee
for index, employee in df.iterrows():
    generate_pdf(employee)
    print(f"Payslip generated for {employee['Name']}")

print("All payslips generated successfully!")

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


    
    
    


