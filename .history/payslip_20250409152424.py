####Payslip generator 
# Import necessary libraries
import pandas as pd

data = pd.read_excel('employees 1.xlsx', sheet_name='Sheet1')

print(data)

# calculating net salary 
data["Net Salary"] = data["Basic salary"] + data["Allowances"] - data["Deductions"]
print(data)



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


    
    
    


