import os
import k
import pandas as pd
from fpdf import FPDF
import yagmail
from datetime import datetime

def load_data_from_excel(file_path):
    """Load employee data from Excel file."""
    df = pd.read_excel(file_path)
    required_columns = ['Employee ID', 'Name', 'Email Address', 'Basic Salary', 'Allowance', 'Deductions']
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in Excel file: {', '.join(missing)}")
    return df

def generate_payslip(row):
    """Generate a PDF payslip for a given employee row."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Title
    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(200, 10, txt="Payslip", ln=True, align="C")
    pdf.ln(10)

    # Employee Information
    pdf.set_font("Arial", size=12)
    pdf.cell(100, 10, f"Employee ID: {row['Employee ID']}", ln=True)
    pdf.cell(100, 10, f"Name: {row['Name']}", ln=True)
    pdf.cell(100, 10, f"Email: {row['Email Address']}", ln=True)
    pdf.cell(100, 10, f"Basic Salary: {row['Basic Salary']:,.2f}", ln=True)
    pdf.cell(100, 10, f"Allowance: {row['Allowance']:,.2f}", ln=True)
    pdf.cell(100, 10, f"Deductions: {row['Deductions']:,.2f}", ln=True)

    # Calculate Net Salary
    net_pay = row['Basic Salary'] + row['Allowance'] - row['Deductions']
    pdf.cell(100, 10, f"Net Pay: {net_pay:,.2f}", ln=True)
    pdf.ln(10)

    # Generate filename
    current_month = datetime.now().strftime('%B_%Y')
    sanitized_name = ''.join(c for c in row['Name'] if c.isalnum())
    file_name = f"Payslip_{sanitized_name}_{current_month}.pdf"

    # Save PDF
    pdf.output(file_name)
    return file_name

def send_email(subject, body, recipient_email, payslip_file):
    """Send the email with the payslip attached."""
    try:
        yag = yagmail.SMTP(user='roymakanjira@gmail.com', password=os.getenv("YAGMAIL_APP_PASSWORD"))
        yag.send(
            to=recipient_email,
            subject=subject,
            contents=body,
            attachments=payslip_file
        )
        print(f"Payslip sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

def main():
    file_path = 'employeesDetails.xlsx'
    data = load_data_from_excel(file_path)

    for index, row in data.iterrows():
        try:
            payslip_file = generate_payslip(row)
            subject = f"Payslip for {row['Name']}"
            body = f"""
Hello {row['Name']},

Please find attached your payslip for the current month.

Best regards,  
Your Company
            """
            send_email(subject, body, row['Email Address'], payslip_file)
        except Exception as e:
            print(f"Error processing {row['Name']}: {e}")
        finally:
            if os.path.exists(payslip_file):
                os.remove(payslip_file)

if __name__ == "__main__":
    main()
