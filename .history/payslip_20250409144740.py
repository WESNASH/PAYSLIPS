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
    pdf.cell(200, 10, txt = f"Basic Salary: {employee['Basic salary']}", ln = )
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
email = input("Enter your email: ")
receiver_email = input("Enter receiver email: ")

subject = input("Enter subject: ")
body = input("Enter body: ")

text = f"Subject: {subject}\n\n{body}"

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()

server.login(email, "Tadiwanash21") 
server.sendmail(email,receiver_email,text) 

print("email has benn sent to" + receiver_email)
    
    
    


