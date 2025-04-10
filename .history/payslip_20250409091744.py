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
    pdf.add_age()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
    pdf.add_page()
    


