####Payslip generator 
# Import necessary libraries
import pandas as pd

data = pd.read_excel("employees.xlsx")

print(data.head())
wesnash = data[data['Name'] == 'Wesley Nash']

