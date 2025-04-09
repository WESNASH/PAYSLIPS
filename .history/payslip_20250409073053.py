####Payslip generator 
# Import necessary libraries
<<<<<<< Tabnine <<<<<<<
import pandas as pd
>>>>>>> Tabnine >>>>>>># {"conversationId":"173ad653-a6d6-4a50-a6b9-beb78e57f5ba","source":"instruct"}

data = pd.read_excel("employees.xlsx")

print(data.head())
wesnash = data[data['Name'] == 'Wesley Nash']

