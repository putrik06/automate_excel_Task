import openpyxl

workbook =  openpyxl.load_workbook("Employees.xlsx")
# Print the workbook properties
# print(workbook.properties)

# Explore the workbook sheetnames
print(workbook.sheetnames) # have 3 sheets ['EmployeeData', 'Salaries', 'Skills']

# which workbook is the active one?
print(workbook.active) # commonly the first is the active one.
# >>> <Worksheet "EmployeeData">

# Reference the workbook to perform other operation
sheet = workbook['EmployeeData']

# # Make a new sheet
workbook.create_sheet('NewSheet')
# save the new worksheet
workbook.save("Employees.xlsx")
# check if the new worksheet is present
print(workbook.sheetnames)

# delete the worksheet
sheet = workbook['NewSheet']
workbook.remove(sheet)

# check if it removed
print(workbook.sheetnames)
workbook.save("Employees.xlsx")