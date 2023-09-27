from openpyxl import Workbook
from openpyxl import load_workbook

# Create a workbook object
# wb = Workbook()

#load existing spreadsheet
wb = load_workbook('hello.xlsx')

# create an active worksheet
ws = wb.active

# set a variable
# num = ws['A2'].value
# letter = ws['B2'].value

# Print something from the spreadsheet
# print(num, letter)

# Grab a whole column
"""
num_list = []
column_a = ws['A']
for cell in column_a:
    #print(f'{cell.value}\n')
    num_list.append(cell.value)

print(f'{ws["A1"].value} in column A are {num_list}')
"""

def sum(_list: list, total=0):
    for value in _list:
        total += int(value)
    return total

num_list = []
range = ws['A2':'A10']

for tuple in range:
    for cell in tuple:
        num_list.append(cell.value)

print(f"{ws['A1'].value} in column A are {num_list}")
print(f"Sum of {ws['A1'].value} in column A are {sum(num_list)}")