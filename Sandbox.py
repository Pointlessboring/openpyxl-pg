##########
#
#

from openpyxl import Workbook

wb = Workbook()             # Creates a Workbook

ws = wb.active              # Created Workbook has ONE sheet. Select it.
ws.title = "MySheet#1"   # Set Sheet name of 1st sheet

for i in range(1,10):
    wb.create_sheet(f"MySheet#{i+1}")   # Create 9 more sheet and set their name

filename = input('filename: ')
filename +='.xlsx'

wb.save(filename)
