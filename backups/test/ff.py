import xlsxwriter

# create a new Excel file
workbook = xlsxwriter.Workbook('example.xlsx')

# create a new worksheet
worksheet = workbook.add_worksheet()

# define the header row
header = ['Name', 'Age', 'Gender']
row = 0
col = 0

# write the header row
for field in header:
    worksheet.write(row, col, field)
    col += 1

# write the data rows
data = [
    {'Name': 'John', 'Age': 1, 'Gender': 'Male'},
    {'Name': 'Sarah', 'Age': 4, 'Gender': 'Female'},
    {'Name': 'Michael', 'Age': 27, 'Gender': 'Male'},
]


for item in data:
    row += 1
    col = 0
    for key, value in item.items():
        worksheet.write(row, col, value)
        col += 1

# close the Excel file
workbook.close()