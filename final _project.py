# clearing the terminal befor each run
import os
clear=lambda: os.system("clear")
clear()
# ---------------------

# importing pyhton list to exell file module
import xlsxwriter

# creat empty list for student names and ther cours 
firstNames=[]
lastNames=[]
coursNames=[]
credits=[]


def ask():
    fN= input('Please enter the student first name: ')
    firstNames.append(fN)
    lN= input('Please enter student last name: ')
    lastNames.append(lN)
    next_step()
    


# input method
def next_step():
    answer=ord(input('Is there any other students? y/n: '))
    if answer==121:
        ask()
    else:
        return





new_list_0 = [['first', 'second'], ['third', 'four'], [1, 2, 65, 8, 5, 6]]
new_list = [['First Names', 'Last Names', "courses", "credits", "Marks"]]

with xlsxwriter.Workbook('test.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    for row_num, data in enumerate(new_list):
        worksheet.write_row(row_num, 0, data)

