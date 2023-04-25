# clearing the terminal befor each run
import os
clear=lambda: os.system("clear")
clear()
# ---------------------

# importing pyhton list to exell file module
import xlsxwriter


# creat empty list for student names and ther cours 
a=[]
student=0


#asking funxtion
def ask():
    globals()["student"]= dict(firstName="" , lastName = "", courses=[], credits=[], scores=[])
    name_ask()
    courses_ask()
    next_step_student()


# ---------getting first and last name--------
def name_ask():
    student["firstName"]=input("Enter the student name: ")
    student["lastName"]=input ("Enter the student last name: ")


#--------------getting courses---------
def courses_ask():
    student["courses"].append(input ("Enter the student course: "))
    student["credits"].append(int(input ("Enter the course credit: ")))
    student["scores"].append(float(input ("Enter the course sores: ")))
    next_step_courses()




 # input method (courses)
def next_step_courses():
    answer=ord(input('Is there any other courses? y/n: '))
    if answer==121:
        courses_ask()
    else:
        return
    


    # input method (courses)
def next_step_student():
    answer=ord(input('Is there any other students? y/n: '))
    if answer==121:
        a.append(student)
        ask()
    else:
        a.append(student)
        print(*a, sep="\n")

# ask()
# make_title_exell()


#--------------------------------Writing Data in to an Exell file---------------------------------------------------
#                ---------------------------------------------------------------------------------
#                              ----------------------------------


#Create an new Excel file and add a worksheet.
workbook=xlsxwriter.Workbook("karnameh_1.xlsx")
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column("A:A", 10)
worksheet.set_column("B:B", 10)
worksheet.set_column("C:C", 40)
worksheet.set_column("D:D", 20)
worksheet.set_column("E:E", 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({"bold":True})


# Write some header.
worksheet.write('A1', "First Name", bold)
worksheet.write('B1', 'Last Name', bold)
worksheet.write('C1', 'Courses', bold)
worksheet.write('D1', 'Credits', bold)
worksheet.write('E1', 'Scores', bold)


workbook.close()



# ---------------------------------Inserting Data in Exell file and make version 2-------------------


#
import openpyxl 
from openpyxl import Workbook
#wb = Workbook()



# open workbook
from openpyxl import load_workbook
wb = load_workbook(filename = 'karnameh_1.xlsx')
ws = wb.active

ws["A4"]=56
c= ws.cell(row=12, column=9, value= 0)

# ws1 =wb.create_sheet("Mysheet")


#Ranges of cells 
c.value = 'hello, world'


wb.save("karnameh_2.xlsx")
#print(wb.sheetnames)


























#












# def make_title_exell():
#     new_list = [['First Names', 'Last Names', "courses", "credits", "Marks"]]
#     with xlsxwriter.Workbook('test.xlsx') as workbook:
#         worksheet = workbook.add_worksheet()

#         for row_num, data in enumerate(new_list):
#             worksheet.write_row(row_num, 0, data)

#         for row_num, data in enumerate(new_list):
#             worksheet.write_row(row_num, 0, data)