# clearing the terminal befor each run
import os
clear=lambda: os.system("clear")
clear()
# ---------------------

# importing pyhton list to exell file module
import xlsxwriter

# creat empty dics for student names and ther cours 
firstStudent= dict(firstName="" , lastName = "", courses=[], credits=[], marks=[])


# asking funxtion
def ask():
    
    print(firstStudent)
    name_ask()
    courses_ask()
    next_step_student()


# getting first and last name
def name_ask():
    firstStudent["firstName"]=input("Enter the student name: ")
    firstStudent["lastName"]=input ("Enter the student last name: ")

# getting courses
def courses_ask():
    firstStudent["courses"].append(input ("Enter the student course: "))
    firstStudent["credits"].append(int(input ("Enter the course credit: ")))
    firstStudent["marks"].append(float(input ("Enter the course mark: ")))
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
        ask()
    else:
        return
    



ask()
print(firstStudent)


































# new_list = [['First Names', 'Last Names', "courses", "credits", "Marks"]]

# with xlsxwriter.Workbook('test.xlsx') as workbook:
#     worksheet = workbook.add_worksheet()

#     for row_num, data in enumerate(new_list):
#         worksheet.write_row(row_num, 0, data)

#     for row_num, data in enumerate(new_list_0):
#         worksheet.write_row(row_num, 0, data)

