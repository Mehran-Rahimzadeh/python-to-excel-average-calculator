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



# asking funxtion
def ask():
    globals()["student"]= dict(firstName="" , lastName = "", courses=[], credits=[], marks=[])
    name_ask()
    courses_ask()
    next_step_student()


# getting first and last name
def name_ask():
    student["firstName"]=input("Enter the student name: ")
    student["lastName"]=input ("Enter the student last name: ")


# getting courses
def courses_ask():
    student["courses"].append(input ("Enter the student course: "))
    student["credits"].append(int(input ("Enter the course credit: ")))
    student["marks"].append(float(input ("Enter the course mark: ")))
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
    



def make_title_exell():
    new_list = [['First Names', 'Last Names', "courses", "credits", "Marks"]]
    with xlsxwriter.Workbook('test.xlsx') as workbook:
        worksheet = workbook.add_worksheet()

        for row_num, data in enumerate(new_list):
            worksheet.write_row(row_num, 0, data)

        for row_num, data in enumerate(new_list):
            worksheet.write_row(row_num, 0, data)




ask()
make_title_exell()


