#import tkinter library
import tkinter as tk
import xlwt

#BUILDING THE FORM WINDOW
root = tk.Tk()
#SETTING TITLE
root.title("Student Form")
#setting the window size
root.geometry("600x400")

#declaring string variable for storing details
surname_var=tk.StringVar()
first_name_var=tk.StringVar()
middle_name_var=tk.StringVar()
matric_number_var=tk.StringVar()
department_var=tk.StringVar()
score_var=tk.StringVar()
# defining a function that will get the details and print them on the screen
def submit():
    surname= surname_var.get()
    first_name = first_name_var.get()
    middle_name = middle_name_var.get()
    matric_number = matric_number_var.get()
    department = department_var.get()
    score = score_var.get()
    print(surname)
    print(first_name)
    print(middle_name)
    print(matric_number)
    print(department)
    print(score)


# use file path inside it
# with open('file.csv;'w') as file:
#   doc = DictReader(file)
#   # doc.append()
    work= xlwt.Workbook()
# Naming the sheet in the excel sheet=
    workbook = work.add_sheet("Student Details")
    workbook.write(0,  0,  'Surname')
    workbook.write(0,  1,  'First Name')
    workbook.write(0,  2,  'Middle Name')
    workbook.write(0,  3,  'Matric Number')
    workbook.write(0,  4,  'Department')
    workbook.write(0,  5,  'Score')


    workbook.write(1,  0, surname)
    workbook.write(1,  1, first_name)
    workbook.write(1,  2, middle_name)
    workbook.write(1,  3, matric_number)
    workbook.write(1,  4, department)
    workbook.write(1,  5, score)


#saving the data to excel

    work.save("StudentDetails.xls")

    surname_var.set("")
    first_name_var.set("")
    middle_name_var.set("")
    matric_number_var.set("")
    department_var.set("")
    score_var.set("")


surname_label = tk.Label(root, text='Surname', font=('Bahnschrift', 18, 'bold'))
first_name_label = tk.Label(root, text='First name', font=('Bahnschrift',18, 'bold'))
middle_name_label = tk.Label(root, text='Middle Name', font=('Bahnschrift', 18, 'bold'))
matric_number_label = tk.Label(root, text='Matric Number', font=('Bahnschrift', 18, 'bold'))
department_label = tk.Label(root, text='Department', font=('Bahnschrift', 18, 'bold'))
score_label = tk.Label(root, text='Score', font=('Bahnschrift', 18, 'bold'))

surname_entry = tk.Entry(root,  textvariable=surname_var, font=('Bahnschrift', 18, 'normal'))
first_name_entry = tk.Entry(root,  textvariable=first_name_var, font=('Bahnschrift', 18, 'normal'))
middle_name_entry = tk.Entry(root,  textvariable=middle_name_var, font=('Bahnschrift', 18, 'normal'))
matric_number_entry = tk.Entry(root, textvariable=matric_number_var, font=('Bahnschrift', 18, 'normal'))
department_entry = tk.Entry(root, textvariable=department_var, font=('Bahnschrift', 18, 'normal'))
score_entry = tk.Entry(root, textvariable=score_var, font=('Bahnschrift', 18, 'normal'))


#submit button
sub_btn = tk.Button(root, text='Submit', command=submit)
#close button
exit_button=tk.Button(root, text='Exit', command=root.destroy)


#placing the label and entry in the required position using grid method
surname_label.grid(row=0, column=0)
first_name_label.grid(row=1, column=0)
middle_name_label.grid(row=2, column=0)
matric_number_label.grid(row=3, column=0)
department_label.grid(row=4, column=0)
score_label.grid(row=5, column=0)

surname_entry.grid(row=0, column=1)
first_name_entry.grid(row=1, column=1)
middle_name_entry.grid(row=2, column=1)
matric_number_entry.grid(row=3, column=1)
department_entry.grid(row=4, column=1)
score_entry.grid(row=5, column=1)

#placing butttons in the window
sub_btn.grid(row=6, column=1)
exit_button.grid(row=7, column=1)
#performing an infinite loop for the window to display
root.mainloop()
