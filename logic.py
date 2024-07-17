import openpyxl
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# Workbook
book = openpyxl.load_workbook('D:\\attendance.xlsx')
ws = book['attendance']
# staff emails
staff_mail=['targyarg13@gmail.com']
# Max amount of missed days
attendance_threshold= 3
# Chooses the sheet
sheet = book['Sheet1']
#counting number of rows / student
r = sheet.max_row
# number of days students have missed
no_of_days= []
# list of students to remind
l1 =[]
# warning messages
m1 = "Warning!!! you can only miss more day for CI class"
m2 = " Warning !!! you can only miss one more day for python class"
m3 = "Warning!!! you can only miss one more day for DM class"

# saves excel on every update
def savefile():
    book.save(r'c:\Users\jorda\OneDrive\Documents\attendance.xlsx')
    print("saved!")

# To track attendance
def sendemail(to_address, subject, body):
    msg = MIMEMultipart()
    msg['From'] = staff_mail
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))



    # for students if missed days
    if missed_days >= attendance_threshold:
        subject = "Attendance warning"
        body = (f"Dear {student_name},\n\n"
                f"You have missed {missed_days} days of class. " 
                "Please be aware that you are approaching the max days allowed to miss."
                "Best Regards, \n"
                "Attendance office")
        send_email(student_email, subject, body)
        print(f"Email sent to {student_name} at {student_email}")




# iterates over the rows in excel
for row in ws.iter_rows(row_num=2,values_only=True,max_row=ws.max_row, min_col=1, max_col=3):
    student_name = row[0]
    mailid = row[1]
    missed_days = row[2]

       

