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
status = no_of_days= []
# list of students to remind
l1 =[]
# warning messages
m1 = "Warning!!! you can only miss more day for CI class"
m2 = " Warning !!! you can only miss one more day for python class"
m3 = "Warning!!! you can only miss one more day for DM class"
class Student:
    def __init__(self, name, email, status):
        self.name= name
        self.email= email
        self.status = status 
       # Marks students late 
    def mark_late(self):
        self.status = 'Late'
    
    # saves excel on every update
    def savefile():
        book.save(r'c:\Users\jorda\OneDrive\Documents\attendance.xlsx')
        print("saved!")

# To send email
    def send_mail(to_address, subject, body):
        msg = MIMEMultipart()
        msg['From'] = staff_mail
        msg['To'] = to_address
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
    
    # for students if missed days
    def notify_late(self, name, student, email,):
        subject = "Attendance warning"
        body = (f"Dear {name},\n\n"
                f"You have missed {no_of_days} days of class. " 
                "Please be aware that you are approaching the max days allowed to miss."
                "Best Regards, \n"
                "Attendance office")
        self.send_email(subject, body)
        print(f"Email sent to {student} at {email}")


def Process_attendance(sheet):
    students= []
# iterates over the rows in excel
    for row in ws.iter_rows(row_num=2,values_only=True,max_row=ws.max_row, min_col=1, max_col=3):
        student_name = row[0]
        mailid = row[1]
        no_of_days= row[2]

       

