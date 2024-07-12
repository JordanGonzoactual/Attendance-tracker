import openpyxl
import smtplib
import email
# Workbook
book = openpyxl.load_workbook('D:\\attendance.xlsx')
# staff emails
staff_mails=['targyarg13@gmail.com', 'largmarg814@gmail.com']
# days limit
days_limit= 12
# Max amount of missed days
attendance_threshold= 13




# saves excel on every update
def savefile():
    book.save(r'c:\Users\jorda\OneDrive\Documents\attendance.xlsx')
    print("saved!")

# To track attendance
def check(no_of_days, row_num, b):
    #to use the globally declared lists and strings
    global staff_mails
    global days_limit
    global attendance_threshold