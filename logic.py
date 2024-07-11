import openpyxl
import smtplib
import email
# Workbook
book = openpyxl.load_workbook('D:\\attendance.xlsx')

# saves excel on every update
def savefile():
    book.save(r'c:\Users\jorda\OneDrive\Documents\attendance.xlsx')
    print("saved!")