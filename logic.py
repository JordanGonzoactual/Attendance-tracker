import openpyxl
import smtplib
import email

# Workbook
wb = Workbook()
# grabs active worksheet
ws = wb.active

def savefile():
    book.save(r'<your-path>\attendance.xlsx')
    print("saved!")