# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows
# INSERT ANSWER FOR QUESTION 1 HERE


while True:
    try:
        directory = single_input('Input Directory Location')+"\\"
        filelist = os.listdir(directory)
        if ('MasterInvoice.pdf' in filelist)&('Order1.xlsx' in filelist)&('Order2.xlsx' in filelist)&('Order3.xlsx' in filelist)&('Order4.xlsx' in filelist)&('ProductData.xlsx' in filelist):
            break
    except:
        show_message('Error','Directory does not contain all files')
filelist
# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}
descriptions = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')
for r, row in enumerate(descriptions['Unit number']):
    product_descriptions[row] = descriptions['Description'][r]
prices = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')
for r, row in enumerate(prices['Unit number']):
    product_prices[row] = float(prices['Price per unit'][r])
product_prices
# INSERT ANSWER FOR QUESTION 3 HERE
df = get_pdf(directory+'MasterInvoice.pdf')
df2 = pd.read_excel('Order2.xlsx')
def get_order_data(invoice_num,filename,directory):
    file = pd.read_excel(directory+filename)
    PDFdictionary = {}
    PDFdictionary['InvoiceNum'] = invoice_num
    for r, row in enumerate(file['Unnamed: 1']):
        if row == 'Order Date':
            date = file['Unnamed: 2'][r].date()
            day = date.day
            month = date.month
            year = date.year
            PDFdictionary['Date'] = f'{month}/{day}/{year}'
            Dateout = f'{month}/{day}/{year}'
        elif row == 'Email':
            emailaddress = file['Unnamed: 2'][r]
        elif row == 'Name':
            Custname = file['Unnamed: 2'][r]
            PDFdictionary['CustomerName'] = file['Unnamed: 2'][r]
        elif row == 'Company Name':
            PDFdictionary['CustomerCompany'] = file['Unnamed: 2'][r]
        elif row == 'Address':
            PDFdictionary['CustomerAddress1'] = f'{file['Unnamed: 2'][14]} {file['Unnamed: 2'][15]} {file['Unnamed: 2'][16]} {file['Unnamed: 2'][17]}'
        elif row == 'Unit Number':
            PDFdictionary['Description0'] = file['Unnamed: 1'][22]
            PDFdictionary['Quantity0'] = file['Unnamed: 2'][22]
            PDFdictionary['Price0'] = product_prices[file['Unnamed: 1'][22]]
            num0 = PDFdictionary['Quantity0']*PDFdictionary['Price0']
            total = num0
            numstr = f'${num0:,.2f}'
            PDFdictionary['Amount0'] = numstr
            num1 = 0
            num2 = 0
            num3 = 0
            num4 = 0
            PDFdictionary['Price0'] = f'${PDFdictionary['Price0']:,.2f}'
            try:
                PDFdictionary['Description1'] = file['Unnamed: 1'][23]
                PDFdictionary['Quantity1'] = file['Unnamed: 2'][23]
                PDFdictionary['Price1'] = product_prices[file['Unnamed: 1'][23]]
                num1 = PDFdictionary['Quantity1']*PDFdictionary['Price1']
                total = num0+num1+num2+num3+num4
                numstr = f'${num1:,.2f}'
                PDFdictionary['Amount1'] = numstr
                PDFdictionary['Price1'] = f'${PDFdictionary['Price1']:,.2f}'
                PDFdictionary['Total'] = total
                PDFdictionary['Description2'] = file['Unnamed: 1'][24]
                PDFdictionary['Quantity2'] = file['Unnamed: 2'][24]
                PDFdictionary['Price2'] = product_prices[file['Unnamed: 1'][24]]
                num2 = PDFdictionary['Quantity2']*PDFdictionary['Price2']
                total = num0+num1+num2+num3+num4
                numstr = f'${num2:,.2f}'
                PDFdictionary['Amount2'] = numstr
                PDFdictionary['Price2'] = f'${PDFdictionary['Price2']:,.2f}'
                PDFdictionary['Total'] = total
                PDFdictionary['Description3'] = file['Unnamed: 1'][25]
                PDFdictionary['Quantity3'] = file['Unnamed: 2'][25]
                PDFdictionary['Price3'] = f'{product_prices[file['Unnamed: 1'][25]]:,.2f}'
                num3 = PDFdictionary['Quantity3']*PDFdictionary['Price3']
                total = num0+num1+num2+num3+num4
                numstr = f'${num3:,.2f}'
                PDFdictionary['Amount3'] = numstr
                PDFdictionary['Price3'] = f'${PDFdictionary['Price3']:,.2f}'
                PDFdictionary['Total'] = total
                PDFdictionary['Description4'] = file['Unnamed: 1'][26]
                PDFdictionary['Quantity4'] = file['Unnamed: 2'][26]
                PDFdictionary['Price4'] = f'{product_prices[file['Unnamed: 1'][26]]:,.2f}'
                num4 = PDFdictionary['Quantity4']*PDFdictionary['Price4']
                total = num0+num1+num2+num3+num4
                PDFdictionary['Total'] = total
                numstr = f'${num4:,.2f}'
                PDFdictionary['Amount4'] = numstr
                PDFdictionary['Price4'] = f'${PDFdictionary['Price4']:,.2f}'
            except:
                num = total
                numstr = f'${num:,.2f}'
                PDFdictionary['Total'] = numstr
                break          
    return PDFdictionary,Custname,emailaddress,Dateout
PDFdictionary, Custname, emailaddress, Dateout = get_order_data(8517,'Order3.xlsx',directory)
print(PDFdictionary)
print(emailaddress)
print(Custname)
print(Dateout)
print(PDFdictionary['Amount2'])
# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num,order_data,directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)
    return
create_invoice(8515,PDFdictionary,directory)
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num,name,email,date,directory):
    gmail_user = f'{email}'
    gmail_password = open('gmail_password.txt').read()
    recipients = email
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},

Attached is your invoice dated {date}.
    
Thank you for your business.
    
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()
    # INSERT ANSWER FOR QUESTION 6 HERE
filenames = filelist
filenum = 8515
for item in filenames:
    if ('Order' in item) and ('xlsx' in item):
        PDFdictionary, Custname, emailaddress, Dateout = get_order_data(filenum,item,directory)
        create_invoice(filenum,PDFdictionary,directory)
        send_email(filenum,Custname,emailaddress,Dateout,directory)
        filenum = filenum+1

filenames