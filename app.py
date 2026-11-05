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
required_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
    missing_files = []
    directory = single_input('Please enter the directory to your files:')
    if not directory.endswith("\\"):
        directory += "\\"
    for file in required_files:
        if not os.path.isfile(directory + file):
            missing_files.append(file)
    if len(missing_files) == 0:
        break
    else:
        show_message('Error', ' Some required files are missing please try again.')



# INSERT ANSWER FOR QUESTION 2 HERE
df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for index, row in df1.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df2 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices = {}
for index, row in df2.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']



#Get email
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read().strip()



# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):

    template = openpyxl.load_workbook(directory + filename, data_only=True)
    sheet = template['OrderForm']
    name = sheet['C14'].value
    email = sheet['C13'].value
    date = sheet['C9'].value.strftime('%m/%d/%Y')
    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = str(name)
    order_data['CustomerCompany'] = str(sheet['C15'].value)
    order_data['CustomerAddress1'] = str(sheet['C16'].value)
    order_data['CustomerAddress2'] = str(sheet['C17'].value) + ', ' + str(sheet['C18'].value) + ' ' + str(sheet['C19'].value)

    total = 0
    row = 24
    item = 0

    while sheet['B' +  str(row)].value is not None:
        unit = sheet['B'+str(row)].value
        quantity = sheet['C' + str(row)].value
        description = product_descriptions[unit]
        price = product_prices[unit]
        amount = price * quantity
        order_data['Description' + str(item)] = str(description)
        order_data['Price' + str(item)] = f'${price:,.2f}'
        order_data['Quantity' + str(item)] = str(quantity)
        order_data['Amount' + str(item)] = f'${amount:,.2f}'
        total += amount
        row += 1
        item += 1

    order_data['Total'] = f'${total:,.2f}'
    template.close()
    return order_data, name, email, date



# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    template = directory + 'MasterInvoice.pdf'
    filename = directory + 'Invoice '+ str(invoice_num) + '.pdf'
    create_pdf(template, filename, order_data)



# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    gmail_user = 'benkharris04@gmail.com'
    recipients = [email]
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
    path_to_file = directory + 'Invoice ' + str(invoice_num) + '.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()



# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num += 1
