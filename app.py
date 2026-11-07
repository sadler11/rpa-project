# INSERT ANSWER FOR QUESTION 1 HERE


while True:
   directory = single_input("Please enter the file path to your folder.")
   directory = directory.replace('"','')
   if directory.endswith('/') != True:
       directory = directory +'/'
   try:
       filenames = os.listdir(directory)
       if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
           break
       else:
           show_message("Error", "The file path you entered is incorrect. Please try again.")
   except:
       show_message("Error", "The file path you entered is incorrect. Please try again.")
print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

ProductData = 'ProductData.xlsx'
full_path = os.path.join(directory, ProductData)

df = pd.read_excel(full_path, sheet_name='Descriptions')
for index, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

print(product_descriptions)

dfp = pd.read_excel(full_path, sheet_name = 'Prices')
for index, row in dfp.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE
import datetime

order_data = {}

def get_order_data(invoice_num, filename, directory):
    order_dict = {}
    template = openpyxl.load_workbook(filename)
    selected_sheet = template['OrderForm']
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    Unit1 = selected_sheet['B24'].value
    Quantity1 = selected_sheet['C24'].value
    Unit2 = selected_sheet['B25'].value
    Quantity2 = selected_sheet['C25'].value
    Unit3 = selected_sheet['B26'].value
    Quantity3 = selected_sheet['C26'].value
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = date
    if Unit3 is not None:
        order_data['Total'] = (product_prices.get(Unit1) * Quantity1) + ((product_prices.get(Unit2) * Quantity2)) + (product_prices.get(Unit3) * Quantity3)
        order_data['Total'] = f'${order_data['Total']:,.2f}'
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}'
        order_data['Price0'] = product_prices.get(Unit1)
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'
        order_data['Description0'] = product_descriptions.get(Unit1)
        order_data['Price1'] = product_prices.get(Unit2)
        order_data['Price1'] = f'${order_data['Price1']:,.2f}'
        order_data['Quantity1'] = Quantity2
        order_data['Amount1'] = product_prices.get(Unit2) * Quantity2
        order_data['Amount1'] = f'${order_data['Amount1']:,.2f}'
        order_data['Description1'] = product_descriptions.get(Unit2)
        order_data['Price2'] = product_prices.get(Unit3)
        order_data['Price2'] = f'${order_data['Price2']:,.2f}'
        order_data['Quantity2'] = Quantity3
        order_data['Amount2'] = product_prices.get(Unit3) * Quantity3
        order_data['Amount2'] = f'${order_data['Amount2']:,.2f}'
        order_data['Description2'] = product_descriptions.get(Unit2)
    elif Unit2 is not None:
        order_data['Total'] = (product_prices.get(Unit1) * Quantity1) + ((product_prices.get(Unit2) * Quantity2))
        order_data['Total'] = f'${order_data['Total']:,.2f}'
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}'
        order_data['Price0'] = product_prices.get(Unit1)
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'
        order_data['Description0'] = product_descriptions.get(Unit1)
        order_data['Price1'] = product_prices.get(Unit2)
        order_data['Price1'] = f'${order_data['Price1']:,.2f}'
        order_data['Quantity1'] = Quantity2
        order_data['Amount1'] = product_prices.get(Unit2) * Quantity2
        order_data['Amount1'] = f'${order_data['Amount1']:,.2f}'
        order_data['Description1'] = product_descriptions.get(Unit2)
    else:
        order_data['Total'] = (product_prices.get(Unit1) * Quantity1)
        order_data['Total'] = f'${order_data['Total']:,.2f}'
        order_data['CustomerAddress1'] = selected_sheet['C16'].value
        order_data['CustomerName'] = selected_sheet['C14'].value
        order_data['CustomerCompany'] = selected_sheet['C15'].value
        order_data['CustomerAddress2'] = f'{selected_sheet['C17'].value}, {selected_sheet['C18'].value}, {selected_sheet['C19'].value}'
        order_data['Price0'] = product_prices.get(Unit1)
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'
        order_data['Quantity0'] = Quantity1
        order_data['Amount0'] = product_prices.get(Unit1) * Quantity1
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'
        order_data['Description0'] = product_descriptions.get(Unit1)
    return order_data, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(get_order_data(8515,'Order1.xlsx',directory))

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf('MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    gmail_user = 'jonessadler@gmail.com'
    gmail_password = open('gmail_password.txt').read()

    recipients = email

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},
    
    Attached is the invoice dated {date}.

    Thank you for your business.
    
    Python Tech Co.
    10 Python Drive
    Pleasant Grove, UT 84062 
    Phone (801) 555-2372
    Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    filename = f'Invoice {invoice_num}.pdf'
    path_to_file = os.path.join(directory, filename)

    path_to_file = os.getcwd()+f'\\Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file, message)
        

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)


    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,'Thomas Holt','jonessadler@gmail.com','11/13/2023',directory)

# INSERT ANSWER FOR QUESTION 6 HERE
import os
filenames = os.listdir(directory)
print(filenames)

number = 8515

for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(number, file, directory)
        create_invoice(number, order_data, directory)
        send_email(number, name, 'jonessadler@gmail.com', date, directory)
        number = number + 1



