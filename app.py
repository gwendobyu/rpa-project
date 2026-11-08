while True:    
    directory = single_input("Please input the directory where the ProductData, MasterInvoice, and Order files are.")
    directory = directory.strip('''"''')
    directory = f'{directory}\\'
    try:
        filenames = os.listdir(directory)
        if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
            break
        else:
            show_message("Error", "The directory you entered does not contain the files. Please try again.")
    except:
            show_message("Error", "The directory you entered does not exist. Please try again.")


descriptionsdf = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for i, row in descriptionsdf.iterrows():
    Description = row['Description']
    unitnumber = row['Unit number']
    product_descriptions[unitnumber] = Description
pricesdf = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
product_prices = {}
for i, row in pricesdf.iterrows():
    ppu = row['Price per unit']
    unitnumber = row['Unit number']
    product_prices[unitnumber] = ppu

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    orderfile = openpyxl.load_workbook(directory+filename, data_only=True)
    order = orderfile['OrderForm']
    order_data = {}
    InvoiceNum = str(invoice_num)
    unitnum0 = order['B24'].value
    Quantity0 = order['C24'].value
    PlainPrice0 = product_prices[unitnum0]
    Price0 = f'${PlainPrice0:,.2f}'
    Description0 = product_descriptions[unitnum0]
    PlainAmount0 = PlainPrice0 * Quantity0
    Amount0 = PlainAmount0
    if pd.isnull(order['B25'].value) == False:
        unitnum1 = order['B25'].value
        Quantity1 = order['C25'].value
        PlainPrice1 = product_prices[unitnum1]
        Price1 = f'${PlainPrice1:,.2f}'
        Description1 = product_descriptions[unitnum1]
        PlainAmount1 = PlainPrice1 * Quantity1
        Amount1 = PlainAmount1
        order_data['Price1'] = Price1
        order_data['Quantity1'] = str(Quantity1)
        order_data['Amount1'] = Amount1
        order_data['Description1'] = Description1
    else: PlainAmount1 = 0
    if pd.isnull(order['B26'].value) == False:
            unitnum2 = order['B26'].value
            Quantity2 = order['C26'].value
            PlainPrice2 = product_prices[unitnum2]
            Price2 = f'${PlainPrice2:,.2f}'
            Description2 = product_descriptions[unitnum2]
            PlainAmount2 = PlainPrice2 * Quantity2
            Amount2 = PlainAmount2
            order_data['Price2'] = Price2
            order_data['Quantity2'] = str(Quantity2)
            order_data['Amount2'] = Amount2
            order_data['Description2'] = Description2
    else: PlainAmount2 =0
    Total = PlainAmount0 + PlainAmount1 + PlainAmount2
    Customeraddress = order['C16'].value
    city = order['C17'].value
    state = order['C18'].value
    zipcode = order['C19'].value
    CustomerAddress2 = f'{city}, {state} {zipcode}'
    CustomerCompany = order['C15'].value
    custname = order['C14'].value
    custemail = order['C13'].value
    orderdate = str(order['C9'].value.strftime('%m/%d/%Y'))
    order_data['InvoiceNum'] = InvoiceNum
    order_data['Date'] = orderdate
    order_data['Total'] = Total
    order_data['CustomerAddress1'] = Customeraddress
    order_data['CustomerName'] = custname
    order_data['CustomerCompany'] = CustomerCompany
    order_data['CustomerAddress2'] = CustomerAddress2
    order_data['Price0'] = Price0
    order_data['Quantity0'] = str(Quantity0)
    order_data['Amount0'] = Amount0
    order_data['Description0'] = Description0
    return order_data, custname, custemail, orderdate

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', f'Invoice {invoice_num}.pdf', order_data)

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    gmail_user = "owca1000aj@gmail.com"
    recipients = email
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.
    
PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062 
Phone: (801) 555-2372
E-mail: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)


    path_to_data = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_data, message)    

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

filenames = os.listdir(directory)

for filename in filenames:
    invoice_num = 8515
    if 'Order' in filename and 'xlsx' in filename:
        orderdata = get_order_data(invoice_num, filename,directory)
        create_invoice(invoice_num, orderdata[0], directory)
        send_email(invoice_num, orderdata[1], orderdata[2], orderdata[3], directory)
        invoice_num = invoice_num + 1

