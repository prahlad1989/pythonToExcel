# Importing libraries 
import imaplib
import datetime
import xlwt
from xlwt import Workbook


import smtplib

import email

user = 'emailaddress'
password = 'password'
imap_url = 'imap.gmail.com'


# -PO number: EWF-2605
# -Name: Jean William
# -Street: 183 Shore Dr
# -City: Salem
# -State: NH
# -Zip code: 03079
# -Phone: ###-###-####
# -SKU: whole line.
# -Qty: 1
# -Net Price: 100.50
# -Net Total Price: 100.05


# Function to get email content part i.e its body part
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)

    # Function to search for a key value pair


def search(key, value, con):
    # result, data = con.search(None, key, value)
    date = (datetime.date.today() - datetime.timedelta(1)).strftime("%d-%b-%Y")
    result, data = con.search(None, ('SEEN'), '(SENTSINCE {0})'.format(date))
    return data


# Function to get the list of emails under this label 
def get_emails(result_bytes):
    msgs = []  # all the email data are pushed inside an array
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)

    return msgs


# this is done to make SSL connnection with GMAIL 
con = imaplib.IMAP4_SSL(imap_url)

# logging the user in 
con.login(user, password)

# calling function to check for email under this label 
con.select('Inbox')

# fetching emails from this user "tu**h*****1@gmail.com" 
msgs = get_emails(search('DATE', datetime.date.today(), con))

# Uncomment this to see what actually comes as data 
# print(msgs) 


# Finding the required content from our msgs 
# User can make custom changes in this part to 
# fetch the required content he / she needs 

# printing them by the order they are displayed in your gmail

# google sheets initializtion
allMails = list()
for msg in msgs[::-1]:
    for sent in msg:
        if type(sent) is tuple:
            content = str(sent[1], "UTF-8")
            message1 = email.message_from_bytes(sent[1])
            try:

                for part in message1.walk():
                    if part.get_content_type() == 'text/plain' and "Purchase Order No#" in part.get_payload() and "Shipping Address:" in part.get_payload():
                        print(str(message1['date']))
                        print(str(message1['from']))
                        print(str(message1['subject']))
                        print(part.get_payload())
                        print()
                        print()
                        emailContent = part.get_payload()
                        emailContent = emailContent.replace("=", "")
                        tokens = emailContent.split("\r\n")
                        count = 0
                        orderInfo = dict()
                        allMails.append(orderInfo)
                        while count < len(tokens):
                        # print(tokens[count])
                            if "Purchase Order No# " in tokens[count]:
                                orderInfo['purchaseOrder'] = tokens[count].replace("Purchase Order No#", "").strip()
                            elif "Shipping Address:" in tokens[count]:
                                orderInfo['name'] = tokens[count + 1]
                                orderInfo['street'] = tokens[count + 2]
                                cityEtc = tokens[count + 3]
                                cityTokens = cityEtc.split(",")

                                orderInfo['city'] = cityTokens[0]
                                orderInfo['state'] = cityTokens[1].strip()
                                orderInfo['pin'] = cityTokens[2].split("-")[0].replace("\r\n", "").strip()
                                orderInfo['phone'] = tokens[count + 4]
                            elif "SKU:" in tokens[count]:
                                orderInfo['sku'] = tokens[count].replace("SKU:", "").strip()
                            elif "Qty:" in tokens[count]:
                                orderInfo['qty'] = tokens[count].replace("Qty:", "").strip()
                            elif "Gross Price:" in tokens[count]:
                                orderInfo['Gross Price:'] = tokens[count].replace("Gross Price:", "").strip()
                            elif "Net Price:" in tokens[count]:
                                orderInfo['Net Price:'] = tokens[count].replace("Net Price:", "").strip()
                            elif "Net Total Price:" in tokens[count]:
                                orderInfo['Net Total Price:'] = tokens[count].replace("Net Total Price:", "").strip()
                            elif "Grand Total:" in tokens[count]:
                                orderInfo['Grand Total:'] = tokens[count].replace("Grand Total:", "").strip()

                            count = count + 1


            except UnicodeEncodeError as e:
                pass

wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('OrdersInfo')
# title
style = xlwt.easyxf('font: bold 1, color red;')
# -PO number: EWF-2605
# -Name: Jean William
# -Street: 183 Shore Dr
# -City: Salem
# -State: NH
# -Zip code: 03079
# -Phone: ###-###-####
# -SKU: whole line.
# -Qty: 1
# -Net Price: 100.50
# -Net Total Price: 100.05
sheet1.write(0, 0, 'PO number', style)
sheet1.write(0, 1, 'Name', style)
sheet1.write(0, 2, 'Street', style)
sheet1.write(0, 3, 'City', style)
sheet1.write(0,4,'State',style)
sheet1.write(0, 5, 'Zip code', style)
sheet1.write(0, 6, 'Phone', style)
sheet1.write(0, 7, 'SKU', style)
sheet1.write(0, 8, 'Qty', style)
sheet1.write(0, 9,'Gross Price', style)
sheet1.write(0, 10, 'Net Price', style)
sheet1.write(0, 11, 'Net Total Price', style)
sheet1.write(0, 12, 'Grand Total', style)

for i in range(0,len(allMails)):
    orderInfo = allMails[i]
    sheet1.write(i+1, 0, orderInfo['purchaseOrder'])
    sheet1.write(i+1, 1, orderInfo['name'])
    sheet1.write(i+1, 2, orderInfo['street'])
    sheet1.write(i+1, 3, orderInfo['city'])
    sheet1.write(i+1,4,orderInfo['state'])
    sheet1.write(i+1, 5, orderInfo['pin'])
    sheet1.write(i+1, 6, orderInfo['phone'])
    sheet1.write(i+1, 7, orderInfo['sku'])
    sheet1.write(i+1, 8, orderInfo['qty'])
    sheet1.write(i+1, 9, orderInfo['Gross Price:'])
    sheet1.write(i+1, 10,  orderInfo['Net Price:'])
    sheet1.write(i+1, 11, orderInfo['Net Total Price:'])
    sheet1.write(i+1, 12, orderInfo['Grand Total:'])


wb.save('orders_info.xls')

