#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun  1 11:10:31 2018

@author: bendaigle
"""
import xlsxwriter
import psycopg2
import datetime
import calendar
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

# Connect to the Sierra database
try:
	conn = psycopg2.connect("dbname='iii' user='<USERNAME>' host='<HOSTNAME>' port='1032' password='<PASSWORD>' sslmode='require'")
	print ("Connected to the database successfully!")
except psycopg2.Error as e:
	print ("Unable to connect to database: " + str(e))
    
# Open a database session
cursor = conn.cursor()

# Execute a database query. Database query is stored locally in an external sql file alongside this script
cursor.execute(open("<PATH_TO_SQL_FILE>","r").read())

# Store the results of the query in a variable called 'rows'
rows = cursor.fetchall()
print (rows[1])

# Get the current month. We list the current month in the report header
month_number = datetime.datetime.now().month
month_name = calendar.month_name[month_number]
year = str(datetime.datetime.now().year)
fulldate = datetime.datetime.now()

# Close the database connection
conn.close()

# Create an Excel workbook and assign that to a variable called 'workbook'
report_filename = 'KenyonInternalUseCounts-' + str(fulldate.year) + str(fulldate.month) + str(fulldate.day)  + '.xlsx'
workbook = xlsxwriter.Workbook(report_filename)
worksheet = workbook.add_worksheet()

# Format cells
formatLabel = workbook.add_format({'bold' : True, 'font_name' : 'Arial', 'font_size' : 12, 'valign' : 'vcenter'})
formatHeader = workbook.add_format({'font_name' : 'Arial', 'font_size' : 12})
formatNumbers = workbook.add_format({'num_format': '0.00', 'font_name' : 'Arial', 'font_size' : 12, 'border' : 1})
formatTable = workbook.add_format({'font_name' : 'Arial', 'font_size' : 12})
formatDates = workbook.add_format({'num_format' : 'mm/dd/yy hh:mm AM/PM', 'font_name' : 'Arial', 'font_size' : 12})

# Set column widths and row heights
worksheet.set_column(0,0,25.00)
worksheet.set_column(1,1,80.00)
worksheet.set_column(2,2,40.00)
worksheet.set_column(3,3,25.00)
worksheet.set_column(4,4,25.00)
worksheet.set_column(5,5,25.00)
worksheet.set_column(6,6,25.00)

worksheet.set_row(0, 25.00)

# Assign column labels
worksheet.write(0,0,'Item Barcode', formatLabel)
worksheet.write(0,1,'Title', formatLabel)
worksheet.write(0,2,'Author', formatLabel)
worksheet.write(0,3,'Call Number', formatLabel)
worksheet.write(0,4,'Location', formatLabel)
worksheet.write(0,5,'Volume', formatLabel)
worksheet.write(0,6,'Date Counted', formatLabel)

# Write data to the worksheet
for rownum, row in enumerate(rows):
	worksheet.set_row(rownum+1, 25.00)
	worksheet.write(rownum+1,0,row[0], formatTable)
	worksheet.write(rownum+1,1,row[1], formatTable)
	worksheet.write(rownum+1,2,row[2], formatTable)
	worksheet.write(rownum+1,3,row[3], formatTable)
	worksheet.write(rownum+1,4,row[4], formatTable)
	worksheet.write(rownum+1,5,row[5], formatTable)
	worksheet.write(rownum+1,6,row[6], formatDates)

workbook.close()

# Configure email sender information
emailhost = '<HOSTNAME>'
emailuser = '<FROM_ADDRESS>'
emailpass = '<EMAIL_PASSWORD>'
emailport = '<EMAIL_PORT>'
emailsubject = 'Monthly Kenyon Internal Use Count Report'
emailmessage = 'Hello,\n\nI\'ve attached the monthly Kenyon Internal Use Count Report. This report contains title-level information about items tracked as internal uses (used but not checked out). If you have questions about the report, please let me know.\n\nThanks!\n\nBen'
emailfrom = '<EMAIL_SENDER_NAME>'
emailto = ['EMAIL_RECIPIENT_ADDRESS_1','EMAIL_RECIPIENT_ADDRESS_2','EMAIL_RECIPIENT_ADDRESS_3']

msg = MIMEMultipart()
msg['From'] = emailfrom
if type(emailto) is list:
    msg['To'] = ', '.join(emailto)
else:
    msg['To'] = emailto
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = emailsubject
msg.attach (MIMEText(emailmessage))
part = MIMEBase('application', "octet-stream")
part.set_payload(open(report_filename,"rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition','attachment; filename=%s' % report_filename)
msg.attach(part)

#Send the email
smtp = smtplib.SMTP(emailhost, emailport)
#for Google connection
smtp.ehlo()
smtp.starttls()
smtp.login(emailuser, emailpass) 
#end for Google connection
smtp.sendmail(emailfrom, emailto, msg.as_string())
smtp.quit()