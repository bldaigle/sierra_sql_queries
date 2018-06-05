#!/usr/bin/env python3

# Import psycopg2 module to enable connections to the posgresql database.
# Import xlsxwriter to create and format an Excel document for our report.
import psycopg2
import xlsxwriter
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
report_filename = 'KenyonStudentFineDetails-' + str(fulldate.year) + str(fulldate.month) + str(fulldate.day)  + '.xlsx'
workbook = xlsxwriter.Workbook(report_filename)

# Create a new worksheet within the the new workbook
worksheet = workbook.add_worksheet()

# Format cells
formatLabel = workbook.add_format({'bold' : True, 'font_name' : 'Arial', 'font_size' : 12, 'valign' : 'vcenter'})
formatHeader = workbook.add_format({'font_name' : 'Arial', 'font_size' : 12})
formatCurrency = workbook.add_format({'num_format' : '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'font_name' : 'Arial', 'font_size' : 12})
formatTable = workbook.add_format({'font_name' : 'Arial', 'font_size' : 12})
formatDates = workbook.add_format({'num_format' : 'mm/dd/yy hh:mm AM/PM', 'font_name' : 'Arial', 'font_size' : 12})

# Set column widths and row heights
worksheet.set_column(0,0,18.00)
worksheet.set_column(1,1,30.00)
worksheet.set_column(2,2,18.00)
worksheet.set_column(3,3,18.00)
worksheet.set_column(4,4,45.00)
worksheet.set_column(5,5,30.00)
worksheet.set_column(6,6,30.00)
worksheet.set_column(7,7,30.00)
worksheet.set_column(8,8,18.00)

# Assign column labels
worksheet.write(0,0,'Student ID', formatLabel)
worksheet.write(0,1,'Name', formatLabel)
worksheet.write(0,2,'Invoice Number', formatLabel)
worksheet.write(0,3,'Charge Type', formatLabel)
worksheet.write(0,4,'Description', formatLabel)
worksheet.write(0,5,'Checkout Date', formatLabel)
worksheet.write(0,6,'Due Date', formatLabel)
worksheet.write(0,7,'Returned Date', formatLabel)
worksheet.write(0,8,'Amount Due', formatLabel)

# Write data to the worksheet
for rownum, row in enumerate(rows):
	worksheet.set_row(rownum+1, 25.00)
	worksheet.write(rownum+1,0,row[0], formatTable)
	worksheet.write(rownum+1,1,row[1], formatTable)
	worksheet.write(rownum+1,2,row[2], formatTable)
	worksheet.write(rownum+1,3,row[3], formatTable)
	worksheet.write(rownum+1,4,row[4], formatTable)
	worksheet.write(rownum+1,5,row[5], formatDates)
	worksheet.write(rownum+1,6,row[6], formatDates)
	worksheet.write(rownum+1,7,row[7], formatDates)
	worksheet.write(rownum+1,8,row[8], formatCurrency)

# Close the workbook
workbook.close()

# Configure email sender information
emailhost = '<EMAIL_HOSTNAME>'
emailuser = '<FROM_ADDRESS>'
emailpass = '<EMAIL_PASSWORD>'
emailport = '<EMAIL_PORT>'
emailsubject = 'Monthly Kenyon Student Fine Details Report'
emailmessage = 'Hello,\n\nI\'ve attached ' + month_name + ' ' + year + ' Student Library Fine Details Report. This report contains detailed information about individual fines incurred by students. If you have questions about the report, please let me know.\n\nThanks!\n\nBen'
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
