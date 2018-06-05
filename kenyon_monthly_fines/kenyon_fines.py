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
report_filename = 'KenyonStudentFines-' + str(fulldate.year) + str(fulldate.month) + str(fulldate.day)  + '.xlsx'
workbook = xlsxwriter.Workbook(report_filename)

# Create a new worksheet within the the new workbook
worksheet = workbook.add_worksheet()

# Format cells
formatLabel = workbook.add_format({'bold' : True, 'font_name' : 'Arial', 'font_size' : 12, 'border' : 2, 'valign' : 'vcenter'})
formatHeader = workbook.add_format({'font_name' : 'Arial', 'font_size' : 12})
formatNumbers = workbook.add_format({'num_format': '0.00', 'font_name' : 'Arial', 'font_size' : 12, 'border' : 1})
formatTable = workbook.add_format({'font_name' : 'Arial', 'font_size' : 12, 'border' : 1})

# Set column widths and row heights
worksheet.set_column(0,0,15.00)
worksheet.set_column(1,1,40.00)
worksheet.set_column(2,2,15.00)


# Insert static header information
worksheet.merge_range('A1:B1', 'KENYON COLLEGE', formatHeader)
worksheet.merge_range('A2:C2', 'AUTHORIZATION OF CHARGES/CREDITS TO STUDENT ACCOUNTS', formatHeader)
worksheet.write(3,0, 'DATE:', formatHeader)
worksheet.write(3,1, month_name + ' ' + year + ' Library Fines', formatHeader)
worksheet.write(4,0, 'DEPARTMENT:', formatHeader)
worksheet.write(4,1, 'Olin & Chalmers Libraries', formatHeader)
worksheet.write(5,0, 'ACCOUNT:', formatHeader)
worksheet.write(5,1, '<ACCOUNT_NUMBER>', formatHeader)
worksheet.merge_range('A8:C8', 'DESCRIPTION OF CHARGES: CHARGE TO STUDENT ACCOUNTS', formatHeader)
worksheet.merge_range('A9:B9', 'APPROVAL: Joan Nielson, Manager of Access Services', formatHeader)

for i in range(11):
	worksheet.set_row(i, 20.00)

# Assign column labels
worksheet.write(10,0,'Student ID', formatLabel)
worksheet.write(10,1,'Name', formatLabel)
worksheet.write(10,2,'Amount', formatLabel)


# Write data to the worksheet
for rownum, row in enumerate(rows):
	worksheet.set_row(rownum+11, 20.00)
	worksheet.write(rownum+11,0,row[0], formatTable)
	worksheet.write(rownum+11,1,row[1], formatTable)
	worksheet.write(rownum+11,2,row[2], formatNumbers)

# Close the workbook
workbook.close()

# Configure email sender information
emailhost = '<HOSTNAME>'
emailuser = '<FROM_ADDRESS>'
emailpass = '<EMAIL_PASSWORD>'
emailport = '<EMAIL_PORT>'
emailsubject = 'Monthly Kenyon Student Fines Report'
emailmessage = 'Hello,\n\nI\'ve attached ' + month_name + ' ' + year + ' Student Library Fines Report. This report contains a summary of total fines owed by students. If you have questions about the report, please let me know.\n\nThanks!\n\nBen'
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