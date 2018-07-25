import os, openpyxl, sys, smtplib

os.chdir('C:\\Users\\peter\\OneDrive\\工作夾')

wb = openpyxl.load_workbook('duesRecords.xlsx')
sheet = wb['Sheet1']
lastCol = sheet.max_column
latestMonth = sheet.cell(row = 1, column = lastCol).value


unpaidMembers = {}
for r in range(2, sheet.max_row + 1):
    payment = sheet.cell(row = r, column = lastCol).value
    if payment != 'login':
        name = sheet.cell(row = r, column = 1).value
        email = sheet.cell(row = r, column = 2).value
        unpaidMembers[name] = email
    

smtpObj = smtplib.SMTP('smtp-mail.outlook.com', 587)
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.login('petertsao@winairjet.com', 'PASSWORD')#PASSWORD



for name, email in unpaidMembers.items():
    body = "Subject: %s Aircraft Log Book Entry Inspection Maintenance Record Update Notice.\n\nDear %s,\nPlease update Aircraft Log Book Entry Inspection Maintenance Record for %s and report to QA Manager Eric (ericyo@winairjet.com) within 2 days as soon as possible. Thank you!\nBR" %(latestMonth, name, latestMonth)
    print('Sending email to %s...' % email)
    sendmailStatus = smtpObj.sendmail('petertsao@winairjet.com', email, body)

    if sendmailStatus != {}:
        print('There was a problem sending email to %s: %s' % (email, sendmailStatus))
smtpObj.quit()    

#Please Note: This message is automatically sent, if you have any problem please reply directly to petertsao@winairjet.com\n
