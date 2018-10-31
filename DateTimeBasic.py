#DateTimeBasic

import datetime, time, csv, os, openpyxl

print('請輸入機尾號 ex: B-##### (***Cal Flight Time***)')
tailNum = input()
print('請輸入資料夾名稱 ex: yyyymm')
folderName = input()

wb = openpyxl.Workbook()

sheet = wb['Sheet']
i = 2
#take off time
for fileName in os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T'):#T means TakeOff

    sheet = wb['Sheet']

    takeOffCsv = open('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T' + '\\' + fileName)
    takeOffReader = csv.reader(takeOffCsv)
    takeOffData = list(takeOffReader)
    
    takeOffTriggerTimeStr = takeOffData[3][0]#get trigger time from csv 
    takeOffTimeStr = takeOffTriggerTimeStr[12:29]#cut str
    takeOffTime = datetime.datetime.strptime(takeOffTimeStr, '%d %b %Y %H:%M')

    #write in a new excel file
    sheet.cell(row = i, column = 1).value = takeOffTime

    takeOffCsv.close()
    i = i + 1
    print(takeOffTime)

print('---')
i = 2
#reverse time
for fileName in os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'R'):#R means Reverse

    sheet = wb['Sheet']

    reverseCsv = open('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'R' + '\\' + fileName)
    reverseReader = csv.reader(reverseCsv)
    reverseData = list(reverseReader)
    
    reverseReportTimeStr = reverseData[4][0]#get trigger time from csv 
    reverseTimeStr = reverseReportTimeStr[18:35]#cut str
    reverseTime = datetime.datetime.strptime(reverseTimeStr, '%d %b %Y %H:%M')

    #write in a new excel file
    sheet.cell(row = i, column = 2).value = reverseTime

    sheet.cell(row = i, column = 3).value = '=B' + str(i) + '-A' + str(i)
    
    reverseCsv.close()
    i = i + 1
    print(reverseTime)

sheet.cell(row = i, column = 3).value = '=SUM(C2:C' + str(i-1) +')'

sheet.column_dimensions['A'].width = 20
sheet.column_dimensions['B'].width = 20
sheet.column_dimensions['C'].width = 20
sheet.cell(row = 1, column = 1).value = 'Take Off Time'
sheet.cell(row = 1, column = 2).value = 'Reverse Time'
sheet.cell(row = 1, column = 3).value = 'Flight Time'


Mmm = folderName[4:]
YYYY = time.strftime('%Y')

wb.save('FlightTime-' + Mmm + ' ' + YYYY + ' ' + tailNum + '.xlsx')    
'''





#開excel 擷取
takeOffTriggerTimeStr = 'TriggerTime 30 OCT 2018 23:58:00'
takeOffTimeStr = takeOffTriggerTimeStr[12:29]
takeOffTime = datetime.datetime.strptime(takeOffTimeStr, '%d %b %Y %H:%M')


reverseTriggerTimeStr = 'TriggerTime 31 OCT 2018 00:08:00'
reverseTimeStr = reverseTriggerTimeStr[12:29]
reverseTime = datetime.datetime.strptime(reverseTimeStr, '%d %b %Y %H:%M')

flightTime = reverseTime - takeOffTime


#print(ReverseTime - TakeOffTime)

#print(earlyTime - datetime.datetime.now())
'''
