import os, csv, openpyxl
'''
for folderName, subfolders, filenames in os.walk('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\5 years trend'):
    #print('the current folder is ' + folderName)
    for subfolder in subfolders:
        print('subfolder of ' + folderName + ': ' + subfolder)
    for filename in filenames:
        print('file inside ' + folderName + ': ' + filename)
    print('')
'''
wb = openpyxl.load_workbook('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\5years.xlsx')
sheet = wb['sheet1']

j = 2 #Because column B is 2
tempList1 = []
tempList2 = []
tempList3 = []
tempList4 = []
tempList5 = []
tempList6 = []
tempList7 = []
allFileNames = []

for fileName in os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\5_years_trend'):
    #print(fileName)

    csvFile = open('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\5_years_trend\\' + fileName)
    csvReader = csv.reader(csvFile)
    csvData = list(csvReader)

    #TGT
    ENG_1_TGT = []
    for i in range(8, 408):
        ENG_1_TGT.append(float(csvData[i][22]))
    ENG_2_TGT = []
    for i in range(8, 407):
        ENG_2_TGT.append(float(csvData[i][47]))

    #LP_VIB
    E1_LP_VIB = []
    for i in range(8, 408):
        E1_LP_VIB.append(float(csvData[i][11]))
    E2_LP_VIB = []
    for i in range(8, 408):
        E2_LP_VIB.append(float(csvData[i][15]))

    #HP_VIB
    E1_HP_VIB = []
    for i in range(8, 408):
        E1_HP_VIB.append(float(csvData[i][10]))
    E2_HP_VIB = []
    for i in range(8, 408):
        E2_HP_VIB.append(float(csvData[i][14]))

    tempList7.append(int(fileName[21:27]))
    allFileNames.append(fileName)

    

    if tempList7[0] == int(fileName[21:27]):
        

        
        tempList1.append(max(ENG_1_TGT))
        tempList2.append(max(ENG_2_TGT))
        tempList3.append(max(E1_LP_VIB))
        tempList4.append(max(E2_LP_VIB))
        tempList5.append(max(E1_HP_VIB))
        tempList6.append(max(E2_HP_VIB))
        tempList7.append(int(fileName[21:27]))
        print(int(fileName[21:27]))
    else:
        print('nextMonth')
        
        
        sheet.cell(row = 4, column = j).value = max(tempList1)
        sheet.cell(row = 5, column = j).value = max(tempList2)
        sheet.cell(row = 6, column = j).value = max(tempList3)
        sheet.cell(row = 7, column = j).value = max(tempList4)
        sheet.cell(row = 8, column = j).value = max(tempList5)
        sheet.cell(row = 9, column = j).value = max(tempList6)
        sheet.cell(row = 3, column = j).value = tempList7[0]
        tempList1 = []
        tempList2 = []
        tempList3 = []
        tempList4 = []
        tempList5 = []
        tempList6 = []
        tempList7 = []
        
        
        tempList1.append(max(ENG_1_TGT))
        tempList2.append(max(ENG_2_TGT))
        tempList3.append(max(E1_LP_VIB))
        tempList4.append(max(E2_LP_VIB))
        tempList5.append(max(E1_HP_VIB))
        tempList6.append(max(E2_HP_VIB))
        tempList7.append(int(fileName[21:27]))
        print(int(fileName[21:27]))

        
        j = j + 1

              
    

    csvFile.close()
    
if allFileNames[len(allFileNames)-1] == fileName:
    sheet.cell(row = 4, column = j).value = max(tempList1)
    sheet.cell(row = 5, column = j).value = max(tempList2)
    sheet.cell(row = 6, column = j).value = max(tempList3)
    sheet.cell(row = 7, column = j).value = max(tempList4)
    sheet.cell(row = 8, column = j).value = max(tempList5)
    sheet.cell(row = 9, column = j).value = max(tempList6)
    sheet.cell(row = 3, column = j).value = tempList7[0]
    


    
wb.save('5566' + '.xlsx')



'''
    sheet.cell(row = 3, column = j).value = int(fileName[21:27])
    sheet.cell(row = 4, column = j).value = max(ENG_1_TGT)
    sheet.cell(row = 5, column = j).value = max(ENG_2_TGT)
    sheet.cell(row = 6, column = j).value = max(E1_LP_VIB)
    sheet.cell(row = 7, column = j).value = max(E2_LP_VIB)
    sheet.cell(row = 8, column = j).value = max(E1_HP_VIB)
    sheet.cell(row = 9, column = j).value = max(E2_HP_VIB)
'''
