#CRUISE

import csv, openpyxl, os

print('請輸入機尾號 ex: B-#####')
tailNum = 'B-99888' #input()
print('請輸入資料夾名稱 ex: yyyymm')
folderName = '201806' #input()

wb = openpyxl.load_workbook('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\B-99888\\Engine_Trend_Cruise-Mmm 2018 B-99888.xlsx')
sheet = wb['Cruise Monitor data']

j = 2 #Because column B is 2
k = 1 #No. 1 is started
for fileName in os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + folderName):
    print('No.: ' + str(k) +' '+ fileName + '\n')
    k = k+1

    exampleFile = open('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\B-99888\\201806\\' + fileName)
    exampleReader = csv.reader(exampleFile)
    exampleData = list(exampleReader)

    '''
    for i in range(8, 508):
        .append(float(exampleData[i][]))
    '''

    #TGT
    ENG_1_TGT = []
    for i in range(8, 508):
        ENG_1_TGT.append(float(exampleData[i][22]))
    ENG_2_TGT = []
    for i in range(8, 508):
        ENG_2_TGT.append(float(exampleData[i][47]))

    #FF
    ENG_1_FF = []
    for i in range(8, 508):
        ENG_1_FF.append(float(exampleData[i][23]))
    ENG_2_FF = []
    for i in range(8, 508):
        ENG_2_FF.append(float(exampleData[i][48]))

    #N1
    ENG_1_N1 = []
    for i in range(8, 508):
        ENG_1_N1.append(float(exampleData[i][31]))
    ENG_2_N1 = []
    for i in range(8, 508):
        ENG_2_N1.append(float(exampleData[i][56]))

    #N2
    ENG_1_N2 = []
    for i in range(8, 508):
        ENG_1_N2.append(float(exampleData[i][33]))
    ENG_2_N2 = []
    for i in range(8, 508):
        ENG_2_N2.append(float(exampleData[i][58]))

    #P30
    ENG_1_P30 = []
    for i in range(8, 508):
        ENG_1_P30.append(float(exampleData[i][34]))
    ENG_2_P30 = []
    for i in range(8, 508):
        ENG_2_P30.append(float(exampleData[i][59]))

    #T30
    ENG_1_T30 = []
    for i in range(8, 508):
        ENG_1_T30.append(float(exampleData[i][21]))
    ENG_2_T30 = []
    for i in range(8, 508):
        ENG_2_T30.append(float(exampleData[i][46]))

    #LP_VIB
    E1_LP_VIB = []
    for i in range(8, 508):
        E1_LP_VIB.append(float(exampleData[i][11]))
    E2_LP_VIB = []
    for i in range(8, 508):
        E2_LP_VIB.append(float(exampleData[i][15]))

    #HP_VIB
    E1_HP_VIB = []
    for i in range(8, 508):
        E1_HP_VIB.append(float(exampleData[i][10]))
    E2_HP_VIB = []
    for i in range(8, 508):
        E2_HP_VIB.append(float(exampleData[i][14]))

    #OIL_P
    E1_OIL_P = []
    for i in range(8, 508):
        E1_OIL_P.append(float(exampleData[i][24]))
    E2_OIL_P = []
    for i in range(8, 508):
        E2_OIL_P.append(float(exampleData[i][49]))

    #OIL_T
    E1_OIL_T = []
    for i in range(8, 508):
        E1_OIL_T.append(float(exampleData[i][25]))
    E2_OIL_T = []
    for i in range(8, 508):
        E2_OIL_T.append(float(exampleData[i][50]))

    
    sheet.cell(row = 4, column = j).value = max(ENG_1_TGT)
    sheet.cell(row = 5, column = j).value = max(ENG_2_TGT)
    sheet.cell(row = 13, column = j).value = max(ENG_1_FF)
    sheet.cell(row = 14, column = j).value = max(ENG_2_FF)
    sheet.cell(row = 23, column = j).value = max(ENG_1_N1)
    sheet.cell(row = 24, column = j).value = max(ENG_2_N1)
    sheet.cell(row = 37, column = j).value = max(ENG_1_N2)
    sheet.cell(row = 38, column = j).value = max(ENG_2_N2)
    sheet.cell(row = 47, column = j).value = max(ENG_1_P30)
    sheet.cell(row = 48, column = j).value = max(ENG_2_P30)
    sheet.cell(row = 57, column = j).value = max(ENG_1_T30)
    sheet.cell(row = 58, column = j).value = max(ENG_2_T30)
    sheet.cell(row = 70, column = j).value = max(E1_LP_VIB)
    sheet.cell(row = 71, column = j).value = max(E2_LP_VIB)
    sheet.cell(row = 80, column = j).value = max(E1_HP_VIB)
    sheet.cell(row = 81, column = j).value = max(E2_HP_VIB)
    sheet.cell(row = 90, column = j).value = max(E1_OIL_P)
    sheet.cell(row = 91, column = j).value = max(E2_OIL_P)
    sheet.cell(row = 103, column = j).value = max(E1_OIL_T)
    sheet.cell(row = 104, column = j).value = max(E2_OIL_T)

    '''
    sheet['B4'] = max(ENG_1_TGT)
    sheet['B5'] = max(ENG_2_TGT)
    sheet['B13'] = max(ENG_1_FF)
    sheet['B14'] = max(ENG_2_FF)
    sheet['B23'] = max(ENG_1_N1)
    sheet['B24'] = max(ENG_2_N1)
    sheet['B37'] = max(ENG_1_N2)
    sheet['B38'] = max(ENG_2_N2)
    sheet['B47'] = max(ENG_1_P30)
    sheet['B48'] = max(ENG_2_P30)
    sheet['B57'] = max(ENG_1_T30)
    sheet['B58'] = max(ENG_2_T30)
    sheet['B70'] = max(E1_LP_VIB)
    sheet['B71'] = max(E2_LP_VIB)
    sheet['B80'] = max(E1_HP_VIB)
    sheet['B81'] = max(E2_HP_VIB)
    sheet['B90'] = max(E1_OIL_P)
    sheet['B91'] = max(E2_OIL_P)
    sheet['B103'] = max(E1_OIL_T)
    sheet['B104'] = max(E2_OIL_T)
    '''
    
    exampleFile.close()
    j = j + 1
    
    '''
    print(': ' + str(max()))
    '''
    
    print('ENG_1_TGT: ' + str(max(ENG_1_TGT)))
    print('ENG_2_TGT: ' + str(max(ENG_2_TGT)))
    print('ENG_1_FF: ' + str(max(ENG_1_FF)))
    print('ENG_2_FF: ' + str(max(ENG_2_FF)))
    print('ENG_1_N1: ' + str(max(ENG_1_N1)))
    print('ENG_2_N1: ' + str(max(ENG_2_N1)))
    print('ENG_1_N2: ' + str(max(ENG_1_N2)))
    print('ENG_2_N2: ' + str(max(ENG_2_N2)))
    print('ENG_1_P30: ' + str(max(ENG_1_P30)))
    print('ENG_2_P30: ' + str(max(ENG_2_P30)))
    print('ENG_1_T30: ' + str(max(ENG_1_T30)))
    print('ENG_2_T30: ' + str(max(ENG_2_T30)))
    print('E1_LP_VIB: ' + str(max(E1_LP_VIB)))
    print('E2_LP_VIB: ' + str(max(E2_LP_VIB)))
    print('E1_HP_VIB: ' + str(max(E1_HP_VIB)))
    print('E2_HP_VIB: ' + str(max(E2_HP_VIB)))
    print('E1_OIL_P: ' + str(max(E1_OIL_P)))
    print('E2_OIL_P: ' + str(max(E2_OIL_P)))
    print('E1_OIL_T: ' + str(max(E1_OIL_T)))
    print('E2_OIL_T: ' + str(max(E2_OIL_T)))
    print('\n')
    
    
wb.save('example_copy.xlsx')






