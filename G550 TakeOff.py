#G550 TakeOff

import csv, openpyxl, os, time

print('請輸入機尾號 ex: B-##### (***適用G550 TakeOff***)')
tailNum = input()
print('請輸入資料夾名稱 ex: yyyymm')
folderName = input()

wb = openpyxl.load_workbook('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\B-99888\\Engine_Trend_TakeOff-Mmm 2018 B-99888.xlsx')
sheet = wb['Cruise Monitor data']

j = 2 #Because column B is 2
k = 1 #No. 1 is started for fileName's No.
m = 1 #Flist is started from 1
for fileName in os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T'):#T means TakeOff
    print('No.: ' + str(k) +' '+ fileName + '\n')
    k = k+1
    
    #切到第1個工作分頁
    sheet = wb['Cruise Monitor data']


    exampleFile = open('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T' + '\\' + fileName)
    exampleReader = csv.reader(exampleFile)
    exampleData = list(exampleReader)


    #TGT
    ENG_1_TGT = []
    for i in range(8, 408):
        ENG_1_TGT.append(float(exampleData[i][22]))
    ENG_2_TGT = []
    for i in range(8, 407):
        ENG_2_TGT.append(float(exampleData[i][47]))
   
    #N1
    ENG_1_N1 = []
    for i in range(8, 408):
        ENG_1_N1.append(float(exampleData[i][31]))
    ENG_2_N1 = []
    for i in range(8, 408):
        ENG_2_N1.append(float(exampleData[i][56]))

    #N2
    ENG_1_N2 = []
    for i in range(8, 408):
        ENG_1_N2.append(float(exampleData[i][33]))
    ENG_2_N2 = []
    for i in range(8, 408):
        ENG_2_N2.append(float(exampleData[i][58]))

    #LP_VIB
    E1_LP_VIB = []
    for i in range(8, 408):
        E1_LP_VIB.append(float(exampleData[i][11]))
    E2_LP_VIB = []
    for i in range(8, 408):
        E2_LP_VIB.append(float(exampleData[i][15]))

    #HP_VIB
    E1_HP_VIB = []
    for i in range(8, 408):
        E1_HP_VIB.append(float(exampleData[i][10]))
    E2_HP_VIB = []
    for i in range(8, 408):
        E2_HP_VIB.append(float(exampleData[i][14]))

    #OIL_P
    E1_OIL_P = []
    for i in range(8, 408):
        E1_OIL_P.append(float(exampleData[i][24]))
    E2_OIL_P = []
    for i in range(8, 408):
        E2_OIL_P.append(float(exampleData[i][49]))

    #OIL_T
    E1_OIL_T = []
    for i in range(8, 408):
        E1_OIL_T.append(float(exampleData[i][25]))
    E2_OIL_T = []
    for i in range(8, 408):
        E2_OIL_T.append(float(exampleData[i][50]))

    #Air Speed
    Air_Speed = []
    for i in range(8, 408):
        Air_Speed.append(float(exampleData[i][80]))
    
    #Altitude
    Altitude = []
    for i in range(8, 408):
        Altitude.append(float(exampleData[i][1]))
    
    sheet.cell(row = 4, column = j).value = max(ENG_1_TGT)
    sheet.cell(row = 5, column = j).value = max(ENG_2_TGT)

    sheet.cell(row = 13, column = j).value = max(ENG_1_N1)
    sheet.cell(row = 14, column = j).value = max(ENG_2_N1)
    sheet.cell(row = 24, column = j).value = max(ENG_1_N2)
    sheet.cell(row = 25, column = j).value = max(ENG_2_N2)

    sheet.cell(row = 37, column = j).value = max(E1_LP_VIB)
    sheet.cell(row = 38, column = j).value = max(E2_LP_VIB)
    sheet.cell(row = 47, column = j).value = max(E1_HP_VIB)
    sheet.cell(row = 48, column = j).value = max(E2_HP_VIB)
    sheet.cell(row = 57, column = j).value = max(E1_OIL_P)
    sheet.cell(row = 58, column = j).value = max(E2_OIL_P)
    sheet.cell(row = 70, column = j).value = max(E1_OIL_T)
    sheet.cell(row = 71, column = j).value = max(E2_OIL_T)
    #Air Speed
    sheet.cell(row = 81, column = j).value = max(Air_Speed)
    #Altitude
    sheet.cell(row = 82, column = j).value = max(Altitude)

  
    
    exampleFile.close()
    j = j + 1
    
 
    
    print('ENG_1_TGT: ' + str(max(ENG_1_TGT)))
    print('ENG_2_TGT: ' + str(max(ENG_2_TGT)))

    print('ENG_1_N1: ' + str(max(ENG_1_N1)))
    print('ENG_2_N1: ' + str(max(ENG_2_N1)))
    print('ENG_1_N2: ' + str(max(ENG_1_N2)))
    print('ENG_2_N2: ' + str(max(ENG_2_N2)))

    print('E1_LP_VIB: ' + str(max(E1_LP_VIB)))
    print('E2_LP_VIB: ' + str(max(E2_LP_VIB)))
    print('E1_HP_VIB: ' + str(max(E1_HP_VIB)))
    print('E2_HP_VIB: ' + str(max(E2_HP_VIB)))
    print('E1_OIL_P: ' + str(max(E1_OIL_P)))
    print('E2_OIL_P: ' + str(max(E2_OIL_P)))
    print('E1_OIL_T: ' + str(max(E1_OIL_T)))
    print('E2_OIL_T: ' + str(max(E2_OIL_T)))
    print('\n')

    #切到第2個工作分頁
    sheet = wb['Flist']
    sheet.cell(row = m, column = 1).value = fileName
    m = m + 1
#Mmm = time.strftime('%b')
Mmm = folderName[4:]
YYYY = time.strftime('%Y')
#切到第1個工作分頁
sheet = wb['Cruise Monitor data']
sheet.cell(row = 1, column = 1).value = tailNum + ' Take Off Performance Data Analyze-' + Mmm + ' ' + YYYY


#Engine_Trend_Cruise-Jul 2018 B-99888
wb.save('Engine_Trend_TakeOff-' + Mmm + ' ' + YYYY + ' ' + tailNum + '.xlsx')

