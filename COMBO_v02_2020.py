import csv, openpyxl, os, time


print('請輸入機尾號 ex: B-##### (*** 適用機型：G550 或 G650 ***)')
tailNum = input()
if tailNum == "B-56789" or tailNum == "B-99988":
    print("*** G650 ***\n")
    print('請輸入資料夾名稱 ex: yyyymm')
    folderName = input()


#CRUISE##CRUISE##CRUISE##CRUISE##CRUISE##CRUISE#

    
    wb = openpyxl.load_workbook('C:\\RCP\\template\\Engine_Trend_Cruise 650 template.xlsx')
    sheet = wb['Cruise Monitor data']

    j = 2 #Because column B is 2
    k = 1 #No. 1 is started for fileName's No.
    m = 1 #Flist is started from 1
    for fileName in os.listdir('C:\\RCP\\' + tailNum + '\\' + folderName + '\\' + 'C'):
        print('No.: ' + str(k) +' '+ fileName + '\n')
        k = k+1

        #切到第1個工作分頁
        sheet = wb['Cruise Monitor data']


        exampleFile = open('C:\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'C' + '\\' + fileName)
        exampleReader = csv.reader(exampleFile)
        exampleData = list(exampleReader)
        n = int(exampleReader.line_num)#最多列數

        #TGT
        ENG_1_TGT = []
        for i in range(8, n):
            ENG_1_TGT.append(float(exampleData[i][24]))
        ENG_2_TGT = []
        for i in range(8, n):
            ENG_2_TGT.append(float(exampleData[i][45]))

        #FF
        ENG_1_FF = []
        for i in range(8, n):
            ENG_1_FF.append(float(exampleData[i][25]))
        ENG_2_FF = []
        for i in range(8, n):
            ENG_2_FF.append(float(exampleData[i][46]))

        #N1
        ENG_1_N1 = []
        for i in range(8, n):
            ENG_1_N1.append(float(exampleData[i][31]))
        ENG_2_N1 = []
        for i in range(8, n):
            ENG_2_N1.append(float(exampleData[i][52]))

        #N2
        ENG_1_N2 = []
        for i in range(8, n):
            ENG_1_N2.append(float(exampleData[i][33]))
        ENG_2_N2 = []
        for i in range(8, n):
            ENG_2_N2.append(float(exampleData[i][54]))

        #P30
        ENG_1_P30 = []
        for i in range(8, n):
            ENG_1_P30.append(float(exampleData[i][34]))
        ENG_2_P30 = []
        for i in range(8, n):
            ENG_2_P30.append(float(exampleData[i][55]))

        #T30
        ENG_1_T30 = []
        for i in range(8, n):
            ENG_1_T30.append(float(exampleData[i][23]))
        ENG_2_T30 = []
        for i in range(8, n):
            ENG_2_T30.append(float(exampleData[i][44]))

        #LP_VIB
        E1_LP_VIB = []
        for i in range(8, n):
            E1_LP_VIB.append(float(exampleData[i][64]))
        E2_LP_VIB = []
        for i in range(8, n):
            E2_LP_VIB.append(float(exampleData[i][68]))

        #HP_VIB
        E1_HP_VIB = []
        for i in range(8, n):
            E1_HP_VIB.append(float(exampleData[i][66]))
        E2_HP_VIB = []
        for i in range(8, n):
            E2_HP_VIB.append(float(exampleData[i][70]))

        #OIL_P
        E1_OIL_P = []
        for i in range(8, n):
            E1_OIL_P.append(float(exampleData[i][26]))
        E2_OIL_P = []
        for i in range(8, n):
            E2_OIL_P.append(float(exampleData[i][47]))

        #OIL_T
        E1_OIL_T = []
        for i in range(8, n):
            E1_OIL_T.append(float(exampleData[i][27]))
        E2_OIL_T = []
        for i in range(8, n):
            E2_OIL_T.append(float(exampleData[i][48]))

        #Mach SP
        Mach_SP = []
        for i in range(8, n):
            Mach_SP.append(float(exampleData[i][2]))
        
        #Altitude
        Altitude = []
        for i in range(8, n):
            Altitude.append(float(exampleData[i][3]))
        
        sheet.cell(row = 3, column = j).value = max(ENG_1_TGT)
        sheet.cell(row = 4, column = j).value = max(ENG_2_TGT)
        sheet.cell(row = 12, column = j).value = max(ENG_1_FF)
        sheet.cell(row = 13, column = j).value = max(ENG_2_FF)
        sheet.cell(row = 22, column = j).value = max(ENG_1_N1)
        sheet.cell(row = 23, column = j).value = max(ENG_2_N1)
        sheet.cell(row = 34, column = j).value = max(ENG_1_N2)
        sheet.cell(row = 35, column = j).value = max(ENG_2_N2)
        sheet.cell(row = 44, column = j).value = max(ENG_1_P30)
        sheet.cell(row = 45, column = j).value = max(ENG_2_P30)
        sheet.cell(row = 54, column = j).value = max(ENG_1_T30)
        sheet.cell(row = 55, column = j).value = max(ENG_2_T30)
        sheet.cell(row = 65, column = j).value = max(E1_LP_VIB)
        sheet.cell(row = 66, column = j).value = max(E2_LP_VIB)
        sheet.cell(row = 75, column = j).value = max(E1_HP_VIB)
        sheet.cell(row = 76, column = j).value = max(E2_HP_VIB)
        sheet.cell(row = 85, column = j).value = max(E1_OIL_P)
        sheet.cell(row = 86, column = j).value = max(E2_OIL_P)
        sheet.cell(row = 96, column = j).value = max(E1_OIL_T)
        sheet.cell(row = 97, column = j).value = max(E2_OIL_T)
        #Mach SP
        sheet.cell(row = 107, column = j).value = max(Mach_SP)
        #Altitude
        sheet.cell(row = 108, column = j).value = max(Altitude)

      
        
        exampleFile.close()
        j = j + 1
        
     
        
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
        print('E1_OIL_P: ' + str(min(E1_OIL_P)))
        print('E2_OIL_P: ' + str(min(E2_OIL_P)))
        print('E1_OIL_T: ' + str(max(E1_OIL_T)))
        print('E2_OIL_T: ' + str(max(E2_OIL_T)))
        print('\n')

        #切到第2個工作分頁
        sheet = wb['Flist']
        sheet.cell(row = m, column = 1).value = fileName
        m = m + 1
    #Mmm = time.strftime('%b')
    Mmm = folderName[4:]
    YYYY = folderName[0:4]
    #切到第1個工作分頁
    sheet = wb['Cruise Monitor data']
    sheet.cell(row = 1, column = 1).value = tailNum + ' Cruise Performance Data Analyze-' + Mmm + ' ' + YYYY


    #Engine_Trend_Cruise-Jul 2018 B-99888
    wb.save('Engine_Trend_Cruise-' + Mmm + ' ' + YYYY + ' ' + tailNum + '.xlsx')





#TAKE OFF##TAKE OFF##TAKE OFF##TAKE OFF##TAKE OFF##TAKE OFF#


    wb = openpyxl.load_workbook('C:\\RCP\\template\\Engine_Trend_TakeOff 650 template.xlsx')
    sheet = wb['Cruise Monitor data']

    j = 2 #Because column B is 2
    k = 1 #No. 1 is started for fileName's No.
    m = 1 #Flist is started from 1
    for fileName in os.listdir('C:\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T'):#T means TakeOff
        print('No.: ' + str(k) +' '+ fileName + '\n')
        k = k+1
        
        #切到第1個工作分頁
        sheet = wb['Cruise Monitor data']


        exampleFile = open('C:\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T' + '\\' + fileName)
        exampleReader = csv.reader(exampleFile)
        exampleData = list(exampleReader)


        #TGT
        ENG_1_TGT = []
        for i in range(8, 408):
            ENG_1_TGT.append(float(exampleData[i][24]))
        ENG_2_TGT = []
        for i in range(8, 407):
            ENG_2_TGT.append(float(exampleData[i][45]))
       
        #N1
        ENG_1_N1 = []
        for i in range(8, 408):
            ENG_1_N1.append(float(exampleData[i][31]))
        ENG_2_N1 = []
        for i in range(8, 408):
            ENG_2_N1.append(float(exampleData[i][52]))

        #N2
        ENG_1_N2 = []
        for i in range(8, 408):
            ENG_1_N2.append(float(exampleData[i][33]))
        ENG_2_N2 = []
        for i in range(8, 408):
            ENG_2_N2.append(float(exampleData[i][54]))

        #LP_VIB
        E1_LP_VIB = []
        for i in range(8, 408):
            E1_LP_VIB.append(float(exampleData[i][64]))
        E2_LP_VIB = []
        for i in range(8, 408):
            E2_LP_VIB.append(float(exampleData[i][68]))

        #HP_VIB
        E1_HP_VIB = []
        for i in range(8, 408):
            E1_HP_VIB.append(float(exampleData[i][66]))
        E2_HP_VIB = []
        for i in range(8, 408):
            E2_HP_VIB.append(float(exampleData[i][70]))

        #OIL_P
        E1_OIL_P = []
        for i in range(8, 408):
            E1_OIL_P.append(float(exampleData[i][26]))
        E2_OIL_P = []
        for i in range(8, 408):
            E2_OIL_P.append(float(exampleData[i][47]))

        #OIL_T
        E1_OIL_T = []
        for i in range(8, 408):
            E1_OIL_T.append(float(exampleData[i][27]))
        E2_OIL_T = []
        for i in range(8, 408):
            E2_OIL_T.append(float(exampleData[i][48]))

        #Air Speed
        Air_Speed = []
        for i in range(8, 408):
            Air_Speed.append(float(exampleData[i][1]))
        
        #Altitude
        Altitude = []
        for i in range(8, 408):
            Altitude.append(float(exampleData[i][3]))
        
        sheet.cell(row = 3, column = j).value = max(ENG_1_TGT)
        sheet.cell(row = 4, column = j).value = max(ENG_2_TGT)

        sheet.cell(row = 13, column = j).value = max(ENG_1_N1)
        sheet.cell(row = 14, column = j).value = max(ENG_2_N1)
        sheet.cell(row = 24, column = j).value = max(ENG_1_N2)
        sheet.cell(row = 25, column = j).value = max(ENG_2_N2)

        sheet.cell(row = 34, column = j).value = max(E1_LP_VIB)
        sheet.cell(row = 35, column = j).value = max(E2_LP_VIB)
        sheet.cell(row = 44, column = j).value = max(E1_HP_VIB)
        sheet.cell(row = 45, column = j).value = max(E2_HP_VIB)
        sheet.cell(row = 55, column = j).value = max(E1_OIL_P)
        sheet.cell(row = 56, column = j).value = max(E2_OIL_P)
        sheet.cell(row = 65, column = j).value = max(E1_OIL_T)
        sheet.cell(row = 66, column = j).value = max(E2_OIL_T)
        #Air Speed
        sheet.cell(row = 76, column = j).value = max(Air_Speed)
        #Altitude
        sheet.cell(row = 77, column = j).value = max(Altitude)

      
        
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
        print('E1_OIL_P: ' + str(min(E1_OIL_P)))
        print('E2_OIL_P: ' + str(min(E2_OIL_P)))
        print('E1_OIL_T: ' + str(max(E1_OIL_T)))
        print('E2_OIL_T: ' + str(max(E2_OIL_T)))
        print('\n')

        #切到第2個工作分頁
        sheet = wb['Flist']
        sheet.cell(row = m, column = 1).value = fileName
        m = m + 1
    #Mmm = time.strftime('%b')
    Mmm = folderName[4:]
    YYYY = folderName[4:]
    #切到第1個工作分頁
    sheet = wb['Cruise Monitor data']
    sheet.cell(row = 1, column = 1).value = tailNum + ' Take Off Performance Data Analyze-' + Mmm + ' ' + YYYY


    #Engine_Trend_Cruise-Jul 2018 B-99888
    wb.save('Engine_Trend_TakeOff-' + Mmm + ' ' + YYYY + ' ' + tailNum + '.xlsx')


#-------------------------------------------------------------------


elif tailNum == "B-99888" or tailNum == "B-90609":
    print("*** G550 ***\n")
    print('請輸入資料夾名稱 ex: yyyymm')
    folderName = input()


#CRUISE##CRUISE##CRUISE##CRUISE##CRUISE##CRUISE#


    wb = openpyxl.load_workbook('C:\\RCP\\template\\Engine_Trend_Cruise 550 template.xlsx')
    sheet = wb['Cruise Monitor data']

    j = 2 #Because column B is 2
    k = 1 #No. 1 is started for fileName's No.
    m = 1 #Flist is started from 1
    for fileName in os.listdir('C:\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'C'):
        print('No.: ' + str(k) +' '+ fileName + '\n')
        k = k+1
        
        #切到第1個工作分頁
        sheet = wb['Cruise Monitor data']


        exampleFile = open('C:\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'C' + '\\' + fileName)
        exampleReader = csv.reader(exampleFile)
        exampleData = list(exampleReader)
        n = int(exampleReader.line_num)#最多列數

        #TGT
        ENG_1_TGT = []
        try:
            for i in range(8, n):
                ENG_1_TGT.append(float(exampleData[i][22]))#
        except ValueError:
            n = n - 1
            for i in range(8, n):
                ENG_1_TGT.append(float(exampleData[i][22]))    
        ENG_2_TGT = []
        try:
            for i in range(8, n):
                ENG_2_TGT.append(float(exampleData[i][47]))#
        except ValueError:
            n = n - 1
            for i in range(8, n):
                ENG_2_TGT.append(float(exampleData[i][47]))

        #FF
        ENG_1_FF = []
        try:
            for i in range(8, n):
                ENG_1_FF.append(float(exampleData[i][23]))#
        except ValueError:
            n = n - 1
            for i in range(8, n):
                ENG_1_FF.append(float(exampleData[i][23]))
        ENG_2_FF = []
        for i in range(8, n):
            ENG_2_FF.append(float(exampleData[i][48]))

        #N1
        ENG_1_N1 = []
        for i in range(8, n):
            ENG_1_N1.append(float(exampleData[i][31]))
        ENG_2_N1 = []
        for i in range(8, n):
            ENG_2_N1.append(float(exampleData[i][56]))

        #N2
        ENG_1_N2 = []
        for i in range(8, n):
            ENG_1_N2.append(float(exampleData[i][33]))
        ENG_2_N2 = []
        for i in range(8, n):
            ENG_2_N2.append(float(exampleData[i][58]))

        #P30
        ENG_1_P30 = []
        for i in range(8, n):
            ENG_1_P30.append(float(exampleData[i][34]))
        ENG_2_P30 = []
        try:
            for i in range(8, n):
                ENG_2_P30.append(float(exampleData[i][59]))#
        except ValueError:
            n = n - 1
            for i in range(8, n):
                ENG_2_P30.append(float(exampleData[i][59]))
            
        #T30
        ENG_1_T30 = []
        for i in range(8, n):
            ENG_1_T30.append(float(exampleData[i][21]))
        ENG_2_T30 = []
        for i in range(8, n):
            ENG_2_T30.append(float(exampleData[i][46]))

        #LP_VIB
        E1_LP_VIB = []
        for i in range(8, n):
            E1_LP_VIB.append(float(exampleData[i][11]))
        E2_LP_VIB = []
        for i in range(8, n):
            E2_LP_VIB.append(float(exampleData[i][15]))

        #HP_VIB
        E1_HP_VIB = []
        for i in range(8, n):
            E1_HP_VIB.append(float(exampleData[i][10]))
        E2_HP_VIB = []
        try:   
            for i in range(8, n):
                E2_HP_VIB.append(float(exampleData[i][14]))#
        except ValueError:
            n = n - 1
            for i in range(8, n):
                E2_HP_VIB.append(float(exampleData[i][14]))

        #OIL_P
        E1_OIL_P = []
        try:
            for i in range(8, n):
                E1_OIL_P.append(float(exampleData[i][24]))#
        except ValueError:
            n = n - 1
            for i in range(8, n):
                E1_OIL_P.append(float(exampleData[i][24]))
        E2_OIL_P = []
        for i in range(8, n):
            E2_OIL_P.append(float(exampleData[i][49]))

        #OIL_T
        E1_OIL_T = []
        for i in range(8, n):
            E1_OIL_T.append(float(exampleData[i][25]))
        E2_OIL_T = []
        try:
            for i in range(8, n):
                E2_OIL_T.append(float(exampleData[i][50]))#
        except:
            n = n - 1
            for i in range(8, n):
                E2_OIL_T.append(float(exampleData[i][50]))

        #Mach SP
        Mach_SP = []
        try:
            for i in range(8, n):
                Mach_SP.append(float(exampleData[i][83]))#
        except ValueError:
            n = n - 1
            for i in range(8, n):
                Mach_SP.append(float(exampleData[i][83]))
        #Altitude
        Altitude = []
        for i in range(8, n):
            Altitude.append(float(exampleData[i][1]))
        
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
        sheet.cell(row = 90, column = j).value = min(E1_OIL_P)
        sheet.cell(row = 91, column = j).value = min(E2_OIL_P)
        sheet.cell(row = 103, column = j).value = max(E1_OIL_T)
        sheet.cell(row = 104, column = j).value = max(E2_OIL_T)
        #Mach SP
        sheet.cell(row = 114, column = j).value = max(Mach_SP)
        #Altitude
        sheet.cell(row = 115, column = j).value = max(Altitude)

      
        
        exampleFile.close()
        j = j + 1
        
     
        
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

        #切到第2個工作分頁
        sheet = wb['Flist']
        sheet.cell(row = m, column = 1).value = fileName
        m = m + 1
    #Mmm = time.strftime('%b')
    Mmm = folderName[4:]
    YYYY = folderName[4:]
    #切到第1個工作分頁
    sheet = wb['Cruise Monitor data']
    sheet.cell(row = 1, column = 1).value = tailNum + ' Cruise Performance Data Analyze-' + Mmm + ' ' + YYYY


    #Engine_Trend_Cruise-Mmm YYYY B-xxxxx
    wb.save('Engine_Trend_Cruise-' + Mmm + ' ' + YYYY + ' ' + tailNum + '.xlsx')


#TAKE OFF##TAKE OFF##TAKE OFF##TAKE OFF##TAKE OFF##TAKE OFF#

    
    wb = openpyxl.load_workbook('C:\\RCP\\template\\Engine_Trend_TakeOff 550 template.xlsx')
    sheet = wb['Cruise Monitor data']

    j = 2 #Because column B is 2
    k = 1 #No. 1 is started for fileName's No.
    m = 1 #Flist is started from 1
    for fileName in os.listdir('C:\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T'):#T means TakeOff
        print('No.: ' + str(k) +' '+ fileName + '\n')
        k = k+1
        
        #切到第1個工作分頁
        sheet = wb['Cruise Monitor data']


        exampleFile = open('C:\\RCP\\'+ tailNum + '\\' + folderName + '\\' + 'T' + '\\' + fileName)
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
        print('E1_OIL_P: ' + str(min(E1_OIL_P)))
        print('E2_OIL_P: ' + str(min(E2_OIL_P)))
        print('E1_OIL_T: ' + str(max(E1_OIL_T)))
        print('E2_OIL_T: ' + str(max(E2_OIL_T)))
        print('\n')

        #切到第2個工作分頁
        sheet = wb['Flist']
        sheet.cell(row = m, column = 1).value = fileName
        m = m + 1
    #Mmm = time.strftime('%b')
    Mmm = folderName[4:]
    YYYY = folderName[4:]
    #切到第1個工作分頁
    sheet = wb['Cruise Monitor data']
    sheet.cell(row = 1, column = 1).value = tailNum + ' Take Off Performance Data Analyze-' + Mmm + ' ' + YYYY


    #Engine_Trend_Cruise-Jul 2018 B-99888
    wb.save('Engine_Trend_TakeOff-' + Mmm + ' ' + YYYY + ' ' + tailNum + '.xlsx')


#-------------------------------------------------------------------    
elif tailNum == "B99888" or tailNum == "B90609" or tailNum == "B56789" or tailNum == "B99988":
    print("The tail number should be B-#####.\nNot B#####.")
    print("PRESS ENTER TO ESCAPE.")
    input()


else:
    print("CANNOT IDENT " + tailNum + ".")
    print("PRESS ENTER TO ESCAPE.")
    input()

