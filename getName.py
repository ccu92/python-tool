import os
print('請輸入機尾號')
tailNum = input()
print('請輸入資料夾名稱')
fileName = input()

for fileName in os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\RCP\\'+ tailNum + '\\' + fileName):
    print(fileName)


print('按Enter離開')
input()






