#make a timeStamp image
from PIL import Image, ImageDraw, ImageFont
import os, time, shutil
from selenium import webdriver

im = Image.new('RGBA', (650, 150), 'white')
draw = ImageDraw.Draw(im)
screenshotTime = time.strftime('%y_%m_%d_%H_%M_%S')
fontFolder = 'C:\\Windows\\Fonts'
arialFont = ImageFont.truetype(os.path.join(fontFolder, 'Arial.ttf'), 32)
draw.text((100, 35), screenshotTime + '_Screenshot', fill = 'red', font = arialFont)
im.save('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\' + 'timeStamp.png')

#make a screenshot
#should install selenium prior 
# place geckodriver.exe in a folder
browser = webdriver.Firefox(executable_path = 'C:\\Users\\peter\\AppData\\Local\\Programs\\Python\\Python36-32\\geckodriver')
browser.minimize_window()

# target web page to take screenshot
browser.get('http://www.caa.gov.tw/BIG5/ad/index.asp?page=1')

# save ('path of folder' + 'time stamp' + 'type of image (.png)')
browser.save_screenshot('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\' + screenshotTime + '.png')
browser.close()

#paste timeStamp on screenshot
os.chdir('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\')

'''
for fileName in [os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\')]:
    fileName.sort(reverse = True)
'''

screenshotWithoutTimeStamp = Image.open(screenshotTime + '.png')
TimeStamp = Image.open('timeStamp.png')
screenshotWithoutTimeStamp.paste(TimeStamp, (1572-650, 200))
screenshotWithoutTimeStamp.save(screenshotTime + '.png')
os.remove("C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\timeStamp.png")#檔案路徑和名稱

'''
#log in localHd
logFile = open('!screenshotLog.txt', 'a')
logFile.write(screenshotTime + ' in localHd,' + '\n')
logFile.close()
'''

# copy screenshot to web hd Y:\
shutil.copy('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\' + screenshotTime + '.png', 'Y:\\AD screenshot\\')


#log in text
logFile = open('!screenshotLog.txt', 'a')
logFile.write(screenshotTime + ' in localHd,' + screenshotTime + ' in webHd,' + '\n')
logFile.close()
# copy log to web hd Y:\
for fileName in [os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\')]:
    fileName.sort(reverse = True)
shutil.copy('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\!screenshotLog.txt', 'Y:\\AD screenshot\\')

