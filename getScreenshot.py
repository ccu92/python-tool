

#make a timeStamp image
from PIL import Image, ImageDraw, ImageFont
import os
import time

im = Image.new('RGBA', (650, 150), 'white')
draw = ImageDraw.Draw(im)

screenshotTime = time.strftime('%y_%m_%d_%H_%M_%S')

fontFolder = 'C:\\Windows\\Fonts'
arialFont = ImageFont.truetype(os.path.join(fontFolder, 'Arial.ttf'), 32)

draw.text((100, 35), screenshotTime + '_Screenshot', fill = 'red', font = arialFont)

im.save('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\' + 'timeStamp.png')





#make a screenshot

#should install selenium prior 
from selenium import webdriver
import time

# place geckodriver.exe in a folder
browser = webdriver.Firefox(executable_path = 'C:\\Users\\peter\\AppData\\Local\\Programs\\Python\\Python36-32\\geckodriver')

browser.minimize_window()

# target web page to take screenshot
browser.get('http://www.caa.gov.tw/BIG5/ad/index.asp?page=1')



# save ('path of folder' + 'time stamp' + 'type of image (.png)')
browser.save_screenshot('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\' + screenshotTime + '.png')


#paste timeStamp on screenshot
import os
os.chdir('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\')
for fileName in [os.listdir('C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\')]:
    fileName.sort(reverse = True)

screenshotWithoutTimeStamp = Image.open((fileName[1]))
TimeStamp = Image.open('timeStamp.png')
screenshotWithoutTimeStamp.paste(TimeStamp, (1572-650, 200))
screenshotWithoutTimeStamp.save(fileName[1])
#os.remove("C:\\Users\\peter\\OneDrive\\工作夾\\screenshot\\timeStamp.png")#檔案路徑和名稱



# web hd


im.save('Y:\\AD screenshot\\' + 'timeStamp.png')

browser.save_screenshot('Y:\\AD screenshot\\' + screenshotTime + '.png')

browser.close()

os.chdir('Y:\\AD screenshot\\')
for fileName in [os.listdir('Y:\\AD screenshot\\')]:
    fileName.sort(reverse = True)

screenshotWithoutTimeStamp = Image.open((fileName[1]))
TimeStamp = Image.open('timeStamp.png')
screenshotWithoutTimeStamp.paste(TimeStamp, (1572-650, 200))
screenshotWithoutTimeStamp.save(fileName[1])




