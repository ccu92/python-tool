import pyautogui, time
from selenium import webdriver
print('Ctrl + C to queit')
'''
try:
    while True:
        x, y = pyautogui.position()
        positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
        print(positionStr, end='')
        print('\b' * len(positionStr), end='', flush=True)
except KeyboardInterrupt:
    print('Done')

browser = webdriver.Firefox(executable_path = 'C:\\Users\\peter\\AppData\\Local\\Programs\\Python\\Python36-32\\geckodriver')
browser.get('http://mail.yonglin.org.tw/')
time.sleep(5)
'''

pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('C:\\Users\\peter\\Desktop\\arrowUp.png')))


pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('C:\\Users\\peter\\Desktop\\settingButton.png')))
time.sleep(1)

pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('C:\\Users\\peter\\Desktop\\autoReplyTickBox.png')))
time.sleep(0.5)

pyautogui.scroll(-2000)

#pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('C:\\Users\\peter\\Desktop\\save.png')))
#time.sleep(1)

#pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('C:\\Users\\peter\\Desktop\\finish.png')))
