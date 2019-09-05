import os, datetime, time
if time.strftime("%y-%m-%d",time.localtime()) < '20-04-01':
    time.sleep(0.001)
else:
    os._exit()
    print('you cannot see this text')
    input()
