from pywinauto import application
import time
import os

os.system('taskkill /IM coStarter* /F /T')
os.system('taskkill /IM CpStart* /F /T')
os.system('taskkill /IM DibServer* /F /T')
os.system('wmic process where "name like \'%coStarter%\'" call terminate')
os.system('wmic process where "name like \'%CpStart%\'" call terminate')
os.system('wmic process where "name like \'%DibServer%\'" call terminate')
time.sleep(5)        
        

app = application.Application()
app.start('D:\CREON\STARTER\coStarter.exe /prj:cp /id:kim3329 /pwd:tlt5656 /pwdcert:Godlovesme60* /autostart')
time.sleep(60)