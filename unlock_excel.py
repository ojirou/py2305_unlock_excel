import win32com.client as win32
import pandas as pd
import subprocess
import datetime
import os
import shutil
FileName='sample.xlsx'
Dir=r'C:\\Users\\user\\Downloads'
OrgFilePath=Dir+'\\'+FileName
#  Backup
today=datetime.date.today()
num=1
BkDir=Dir+'\\{:%y%m%d}Backup'.format(today,num)
os.mkdir(BkDir)
shutil.move(OrgFilePath, BkDir)
#  Open
BkFilePath=BkDir+'\\'+FileName
Pw='〇〇'
excel = win32.gencache.EnsureDispatch('Excel.Application')
#excel = win32.DispatchEx('Excel.Application')
wb = excel.Workbooks.Open(BkFilePath, Password=Pw)