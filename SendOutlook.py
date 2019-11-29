# *===================================*
# -*- coding: utf-8 -*-
# * Time : 2019/11/29 10:39
# * Author : zhangsf
# *===================================*
from importlib import reload

import win32com.client as win32
import warnings
import pythoncom
import  sys
reload(sys)
# sys.setdefaultencoding("utf-8")
warnings.filterwarnings('ignore')
pythoncom.CoInitialize()
def sendmail():
    sub = 'outlook python mail test'
    body = 'my test\r\n my python mail'
    outlook = win32.Dispatch('outlook.application')
    receivers = ['xxx@xxx.com']
    mail = outlook.CreateItem(0)
    mail.To = receivers[0]
    mail.Subject = sub.encode('utf-8').decode('utf-8')
    mail.Body = body.encode('utf-8').decode('utf-8')
    mail.Attachments.Add('E:\code\python\SendEmail\/text1.txt')
    mail.Send()
sendmail()
