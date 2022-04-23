#Neccessary Imports
from ctypes import alignment
from turtle import right
from numpy import pi
import pandas as pd
import datetime
import win32com.client as win32

x = datetime.datetime.now()
today = x.strftime("%A-%d-%m-%Y")

target = r'C:\Users\jagtapawan\Downloads\ADH' +str(today)+ '.xlsx'
df = pd.read_excel(target)
df.fillna('',inplace=True)

AVSonMD = df[(df["NewVersion"] == '') & df["DeviceName"].map(lambda x: x.startswith('wl'))]
print(AVSonMD.__getitem__("DeviceName"))

AVUpdated = df[(df["NewVersion"] != '') & df["DeviceName"].map(lambda x: x.startswith('wl'))]
print(AVUpdated.__getitem__("DeviceName"))

# Open Outlook Application
outloook = win32.Dispatch('outlook.application')

# Create Mail Item
mail = outloook.CreateItem(0)
x = datetime.datetime.now()
today = x.strftime("%A-%d-%m-%Y")

# Fill the Paramters

mail.To = 'dummy@dummy.com'
mail.CC = 'dummy@dummy.com'
mail.Subject = ' *** WELLA : AV Detailed Report on {} ***'.format(str(today))
mail.HTMLBody = "<p> Hi All<br><br> Please find the Inventory "+str(today)+"<br><br><br>Count of AV Latest Signature not updated : "+str(len(AVSonMD.__getitem__("DeviceName").to_dict()))+"<br>AV Latest Signature not updated :<br>"+str(AVSonMD.__getitem__("DeviceName").to_dict())+"</p>"

# Send the email
mail.Send()
