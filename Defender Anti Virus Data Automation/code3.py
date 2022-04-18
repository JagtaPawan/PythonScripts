## Neccesaary Imports

from ctypes import alignment
from itertools import count
from turtle import left, right
from numpy import left_shift, pi
import pandas as pd
import datetime
import win32com.client as win32 #pip3 install openpyxl

x = datetime.datetime.now()
today = x.strftime("%A-%d-%m-%Y")

target = r'C:\Users\jagtapawan\Downloads\Machines' +str(today)+ '.xlsx'
df = pd.read_excel(target)
df.fillna('',inplace=True)

## Counts of Machines

Counts = df['Device Name'].count()
print("\n Total Number of Machines with AV Disbaled are :", Counts)

## MachineCount

pivot1111 = df.pivot_table(index=['OS Platform'], values=['Device Name'], aggfunc='count')
pivot1111.fillna('',inplace=True)
pivot1111.to_csv('MachineCount.csv')
print("\n Total Number of Machines with AV Disbaled are by OS Platform : \n \n", pivot1111)

outloook = win32.Dispatch('outlook.application') # Open Outlook Application

mail = outloook.CreateItem(0) # Create Mail Item

x = datetime.datetime.now()
today = x.strftime("%A-%d-%m-%Y")

# Fill the Paramters

mail.To = 'dummy@gmail.com'
mail.CC = 'dummy@gmail.com'
mail.Subject = ' *** AV Detailed Report on {} ***'.format(str(today))

mail.HTMLBody = "<p> Hi<br><br>  Please find the Inventory "+str(today)+" <br><br> Total Number of Machines with AV disabled are <b>"+str(Counts)+"</b> <br> Total Number of Machines with AV disabled are by OS Platform :"+str(pivot1111.to_html())+"</p>"

# Send the email
mail.Send()
