## Neccesaary Imports

import os
import glob
import datetime
from pyexcel.cookbook import merge_all_to_a_book #pip3 install pyexcel pyexcel-xlsx

x = datetime.datetime.now()

today = x.strftime("%A-%d-%m-%Y")
print(today)

folder_path = r'C:\\Users\\jagtapawan\\Downloads\\' # Location of folder to perform activity
file_type = '\*csv' # check for all file type where csv - Comma Seperated Value
files = glob.glob(folder_path + file_type) # add the folder path and file type to 'files' varibale
latest_file = max(files, key=os.path.getctime) # this will fetch the latest file in the particular folder location

print (latest_file) # output the latest csv file in particual folder path

old_name = latest_file # save the latest file name to Old file variable. 
new_name = "C:\\Users\\jagtapawan\\Downloads\\Machines"+today+".csv" # rename the file by appending Machines + today's date + format that is csv

os.rename(old_name,new_name) # rename latest file in that particular folder with new name consisting of Machines + today's date + format that is csv

merge_all_to_a_book(glob.glob("C:\\Users\\jagtapawan\\Downloads\\Machines"+today+".csv"), "C:\\Users\\jagtapawan\\Downloads\\Machines"+today+".xlsx") # Since we cannot save data in csv format or make changes to csv file, convert the file to Excel and perform operations on the excel.
