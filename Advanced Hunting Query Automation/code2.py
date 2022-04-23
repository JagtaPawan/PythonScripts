## Neccesaary Imports
import os
import glob
import datetime
from pyexcel.cookbook import merge_all_to_a_book #pip3 install pyexcel pyexcel-xlsx
x = datetime.datetime.now()

today = x.strftime("%A-%d-%m-%Y")
print(today)

folder_path = r'C:\\Users\\jagtapawan\\Downloads\\'
file_type = '\*csv'
files = glob.glob(folder_path + file_type)
latest_file = max(files, key=os.path.getctime)

print (latest_file)

old_name = latest_file
new_name = "C:\\Users\\jagtapawan\\Downloads\\ADH"+today+".csv"

os.rename(old_name,new_name)
merge_all_to_a_book(glob.glob("C:\\Users\\jagtapawan\\Downloads\\ADH"+today+".csv"), "C:\\Users\\jagtapawan\\Downloads\\ADH"+today+".xlsx")
