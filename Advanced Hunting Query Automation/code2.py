## Neccesaary Imports
import os
import glob
import datetime
from pyexcel.cookbook import merge_all_to_a_book #pip3 install pyexcel pyexcel-xlsx
x = datetime.datetime.now()

today = x.strftime("%A-%d-%m-%Y")
print(today)