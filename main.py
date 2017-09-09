# import datetime, xlrd	#module for read from xls and xlsx file
# import math				#module for exp (methods: calc_index_proj and lider_index)
# import os 				#module for get files in dir(method: get_files)
# import sys 				#module for get name of system for get path(method: get_symbol)

from task import *
    
exel_files = get_files()
names = [] 
for file in exel_files:
    sheet = open_file(file)         #get first sheet of book
    get_names(sheet, names)

#generate a dict with the names of employees and effectiveness = 0
eff_workers = {name: 0 for name in names}
for file in exel_files:
    sheet = open_file(file)
    for name in eff_workers:
        value = eff_index(sheet, name)
        if (value != 0):
            eff_workers[name] += value


#sorted dict of employes by effectiveness
sorted_data = sorted(eff_workers.items(), key=lambda x:x[1], reverse=True)
for data in sorted_data:
    print(data[0])