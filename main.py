# import datetime, xlrd	#module for read from xls and xlsx file
# import math				#module for exp (methods: calc_index_proj and lider_index)
# import os 				#module for get files in dir(method: get_files)
# import sys 				#module for get name of system for get path(method: get_symbol)

from task import *

exel_files = get_files()
names = [] 
for file in exel_files:
	sheet = open_file(file)			#get first sheet of book
	get_names(sheet, names)
	

eff_workers = []
for i in range(len(names)):
	eff_workers.append([])
	for j in range(1):
		eff_workers[i].append(names[i])
		eff_workers[i].append(0)


for file in exel_files:
	sheet = open_file(file)	
	for i in range(len(names)):
		num = 0
		value = eff_index(sheet, names[i])
		if (value != 0):
			eff_workers[i][1] = eff_workers[i][1] + value
			num = num + 1

sort_by_index(eff_workers)

for i in range(len(eff_workers)):
	names[i] = eff_workers[i][0]
print(names)
