import os
from openpyxl import load_workbook
import time
import parse_module
import data_module

startTime = time.time()

directory = "reports"
absDirectory = os.path.abspath(directory)
files = os.listdir(absDirectory)

wb = load_workbook('./first.xlsx')
sheet = wb['Лист1']

for file in files:
	os.chdir(absDirectory)
	
	if "ЦО" in file or "ХВС" in file:
		continue

	parse_module.read_from_pdf(file)

	data_module.insert_in_excel(sheet, parse_module.string)

	wb.save("first.xlsx")

	print("RunTime: "+str(time.time()-startTime))