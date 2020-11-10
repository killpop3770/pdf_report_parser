import camelot
import json
import os
from openpyxl import load_workbook

# files = os.listdir(".")
# print(files)
# print("\n")

#Добавить адреса

files = {'1.pdf', #ГВС
			#'2.pdf', #ЦО
			'3.pdf', #ГВС
			#'4.pdf', #ЦО
			'5.pdf', #ГВС
			#'6.pdf', #ЦО
			'7.pdf'} #ГВС

row = 0

for file in files:
	row+=1
	print("File number: " +str(row))

	tables = camelot.read_pdf(file)

	tables.export('foo.json', f='json', compress=False) # json, excel, html

	wb = load_workbook('./first.xlsx')
	sheet = wb['Лист1']

	with open('foo-page-1-table-1.json') as report:
		data = json.loads(report.read())
		count = 0

		# TWO ALTERNATIVES FOR ANALYZE
		# for j in range(11):
		# 	if "tпод" in data[0].get(str(j)):

		
		print("Часы наработки: "+str(data[-1].get('10'))+"\n")

		print("Средняя температура подачи: "+str(data[-1].get('5'))+"\n")

		for i in range(2, len(data)-1):
			T = data[i].get('5')
			Tflow = T.replace(',', '.')
			if float(Tflow) < 60.0:
				count+=1
		print("Дней с температурой менее 60.0: "+str(count)+"\n")

		position = str(row)

		Hours = sheet[f"B{position}"]
		Tflow = sheet[f"C{position}"]
		Days = sheet[f"D{position}"]

		Hours.value = data[-1].get('10')
		Tflow.value = data[-1].get('5')
		Days.value = count

		wb.save("first.xlsx")
