import camelot
import json
import os
from openpyxl import load_workbook

directory = "reports"
wb = load_workbook('./first.xlsx')
sheet = wb['Лист1']
files = os.listdir(os.path.abspath(directory))

for file in files:
	os.chdir(directory)
	tables = camelot.read_pdf(file)
	os.chdir("..")
	tables.export('foo.json', f='json', compress=False)

	string = " ".join(file.split("_")[:-4])
	string = string.replace("д. ", "д.")
	string = str(string.replace("к. ", "корп.")[5:])
	print("File name: " +string)

	with open('foo-page-1-table-1.json') as report:
		data = json.loads(report.read()) 
		count = 0

		# TWO ALTERNATIVES FOR ANALYZE
		# for j in range(11):
		# 	if "tпод" in data[0].get(str(j)):

		for nameCell in range(3, sheet.max_row):
			address = str(sheet[f'B{nameCell}'].value)
			if string == address:
				print("Status: ok!")

				print("Operating hours: "+str(data[-1].get('10')))

				print("Average Tflow: "+str(data[-1].get('5')))

				for i in range(2, len(data)-1):
					T = data[i].get('5')
					Tflow = T.replace(',', '.')
					if float(Tflow) < 60.0:
						count+=1
				print("Days with Tflow before 60.0: "+str(count)+"\n")

				position = str(nameCell)

				hours = sheet[f"D{position}"]
				Tflow = sheet[f"E{position}"]
				days = sheet[f"F{position}"]

				hours.value = data[-1].get('10')
				Tflow.value = data[-1].get('5')
				days.value = count

	wb.save("first.xlsx")