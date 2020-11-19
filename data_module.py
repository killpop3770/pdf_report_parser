import camelot
import json
from openpyxl import load_workbook

def insert_in_excel(sheet, string):
	for nameCell in range(3, sheet.max_row):
		address = str(sheet[f'B{nameCell}'].value)
		if string == address:

				with open('foo-page-1-table-1.json') as report:
					data = json.loads(report.read()) 
					count = 0

					try:
						for i in range(2, len(data)-1):
							T = data[i].get('5')
							if T != "" or T is not None:
									Tflow = T.replace(',', '.')
									if float(Tflow) < 60.0:
										count+=1
							else:
								Tflow = 60.0
					except Exception as e:
						print("Bad data\n")
					finally:
						Tflow = 60.0

					position = str(nameCell)

					hours = sheet[f"D{position}"]
					Tflow = sheet[f"E{position}"]
					days = sheet[f"F{position}"]

					hours.value = data[-1].get('10')
					Tflow.value = data[-1].get('5')
					days.value = count