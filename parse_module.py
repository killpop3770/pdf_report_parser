import camelot
import json
import os

string = ""

def read_from_pdf(file):
	tables = camelot.read_pdf(file)
	os.chdir("..")
	tables.export('foo.json', f='json', compress=False)

	string = " ".join(file.split("_")[:-4])
	string = string.replace("д. ", "д.")
	string = str(string.replace("к. ", "корп.")[5:])
	print("File name: "+string)