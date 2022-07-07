from openpyxl import load_workbook

#Utility functions

def changeActive(name, workbook):
	for i,v in enumerate(workbook.sheetnames):
		if v == name:
			wb.active = i
			break


#Main

wb = load_workbook('Customer General Code #2  6-28-2022.xlsx', data_only=True, read_only=True)

changeActive('Sheet2', wb)
sheet = wb.active


print(sheet['E2'].value)