from openpyxl import load_workbook
from fuzzywuzzy import fuzz

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

i = 2

with open('test.txt', 'w') as f:
	while i <= 200:
		count = 0
		for j in range(1, 11):
			one = sheet[f'E{i}'].value.replace('-', '').replace('_', '')
			two = sheet[f'E{i+j}'].value.replace('-', '').replace('_', '')
			onearr = ''.join([i for i in [char for char in one] if i.isdigit()])
			twoarr = ''.join([i for i in [char for char in two] if i.isdigit()])
			if fuzz.token_sort_ratio(one, two) >= 80 and fuzz.token_sort_ratio(onearr, twoarr) >= 85 or fuzz.token_sort_ratio(onearr, twoarr) >= 90:
				f.write(f"{i}: {sheet[f'E{i}'].value}\n")
				print(i)
				count += 1
			else:
				f.write(f'{i}: NOT MATCH')

		if count > 0:
			i += count
		else:
			i += 1

one = sheet['E7'].value.replace('-', '').replace('_', '')
two = sheet['E8'].value.replace('-', '').replace('_', '')
print(one, two)

print(fuzz.token_sort_ratio(one, two))

onearr = ''.join([i for i in [char for char in one] if i.isdigit()])
twoarr = ''.join([i for i in [char for char in two] if i.isdigit()])
print(fuzz.token_sort_ratio(onearr, twoarr))
print(onearr, twoarr)
#2527