import xlrd
import re
import datetime
import json

xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True
workbook = xlrd.open_workbook(r"C:\Users\kalee\Downloads\Python Skill Test.xlsx")
worksheet = workbook.sheet_by_index(0)
# print(worksheet)
nrows = worksheet.nrows
ncols = worksheet.ncols
# elements = [worksheet.cell(row_index, col_index).value for row_index in range(nrows) for col_index in range(ncols)]
headerfields = ["Quote Number", "Date", "Ship To", "Ship From", "Name"]
columnfields = ["LineNumber", "PartNumber", "Description", "Price"]
seperator = re.compile('^-----------*$')
output = {}
keys = {}
dict_list = []
flag = 1
dic = {}

def changeDate_format(value):
	date = datetime.datetime(1899, 12, 30)
	get_ = datetime.timedelta(int(value))
	get_col2 = str(date + get_)[:10]
	d = datetime.datetime.strptime(get_col2, '%Y-%m-%d')
	get_date = d.strftime('%Y-%m-%d')
	return get_date

for row_index in range(nrows):
	# ignore all empty rows
	if all([worksheet.cell(row_index, col_index).ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) for col_index in range(ncols)]):
		continue
	for col_index in range(ncols):
		cell = worksheet.cell(row_index, col_index)

		# check for separator to exist 
		if col_index == 1 and seperator.match(cell.value):               
			break
		# ignore empty columns
		if cell.ctype in  (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
			continue
		# check for Name: field
		if isinstance(cell.value, str) and cell.value.startswith("Name:"):
			key, val = re.split(r':\s*', cell.value, 1)
			output[key]= val
			continue
		# check for labels and correspoding values
		if flag and cell.value in headerfields:
			output[cell.value] = 0
			flag = 0
			continue
		else:
			if cell.ctype == xlrd.XL_CELL_DATE:
				output[list(output)[-1]] = changeDate_format(cell.value)
				flag = 1
				continue
			elif worksheet.cell(row_index, col_index-1).value == list(output)[-1]:
				output[list(output)[-1]] = cell.value
				flag = 1
				continue

		# check for Columns item list
		#columns header
		if cell.value == columnfields[0]:
			# print('flag', flag)
			keys[col_index] = cell.value
			flag = 0
		elif cell.value in columnfields[1:]:
			if flag:
				print('first column item is not LineNumber')
				break
			keys[col_index] = cell.value
		#column items
		else:
			# print(cell.value)
			if col_index in keys and keys[col_index] in columnfields:
				dic[keys[col_index]]= cell.value
				continue			
	else:
		if dic:
			dict_list.append(dic)
			dic = {}
	# Exist if separator is found and first column item (LineNumber) not found
	if (col_index == 1 and seperator.match(cell.value)) or (cell.value in columnfields and flag):
		break


#error message
for field in headerfields:
	if field not in output:
		print(f'Error message: {field} is not in headerfields')
	

output['items'] = dict_list


print(json.dumps(output, indent = 4))