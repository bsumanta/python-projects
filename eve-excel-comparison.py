import sys
import xlrd
from xlwt import easyxf
from xlutils.copy import copy

file_expected=sys.argv[1]
file_actual=sys.argv[2]

book_expected = xlrd.open_workbook(file_expected)
sheet_expected = book_expected.sheet_by_index(0)

book_actual = xlrd.open_workbook(file_actual,formatting_info=True)
sheet_actual = book_actual.sheet_by_index(0)


#read the fields name
fields = []
for c in range(7,sheet_expected.ncols):
	fields.append(sheet_expected.cell(1,c))


#reads all the required columns of expected sheet
expected_field_dict = {}

for c in range(7,sheet_expected.ncols):
	header = sheet_expected.cell(1,c)
	col_data=[]

	for r in range(2,sheet_expected.nrows):
		col_data.append(str(sheet_expected.cell(r,c)))

	expected_field_dict[header] = col_data

#reads all the required cols of the actual sheet
actual_field_dict = {}

for c in range(7,sheet_actual.ncols):
	header = sheet_actual.cell(1,c)
	col_data = []

	for r in range(2,sheet_actual.nrows):
		col_data.append(str(sheet_actual.cell(r,c)))

	actual_field_dict[header] = col_data

#creates the check cols for each field
field_check_dict = {}
for f in fields:

	check_list  = []
	expected_list  = expected_field_dict[f]
	actual_list  = actual_field_dict[f]
	for i in range(0,len(expected_list)):
		check = True
		if expected_list[i]==actual_list[i]:
			check = True
		else:
			check = False

		check_list.append(check)

	field_check_dict[f] = check_list


# write the new check cols
wb = copy(book_actual)
w_sheet = wb.get_sheet(0)

for f in fields:
	field = field_check_dict[f]
	w_sheet.write(1,w_sheet.ncols+1,f+" check")
	for r in range(2,w_sheet.nrows):
		w_sheet.write(r,w_sheet.ncols+1,field[r-2])

wb.save(file_actual)

