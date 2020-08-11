
import xlrd


file_name = "Swipe Records-2020-08-11_164350_68.xls"
sheet = "ExelData"


xl_workbook  = xlrd.open_workbook(file_name)

xl_sheet = xl_workbook.sheet_by_index(0)
print ('Sheet name: %s' % xl_sheet.name)

# Pull the first row by index
#  (rows/columns are also zero-indexed)
#
row = xl_sheet.row(0)  # 1st row

# Print 1st row values and types

row_name = {}
from xlrd.sheet import ctype_text   

print('(Column #) type:value')
for idx, cell_obj in enumerate(row):
    cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
    #print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))
    row_name[idx]=cell_obj.value

# Print all values, iterating through rows and columns
#



cell_value = {}

num_cols = xl_sheet.ncols   # Number of columns
for row_idx in range(0, xl_sheet.nrows):    # Iterate through rows
    #print ('-'*40)
    #print ('Row: %s' % row_idx)   # Print row number
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))
        print (type (cell_obj))
        #cell_value[col_idx] = str(cell_obj.split(":"))

print ("dict print")

#for key, value in row_name.items():
#    print (key, value)

for key, value in cell_value.items():
    print (key, value)