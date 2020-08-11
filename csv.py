
import xlrd
def get_table(file_path):
    #file_path = "Swipe Records-2020-08-11_164350_68.xls"
    sheet = "ExelData"
    xl_workbook  = xlrd.open_workbook(file_path)
    xl_sheet = xl_workbook.sheet_by_index(0)
    print ('Sheet name: %s' % xl_sheet.name)

    # Pull the first row by index
    #  (rows/columns are also zero-indexed)

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

    table = {}                   # main dict
                 # temp dict

    num_cols = xl_sheet.ncols   # Number of columns
    i = 0
    for row_idx in range(0, xl_sheet.nrows):    # Iterate through rows
        #
        #print ('Row: %s' % row_idx)   # Print row number
        cell_value = {} 
        for col_idx in range(0, num_cols):  # Iterate through columns
            cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
            #print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))
            #print ('-'*40)
            cell_value[row_name[col_idx]] = cell_obj.value #add cell value to temp dict  
            if col_idx==7 :
                if col_idx > 0 :    # don't write first row
                    table[i] = cell_value # than put that vaules to the main dictionary
                    i=i+1    
    
    return table





def main():
    file_path = "Swipe Records-2020-08-11_164350_68.xls"
    table = get_table(file_path)
    for key, value in table.items():
            print (key, value)

if __name__ == "__main__":
    main()        
