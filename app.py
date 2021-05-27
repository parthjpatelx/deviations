# Reading an excel file using Python
import xlrd, datetime

# Give the location of the file
loc = ('/Users/parth/drugshipments.xls')

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
print(sheet.cell_value(0, 0))

for i in range(sheet.nrows):
    #https://www.kite.com/python/answers/how-to-convert-an-excel-date-to-a-string-in-python
    xl_date = int(sheet.cell_value(i,1))
    datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
    date_object = datetime_date.date()
    string_date = date_object.isoformat()
    print(string_date)
