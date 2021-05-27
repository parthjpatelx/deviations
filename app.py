# Reading an excel file using Python
import xlrd, datetime
import xlwt
from xlwt import Workbook


# Give the location of the file- https://www.geeksforgeeks.org/reading-excel-file-using-python/
loc = ('/Users/parth/drugshipments.xls')


# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


#create a new workbook to write
writing_wb = Workbook()
sheet1 = writing_wb.add_sheet('Deviations')


for i in range(sheet.nrows):
    #https://www.kite.com/python/answers/how-to-convert-an-excel-date-to-a-string-in-python
    xl_date = int(sheet.cell_value(i,1))
    datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
    date_object = datetime_date.date()

    string_date = date_object.isoformat()
    sheet1.write(i, 0, f"Date : {string_date}")



writing_wb.save('deviations.xls')
