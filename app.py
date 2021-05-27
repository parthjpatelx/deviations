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
    for j in range(sheet.ncols):
        if j == 1:
            #cast the date string into an integer
            xl_date = int(sheet.cell_value(i,1))

            #convert integer into a datetime object
            datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
            date_object = datetime_date.date()

            #convert datetime object into string
            string_date = date_object.isoformat()
            sheet1.write(i, j, f"dope : {string_date}")
        else:
            sheet1.write(i,j, "hello")


writing_wb.save('deviations.xls')


#references:
##https://www.kite.com/python/answers/how-to-convert-an-excel-date-to-a-string-in-python
#https://www.geeksforgeeks.org/reading-excel-file-using-python/
