# Reading an excel file using Python
import xlrd, datetime
import xlwt
from xlwt import Workbook

def write_deviation(last_name, date_string, cycle_name, IRB_number):
    my_deviation = f"Patient {last_name} was assessed on {date_string} for {cycle_name} on {IRB_number}. At this visit, the patient was dispensed drug. However, since the patient was unabel to travel to clinic due to circumstances surrounding COVID-19, the drug was shipped to them. This event constitues a protocol deviation."

    return my_deviation



# Give the location of the file- https://www.geeksforgeeks.org/reading-excel-file-using-python/
loc = ('/Users/parth/drugshipments.xls')


# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


#create a new workbook to write
writing_wb = Workbook()
sheet1 = writing_wb.add_sheet('Deviations')

all_deviations= []

IRB_number = "19-257"
LAST_COLUMN = 6

for i in range(sheet.nrows):
    MRN = cycle_name = last_name = string_date = ""
    for j in range(sheet.ncols):
        #if the column represents a date, we need to treat it differently
        if j == 1:
            #cast the date string into an integer
            xl_date = int(sheet.cell_value(i,1))

            #convert integer into a datetime object
            datetime_date = xlrd.xldate_as_datetime(xl_date, 0)
            date_object = datetime_date.date()

            #convert datetime object into string
            string_date = date_object.isoformat()
            sheet1.write(i, j, string_date)
        else:
            cell_value = sheet.cell_value(i,j)
            if j == 0:
                MRN = cell_value
            if j == 2:
                cycle_name = cell_value
            if j == 3:
                last_name = cell_value

            sheet1.write(i,j, cell_value)

    deviation = write_deviation(last_name, string_date, cycle_name, IRB_number)
    all_deviations.append(deviation)

number_deviations = len(all_deviations)

for n in range(number_deviations):
    current_deviation = all_deviations[n]
    sheet1.write(n,LAST_COLUMN, current_deviation)




writing_wb.save('deviations2.xls')




#references:
##https://www.kite.com/python/answers/how-to-convert-an-excel-date-to-a-string-in-python
#https://www.geeksforgeeks.org/reading-excel-file-using-python/

