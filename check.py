import xlsxwriter
import xlrd

file_location = "provide the path in local with file name and extension"
save_location = "E:\Assign\HCI\DataForTensorFlow\VCExtract\Mode2\ "

# creating object to read excel.
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)

# Creating a list with the selective column data you want to extract
velocity = [sheet.cell_value(row, 4) for row in range(sheet.nrows)]
lanePos = [sheet.cell_value(row, 6) for row in range(sheet.nrows)]
speed = [sheet.cell_value(row, 10) for row in range(sheet.nrows)]
print(len(velocity), len(lanePos))

# testing if the header is correct
print(velocity[0])
print(lanePos[0])
print(speed[0])

# initializing object for writing with the name you want to save it
workbook = xlsxwriter.Workbook(save_location + 'VCExtractSt30mod3.xlsx')
sheet = workbook.add_worksheet()

# writing all the row values for a column recursively
for i in range(len(velocity)):
    sheet.write_string(i, 0, str(velocity[i]))
    sheet.write_string(i, 1, str(lanePos[i]))
    sheet.write_string(i, 2, str(speed[i]))
                            
# This is an important statement as this will save the file and 
# will de-instantiate the object
workbook.close()
