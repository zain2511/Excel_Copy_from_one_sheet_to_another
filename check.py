import xlsxwriter
import xlrd

file_location = "E:\Assign\HCI\DataForTensorFlow\VCExtract\Mode2\Book1.xlsx"
save_location = "E:\Assign\HCI\DataForTensorFlow\VCExtract\Mode2\ "

workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)

#data = [sheet.cell_value(0, col) for col in range(sheet.ncols)]
velocity = [sheet.cell_value(row, 4) for row in range(sheet.nrows)]
lanePos = [sheet.cell_value(row, 6) for row in range(sheet.nrows)]
speed = [sheet.cell_value(row, 10) for row in range(sheet.nrows)]
steer = [sheet.cell_value(row, 11) for row in range(sheet.nrows)]
accel = [sheet.cell_value(row, 12) for row in range(sheet.nrows)]
brake = [sheet.cell_value(row, 13) for row in range(sheet.nrows)]
longAccel = [sheet.cell_value(row, 25) for row in range(sheet.nrows)]
headwayTime = [sheet.cell_value(row, 30) for row in range(sheet.nrows)]
headwayDist = [sheet.cell_value(row, 31) for row in range(sheet.nrows)]
print(len(velocity), len(lanePos))

print(velocity[0])
print(lanePos[0])
print(speed[0])
print(steer[0])
print(accel[0])
print(brake[0])
print(longAccel[0])
print(headwayTime[0])
print(headwayDist[0])

workbook = xlsxwriter.Workbook(save_location + 'VCExtractSt30mod3.xlsx')
sheet = workbook.add_worksheet()

for i in range(len(velocity)):
    sheet.write_string(i, 0, str(velocity[i]))
    sheet.write_string(i, 1, str(lanePos[i]))
    sheet.write_string(i, 2, str(speed[i]))
    sheet.write_string(i, 3, str(steer[i]))
    sheet.write_string(i, 4, str(accel[i]))
    sheet.write_string(i, 5, str(brake[i]))
    sheet.write_string(i, 6, str(longAccel[i]))
    sheet.write_string(i, 7, str(headwayTime[i]))
    sheet.write_string(i, 8, str(headwayDist[i]))
                            


#workbook.save(save_location + 'VCExtractSt1mod3.xlsx')
workbook.close()
#sheet = wb.sheets()[0]
#print(sheet.row(0))