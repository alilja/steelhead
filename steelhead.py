import xlrd

filename = "data.xlsx"

COL_DATA = 0
COL_TIME = 1

workbook = xlrd.open_workbook(filename)
try:
    sheet = workbook.sheet_by_name([s for s in workbook.sheet_names() if "reduced" in s][0])
except IndexError:
    print "Cannot find \"reduced\" worksheet."
    sys.quit()

i = 0
width = 1
maxes = []
while i < sheet.nrows:
    i += 1
    curr_row = sheet.row(i)
    next_row = sheet.row(i + 1)
    prev_row = sheet.row(i - 1)

    if curr_row[COL_DATA] > prev_row[COL_DATA] and curr_row[COL_DATA] > next_row[COL_DATA]:
        print curr_row
