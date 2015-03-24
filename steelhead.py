import xlrd

filename = "data.xlsx"

COL_DATA = 0
COL_TIME = 1

print "Loading \"{0}\"".format(filename)
workbook = xlrd.open_workbook(filename)
try:
    sheet = workbook.sheet_by_name([s for s in workbook.sheet_names() if "reduced" in s][0])
except IndexError:
    print "Cannot find \"reduced\" worksheet."
    sys.quit()

width = 5
i = width
maxes = []
while i < sheet.nrows:
    i += 1
    curr_row = sheet.row(i)
    next_row = sheet.row(i + 1)
    prev_row = sheet.row(i - 1)

    if curr_row[COL_DATA].value > prev_row[COL_DATA].value and curr_row[COL_DATA].value > next_row[COL_DATA].value:
        print prev_row[COL_DATA]
        print curr_row[COL_DATA]
        print next_row[COL_DATA]
        print "-"*80
