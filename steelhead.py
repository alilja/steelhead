import sys
import os
from os import path

import xlrd


class DataFile(object):
    COL_DATA = 0
    COL_TIME = 1

    def __init__(self, filename):
        self.filename = filename[:filename.find(".")]

        print "Loading \"{0}\"".format(filename)
        workbook = xlrd.open_workbook(filename)
        try:
            self.sheet = workbook.sheet_by_name([s for s in workbook.sheet_names() if "reduced" in s][0])
        except IndexError:
            raise IndexError("Cannot find \"reduced\" worksheet.")

    def find_spikes(self, width=5):
        i = width
        maxes = []
        time_series = []
        while i < self.sheet.nrows - width - 1:
            i += 1
            curr_row = self.sheet.row(i)

            # check to see if there's a break in the data
            next_time = self.sheet.row(i + 1)[DataFile.COL_TIME].value
            if next_time - curr_row[DataFile.COL_TIME].value > 8000:
                time_series.append(maxes)
                maxes = []

            next_max = max([self.sheet.row(i + n + 1)[DataFile.COL_DATA].value for n in range(width)])
            prev_max = max([self.sheet.row(i - n - 1)[DataFile.COL_DATA].value for n in range(width)])

            if curr_row[DataFile.COL_DATA].value > prev_max and curr_row[DataFile.COL_DATA].value > next_max:
                maxes.append((curr_row[DataFile.COL_DATA], curr_row[DataFile.COL_TIME]))
        return time_series

    def build_spreadsheet(self, time_series, output_dir=None):
        if not output_dir:
            output_dir = "output"
        total = len(time_series)
        for i, time in enumerate(time_series):
            output = ["Compression Depth, Timestamp"]
            for data_row in time:
                output.append(", ".join([
                    str(data_row[DataFile.COL_DATA].value),
                    str(data_row[DataFile.COL_TIME].value)
                ]))

            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            filename = "{0}/{1}_{2}_{3}.csv".format(output_dir, self.filename, i, total)
            with open(filename, "w") as f:
                f.write("\n".join(output))
                print "Outputted data to {0}.".format(filename)

if __name__ == "__main__":
    from optparse import OptionParser
    from re import search
    parser = OptionParser(usage="usage: %prog [options]")
    parser.add_option(
        "-i",
        "--input",
        help=(
            "The input directory containing the files you"
            "want to process."
        )
    )
    parser.add_option(
        "-o",
        "--output",
        help="The output directory where processed files"
        "should go."
    )
    (options, args) = parser.parse_args()

    if not options.input:
        options.input = "."
    files = [f for f in os.listdir(options.input) if path.isfile(f)]
    print files
    for file in files:
        if not search(r"\.~.+\.xls", file) and ".xls" in file:  # ignore temporary files
            data = DataFile(file)
            data.build_spreadsheet(data.find_spikes(), output_dir=options.output)
