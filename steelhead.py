import sys
import os
from os import path

import xlrd

from pprint import pprint


class DataFile(object):
    COL_DATA = 0
    COL_TIME = 1

    def __init__(self, filename, verbose=False):
        print "Loading \"{0}\"".format(filename)
        workbook = xlrd.open_workbook(filename)
        try:
            self.sheet = workbook.sheet_by_name([s for s in workbook.sheet_names()][1])
        except IndexError:
            raise IndexError("Cannot find second worksheet.")

        filename = path.basename(filename)
        self.filename = filename[:filename.find(".")]

        self.verbose = verbose

    def find_spikes(self, width=5):
        i = width
        maxes = []
        time_series = []
        while i < self.sheet.nrows - width - 1:
            i += 1
            curr_row = self.sheet.row(i)

            # check to see if there's a break in the data
            next_time = self.sheet.row(i + 1)[DataFile.COL_TIME].value
            skip_check = next_time - curr_row[DataFile.COL_TIME].value
            if skip_check > 5000:
                if maxes != []:
                    time_series.append(maxes)
                    maxes = []

            next_max = max([self.sheet.row(i + n + 1)[DataFile.COL_DATA].value for n in range(width)])
            prev_max = max([self.sheet.row(i - n - 1)[DataFile.COL_DATA].value for n in range(width)])

            if curr_row[DataFile.COL_DATA].value > prev_max and curr_row[DataFile.COL_DATA].value > next_max:
                maxes.append((curr_row[DataFile.COL_DATA].value, curr_row[DataFile.COL_TIME].value))
        time_series.append(maxes)
        return time_series

    def find_averages(self, time_series, length=15000):
        all_averages = []
        for data in time_series:
            print data
            i = 0
            time_data = [stamp for depth, stamp in data]
            averages = []
            while i < len(data):
                target_time = time_data[i] + length
                try:
                    end = DataFile._find_closet_time(time_data[i:], target_time)
                    average_section = data[i: i + end]
                except IndexError:
                    average_section = data[i: -1]


                total = 0
                for depth, stamp in average_section:
                    total += depth

                if len(average_section) == 0:
                    break

                averages.append((average_section[0][1], (average_section[0][1] - time_data[0])/ 1000, total / len(average_section)))
                i += end
            if len(average_section) > 0:
                averages.append((average_section[-1][1], (average_section[-1][1] - time_data[0])/ 1000, 0))
            all_averages.append(averages)
            print "="*60
        return all_averages


    @staticmethod
    def _find_closet_time(data, time):
        for i, num in enumerate(data):
            if num < time:
                if data[i + 1] > time:
                    return i
            elif num == time:
                return i

    def build_spreadsheet(self, time_series, header, output_dir=None):

        total = len(time_series)
        for i, time in enumerate(time_series):
            output = [header]
            for data_row in time:
                output.append(", ".join([str(item) for item in data_row]))

            if not output_dir:
                output_dir = "output"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            filename = "{0}/{1}_{2}_{3}.csv".format(output_dir, self.filename, i + 1, total)
            with open(filename, "w") as f:
                f.write("\n".join(output))
                print "Outputted data to {0}.".format(filename)

    def build_spreadsheet_all(self, time_series, header, output_dir):
        output = [header]
        for cycle in time_series:
            for averages in cycle:
                output.append(", ".join([str(item) for item in averages]))
            output.append(","*len(averages))

        if not output_dir:
            output_dir = "output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        filename = "{0}/{1}_ALL.csv".format(output_dir, self.filename,)
        with open(filename, "w") as f:
            f.write("\n".join(output))






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
    parser.add_option(
        "-t",
        "--time",
        help="The time to measure averages over.",
        type="float",
        default=15000
    )
    (options, args) = parser.parse_args()

    print "\n\n\n"
    if not options.input:
        options.input = "."
    files = ["{0}/{1}".format(options.input, f) for f in os.listdir(options.input) if path.isfile("{0}/{1}".format(options.input, f))]
    for file in files:
        if not path.basename(file).startswith('.') and ".xls" in file:  # ignore temporary files
            data = DataFile(file, True)
            average_data = data.find_spikes()
            pprint(average_data)
            averages = data.find_averages(average_data, length=options.time)
            data.build_spreadsheet(
                averages,
                header="Timestamp, Time Since Cycle Start, Compression Depth",
                output_dir=options.output
            )

            data.build_spreadsheet_all(
                averages,
                header="Timestamp, Time Since Cycle Start, Compression Depth",
                output_dir=options.output
            )