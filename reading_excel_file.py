#!/usr/bin/env python
"""
Your task is as follows:
- read the provided Excel file
- find and return the min, max and average values for the COAST region
- find and return the time value for the min and max entries
- the time values should be returned as Python tuples

Please see the test function for the expected return format
"""

import xlrd
from zipfile import ZipFile
datafile = "2013_ERCOT_Hourly_Load_Data.xls"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    hour_end_list = []
    coast_list = []
    sheet_data = [[sheet.cell_value(r, col)
                for col in range(sheet.ncols)]
                    for r in range(sheet.nrows)]
    for row in range(sheet.nrows):
        if row != 0:
            hour_end_list.append(xlrd.xldate_as_tuple(sheet.cell_value(row, 0), 0))
            coast_list.append(sheet.cell_value(row, 1))


    max_load_index = coast_list.index(max(coast_list))
    min_load_index = coast_list.index(min(coast_list))
    avgcoast = sum(coast_list) / float(len(coast_list))

    data = {
            'maxtime': hour_end_list[max_load_index],
            'maxvalue': coast_list[max_load_index],
            'mintime': hour_end_list[min_load_index],
            'minvalue': coast_list[min_load_index],
            'avgcoast': avgcoast
    }
    return data


def test():
    open_zip(datafile)
    data = parse_file(datafile)

    assert data['maxtime'] == (2013, 8, 13, 17, 0, 0)
    assert round(data['maxvalue'], 10) == round(18779.02551, 10)


test()
