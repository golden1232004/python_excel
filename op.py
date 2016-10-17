
#-*-coding:utf-8-*-

import os
import sys
import argparse
import xlsxwriter
import xlrd
from lxml import etree


def read_excel(filename):
    if not os.path.isfile(filename):
        print "exel file %s is not exist"%filename
        return None
    wb = xlrd.open_workbook(filename)
    #table = wb.sheet_by_index(0)
    sheets = wb.sheets()
    return sheets


def save_excel(filename):
    workbook = xlsxwriter.Workbook(date+'.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0, 11, 14)


def main(args):
    filename = args.filename
    sheets = read_excel(filename)
    for s in sheets:
        print 'Sheet:',s.name
        for row in range(s.nrows):
            values = []
            for col in range(s.ncols):
                v = s.cell(row, col).value
                if isinstance(v, float):
                    v = str(v)
                values.append(v)
            print ','.join(values)

        print
    pass

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('--filename',
                        help='excel file path',
                        type=str,
                        default='')
    args = parser.parse_args()
    main(args)
