#-------------------------------------------------------------------------------
# Name:        main
# Purpose:
#
# Author:      Lai
#
# Created:     25/11/2015
#-------------------------------------------------------------------------------
import xlwt, xlrd, openpyxl
from xlutils.copy import copy
from template_search import Workbook_Template

def main():
    path = r"D:\python task\draw_data_and_post\post\tmp.xls"
##    rb = xlrd.open_workbook(path,formatting_info=True)
##    while s in wb_s.sheets():
##    wb = copy(rb)
    a = Workbook_Template(path)

    b = a.get_fill_pos('TX','5G','Standard')
    for i in b[0]:
        for j in b[0][i].values():
            print i,j
    print b[1]
##    wb.save(r"D:\python task\a.xls")

if __name__ == '__main__':
    main()
