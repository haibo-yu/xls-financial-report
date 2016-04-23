#!/usr/bin/Python
# -*- coding: utf-8 -*-
# __author__ = "haibo"

import xlrd


def test():
    data = xlrd.open_workbook('demo.xlsx')  # 打开demo.xls
    sheet_names = data.sheet_names()  # 获取xls文件中所有sheet的名称
    for name in sheet_names:
        print name
    table = data.sheet_by_name(sheet_names[0])
    print table.nrows
    print table.ncols
    for row_value in table.row_values(0):
        print row_value
    # 循环行,得到索引的列表
    for rownum in range(table.nrows):
        print table.row_values(rownum)


if __name__ == "__main__":
    test()
