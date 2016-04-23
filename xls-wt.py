#!/usr/bin/Python
# -*- coding: utf-8 -*-
# __author__ = "haibo"

import xlwt


def test():
    file = xlwt.Workbook()  # 注意这里的Workbook首字母是大写
    # table = file.add_sheet(u'写入测试表')  # 新建一个sheet
    # 如果对一个单元格重复操作，会引发
    # returns error:
    # Exception: Attempt to overwrite cell:
    # sheetname=u'sheet 1' rowx=0 colx=0
    # 所以在打开时加cell_overwrite_ok=True解决
    table = file.add_sheet(u'写入测试表', cell_overwrite_ok=True)
    table.write(0, 0, 'test')  # 写入数据table.write(行,列,value)
    # 另外，使用style
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = 'Times New Roman'
    font.bold = True
    style.font = font  # 为样式设置字体
    table.write(0, 0, 'some bold Times text', style)  # 使用样式

    file.save('demo-wt.xls')  # 保存文件

if __name__ == "__main__":
    test()
