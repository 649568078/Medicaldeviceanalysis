import xlrd
import re
import xlwt
from xlutils.copy import copy
import os

workbook = xlrd.open_workbook('待剔除condit.xls')  # 打开主表复制表格
table = workbook.sheets()[0]  # 获取表
nrows = table.nrows  # 获取该sheet中的有效行数
ncols1 = table.col_values(0)[1:]  # 获取第一列

new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格

style = xlwt.easyxf('pattern: pattern solid, fore_colour ice_blue')


for item in list1:  # 遍历有逗号的condition
    item_withoutcomma = item[1].split(',', 1)[0]
    # print(item_withoutcomma)
    for i in ncols1:  # 遍历总列表
        if item_withoutcomma == i:  # 如果总表里有这么一个单词和逗号condition分割后一样
            new_worksheet.write(item[0] - 1, 1, i, style)  # 用总列表中的单独单词
            print(item[1] + '找到原版对应')
            break
        else:
            new_worksheet.write(item[0] - 1, 1, item[1].split(',', 1)[0])  # 使用切分后的

new_workbook.save('Excel_test.xls')  # 保存工作簿