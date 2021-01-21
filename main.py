import xlrd
import re
import xlwt
from xlutils.copy import copy
import os

def make_sample_excle(path):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_workbook.save(os.getcwd() + "\\" + 'Excel_test.xls')  # 保存工作簿


#打开表格
data = xlrd.open_workbook('conditionstr_chinese.xlsx')
table = data.sheets()[0] #获取表
nrows = table.nrows  #获取该sheet中的有效行数
ncols1 = table.col_values(0)[1:]#获取第一列


list1 = []
for c in ncols1:
    matched_c = re.match(r'(.*?), *',c)
    if matched_c != None:
        list1.append(c)

print(len(list1))
print('带逗号的数据'+ str(list1))

make_sample_excle('conditionstr_chinese.xlsx')# 新建主表复制表格

workbook = xlrd.open_workbook('Excel_test.xls')  # 打开主表复制表格
new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格

for item in list1:#遍历有逗号的condition
    count = 0
    for i in ncols1:#遍历总列表
        count += 1
        if ',' not in i and i in item:#如果总列表中的单词没有逗号,并且没有逗号的单词在有逗号的condition中
            new_worksheet.write(0, 1, label="修改后的condition")
            new_worksheet.write(count, 1, i)  # 追加写入数据
            print("xls格式表格【追加】写入数据成功！")


new_workbook.save('Excel_test.xls')  # 保存工作簿

