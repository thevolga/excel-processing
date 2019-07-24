from docx import Document
import re
import os
import xlrd
import xlwt
import openpyxl
from xlrd.book import Book
from xlrd.sheet import Sheet
from xlrd.sheet import Cell
from xlutils.copy import copy  # 导入excel复制模块
import xlsxwriter

excel_path = r'C:\Users\lenovo\Downloads\打卡记录表.xlsx'
checkin_info = xlrd.open_workbook(excel_path)  # 打开excel
table = checkin_info.sheets()[0] # 读取第一个sheet
nrows = table.nrows  # 获取表的行数 19
ncols = table.ncols  # 获取表的列数 3
# print(nrows,ncols)


def generate_all(name, frequency):
    workbook = xlsxwriter.Workbook('test_data.xlsx')
    worksheet = workbook.add_worksheet()
    row = 1
    col = 0
    for i in range(len(name)):
        worksheet.write(row, col, name[i])
        row += 1
    
    row = 1
    col = 1
    for i in range(len(frequency)):
        worksheet.write(row, col, frequency[i])
        row += 1
    
    workbook.close()


result_dic={}

def count_freq(result):
    
    for item_str in result:
        if item_str not in result_dic:
            result_dic[item_str]=1
        else:
            result_dic[item_str]+=1
    # print(result_dic)

frequency_list = []
for i in range(1, nrows):#第0行为表头
    alldata = table.row_values(i)#循环输出excel表中每一行，即所有数据
    name = alldata[0]#取出表中第二列数据
    frequency_list.append(name)

# print(frequency_list)
count_freq(frequency_list)
name = []
for key in result_dic.keys():
    #    print(key+':'+frequency_list[key])
    name.append(key)

frequency = []
for value in result_dic.values():
    frequency.append(value)



print(frequency_list)
generate_all(name, frequency)



