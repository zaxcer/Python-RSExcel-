#!/usr/bin/env python3
#-*- coding: utf-8 -*-

import xlrd
import xlwt
import os
import sys
import re
 

# 从某Excel文件中提取特定行写入数组
class DetectStringFromExcel(object):

    def __init__(self, file_path, bad_guys):
        self.file = xlrd.open_workbook(file_path)
        self.sheets_num = self.file.nsheets - 2  # 获取sheets数量，
        self.bad_guys = bad_guys  # 需要探测的词汇
        self.results = []  # 返回结果为一个List, 其中元素也是List（某行内容）

    # 获得匹配结果，返回列表
    def run(self):
        for n in range(self.sheets_num):
            sheet = self.file.sheet_by_index(n)
            self.results.extend(self._detect(sheet))

        # 获得表头
        head = self.file.sheet_by_index(0).row_values(0)
        head.insert(0, '项目')
        self.results.insert(0, head)

        return self.results

    def _detect(self, sheet):
        # 先侦测'状态'在第几列!
        if '状态' and '下单日期' in sheet.row_values(0):
            col_status = sheet.row_values(0).index('状态')
            col_date = sheet.row_values(0).index('下单日期')
        else:
            col_status = 11  # 没找到就猜一个列数
            col_date = 7

        temp_list = []
        for row in range(1, sheet.nrows):
            status = sheet.col(col_status)[row].value  # 获取交货情况
            if status in self.bad_guys:
                the_bad_row = sheet.row_values(row)
                
                for d in range(col_date, col_date+4):
                    if the_bad_row[d]:
                        date = xlrd.xldate.xldate_as_datetime(the_bad_row[d], 0)  # 转换为看得懂的时间
                        the_bad_row[d] = date.strftime('%Y/%m/%d')

                the_bad_row.insert(0, sheet.name)  # 在行开头添加sheet name
                temp_list.append(the_bad_row)
            


        return temp_list


# 将内容写进新Excel
class WriteToNewExcel(object):

    def __init__(self, list_to_write, path='/Users/Zhangchi/Desktop/Results.xls'):
        self.list_to_write = list_to_write
        self.file = xlwt.Workbook()
        self.sheet = self.file.add_sheet('跟踪结果', cell_overwrite_ok=True)
        self.path = path
        self.rows = len(self.list_to_write)

    def writein(self):

        # 按行顺序写入数据
        for row in range(self.rows):
            self._writeonerow(row+1, self.list_to_write[row])
        # 设置表的列宽
        self.sheet.col(0).width = 3000
        self.sheet.col(2).width = 3000
        self.sheet.col(3).width = 4500
        self.sheet.col(4).width = 6000
        self.sheet.col(8).width = 3000
        self.sheet.col(9).width = 3000
        self.sheet.col(10).width = 3000
        self.sheet.col(11).width = 3000
        self.sheet.col(12).width = 3000

        self.file.save(self.path)

    def _writeonerow(self, whichrow, whattowrite):
        cols = len(whattowrite)
        for col in range(cols):
            self.sheet.write(whichrow, col, whattowrite[col])

# 鉴别所需Excel
def detect_name(allfilelist):
    filelist = []
    for name in allfilelist:
        if re.search(r'^(EMCxxAxxx采购清单).*', name):
            filelist.append(name)
    return filelist

if __name__ == '__main__':

    bad_guys=('催促交货', '未订货', '未交货', '延迟交货')

    print('请将脚本放入Excel所在文件夹...')
    print('请将Excel命名为"EMCxxAxxx采购清单..."')
    print('请务必保证Excel内探测内容在"状态"字样下...')
    print('默认探测内容为:', bad_guys)
    print('------------------')
    ok = input('按Enter键以继续...')

    # 全部文件
    abspath = sys.path[0]
    
    allfilelist = os.listdir(abspath)
    
    # 鉴别所需文件
    filelist = detect_name(allfilelist)
    print('采购清单Excel文件为: ')
    for excel in filelist:
        print(excel)

    # 匹配特定行，合并写入暂存列表
    newlist = []
    for excel in filelist:
        excel = abspath + '/' + excel  # 绝对路径 
        f = DetectStringFromExcel(excel, bad_guys)
        detected_list = f.run()
        newlist.extend(detected_list)
    print('匹配字段完成...')

    # 从暂存列表写入新Excel
    newpath = abspath + '\Results.xls'
    wt = WriteToNewExcel(newlist, path=newpath)
    wt.writein()
    print('写入新文件完成...')

    #打开新文件Excel
    print('正在打开新文件...')
    os.system(newpath)
