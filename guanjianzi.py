#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：pythonProject
@File    ：gjianzi.py
@IDE     ：PyCharm
@Author  ：Suykm
@Date    ：2023-01-23 10:52
'''
import os
import pandas as pd
from openpyxl import Workbook
import glob
import warnings

print("                                 \n"
      "             关键字查询            \n"
      "                                 \n"
      "     ██╗   ██╗██╗  ██╗███╗   ███╗\n"
      "     ╚██╗ ██╔╝██║ ██╔╝████╗ ████║\n"
      "      ╚████╔╝ █████╔╝ ██╔████╔██║\n"
      "       ╚██╔╝  ██╔═██╗ ██║╚██╔╝██║\n"
      "        ██║   ██║  ██╗██║ ╚═╝ ██║\n"
      "        ╚═╝   ╚═╝  ╚═╝╚═╝     ╚═╝\n"
      "                          Version 0.02 \n")


def data_s(ak, gjz):
    directory_path = ak
    id_number = gjz
    global data_count
    new_file = Workbook()
    sheet = new_file.active

    excel_files = []
    for file in glob.glob(directory_path + '**/*.xlsx', recursive=True):
        excel_files.append(file)
        nb = len(excel_files)
    ks = "扫描到 " + str(nb) + " 个excel文件"
    print('++' + ks + '++')
    for num in excel_files:
        #print(num)
        try:
            excel_file = pd.read_excel(num)
            column_names = excel_file.columns
            warnings.filterwarnings("ignore")
            # name ='姓名'
            # id_number = '身份证号'
            for column_name in column_names:
                # if column_name == name or column_name == id_number:
                if column_name == id_number:
                    data_count = excel_file.shape[0]
                    # print("关键字：" + column_name)
                    # print("关键字总条数：" + str(data_count))
                    # print("文件所在位置：" + num)
                    file_name = os.path.basename(num)
                    file_name = os.path.split(num)[-1]
                    file_stats = os.stat(file_name)
                    print(f'文件大小 {file_stats.st_size / (1024 * 1024)}')
                    file_data = "文件名：" + file_name, "关键字：" + column_name, "文件总条数：" + str(data_count), "文件所在位置：" + num
                    print(file_data)
        except:
            print('----UserWarning兼容性警告----')


if __name__ == '__main__':
    print("请输入盘符：(如 D:/ )")
    ak = input()
    print("请输入关键字：(如 姓名 或者 电话)")
    gjz = input()
    data_s(ak,gjz)
