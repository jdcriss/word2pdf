import os.path

import pdfplumber
import pandas as pd
import time
from time import ctime
import psutil as ps
# import threading
import gc

pdf = pdfplumber.open(r"C:\Users\hml\Documents\WeChat Files\JD_criss\FileStorage\File\2022-11\RB03B20221130C.pdf")
N = len(pdf.pages)
print('总共有', N, '页')

def pdf2exl(i):  # 读取了第i页，第i页是有表格的，
    print('正在输出第', str(i + 1), '页表格')
    p0 = pdf.pages[i]
    try:
        res_df = pd.DataFrame()
        tables = p0.extract_tables()
        for table in tables:
            # print(table)
            # 单元格清洗
            for row in table:
                row_list = [cell.replace('\n','') if cell else None for cell in row]
                row_list = [row_list]
                # print(row_list)
                res_df = res_df.append(row_list, ignore_index=True)
                # print(res_df)
        excel_path = "表格.xlsx"
        if not os.path.exists(excel_path):
            res_df.to_excel(excel_path, index=False, header=True)
        else:
            with pd.ExcelWriter(excel_path, mode='a') as i:
                print("2:",res_df)
                res_df.to_excel(i, index=False,header=True)
        pdf.close()
        # df.to_excel(r"C:\Users\hml\Desktop\pdf\Model" + str(i + 1) + ".xlsx")
        # df.info(memory_usage='deep')
    except Exception as e:
        print(e)
        # print('第' + str(i + 1) + '页无表格，或者检查是否存在表格')
        pass

    # print('目前内存占用率是百分之',str(ps.virtual_memory().percent),'    第',str(i+1),'页输出完毕')
    # print('\n\n\n')
    # time.sleep(5)


def dojob1():  # 此函数  直接循环提取PDF里面各个页面的表格
    for i in range(0, N):
        pdf2exl(i)

dojob1()