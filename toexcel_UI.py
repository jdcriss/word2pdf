import datetime

from pdf2docx import Converter
import PySimpleGUI as sg
import os.path
import pdfplumber
import pandas as pd

def pdf2word(file_path):
    file_name = file_path.split('.')[0]
    doc_file = f'{file_name}.docx'
    p2w = Converter(file_path)
    p2w.convert(doc_file, start=0, end=None)
    p2w.close()
    return doc_file

def pdf2exl(path):  # 此函数直接循环提取PDF里面各个页面的表格
    pdf = pdfplumber.open(path)
    lens = len(pdf.pages)
    for i in range(0, lens):
        print('正在输出第', str(i + 1), '页表格')
        p0 = pdf.pages[i]
        try:
            res_df = pd.DataFrame()
            tables = p0.extract_tables()
            for table in tables:
                # 单元格清洗
                for row in table:
                    row_list = [cell.replace('\n', '') if cell else None for cell in row]
                    row_list = [row_list]
                    res_df = res_df.append(row_list, ignore_index=True)
            name = path.split('.')[0]
            excel_path = name+'-'+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+".xlsx"
            if not os.path.exists(excel_path):
                res_df.to_excel(excel_path, index=False, header=True)
            else:
                with pd.ExcelWriter(excel_path, mode='a') as i:
                    res_df.to_excel(i, index=False, header=True)
            print('文件保存位置 : ', excel_path)
            pdf.close()

        except Exception as e:
            print(e)
            pass
        # print('目前内存占用率是百分之',str(ps.virtual_memory().percent),'    第',str(i+1),'页输出完毕')

if __name__ == '__main__':
    # 选择主题
    sg.theme('DarkAmber')
    layout = [
        [sg.Text('pdfToExcel', font=('微软雅黑', 12)),
         sg.Text('', key='filename', size=(50, 1), font=('微软雅黑', 10))],
        [sg.Output(size=(80, 10), font=('微软雅黑', 10))],
        [sg.FilesBrowse('选择文件', key='file', target='filename'), sg.Button('开始转换'), sg.Button('退出')]]
    # 创建窗口
    window = sg.Window("pdf表格转excel", layout, font=("微软雅黑", 15), default_element_size=(50, 1))
    # 事件循环
    while True:
        # 窗口的读取，有两个返回值（1.事件；2.值）
        event, values = window.read()
        print(event, values)
        if event == "开始转换":
            if values['file'] and values['file'].split('.')[1] == 'pdf':
                pdf2exl(values['file'])
                print('文件个数 ：1')
                print('\n' + '转换成功！' + '\n')
            elif values['file'] and values['file'].split(';')[0].split('.')[1] == 'pdf':
                files_list = values['file'].split(';')
                lens = len(files_list)
                print('文件个数 ：{}'.format(lens),'\n')
                for i in range(lens):
                    sg.one_line_progress_meter('进度条', i+1, lens, '当前进度', orientation='h', bar_color=('#AAFFAA',
                                                                                                   '#FFAAFF'))
                    print('当前文件 ：', files_list[i])
                    pdf2exl(files_list[i])
                    print('转换成功！' + '\n')
            else:
                print('请选择pdf格式的文件哦!')
        if event in (None, '退出'):
            break

    window.close()
