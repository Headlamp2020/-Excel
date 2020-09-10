"""文件夹内多个EXCEL表格（）多个sheet合并方法:
1、所有表格共同sheet合并(concat）成一个大表格 sheet_in_one
2、所有表格共同的sheet放在一个表格的多个sheet all_in_one
Created on 20200910
@author: Gao dengke """

import pandas as pd
import os

folder = r'C:\Users\gaodengke\Desktop\待合并'
file_list = os.listdir(folder)
record_path = r'C:\Users\gaodengke\Desktop\合并后\日志.xlsx'
report_path = r'C:\Users\gaodengke\Desktop\合并后\体检报告.xlsx'
all_path = r'C:\Users\gaodengke\Desktop\合并后\整个报告.xlsx'


def all_in_one():
    df_record = pd.DataFrame()
    df_report = pd.DataFrame()
    with pd.ExcelWriter(all_path) as writer:
        for file in file_list:
            file_path = os.path.join(folder, file)
            df_report1 = pd.read_excel(file_path, '体检报告').dropna(how='all')
            df_record1 = pd.read_excel(file_path, '日志').dropna(how='all')
            df_report = pd.concat([df_report, df_report1]).drop_duplicates()
            df_record = pd.concat([df_record, df_record1]).drop_duplicates()
            df_report.to_excel(writer, '体检报告')
            df_record.to_excel(writer, '日志')


def sheet_in_one():
    dict = {record_path: '日志', report_path: '体检报告'}
    for path, sheet_name in dict.items():
        with pd.ExcelWriter(path) as writer:
            for file in file_list:
                file_path = os.path.join(folder, file)
                df_record = pd.read_excel(file_path, sheet_name=sheet_name).dropna(how='all')
                df_record.to_excel(writer, sheet_name=file[0:5])


if __name__ == '__main__':
    all_in_one()
    sheet_in_one()
