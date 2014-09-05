# coding=gb2312
__author__ = 'ziqli'
import os
from xlrd import open_workbook
import xlwt

tittle_list = [u'名字', u'单位', u'年龄']
key_col_name = u'单位'

cur_dir = os.path.dirname(__file__)
wb = open_workbook(os.path.join(cur_dir, u'职工信息.xls'))
out_wb = xlwt.Workbook(encoding='gb2312')

sheet = wb.sheet_by_name(u'信息')
cols = sheet.ncols
rows = sheet.nrows


def get_tittle_row():
    titles = []
    for column in range(0, cols):
        titles.append(sheet.col_values(column, 0, 1)[0])
    return titles


def get_column_idx_by_tittle(tittle_str):
    tittles = get_tittle_row()
    return list.index(tittles, tittle_str)


def get_column_ids_by_names(tittle_list):
    tittles = get_tittle_row()
    ids = [list.index(tittles, tittle) for tittle in tittle_list]
    return ids


def get_unique_value_from_column(tittle):
    idx = get_column_idx_by_tittle(tittle)
    values = list()
    for row in range(1, rows):
        values.append(sheet.row_values(row, idx, idx + 1)[0])
    return set(values)

column_index = get_column_idx_by_tittle(key_col_name)
for category in get_unique_value_from_column(key_col_name):
    out_sheet = out_wb.add_sheet(category)
    out_row_id = 0
    for row in range(rows):
        if row == 0 or sheet.cell(row, column_index).value == category:
            out_column_id = 0
            for col in get_column_ids_by_names(tittle_list):
                out_sheet.row(out_row_id).write(out_column_id, sheet.cell(row, col).value)
                out_column_id += 1
            out_sheet.flush_row_data()
            out_row_id += 1

out_wb.save('output.xls')
