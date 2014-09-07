# coding=gb2312

__author__ = 'ziqli'
import os
import re
from xlrd import open_workbook
from xlwt import *

in_title_list = [u'序号', u'姓名', u'工号', u'入职日期', u'离职日期']
out_tittle_list = [u'序号', u'姓名', u'工号', u'工作单位', u'入职日期', u'合同状态', u'离职日期', u'金额/元', u'备注']
date_tittle_list = [u'入职日期', u'离职日期']
key_col_name = u'第三方'
xls_to_be_processed = u'联胜7月外包费用核对-承天（8-14）.xls'
current_company = u'联胜'


alignment = Alignment()
alignment.horz = Alignment.HORZ_CENTER
common_style = XFStyle()
common_style.alignment = alignment

date_style = XFStyle()
date_style.num_format_str = 'YYYY/MM/DD'
date_style.alignment = alignment

sheet_name_from_value = {}

cur_dir = os.path.dirname(__file__)
wb = open_workbook(os.path.join(cur_dir, xls_to_be_processed))
#print wb.encoding
#print wb.codepage
out_wb = Workbook(encoding='utf-16le')

sheet = wb.sheet_by_index(0)
cols = sheet.ncols
rows = sheet.nrows


def get_tittle_row():
    titles = []
    for column in range(0, cols):
        titles.append(sheet.col_values(column, 0, 1)[0])
    return titles


def get_column_idx_by_tittle(title_str):
    # to do: move get_tittle_row() out
    tittles = get_tittle_row()
    return list.index(tittles, title_str)


def get_cell_by_tittle(row_idx, title_str):
    tittles = get_tittle_row()
    col_idx = list.index(tittles, title_str)
    content = sheet.cell(row_idx, col_idx).value
    return content


def get_column_ids_by_names(title_list):
    tittles = get_tittle_row()
    ids = [list.index(tittles, tittle) for tittle in title_list]
    return ids


def get_sheet_name_dict_from_column(title):
    idx = get_column_idx_by_tittle(title)
    valid_style = re.compile(ur'(\w+)\W*[\uff08]\W*(\w+)\W*[\uff09]', re.UNICODE)
    for row in range(1, rows):
        content = sheet.row_values(row, idx, idx + 1)[0]
        if type(content) == unicode:
            #print content
            matched_content = valid_style.match(content)
        else:
            matched_content = None
        if matched_content:
            sheet_name = "%s-%s%s" % (matched_content.group(1), current_company, matched_content.group(2))
            sheet_name_from_value[content] = sheet_name
    print sheet_name_from_value and "OK" or "NULL"
    return sheet_name_from_value

date_tittle_idx = get_column_ids_by_names(date_tittle_list)
column_index = get_column_idx_by_tittle(key_col_name)
for category in get_sheet_name_dict_from_column(key_col_name).keys():
    print category, sheet_name_from_value[category]
    out_sheet = out_wb.add_sheet(sheet_name_from_value[category])
    out_row_id = 0
    for row in range(rows):
        if row == 0 or sheet.cell(row, column_index).value == category:
            out_column_id = 0
            for col in get_column_ids_by_names(in_title_list):
                if col in date_tittle_idx:
                    if not out_row_id:
                        def get_width(num_characters):
                            return int((1+num_characters) * 256)
                        #2014/11/02 contains 11 num_chars
                        out_sheet.col(out_column_id).width = get_width(11)
                    out_sheet.row(out_row_id).write(out_column_id, sheet.cell(row, col).value, style=date_style)
                elif col == get_column_idx_by_tittle(u'序号') and out_row_id:
                    out_sheet.row(out_row_id).write(out_column_id, out_row_id, style=common_style)
                else:
                    out_sheet.row(out_row_id).write(out_column_id, sheet.cell(row, col).value, style=common_style)
                out_column_id += 1
            out_sheet.flush_row_data()
            out_row_id += 1

out_wb.save('output.xls')
