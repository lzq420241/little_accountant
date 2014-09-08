# coding=gb2312

__author__ = 'ziqli'
# todo: 1. sequence_number start from 3 instead of 1
# todo: 2. in_bundle info should refer previous month data
# todo: 3. add three rows to each sheet
# todo: 4. add a summary sheet

import os
import re

from xlrd import *
from xlwt import *

from personnel import *
from calculator import get_first_day_of_last_month, is_date_in_last_month


in_title_list = [u'序号', u'姓名', u'工号', u'入职日期', u'离职日期']
out_title_list = [u'序号', u'姓名', u'工号', u'工作单位', u'入职日期', u'合同状态', u'离职日期', u'金额/元', u'备注']
date_title_list = [u'入职日期', u'离职日期']
key_col_name = u'第三方'
xls_to_be_processed = u'联胜7月外包费用核对-承天（8-14）.xls'
current_company = u'联胜'
col_len = len(out_title_list)
BUNDLE = 30

alignment = Alignment()
alignment.horz = Alignment.HORZ_CENTER

left_alignment = Alignment()
left_alignment.horz = Alignment.HORZ_LEFT

borders = Borders()
borders.left = Borders.THIN
borders.right = Borders.THIN
borders.top = Borders.THIN
borders.bottom = Borders.THIN

font = Font()
font.name = 'SimSun'
font.height = 200

big_font = Font()
big_font.name = 'SimSun'
# 16 point
big_font.height = 320

table_title_style = XFStyle()
table_title_style.alignment = alignment
table_title_style.font = big_font

tab_time_style = XFStyle()
tab_time_style.alignment = alignment
tab_time_style.font = font

common_style = XFStyle()
common_style.alignment = alignment
common_style.borders = borders
common_style.font = font

table_end_style = XFStyle()
table_end_style.alignment = left_alignment
table_end_style.font = font

date_style = XFStyle()
date_style.num_format_str = 'YYYY/MM/DD'
date_style.alignment = alignment
date_style.borders = borders
date_style.font = font

# key:category value:sheet_name
sheet_name_from_value = {}
# key:category value:personnel list
personnel_from_value = {}
# key:category value:list [<int>aboard_num, <int>dismiss_num, <str>calc_desc]
sheet_statistics = {}

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


def get_column_idx_by_title(title_str):
    # to do: move get_tittle_row() out
    tittles = get_tittle_row()
    return list.index(tittles, title_str)


def get_cell_by_title(row_idx, title_str):
    tittles = get_tittle_row()
    col_idx = list.index(tittles, title_str)
    content = sheet.cell(row_idx, col_idx).value
    return content


def get_column_ids_by_names(title_list):
    tittles = get_tittle_row()
    ids = [list.index(tittles, tittle) for tittle in title_list]
    return ids


def get_sheet_name_dict_from_column(title):
    idx = get_column_idx_by_title(title)
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
    #print sheet_name_from_value and "OK" or "NULL"
    return sheet_name_from_value


def get_table_title(sht_name):
    # todo: what if len of current_company is more than 2
    valid_style = re.compile(ur'\w+[\u002d](\w\w)(\w+)', re.UNICODE)
    matched_content = valid_style.match(sht_name)
    table_title = u"%s提成（%s）" % (matched_content.group(1), matched_content.group(2))
    return table_title


date_tittle_idx = get_column_ids_by_names(date_title_list)
key_column_index = get_column_idx_by_title(key_col_name)
sequence_col_idx = get_column_idx_by_title(u'序号')
aboard_col_idx = get_column_idx_by_title(u'入职日期')
dismission_col_idx = get_column_idx_by_title(u'离职日期')
status_col_idx = get_column_idx_by_title(u'离职方式')


def personnel_initializer():
    for category in get_sheet_name_dict_from_column(key_col_name).keys():
        preliminary_personnel_list = list()
        for row in range(1, rows):
            if sheet.cell(row, key_column_index).value == category:
                sequence = sheet.cell(row, sequence_col_idx).value
                aboard_date = sheet.cell(row, aboard_col_idx).value
                status = sheet.cell(row, status_col_idx).value
                dismission_date = sheet.cell(row, dismission_col_idx).value
                person = Personnel(sequence, aboard_date, status, dismission_date, False, False)
                if person.valid:
                    preliminary_personnel_list.append(person)

        personnel_list = list()
        bundle_list = update_bundle_info(preliminary_personnel_list)
        # todo: not elegant, re-design after learning Design Pattern.
        for p in preliminary_personnel_list:
            if p in bundle_list:
                person = Personnel(p.sequence, p.aboard_date, p.status, p.dismission_date, True)
            else:
                person = Personnel(p.sequence, p.aboard_date, p.status, p.dismission_date, False)
            personnel_list.append(person)

        sheet_statistics[category] = get_sheet_statistics(personnel_list)
        # print sheet_statistics[category]
        personnel_from_value[category] = personnel_list


def get_sheet_statistics(persons):
    assert type(persons) == list
    aboard_number = len(filter((lambda p: is_date_in_last_month(p.aboard_date)), persons))
    dismiss_number = len(filter((lambda p: is_date_in_last_month(p.dismission_date)), persons))
    calc_stat = list()
    com_list = [p.commission for p in persons]
    com_set = set(com_list)
    if 0.00 in com_set:
        com_set.remove(0.00)
    for c in com_set:
        cnt = com_list.count(c)
        calc_stat.append(u"%s人*%s元/人" % (cnt, c))
    return [aboard_number, dismiss_number, '+'.join(calc_stat)]


def update_bundle_info(persons):
    # search though persons,
    assert type(persons) == list
    aboard_month_ids = [p.aboard_month_id for p in persons if p.care_bundle]
    need_update_list = [idx for idx in set(aboard_month_ids) if aboard_month_ids.count(idx) >= BUNDLE]
    return [p for p in persons if p.aboard_month_id in need_update_list]


def draw_start_rows():
    tab_title = get_table_title(sheet_name_from_value[category])
    time_info = get_first_day_of_last_month()
    desc = u'所属时间：%s年%s月' % (time_info.year, time_info.month)
    # print tab_title
    out_sheet.write_merge(0, 0, 0, col_len - 1, tab_title, table_title_style)
    out_sheet.row(1).write(col_len - 1, desc, style=tab_time_style)
    out_sheet.flush_row_data()


def draw_ending_rows():
    aboard_num, dismiss_num, calc_info = sheet_statistics[category]
    out_sheet.write_merge(out_row_id + 2, out_row_id + 2, 0, col_len - 1, u'本月新入职：%s人' % aboard_num, table_end_style)
    out_sheet.write_merge(out_row_id + 3, out_row_id + 3, 0, col_len - 1, u'本月离职：%s人' % dismiss_num, table_end_style)
    out_sheet.write_merge(out_row_id + 4, out_row_id + 4, 0, col_len - 1, u'实际结算:%s=%s元'
                          % (calc_info, commission_sum), table_end_style)
    out_sheet.flush_row_data()


def draw_last_row(row_id, paid_sum):
    for col_idx in range(0, col_len):
        if out_title_list[col_idx] == u'姓名':
            out_sheet.row(row_id).write(col_idx, u'合计', style=common_style)
        elif out_title_list[col_idx] == u'金额/元':
            out_sheet.row(row_id).write(col_idx, paid_sum, style=common_style)
        else:
            out_sheet.row(row_id).write(col_idx, '', style=common_style)


def get_inserted_column_dict():
    inserted_column_dict = {u'工作单位': current_company, u'合同状态': person.status,
                            u'金额/元': person.commission, u'备注': person.comment}
    return inserted_column_dict


def draw_inserted_column_dict():
    inserted_cols = get_inserted_column_dict()
    if title == u'备注':
        width = len(inserted_cols[title]) * 12 / 7 + 1
        out_sheet.col(out_column_id).width = width * 256
    out_sheet.row(out_row_id).write(out_column_id, inserted_cols[title], style=common_style)


personnel_initializer()
for category in get_sheet_name_dict_from_column(key_col_name).keys():
    out_sheet = out_wb.add_sheet(sheet_name_from_value[category])

    draw_start_rows()

    for idx, title in enumerate(out_title_list):
        out_sheet.row(2).write(idx, title, style=common_style)

    out_row_id = 2
    for row in [int(p.sequence) for p in personnel_from_value[category]]:
        out_column_id = 0

        # offset between row id and corresponding content index
        # in personnel_from_value[category]
        out_row_offset = 3
        out_row_id += 1
        for title in out_title_list:
            if title in in_title_list:
                col = get_column_idx_by_title(title)
                if col in date_tittle_idx:
                    # YYYY/MM/DD contains 11 num_chars,
                    # good estimate 12*256 as the unit is 1/256 width of the char '0'
                    out_sheet.col(out_column_id).width = 12 *256
                    out_sheet.row(out_row_id).write(out_column_id, sheet.cell(row, col).value, style=date_style)
                elif col == get_column_idx_by_title(u'序号'):
                    out_sheet.row(out_row_id).write(out_column_id, out_row_id, style=common_style)
                else:
                    out_sheet.row(out_row_id).write(out_column_id, sheet.cell(row, col).value, style=common_style)
            else:
                person = personnel_from_value[category][out_row_id - out_row_offset]
                draw_inserted_column_dict()
            out_column_id += 1
        out_sheet.flush_row_data()

    commission_sum = sum([p.commission for p in personnel_from_value[category]])
    draw_last_row(out_row_id + 1, commission_sum)
    draw_ending_rows()


out_wb.save('output.xls')
