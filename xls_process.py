# coding=gb2312

__author__ = 'ziqli'
# todo: 1. sort 'contract status' normal before dismiss, care for sequence
# todo: 2. add three rows to each sheet
# todo: 3. add a summary sheet

import os
import re

from xlrd import *
from xlwt import *

from personnel import *
from calculator import *


in_title_list = [u'序号', u'姓名', u'工号', u'入职日期', u'离职日期']
out_title_list = [u'序号', u'姓名', u'工号', u'工作单位', u'入职日期', u'合同状态', u'离职日期', u'金额/元', u'备注']
date_title_list = [u'入职日期', u'离职日期']
key_col_name = u'第三方'
xls_to_be_processed = u'联胜7月外包费用核对-承天（8-14）.xls'
reference_xls = u'2014年7月联胜（承天）第三方提成8-21.xls'
calc_desc = u'2014.2.1开始当月入职满30人，支付100元/人/月，不满30人，支付80元/人/月，共支付1年。员工入职当月满7天以上' \
            u'办理正常离职手续的，一次性支付200元/人。员工非入职月离职，正常办理离职手续且离职当月满16天以上（含16天）' \
            u'按离职当月实际在职天数计算提成。\n' \
            u'2014.03.20至2014.03.31期间由第三方输送联胜的人员支付方式调整如下：' \
            u'取消第三方输送人数限制，每输送一人支付100元/人/月，支付时间为一年。\n' \
            u'从2014年7月1日起至2014年12月31日，不管入职员工人数数量的多少，按100元/人/月的计算方式计算费用。'
corp_sign = u'东莞市承天人力资源管理咨询有限公司'
current_company = u'联胜'
summary_sheet_name = u'提成汇总'
col_len = len(out_title_list)
BUNDLE = 30

alignment = Alignment()
alignment.horz = Alignment.HORZ_CENTER
left_alignment = Alignment()
left_alignment.horz = Alignment.HORZ_LEFT
right_alignment = Alignment()
right_alignment.horz = Alignment.HORZ_RIGHT

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

center_big_font = XFStyle()
center_big_font.alignment = alignment
center_big_font.font = big_font

center_no_border = XFStyle()
center_no_border.alignment = alignment
center_no_border.font = font

center_border = XFStyle()
center_border.alignment = alignment
center_border.borders = borders
center_border.font = font

center_border_num = XFStyle()
center_border_num.alignment = alignment
center_border_num.borders = borders
center_border_num.font = font
center_border_num.num_format_str = '0.00'

left_align = XFStyle()
left_align.alignment = left_alignment
left_align.alignment.wrap = 1
left_align.font = font

right_align = XFStyle()
right_align.alignment = right_alignment
right_align.alignment.wrap = 1
right_align.font = font
right_align.num_format_str = '0.00'

date_style = XFStyle()
date_style.num_format_str = 'YYYY/MM/DD'
date_style.alignment = alignment
date_style.borders = borders
date_style.font = font

# key:category value:sheet_name
sheet_name_from_value = {}
# key:category value:personnel list
personnel_from_value = {}
# key:category value:sequence list
sequence_from_value = {}
# key:category value:list [<int>aboard_num, <int>dismiss_num, <str>calc_desc]
sheet_statistics = {}
# key:third party value:total commission
commission_for_third_party = {}
# worker_ids need to be retained
valid_worker_ids = set()
# worker_ids in bundle
worker_ids_in_bundle = set()

cur_dir = os.path.dirname(__file__)
wb = open_workbook(os.path.join(cur_dir, xls_to_be_processed))
# print wb.encoding
#print wb.codepage
out_wb = Workbook(encoding='utf-16le')

sheet = wb.sheet_by_index(0)
cols = sheet.ncols
# do not calc last stat line
rows = sheet.nrows - 1


# this func get all worker_ids that are in a bundle
# from clearing xls of previous month
# (more than 30 were hire during a month)
def get_pre_bundle_worker_ids():
    global worker_ids_in_bundle
    ref_wb = open_workbook(os.path.join(cur_dir, reference_xls))
    for sheet_index in range(ref_wb.nsheets):
        cur_sheet = ref_wb.sheet_by_index(sheet_index)
        if cur_sheet.ncols < col_len:
            continue
        aboard_2_months = set()
        row_idx = 3
        comment = cur_sheet.cell(row_idx, col_len - 1).value
        while comment:
            worker_id = cur_sheet.cell(row_idx, worker_id_col_idx).value
            aboard_date = cur_sheet.cell(row_idx, aboard_col_idx).value
            # search for '[^不]满30'
            if re.match(ur'.+[^\u4e0d][\u6ee1]%s.+' % BUNDLE, comment, re.UNICODE):
                worker_ids_in_bundle.add(worker_id)
            if get_month_id(aboard_date) == get_last_month_id() - 1:
                aboard_2_months.add(worker_id)
            row_idx += 1
            comment = cur_sheet.cell(row_idx, col_len - 1).value
        if len(aboard_2_months) >= BUNDLE:
            worker_ids_in_bundle += aboard_2_months


# this func get all worker_ids that are valid
# from upstream xls
# (already paid for 12 month or for the dismission)
def get_valid_worker_ids():
    for row in range(1, rows):
        aboard_date = get_date(sheet.cell(row, aboard_col_idx).value)
        dismission_date = get_date(sheet.cell(row, dismission_col_idx).value)
        if is_record_valid(aboard_date, dismission_date):
            valid_worker_id = sheet.cell(row, worker_id_col_idx).value
            valid_worker_ids.add(valid_worker_id)


# this func add worker_ids in bundle to worker_ids_in_bundle for new comer
# from upstream xls
def add_worker_ids_for_new_comer():
    global worker_ids_in_bundle
    for category in get_sheet_name_dict_from_column(key_col_name).keys():
        tmp_set = set()
        for row in range(1, rows):
            if sheet.cell(row, key_column_index).value == category:
                aboard_date = sheet.cell(row, aboard_col_idx).value
                if is_date_in_last_month(aboard_date) and not is_in_span(aboard_date):
                    worker_id = sheet.cell(row, worker_id_col_idx).value
                    tmp_set.add(worker_id)
        if len(tmp_set) >= BUNDLE:
            worker_ids_in_bundle += tmp_set


up_stream_xls_titles = []
for column in range(0, cols):
    up_stream_xls_titles.append(sheet.col_values(column, 0, 1)[0])


def get_column_idx_by_title(title_str):
    return list.index(up_stream_xls_titles, title_str)


def get_cell_by_title(row_idx, title_str):
    col_idx = get_column_idx_by_title(title_str)
    content = sheet.cell(row_idx, col_idx).value
    return content


def get_column_ids_by_names(title_list):
    ids = [get_column_idx_by_title(t) for t in title_list]
    return ids


def get_sheet_name_dict_from_column(title):
    idx = get_column_idx_by_title(title)
    #search for '来源（第三方）'
    valid_style = re.compile(ur'(\w+)\W*[\uff08\u0028]\W*(\w+)\W*[\uff09\u0029]', re.UNICODE)
    for row in range(1, rows):
        content = sheet.row_values(row, idx, idx + 1)[0]
        if type(content) == unicode:
            #print content
            matched_content = valid_style.match(content)
        else:
            matched_content = None
        if matched_content:
            # sheet_name = "%s-%s%s" % (matched_content.group(1), current_company, matched_content.group(2))
            sheet_name = (matched_content.group(1), matched_content.group(2))
            sheet_name_from_value[content] = sheet_name
    #print sheet_name_from_value and "OK" or "NULL"
    return sheet_name_from_value


date_title_idx = get_column_ids_by_names(date_title_list)
key_column_index = get_column_idx_by_title(key_col_name)
sequence_col_idx = get_column_idx_by_title(u'序号')
worker_id_col_idx = get_column_idx_by_title(u'工号')
aboard_col_idx = get_column_idx_by_title(u'入职日期')
dismission_col_idx = get_column_idx_by_title(u'离职日期')
status_col_idx = get_column_idx_by_title(u'离职方式')


def personnel_initializer():
    global worker_ids_in_bundle
    get_pre_bundle_worker_ids()
    add_worker_ids_for_new_comer()
    get_valid_worker_ids()
    # print worker_ids_in_bundle
    worker_ids_in_bundle &= valid_worker_ids
    for category in get_sheet_name_dict_from_column(key_col_name).keys():
        personnel_list = list()
        sequence_list = list()
        for row in range(1, rows):
            if sheet.cell(row, key_column_index).value == category:

                sequence = sheet.cell(row, sequence_col_idx).value
                worker_id = sheet.cell(row, worker_id_col_idx).value
                aboard_date = sheet.cell(row, aboard_col_idx).value
                status = sheet.cell(row, status_col_idx).value
                dismission_date = sheet.cell(row, dismission_col_idx).value

                is_in_bundle = worker_id in worker_ids_in_bundle
                is_valid = worker_id in valid_worker_ids
                if is_valid:
                    # print worker_id, get_date(dismission_date), status
                    person = Personnel(aboard_date, status, dismission_date, is_in_bundle)
                    sequence_list.append(int(sequence))
                    personnel_list.append(person)
        if sequence_list:
            tmp_zip = zip(sequence_list, personnel_list)
            sorted_zip = sorted(tmp_zip, key=lambda x: x[1].weight)
            sequence_list, personnel_list = zip(*sorted_zip)
        # information used by ending rows
        sheet_statistics[category] = get_sheet_statistics(personnel_list)
        # print sheet_statistics[category]
        personnel_from_value[category] = personnel_list
        sequence_from_value[category] = sequence_list


def get_sheet_statistics(persons):
    assert type(persons) == list or type(persons) == tuple
    aboard_number = len(filter((lambda p: is_date_in_last_month(p.aboard_date)), persons))
    dismiss_number = len(filter((lambda p: p.status != u'在职'), persons))
    calc_stat = list()
    com_list = [p.commission for p in persons]
    com_set = set(com_list)
    if 0.00 in com_set:
        com_set.remove(0.00)
    for c in com_set:
        cnt = com_list.count(c)
        calc_stat.append(u"%s人*%s元/人" % (cnt, c))
    return [aboard_number, dismiss_number, '+'.join(calc_stat)]


def draw_start_rows():
    tab_title = u'%s提成（%s）' % (current_company, third_party)
    time_info = get_first_day_of_last_month()
    desc = u'所属时间：%s年%s月' % (time_info.year, time_info.month)
    # print tab_title
    out_sheet.write_merge(0, 0, 0, col_len - 1, tab_title, center_big_font)
    out_sheet.row(1).write(col_len - 1, desc, style=center_no_border)
    out_sheet.col(col_len - 1).width = 10 * 12 / 7 * 256
    out_sheet.flush_row_data()


def draw_ending_rows():
    aboard_num, dismiss_num, calc_info = sheet_statistics[category]
    out_sheet.write_merge(out_row_id + 2, out_row_id + 2, 0, col_len - 1, u'本月新入职：%s人' % aboard_num, left_align)
    out_sheet.write_merge(out_row_id + 3, out_row_id + 3, 0, col_len - 1, u'本月离职：%s人' % dismiss_num, left_align)
    out_sheet.write_merge(out_row_id + 4, out_row_id + 4, 0, col_len - 1, u'实际结算:%s=%s元'
                          % (calc_info, commission_sum), left_align)
    out_sheet.flush_row_data()


def draw_signature():
    import datetime

    now = datetime.date.today()

    out_sheet.write_merge(out_row_id + 5, out_row_id + 5, 0, col_len - 1, calc_desc, left_align)
    out_sheet.row(out_row_id + 5).height_mismatch = True
    out_sheet.row(out_row_id + 5).height = 256 * 4
    out_sheet.write_merge(out_row_id + 6, out_row_id + 6, 5, col_len - 1, corp_sign, center_no_border)
    out_sheet.write_merge(out_row_id + 7, out_row_id + 7, 5, col_len - 1, u'%s年%s月%s日'
                          % (now.year, now.month, now.day), center_no_border)

    out_sheet.flush_row_data()


def draw_last_row(row_id, paid_sum):
    for col_idx in range(0, col_len):
        if out_title_list[col_idx] == u'姓名':
            out_sheet.row(row_id).write(col_idx, u'合计', style=center_border)
        elif out_title_list[col_idx] == u'金额/元':
            out_sheet.row(row_id).write(col_idx, paid_sum, style=center_border_num)
        else:
            out_sheet.row(row_id).write(col_idx, '', style=center_border)


def get_inserted_column_dict():
    inserted_column_dict = {u'工作单位': current_company, u'合同状态': person.status,
                            u'金额/元': person.commission, u'备注': person.comment}
    return inserted_column_dict


def draw_inserted_column_dict():
    global max_width_of_comment
    inserted_cols = get_inserted_column_dict()
    if title == u'备注':
        width = len(inserted_cols[title]) * 12 / 7
        if width > max_width_of_comment:
            max_width_of_comment = width
        out_sheet.col(out_column_id).width = max_width_of_comment * 256
    out_sheet.row(out_row_id).write(out_column_id, inserted_cols[title], style=center_border)


personnel_initializer()

import locale

cate_list = get_sheet_name_dict_from_column(key_col_name).keys()
# this reads the environment and inits the right locale
locale.setlocale(locale.LC_ALL, "")

cate_list.sort(cmp=locale.strcoll)

for category in cate_list:
    max_width_of_comment = 0
    source = sheet_name_from_value[category][0]
    third_party = sheet_name_from_value[category][1]
    sheet_name = u"%s-%s%s" % (source, current_company, third_party)
    out_sheet = out_wb.add_sheet(sheet_name)

    draw_start_rows()

    for idx, title in enumerate(out_title_list):
        out_sheet.row(2).write(idx, title, style=center_border)

    out_row_id = 2
    out_row_offset = 3

    for raw_row_id, row in enumerate(sequence_from_value[category]):
        out_column_id = 0

        # offset between raw_row_id and corresponding content index
        # in personnel_from_value[category]
        out_row_id = raw_row_id + out_row_offset
        for title in out_title_list:
            if title in in_title_list:
                col = get_column_idx_by_title(title)
                if col in date_title_idx:
                    # YYYY/MM/DD contains 11 num_chars,
                    # good estimate 12*256 as the unit is 1/256 width of the char '0'
                    out_sheet.col(out_column_id).width = 12 * 256
                    out_sheet.row(out_row_id).write(out_column_id, sheet.cell(row, col).value, style=date_style)
                elif col == get_column_idx_by_title(u'序号'):
                    out_sheet.row(out_row_id).write(out_column_id, raw_row_id + 1, style=center_border)
                else:
                    out_sheet.row(out_row_id).write(out_column_id, sheet.cell(row, col).value, style=center_border)
            else:
                person = personnel_from_value[category][raw_row_id]
                draw_inserted_column_dict()
            out_column_id += 1
        out_sheet.flush_row_data()

    commission_sum = sum([p.commission for p in personnel_from_value[category]])
    commission_for_third_party[third_party] = commission_sum
    draw_last_row(out_row_id + 1, commission_sum)
    draw_ending_rows()
    draw_signature()

# generate the summary sheet
out_sheet = out_wb.add_sheet(summary_sheet_name)
rowx = 0
for third_party, commission in commission_for_third_party.items():
    out_sheet.row(rowx).write(0, third_party, style=left_align)
    out_sheet.row(rowx).write(1, commission, style=right_align)
    rowx += 1
out_sheet.row(rowx).write(1, sum(commission_for_third_party.values()), style=right_align)
out_sheet.flush_row_data()

out_wb.save('output.xls')