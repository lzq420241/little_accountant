__author__ = 'liziqiang'

import datetime


def date_from_string(date_str):
    """
    This Class create a date object from a string.
    string format can be either of these:
    YYYY/MM/DD or [M]M/[D]D/YY
    """
    assert(type(date_str) == str)
    date_list = date_str.split('/')

    if len(date_list) != 3:
        raise Exception("date_str %s format error." % date_str)

    def get_date():
        first_part = int(date_list[0])
        second_part = int(date_list[1])
        third_part = int(date_list[2])
        if first_part > 99:
            year = first_part
            month = second_part
            day = third_part
        else:
            month = first_part
            day = second_part
            year = third_part + 2000

        return datetime.date(year, month, day)

    return get_date()


def get_interval_days(date1, date2):
    if type(date1) == str and type(date2) == str:
        date1 = date_from_string(date1)
        date2 = date_from_string(date2)
    if type(date1) != datetime.date or type(date2) != datetime.date:
        raise Exception("date1 or date2 %s %s type error, expect both be str or date, actual %s %s ."
                        % (date1, date2, type(date1), type(date1)))
    duration = date1 > date2 and date1 - date2 or date2 - date1
    return duration.days