__author__ = 'liziqiang'


# change system time to 2014 Sep to get the UT run properly.
import datetime

__all__ = ['date_from_string', 'string_from_date', 'get_interval_days',
           'get_interval_months_since_now', 'is_date_in_last_month',
           'get_spans', 'is_in_span', 'get_first_day_of_last_month',
           'get_last_day_of_last_month', 'get_days_of_last_month',
           'get_days_of_month', 'get_month_id', 'get_last_month_id',
           'get_first_day_of_cur_month', 'get_date']
hard_times_for_employer = ['3/20-3/31', '7/1/14-12/31/14']


def get_spans():
    """
    :return: return list of list like [date_start date_end]
    """
    current_year = datetime.date.today().year - 2000
    last_year = current_year - 1

    spans = []

    for span in hard_times_for_employer:
        span_for_current = []
        span_for_last = []
        for d in span.split('-'):
            if d.count('/') == 1:
                span_for_current.append(date_from_string(d + '/%s' % current_year))
                span_for_last.append(date_from_string(d + '/%s' % last_year))
            elif d.count('/') == 2:
                span_for_current.append(date_from_string(d))
            else:
                raise Exception("date_str %s format error." % d)
        spans.append(span_for_current)
        if span_for_last:
            spans.append(span_for_last)

    return spans


def get_date_from_xldate(xls_date):
    import xlrd

    (y, m, d, hh, mm, ss) = xlrd.xldate_as_tuple(xls_date, 0)
    return datetime.date(y, m, d)


def is_in_span(date):
    spans = get_spans()
    date = get_date(date)
    for s in spans:
        assert (type(s) == list)
        assert (len(s) >= 2)
        begin = s[0]
        end = s[1]
        if begin <= date <= end:
            return True, [string_from_date(begin), string_from_date(end)]
    return False, None


def get_month_id(para):
    tmp = get_date(para)
    assert tmp
    return tmp.year * 12 + tmp.month


def get_last_month_id():
    today = datetime.date.today()
    return get_month_id(today) - 1


# if para is False, return None
def get_date(para):
    if not para or type(para) == datetime.date:
        return para
    if type(para) == str:
        return date_from_string(para)
    if type(para) == float:
        return get_date_from_xldate(para)
    else:
        raise Exception("para %s type error, expect both be str or date, actual %s ." % (para, type(para)))


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
            assert (0 < first_part < 13)
            assert (0 < second_part < 32)
            month = first_part
            day = second_part
            year = third_part + 2000

        return datetime.date(year, month, day)

    return get_date()


def is_date_in_last_month(date_str):
    if not date_str:
        return False
    in_date_month_id = get_month_id(date_str)
    return in_date_month_id == get_last_month_id()


def string_from_date(date_obj):
    """
    This Class create a string from a date object.
    string format will be:
    YYYY.MM.DD
    """
    assert (type(date_obj) == datetime.date)
    return date_obj.strftime("%Y.%m.%d")


def get_interval_days(date1, date2):
    date1 = get_date(date1)
    date2 = get_date(date2)
    duration = date1 > date2 and date1 - date2 or date2 - date1
    return duration.days


def get_interval_months_since_now(date1):
    if not date1:
        return 0
    now = datetime.date.today()
    interval = get_month_id(now) - get_month_id(date1)
    assert interval >= 0
    return interval


def get_first_day_of_last_month():
    cur = datetime.date.today()
    d = 1
    m = cur.month - 1
    y = cur.year
    if not m:
        m = 12
        y -= 1
    return datetime.date(y, m, d)


def get_first_day_of_cur_month():
    cur = datetime.date.today()
    d = 1
    m = cur.month
    y = cur.year
    return datetime.date(y, m, d)


def get_last_day_of_last_month():
    cur = datetime.date.today()
    last_day_of_last_month = datetime.date(cur.year, cur.month, 1) - datetime.timedelta(days=1)
    return last_day_of_last_month


def get_days_of_last_month():
    return get_last_day_of_last_month().day


def get_days_of_month(para):
    cur_date = get_date(para)
    if cur_date.month == 12:
        next_month = 1
        next_year = cur_date.year + 1
    else:
        next_month = cur_date.month + 1
        next_year = cur_date.year
    last_day_of_cur_month = datetime.date(next_year, next_month, 1) - datetime.timedelta(days=1)
    return last_day_of_cur_month.day