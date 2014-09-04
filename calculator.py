__author__ = 'liziqiang'

import datetime

__all__ = ['date_from_string', 'get_interval_days', 'get_interval_months_since_now',
           'get_spans', 'is_in_span', 'get_first_day_of_last_month',
           'get_last_day_of_last_month', 'get_days_of_last_month']
hard_times_for_employer = ['3/20-3/31', '7/1-12/31']


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
            span_for_current.append(date_from_string(d + '/%s' % current_year))
            span_for_last.append(date_from_string(d + '/%s' % last_year))
        spans.append(span_for_current)
        spans.append(span_for_last)

    return spans


def is_in_span(date):
    spans = get_spans()
    date = _get_date(date)
    for s in spans:
        assert (type(s) == list)
        assert (len(s) == 2)
        start = s[0]
        end = s[1]
        if start <= date <= end:
            return True
    return False


def _get_date(para):
    if type(para) == str:
        return date_from_string(para)
    if type(para) != datetime.date:
        raise Exception("para %s type error, expect both be str or date, actual %s ." % (para, type(para)))
    return para


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


def get_interval_days(date1, date2):
    date1 = _get_date(date1)
    date2 = _get_date(date2)
    duration = date1 > date2 and date1 - date2 or date2 - date1
    return duration.days


def get_interval_months_since_now(date1):
    date1 = _get_date(date1)

    now = datetime.date.today()
    if date1 > now:
        raise Exception("date1 %s value error, ahead of current %s ." % (date1, now))

    month1 = date1.year * 12 + date1.month
    month2 = now.year * 12 + now.month
    return month2 - month1


def get_first_day_of_last_month():
    cur = datetime.date.today()
    d = 1
    m = (cur.month + 11) % 12
    y = cur.year
    if m == 1:
        y -= 1
    return datetime.date(y, m, d)


def get_last_day_of_last_month():
    cur = datetime.date.today()
    last_day_of_last_month = datetime.date(cur.year, cur.month, 1) - datetime.timedelta(days=1)
    return last_day_of_last_month


def get_days_of_last_month():
    return get_last_day_of_last_month().day