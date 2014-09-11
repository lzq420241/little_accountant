# coding=gb2312
__author__ = 'liziqiang'

import unittest

from calculator import *
from personnel import *


# __all__ = ['TestCalculator', 'TestPersonnel']


class TestCalculator(unittest.TestCase):
    def setUp(self):
        pass

    def tearDown(self):
        pass

    def test_date_compare(self):
        str1 = '2014/7/28'
        str2 = '6/30/15'
        date1 = date_from_string(str1)
        date2 = date_from_string(str2)
        self.assertTrue(date1 < date2)

    def test_string_from_date(self):
        date1 = date_from_string('6/3/15')
        self.assertEqual(string_from_date(date1), '2015.06.03')

    def test_date_duration_date(self):
        str1 = '2014/7/28'
        str2 = '6/30/14'
        date1 = date_from_string(str1)
        date2 = date_from_string(str2)
        self.assertEqual(get_interval_days(date1, date2), 28)
        self.assertEqual(get_interval_days(date2, date1), 28)

    def test_is_date_in_last_month(self):
        str1 = '2014/7/28'
        str2 = '8/30/13'
        str3 = '2014/8/1'
        self.assertFalse(is_date_in_last_month(str1))
        self.assertFalse(is_date_in_last_month(str2))
        self.assertTrue(is_date_in_last_month(str3))

    def test_date_duration_str(self):
        str1 = '2014/7/28'
        str2 = '6/30/14'
        self.assertEqual(get_interval_days(str1, str2), 28)
        self.assertEqual(get_interval_days(str2, str1), 28)

    def test_date_duration_leap_year(self):
        str1 = '2000/2/28'
        str2 = '3/1/00'
        self.assertEqual(get_interval_days(str1, str2), 2)
        self.assertEqual(get_interval_days(str2, str1), 2)

    def test_date_duration_month(self):
        str1 = '2013/7/28'
        self.assertEqual(get_interval_months_since_now(str1), 14)

    def test_get_date_from_xldate(self):
        d = 41517.0
        self.assertEqual(get_date(d), get_date('8/31/13'))

    def test_get_last_month_id(self):
        import datetime

        today = datetime.date.today()
        expect_id = today.year * 12 + today.month - 1
        self.assertEqual(get_last_month_id(), expect_id)

    def test_get_spans(self):
        self.assertEqual(len(get_spans()), 3)

    def test_is_in_span(self):
        self.assertTrue(is_in_span('2013/3/20')[0])
        self.assertTrue(is_in_span('9/20/14')[0])
        self.assertFalse(is_in_span('9/20/12')[0])
        self.assertFalse(is_in_span('2015/3/20')[0])

    def test_get_days_of_last_month(self):
        self.assertTrue(get_days_of_last_month() == 31)

    def test_get_days_of_month(self):
        self.assertEqual(get_days_of_month('2000/2/3'), 29)
        self.assertEqual(get_days_of_month('2010/12/3'), 31)


class TestPersonnel(unittest.TestCase):
    def setUp(self):
        # aboard_date, status, dismission_date
        aboard_date = '7/20/14'
        status = ''
        dismission_date = ''
        self.test_class = Personnel(aboard_date, status, dismission_date)

    def tearDown(self):
        pass

    def test_normal_worker_more_than_a_month(self):
        self.assertEqual(self.test_class.commission, BONUS_COMMISSION)

    def test_normal_worker_less_than_twelve_month(self):
        self.test_class = Personnel('10/20/13', '', '')
        self.assertEqual(self.test_class.commission, NORMAL_COMMISSION)

    def test_over_worker_paid_month(self):
        self.test_class = Personnel('8/25/13', '', '')
        self.assertEqual(self.test_class.paid_month, 13)

    def test_normal_worker_paid_month(self):
        self.test_class = Personnel('8/26/13', '', '')
        expect_comment = u'这是第12个月支付，按80元/人支付，共支付一年（入职批次不满30人）'
        self.assertEqual(self.test_class.paid_month, 12)
        self.assertEqual(self.test_class.comment, expect_comment)

    def test_normal_worker_less_than_six_month(self):
        self.test_class = Personnel('5/20/14', '', '')
        self.assertEqual(self.test_class.commission, NORMAL_COMMISSION)

    def test_normal_worker_less_than_a_month(self):
        self.test_class = Personnel('8/25/14', '', '')
        self.assertEqual(self.test_class.commission, BONUS_COMMISSION)

    def test_normal_worker_less_than_seven_day(self):
        self.test_class = Personnel('8/26/14', '', '')
        self.assertEqual(self.test_class.commission, 0)

    def test_dismiss_worker_less_than_a_month(self):
        self.test_class = Personnel('8/26/13', u'试用期不合格', '8/20/14')
        expect_value = round(NORMAL_COMMISSION * 19.0 / 31, 2)
        self.assertEqual(self.test_class.commission, expect_value)

    def test_dismiss_worker_come_and_go_in_a_month(self):
        self.test_class = Personnel('8/13/14', u'试用期不合格', '8/20/14')
        expect_comment = u'员工入职月离职，正常办理离职手续，且离职月工作满%s天，一次性支付%s元/人' \
                         % (LEAST_DAYS_FOR_SPEC_COMMISSION, SPECIAL_COMMISSION)
        expect_value = SPECIAL_COMMISSION
        self.assertEqual(self.test_class.commission, expect_value)
        self.assertEqual(self.test_class.comment, expect_comment)

    def test_dismiss_worker_come_and_go_in_a_month_less_than_least(self):
        self.test_class = Personnel('8/14/14', u'试用期不合格', '8/20/14')
        expect_comment = u'员工入职月离职，正常办理离职手续，且离职月工作不满%s天，不计算提成' \
                         % LEAST_DAYS_FOR_SPEC_COMMISSION
        expect_value = 0
        self.assertEqual(self.test_class.commission, expect_value)
        self.assertEqual(self.test_class.comment, expect_comment)

    def test_normal_worker_less_than_sixteen_day(self):
        self.test_class = Personnel('8/26/13', u'试用期不合格', '8/16/14')
        expect_value = 0
        self.assertEqual(self.test_class.commission, expect_value)

    def test_auto_dismiss_worker_more_than_a_month(self):
        self.test_class = Personnel('8/26/13', u'自动离职', '8/20/14')
        expect_value = 0
        self.assertEqual(self.test_class.commission, expect_value)