# coding=gb2312
__author__ = 'liziqiang'

import unittest

from calculator import *
from personnel import *


__all__ = ['TestCalculator']


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

    def test_date_duration_date(self):
        str1 = '2014/7/28'
        str2 = '6/30/14'
        date1 = date_from_string(str1)
        date2 = date_from_string(str2)
        self.assertEqual(get_interval_days(date1, date2), 28)
        self.assertEqual(get_interval_days(date2, date1), 28)

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

    def test_get_spans(self):
        self.assertEqual(len(get_spans()), 4)

    def test_is_in_span(self):
        self.assertTrue(is_in_span('2013/3/20'))
        self.assertTrue(is_in_span('9/20/14'))
        self.assertFalse(is_in_span('9/20/12'))
        self.assertFalse(is_in_span('2015/3/20'))

    def test_get_days_of_last_month(self):
        self.assertTrue(get_days_of_last_month() == 31)


class TestPersonnel(unittest.TestCase):
    def setUp(self):
        # name, job_id, company, aboard_date, status, dismission_date
        name = '张三'
        jod_id = 'LK141717'
        company = '生产一部 强化课'
        aboard_date = '7/20/14'
        status = ''
        dismission_date = ''
        self.test_class = Personnel(name, jod_id, company, aboard_date, status, dismission_date)

    def tearDown(self):
        pass

    def test_normal_worker_more_than_a_month(self):
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        self.assertTrue(self.test_class.valid)
        self.assertEqual(self.test_class.commission, BONUS_COMMISSION)

    def test_normal_worker_more_than_twelve_month(self):
        self.test_class.aboard_date = date_from_string('3/20/13')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        self.assertFalse(self.test_class.valid)

    def test_normal_worker_less_than_twelve_month(self):
        self.test_class.aboard_date = date_from_string('10/20/13')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        self.assertTrue(self.test_class.valid)
        self.assertEqual(self.test_class.commission, BONUS_COMMISSION)

    def test_normal_worker_less_than_six_month(self):
        self.test_class.aboard_date = date_from_string('5/20/14')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        self.assertTrue(self.test_class.valid)
        self.assertEqual(self.test_class.commission, NORMAL_COMMISSION)

    def test_normal_worker_less_than_a_month(self):
        self.test_class.aboard_date = date_from_string('8/25/14')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        self.assertTrue(self.test_class.valid)
        self.assertEqual(self.test_class.commission, BONUS_COMMISSION)

    def test_normal_worker_less_than_seven_day(self):
        self.test_class.aboard_date = date_from_string('8/26/14')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        self.assertTrue(self.test_class.valid)
        self.assertEqual(self.test_class.commission, 0)

    def test_dismiss_worker_more_than_a_month(self):
        self.test_class.status = '试用期不合格'
        self.test_class.dismission_date = date_from_string('7/20/14')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        self.assertFalse(self.test_class.valid)

    def test_dismiss_worker_more_than_a_month(self):
        self.test_class.status = '试用期不合格'
        self.test_class.dismission_date = date_from_string('8/20/14')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        expect_value = round(BONUS_COMMISSION * 19.0 / 31, 2)
        self.assertEqual(self.test_class.commission, expect_value)

    def test_normal_worker_less_than_sixteen_day(self):
        self.test_class.status = '试用期不合格'
        self.test_class.dismission_date = date_from_string('8/16/14')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        expect_value = 0
        self.assertEqual(self.test_class.commission, expect_value)

    def test_auto_dismiss_worker_more_than_a_month(self):
        self.test_class.status = '自动离职'
        self.test_class.dismission_date = date_from_string('8/20/14')
        self.test_class.update_valid_info()
        self.test_class.get_commission()
        expect_value = 0
        self.assertEqual(self.test_class.commission, expect_value)