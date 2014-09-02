__author__ = 'liziqiang'

import unittest

from calculator import *


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