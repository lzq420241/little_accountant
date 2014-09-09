# coding=gb2312
__author__ = 'liziqiang'
from calculator import *

NORMAL_COMMISSION = 80
BONUS_COMMISSION = 100
SPECIAL_COMMISSION = 200
LEAST_DAYS_FOR_NOR_COMMISSION = 7
LEAST_DAYS_FOR_SPEC_COMMISSION = 7
LEAST_DAYS_FOR_DISMISS_COMMISSION = 16

__all__ = ['Personnel', 'NORMAL_COMMISSION', 'BONUS_COMMISSION', 'SPECIAL_COMMISSION',
           'LEAST_DAYS_FOR_NOR_COMMISSION', 'LEAST_DAYS_FOR_DISMISS_COMMISSION',
           'LEAST_DAYS_FOR_SPEC_COMMISSION', 'is_record_valid']


def is_record_valid(aboard_date, dismission_date):
    paid_month = get_paid_month(aboard_date)
    month_since_dismission = get_interval_months_since_now(dismission_date)
    if paid_month > 12 or month_since_dismission > 1:
        return False
    return True


def get_paid_month(aboard_date):
    aboard_date = get_date(aboard_date)
    assert aboard_date
    paid_month = get_interval_months_since_now(aboard_date)
    days_worked_in_aboard_month = get_days_of_month(aboard_date) - aboard_date.day + 1
    if days_worked_in_aboard_month < LEAST_DAYS_FOR_NOR_COMMISSION:
        paid_month -= 1
    return paid_month


class Personnel():
    def __init__(self, sequence, aboard_date, status, dismission_date,
                 is_in_a_large_bundle=False):

        # for the convenience of locate cell row id of the input xls file
        self.sequence = sequence

        self.aboard_date = get_date(aboard_date)
        self.status = status

        self.dismission_date = get_date(dismission_date)
        self.commission = 0
        self.base_commission = 0
        self.comment = ""
        self.period = list()
        self.paid_month = 0
        # aid for monthly statistics, comment generation, commission calculation
        self.aboard_month_id = get_month_id(self.aboard_date)
        if self.dismission_date:
            self.dismission_month_id = get_month_id(self.dismission_date)

        self.care_bundle = True
        self.is_in_a_large_bundle = is_in_a_large_bundle

        if not status and not self.dismission_date:
            status = u'��ְ'
        self.status = status
        self.get_paid_month()
        self.get_commission()
        self.get_comment()

    def get_commission(self):
        if self.dismission_date and self.status != u'�Զ���ְ':
            self.get_commission_for_dismission()
        elif not self.dismission_date and self.__days_worked_last_month() >= LEAST_DAYS_FOR_NOR_COMMISSION:
            self.get_base_commission()
            self.commission = self.base_commission

    def get_base_commission(self):
        if is_in_span(self.aboard_date)[0]:
            self.base_commission = BONUS_COMMISSION
        elif self.is_in_a_large_bundle:
            self.base_commission = BONUS_COMMISSION
        else:
            self.base_commission = NORMAL_COMMISSION

    def get_commission_for_dismission(self):
        if not self.base_commission:
            self.get_base_commission()
        worked_days = self.__days_worked_last_month()

        if self.dismission_month_id == self.aboard_month_id and worked_days >= LEAST_DAYS_FOR_NOR_COMMISSION:
            self.commission = SPECIAL_COMMISSION
        elif self.dismission_month_id != self.aboard_month_id and worked_days >= LEAST_DAYS_FOR_DISMISS_COMMISSION:
            commission = self.base_commission * 1.0 * worked_days / get_days_of_last_month()
            self.commission = round(commission, 2)

    def get_start_work_day(self):
        tmp_day = get_first_day_of_last_month()
        return self.aboard_date > tmp_day and self.aboard_date or tmp_day

    def get_paid_month(self):
        self.paid_month = get_interval_months_since_now(self.aboard_date)
        days_worked_in_aboard_month = get_days_of_month(self.aboard_date) - self.aboard_date.day + 1
        if days_worked_in_aboard_month < LEAST_DAYS_FOR_NOR_COMMISSION:
            self.paid_month -= 1

    def __days_worked_last_month(self):
        if self.dismission_date:
            return get_interval_days(self.dismission_date, self.get_start_work_day())
        else:
            return get_interval_days(get_last_day_of_last_month(), self.get_start_work_day()) + 1

    def get_comment(self):
        if self.status == u'��ְ':
            self.comment = self.get_comment_for_on_work()
        elif self.status != u'�Զ���ְ':
            self.comment = self.get_comment_for_dismission()
        else:
            self.comment = u'��������ְ���������'

    def get_comment_for_on_work(self):
        # todo: better get comment from basic info other than commission
        # todo: may fail due to requirement changed.
        if self.commission == BONUS_COMMISSION:
            if self.is_in_a_large_bundle:
                reason = u'����ְ������30�ˣ�'
            else:
                flag, period = is_in_span(self.aboard_date)
                assert flag
                reason = u'��%s��%s�ڼ���ְ��' % (period[0], period[1])
        elif self.commission == NORMAL_COMMISSION:
            reason = u'����ְ���β���30�ˣ�'
        else:
            return u'��ְ���²���%s��' % LEAST_DAYS_FOR_NOR_COMMISSION
        comment = u'���ǵ�%s����֧������%sԪ/��֧������֧��һ��%s' \
                  % (self.paid_month, self.commission, reason)
        return comment

    def get_comment_for_dismission(self):
        if self.commission == SPECIAL_COMMISSION:
            desc = u'Ա����ְ����ְ������������ְ����������ְ�¹�����%s�죬һ����֧��%sԪ/��' \
                   % (LEAST_DAYS_FOR_SPEC_COMMISSION, SPECIAL_COMMISSION)
        elif self.commission > 0:
            desc = u'Ա������ְ����ְ������������ְ����������ְ�¹�����%s�죬��ʵ����ְ�����������' \
                   % LEAST_DAYS_FOR_DISMISS_COMMISSION
        elif self.dismission_month_id == self.aboard_month_id:
            desc = u'Ա����ְ����ְ������������ְ����������ְ�¹�������%s�죬���������' \
                   % LEAST_DAYS_FOR_SPEC_COMMISSION
        else:
            desc = u'Ա������ְ����ְ������������ְ����������ְ�¹�������%s�죬���������' \
                   % LEAST_DAYS_FOR_DISMISS_COMMISSION
        return desc

    def __str__(self):
        print self.commission, self.comment
        return '\n'