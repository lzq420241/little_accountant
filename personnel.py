# coding=gb2312
__author__ = 'liziqiang'
from calculator import *

NORMAL_COMMISSION = 80
BONUS_COMMISSION = 100
SPECIAL_COMMISSION = 200
LEAST_DAYS_FOR_NOR_COMMISSION = 7
LEAST_DAYS_FOR_DISMISS_COMMISSION = 16

__all__ = ['Personnel', 'NORMAL_COMMISSION', 'BONUS_COMMISSION', 'SPECIAL_COMMISSION']


class Personnel():
    def __init__(self, aboard_date, status, dismission_date):
        self.valid = True
        self.aboard_date = date_from_string(aboard_date)
        self.status = status and status or u'在职'
        self.dismission_date = None
        if dismission_date:
            assert self.status != u'在职'
            self.dismission_date = date_from_string(dismission_date)
        self.commission = 0
        self.base_commission = 0
        self.comment = ""
        self.period = list()
        self.paid_month = 0
        self.is_in_a_large_bundle = False
        # todo: remove comments below after release
        # self.update_valid_info()
        #self.get_commission()

    def set_bundle_flag(self):
        self.is_in_a_large_bundle = True

    def update_valid_info(self):
        self.get_paid_month()

        month_since_dismission = 0
        if self.dismission_date:
            month_since_dismission = get_interval_months_since_now(self.dismission_date)
        if self.paid_month > 12 or (month_since_dismission and month_since_dismission != 1):
            self.valid = False

    def get_commission(self):
        if self.valid:
            if self.dismission_date and self.status != u'自动离职':
                self.get_commission_for_dismission()
            elif not self.dismission_date and self.__days_worked_last_month() >= LEAST_DAYS_FOR_NOR_COMMISSION:
                self.get_base_commission()
                self.commission = self.base_commission

    def get_base_commission(self):
        if is_in_span(self.aboard_date) or self.is_in_a_large_bundle:
            self.base_commission = BONUS_COMMISSION
        else:
            self.base_commission = NORMAL_COMMISSION

    def get_commission_for_dismission(self):
        if not self.base_commission:
            self.get_base_commission()
        worked_days = self.__days_worked_last_month()

        if self.dismission_date.month == self.aboard_date.month and worked_days >= LEAST_DAYS_FOR_NOR_COMMISSION:
            self.commission = SPECIAL_COMMISSION
        elif self.dismission_date.month != self.aboard_date.month and worked_days >= LEAST_DAYS_FOR_DISMISS_COMMISSION:
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

    # being called under condition that personnel.valid is True
    # no need to consider situation that self.dismission_date.month is same with current month
    def __days_worked_last_month(self):
        if self.dismission_date:
            return get_interval_days(self.dismission_date, self.get_start_work_day())
        else:
            return get_interval_days(get_last_day_of_last_month(), self.get_start_work_day()) + 1

    def get_comment(self):
        if self.valid:
            if self.status == u'在职':
                self.comment = self.get_comment_for_on_work()
            elif self.status != u'自动离职':
                self.comment = self.get_comment_for_dismission()
            else:
                self.comment = u'非正常离职不计算提成'

    def get_comment_for_on_work(self):
        # todo: better get comment from basic info other than commission
        # todo: may fail due to requirement changed.
        if self.commission == BONUS_COMMISSION:
            if self.is_in_a_large_bundle:
                reason = u'（入职批次满30人）'
            else:
                flag, period = is_in_span(self.aboard_date)
                assert flag
                reason = u'（%s至%s期间入职）' % (period[0], period[1])
        elif self.commission == NORMAL_COMMISSION:
            reason = u'（入职批次不满30人）'
        else:
            return u'入职当月不满%s天' % LEAST_DAYS_FOR_NOR_COMMISSION
        comment = u'这是第%s个月支付，按%s元/人支付，共支付一年%s' \
                  % (self.paid_month, self.commission, reason)
        return comment

    def get_comment_for_dismission(self):
        if self.commission == SPECIAL_COMMISSION:
            desc = u'（员工入职月离职，正常办理离职手续，且离职月工作满%s天，一次性支付%s元/人）' \
                   % (LEAST_DAYS_FOR_DISMISS_COMMISSION, SPECIAL_COMMISSION)
        elif self.commission > 0:
            desc = u'（员工非入职月离职，正常办理离职手续，且离职月工作满%s天，按实际在职天数计算提成）' \
                   % LEAST_DAYS_FOR_DISMISS_COMMISSION
        elif self.dismission_date.month == self.aboard_date.month:
            desc = u'（员工入职月离职，正常办理离职手续，且离职月工作不满%s天，不计算提成）' \
                   % LEAST_DAYS_FOR_DISMISS_COMMISSION
        else:
            desc = u'（员工非入职月离职，正常办理离职手续，且离职月工作不满%s天，不计算提成）' \
                   % LEAST_DAYS_FOR_DISMISS_COMMISSION
        return desc