__author__ = 'liziqiang'
from calculator import *

NORMAL_COMMISSION = 80
SPECIAL_COMMISSION = 100


class Personnel():
    # todo: unicode support
    def __init__(self, name, job_id, company, aboard_date, status, dismission_date):
        self.valid = True
        self.name = name
        self.job_id = job_id
        self.company = company
        self.aboard_date = date_from_string(aboard_date)
        self.status = status
        self.dismission_date = None
        if dismission_date:
            assert self.status
            self.dismission_date = date_from_string(dismission_date)
        self.commission = 0
        self.base_commission = 0
        self.comment = ""
        self.is_in_a_large_bundle = False
        self.update_valid_info()

    def set_bundle_flag(self):
        self.is_in_a_large_bundle = True

    def update_valid_info(self):
        month_since_aboard = get_interval_months_since_now(self.aboard_date)
        month_since_dismission = 0
        if self.dismission_date:
            month_since_dismission = get_interval_months_since_now(self.dismission_date)
        if month_since_aboard > 12 or month_since_dismission > 1:
            self.valid = False

    def get_commission(self):
        if self.valid:
            if self.dismission_date:
                self.get_commission_for_dismission()
            else:
                self.get_base_commission()

    def get_base_commission(self):
        if is_in_span(self.aboard_date) or self.is_in_a_large_bundle:
            self.base_commission = SPECIAL_COMMISSION
        else:
            self.base_commission = NORMAL_COMMISSION

    def get_commission_for_dismission(self):
        # todo: add member to record worked month
        if not self.base_commission:
            self.get_base_commission()
        worked_days_last = get_interval_days(self.dismission_date, get_first_day_of_last_month())
