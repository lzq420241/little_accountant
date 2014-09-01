__author__ = 'liziqiang'
from calculator import date_from_string


class Personnel():
    # to do: unicode support
    def __init__(self, name, job_id, company, aboard_date, status, dismission_date):
        self.name = name
        self.job_id = job_id
        self.company = company
        self.aboard_date = date_from_string(aboard_date)
        self.status = status
        self.dismission_date = date_from_string(dismission_date)
        self.commission = 0
        self.comment = ""