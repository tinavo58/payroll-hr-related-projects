#!/usr/bin/env python
"""
.func to calc period of continuos service
.func to determine notice period
"""
import datetime
from dateutil.relativedelta import relativedelta


def convert_string_to_date(date_input):
    return datetime.datetime.strptime(date_input, '%d/%m/%Y')


def calc_period_of_continous_service(startDate: str, termDate: str):
    start_date = convert_string_to_date(startDate)
    termDate = convert_string_to_date(termDate)

    diff = relativedelta(termDate, start_date)
    return diff.years, diff.months, diff.days


def calc_min_notice_period(startDate, termDate):
    year, month, day = calc_period_of_continous_service(startDate, termDate)
    if year <= 1 and month == day == 0:
        print('1 week')
    elif (year <= 3 and month == day == 0) or (month != 0) or (day != 0):
        print('2 weeks')
    elif (year <= 5 and month == day == 0) or (month != 0) or (day != 0):
        print('3 weeks')
    else:
        print('4 weeks')


if __name__ == '__main__':
    calc_min_notice_period('7/05/2020', '08/05/2023')
