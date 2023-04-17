import pandas as pd

from datetime import date
from dateutil.relativedelta import relativedelta
# https://towardsdatascience.com/the-easiest-way-to-identify-holidays-in-python-58333176af4f
from workalendar.europe import Germany

DATE_FORMAT = '%d.%m.%Y'


def highlight_max(s):
    is_large = s.nlargest(6).values
    return ['background-color: #e06666' if v in is_large else '' for v in s]


def format_digits(val):
    return 'number-format: #,##0.0000'


class ExcelWriter:

    def __init__(self, daily_rsls):
        self.daily_rsls = daily_rsls
        self.last_stock_exchange_days_of_the_week = ExcelWriter.find_last_stock_exchange_days_of_the_week()
        print(self.last_stock_exchange_days_of_the_week)
        self.weekly_rsls = self.to_weekly_rsls()

    @staticmethod
    def find_last_stock_exchange_days_of_the_week():
        last_quarter = date.today() - relativedelta(months=3)
        last_fridays_of_the_week = pd.date_range(start=str(last_quarter), end=str(date.today()), freq='W-FRI')

        last_stock_exchange_days_of_the_week = list()
        for last_friday_of_the_week in last_fridays_of_the_week:
            if Germany().is_working_day(ExcelWriter.to_date(last_friday_of_the_week)):
                last_stock_exchange_days_of_the_week.append(last_friday_of_the_week.strftime(DATE_FORMAT))
            else:
                print("Found public holiday: " + last_friday_of_the_week.strftime(DATE_FORMAT))
                prior_stock_exchange_day_of_the_week = ExcelWriter.find_prior_stock_exchange_day(last_friday_of_the_week)
                last_stock_exchange_days_of_the_week.append(prior_stock_exchange_day_of_the_week.strftime(DATE_FORMAT))

        return last_stock_exchange_days_of_the_week

    @staticmethod
    def find_prior_stock_exchange_day(stock_exchange_day):
        de_calendar = Germany()
        for prior_day in range(1, 5):
            prior_stock_exchange_day = stock_exchange_day - pd.DateOffset(prior_day)
            if de_calendar.is_working_day(ExcelWriter.to_date(prior_stock_exchange_day)):
                print("Change stock exchange day to: " + prior_stock_exchange_day.strftime(DATE_FORMAT))
                return stock_exchange_day

        raise ValueError("Oops, no working day found")

    @staticmethod
    def to_date(exchange_day):
        return date(exchange_day.year, exchange_day.month, exchange_day.day)

    def to_weekly_rsls(self):
        weekly_rls = self.daily_rsls.filter(self.last_stock_exchange_days_of_the_week)
        return weekly_rls[weekly_rls.columns[::-1]]

    def write_weekly(self):
        writer = pd.ExcelWriter("hdax.xlsx", engine='xlsxwriter')

        # write data to excel
        self.weekly_rsls.sort_values(self.last_stock_exchange_days_of_the_week[-1], ascending=False) \
            .style \
            .apply(highlight_max) \
            .applymap(format_digits) \
            .to_excel(writer, sheet_name="rsl")

        # Modify layout
        worksheet = writer.sheets['rsl']
        worksheet.set_column(0, 0, 40)  # company names - longest at the moment Münchener Rückversicherungs-Gesellschaft AG

        column_length = 12  # only the length of date is relevant dd.mm.yyyy
        for column in self.weekly_rsls:
            col_idx = 1 + self.weekly_rsls.columns.get_loc(column)
            # print(col_idx)
            worksheet.set_column(col_idx, col_idx, column_length)

        writer.close()
