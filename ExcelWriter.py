import pandas as pd

from datetime import date, datetime
from dateutil.relativedelta import relativedelta

DATE_FORMAT = '%d.%m.%Y'


def highlight_max(s):
    is_large = s.nlargest(6).values
    return ['background-color: #e06666' if v in is_large else '' for v in s]


def format_digits(val):
    return 'number-format: #,##0.0000'


class ExcelWriter:

    def __init__(self, daily):
        self.daily = daily

        self.fridays = self.find_fridays()
        self.weekly = self.to_weekly()

    @staticmethod
    def find_fridays():
        last_quarter = date.today() - relativedelta(months=3)
        return pd.date_range(start=str(last_quarter), end=str(date.today()), freq='W-FRI').strftime(
            DATE_FORMAT).tolist()

    def to_weekly(self):
        dates = list()
        for friday in self.fridays:
            if friday in self.daily.columns:
                # print(friday)
                dates.append(friday)
            else:
                friday_not_found = datetime.strptime(friday, DATE_FORMAT)

                previous_day = friday_not_found - relativedelta(days=1)
                end_date = friday_not_found - relativedelta(days=3)

                while previous_day >= end_date:
                    this_day = previous_day.strftime(DATE_FORMAT)
                    if this_day in self.daily.columns:
                        print("Use a previous day: " + this_day)
                        dates.append(this_day)
                        break

                    previous_day -= relativedelta(days=1)

        tmp_weekly = self.daily.filter(dates)
        return tmp_weekly[tmp_weekly.columns[::-1]]

    def write_weekly(self):
        writer = pd.ExcelWriter("hdax.xlsx", engine='xlsxwriter')

        # write data to excel
        self.weekly.sort_values(self.fridays[-1], ascending=False) \
            .style \
            .apply(highlight_max) \
            .applymap(format_digits) \
            .to_excel(writer, sheet_name="rsl")

        # Modify layout
        worksheet = writer.sheets['rsl']
        worksheet.set_column(0, 0, 40)  # company names - longest at the moment Münchener Rückversicherungs-Gesellschaft AG

        column_length = 12  # only the length of date is relevant dd.mm.yyyy
        for column in self.weekly:
            col_idx = 1 + self.weekly.columns.get_loc(column)
            # print(col_idx)
            worksheet.set_column(col_idx, col_idx, column_length)

        writer.close()
