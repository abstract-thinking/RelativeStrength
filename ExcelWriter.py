import pandas as pd

from datetime import date, datetime
from dateutil.relativedelta import relativedelta

DATE_FORMAT = '%d.%m.%Y'


def highlight_max(s):
    is_large = s.nlargest(6).values
    return ['background-color: red' if v in is_large else '' for v in s]


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
        for day in self.fridays:
            if day in self.daily.columns:
                print(day)
                dates.append(day)
            else:
                date_not_found = datetime.strptime(day, DATE_FORMAT)
                new_date = (date_not_found - relativedelta(days=1)).strftime(DATE_FORMAT)
                if new_date in self.daily.columns:
                    print("Not a Friday: " + new_date)
                    dates.append(new_date)
                else:
                    print("Better logic needed!")

        tmp_weekly = self.daily.filter(dates)
        return tmp_weekly[tmp_weekly.columns[::-1]]

    def write_weekly(self):
        writer = pd.ExcelWriter("hdax.xlsx", engine='xlsxwriter')
        self.weekly.sort_values(self.fridays[-1], ascending=False).style.apply(highlight_max).to_excel(
            writer, sheet_name="rsl", float_format='%06.4f')
        worksheet = writer.sheets['rsl']

        worksheet.set_column(0, 0, 40)  # index
        for column in self.weekly:
            # column_length = max(weekly[column].astype(str).map(len).max(), len(column))
            column_length = 13  # only the length of date is relevant
            col_idx = 1 + self.weekly.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, column_length)

        writer.save()