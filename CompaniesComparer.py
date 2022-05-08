from datetime import date
from dateutil.relativedelta import relativedelta, FR

import requests
import xlrd

from CompaniesMismatchException import CompanyMismatchException

URL_TEMPLATE = 'https://www.dax-indices.com/documents/dax-indices/Documents/Resources/WeightingFiles/Composition/{0:%Y}/{0:%B}/HDAX_ICR.{0:%Y%m%d}.xls'


class CompaniesComparer:

    def __init__(self):
        self.international_securities_identification_numbers = self.fetch_hdax()

    @staticmethod
    def fetch_hdax():
        day = date.today()
        url = URL_TEMPLATE.format(day)
        print(url)
        response = requests.get(url)
        if response.status_code == 404:
            # Be prepared that maybe during the week the excel sheet does not exist, use the last Friday
            print('Excel sheet not found. Retry!')
            day = date.today() + relativedelta(weekday=FR(-1))
            url = URL_TEMPLATE.format(day)
            print(url)
            response = requests.get(url)

        workbook = xlrd.open_workbook(file_contents=response.content)
        worksheet = workbook.sheet_by_index(1)
        return worksheet.col_values(5, 106)

    def compare(self, companies):
        for company in companies:
            for isin in self.international_securities_identification_numbers:
                if company["isin"] == isin:
                    self.international_securities_identification_numbers.remove(isin)
                    continue

        if self.international_securities_identification_numbers:
            print("HDAX has changed!")
            print(self.international_securities_identification_numbers)
            raise CompanyMismatchException()

        print('HDAX is up-to-date')
