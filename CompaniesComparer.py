from datetime import date
from dateutil.relativedelta import relativedelta, FR

import requests
import xlrd

from CompaniesMismatchException import CompanyMismatchException

URL_TEMPLATE = 'https://www.dax-indices.com/documents/dax-indices/Documents/Resources/WeightingFiles/Composition/{0:%Y}/{0:%B}/HDAX_ICR.{0:%Y%m%d}.xls'
# https://www.stoxx.com/data-index-details?symbol=HKDX
HDAX_COMPOSITION_URL = 'https://www.stoxx.com/documents/stoxxnet/Documents/Indices/Current/Composition_Files/icr_hkdx.xls'

class CompaniesComparer:

    def __init__(self):
        self.international_securities_identification_numbers = self.fetch_hdax()

    @staticmethod
    def fetch_hdax():
        response = requests.get(HDAX_COMPOSITION_URL, verify=False)
        response.raise_for_status()

        workbook = xlrd.open_workbook(file_contents=response.content)
        worksheet = workbook.sheet_by_index(1)

        isins = list()
        for i in range(worksheet.nrows):
            row = worksheet.row_values(i)
            # print(row)
            if row[1] == 'HDAX PERFORMANCE-INDEX':
                isins.append(row[5])

  #      return worksheet.col_value#s(5, 106)
        print("Found number of isins: " + str(len(isins)))
        return isins

    def compare(self, companies):
        self.find_new_companies_inside_hdax(companies)
        self.find_removed_companies_from_hdax(companies)
        print('HDAX is up-to-date')

    def find_new_companies_inside_hdax(self, companies):
        current_isins = self.international_securities_identification_numbers.copy()
        for company in companies:
            for isin in current_isins:
                if company["isin"] == isin:
                    # print('Remove current ' + str(company))
                    current_isins.remove(isin)
                    continue

        if current_isins:
            print("HDAX has changed!")
            print('New companies inside the HDAX')
            print(current_isins)
            raise CompanyMismatchException()

    def find_removed_companies_from_hdax(self, companies):
        removed_companies = companies.copy()
        for isin in self.international_securities_identification_numbers:
            for company in removed_companies:
                if company["isin"] == isin:
                    removed_companies.remove(company)
                    continue

        if removed_companies:
            print("HDAX has changed!")
            print('Removed companies from the HDAX')
            removed = list()
            for removed_company in removed_companies:
                removed.append(removed_company['isin'])
            print(removed)
            print(removed_companies)
            raise CompanyMismatchException()
