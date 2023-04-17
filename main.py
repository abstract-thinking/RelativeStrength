from CompaniesComparer import CompaniesComparer
from CompaniesLoader import CompaniesLoader
from ExcelWriter import ExcelWriter
from RelativeStrenthLevyCalculator import RelativeStrengthLevyCalculator

if __name__ == "__main__":
    companies = CompaniesLoader.load()
    CompaniesComparer().compare(companies)

    daily_rsls = RelativeStrengthLevyCalculator(companies).calculate()
    ExcelWriter(daily_rsls).write_weekly()
    print("Wrote result to file. Done!")
