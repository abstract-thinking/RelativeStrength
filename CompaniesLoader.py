import yaml
from yaml.loader import SafeLoader


class CompaniesLoader:

    @staticmethod
    def load():
        with open('./hdax.yml', encoding="utf-8") as f:
            companies = yaml.load(f, Loader=SafeLoader)
        return companies
