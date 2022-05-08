import yaml
from yaml.loader import SafeLoader


class CompaniesLoader:

    @staticmethod
    def load():
        with open('./hdax.yml') as f:
            companies = yaml.load(f, Loader=SafeLoader)
        return companies
