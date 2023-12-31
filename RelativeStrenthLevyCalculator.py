import pandas as pd
import pandas_ta as ta
import yfinance as yf

DATE_FORMAT = '%d.%m.%Y'
PERIOD_IN_DAYS = 130


class RelativeStrengthLevyCalculator:

    def __init__(self, companies):
        self.companies = companies

    def calculate(self):
        frames = list()
        for company in self.companies:
            print("Calculating RSL for " + company["name"])
            # data = yf.download(tickers=fruits, period='1y', interval="1d", group_by='date')
            yahoo_symbol = company["symbols"]["yahoo"]
            data = yf.Ticker(yahoo_symbol)
            history = data.history(period='1y', interval="1d")
            if history.size < PERIOD_IN_DAYS:
                print("Too less data for " + yahoo_symbol)
                continue

            # TODO: This sucks: When a dividend is paid the price are missing
            # 09. Mai 2023 64,02 64,30 63,36 64,30 64,30 153.609
            # 08. Mai 2023 - -    -    -    -    -
            # 08. Mai 2023 1.45 Dividende
            try:
                close = history.Close.copy()
                close.dropna(inplace=True)
                rsl = close / ta.sma(close, length=PERIOD_IN_DAYS)
            except TypeError:
                print("Could not calculate RSL for " + yahoo_symbol)
                continue
            rsl.name = company["name"]
            rsl.dropna(inplace=True)
            rsl.index = rsl.index.strftime(DATE_FORMAT)

            print("Calculated RSL for " + company["name"] )
            frames.append(rsl.to_frame().transpose())

        result = pd.concat(frames)
        result = result[result.columns[::-1]]
        #print(result)
        return result
