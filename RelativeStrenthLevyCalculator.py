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
                print("No less data for " + yahoo_symbol)
                continue

            try:
                rsl = history.Close / ta.sma(history.Close, length=PERIOD_IN_DAYS)
            except TypeError:
                print("Could not calculate RSL for " + yahoo_symbol)
                continue
            rsl.name = company["name"]
            rsl.dropna(inplace=True)
            rsl.index = rsl.index.strftime(DATE_FORMAT)

            frames.append(rsl.to_frame().transpose())

        result = pd.concat(frames)
        result = result[result.columns[::-1]]
        #print(result)
        return result
