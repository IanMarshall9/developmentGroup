import pandas as pd
import yfinance as yf
import time
import datetime
from openpyxl import load_workbook


def get_dates():
    # Get today's date
    today = datetime.date.today() #.strftime('%Y-%m-%d')
    market_close = datetime()

    #Check if weekday
    if today.weekday() < 5:

    yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
    yesterday = yesterday.strftime('%Y-%m-%d')

    return today, yesterday

data = pd.ExcelFile("Combined BoA-ML Portfolios 6-30-20 v.1.xlsx")
sheet3 = data.parse("Sheet3")

ticker_syms = sheet3["Ticker Symbols"][1:220].str.strip()

test = sheet3["Ticker Symbols"][1:10].str.strip()

aapl = yf.Ticker(test[1]).history(period="1h")

print(aapl["Close"])








