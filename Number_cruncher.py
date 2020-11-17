import pandas as pd
import yfinance as yf
import time
import datetime
from openpyxl import load_workbook


# Function ensures api requests reflect the most up to dat results
def get_dates():
    # Get today's date
    today = datetime.date.today()
    now = datetime.datetime.now()
    yesterday = now - datetime.timedelta(days=1)
    # Create time object for after market close
    market_close = datetime.time(21, 0, 0)

    # TODO: Fix this so it's more useful
    # Check if weekday
    if today.weekday() < 5:
        # Check if current time is after market close
        return today.strftime('%Y-%m-%d'), yesterday

    # If it's the weekend, check if saturday
    elif today.weekday() == 5:
        return today.strftime('%Y-%m-%d'), yesterday
    # If we get here it must be sunday
    else:
        friday = now - datetime.timedelta(days=2)
        return today.strftime('%Y-%m-%d'), friday.strftime('%Y-%m-%d')


# TODO: Catch errors here.
def load_file():
    # Load file into data reader
    ExelFile = pd.ExcelFile("Combined BoA-ML Portfolios 6-30-20 v.1.xlsx")
    # Select sheet to load desired data
    sheet3 = ExelFile.parse("Sheet3")
    # Select column and amount of rows to read into dataframe
    data = sheet3["Ticker Symbols"][:408]
    # Strip extra whitespace, remove duplicates
    cleaned_data = data.dropna().str.strip().drop_duplicates()

    return cleaned_data

def download_stockPrices(cleaned_data):
    end, start = get_dates()

    all_Tks_Raw = yf.download(list(cleaned_data), start=start, end=end, progress=True)
    all_Tks = all_Tks_Raw["Adj Close"].transpose()

    all_Tks.to_csv("Adjusted Close.csv")

if __name__ == "__main__":

    data = load_file()
    download_stockPrices(data)







