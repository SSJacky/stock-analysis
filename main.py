import os
import requests
from io import StringIO
import pandas as pd
import datetime
import time


def crawl_price(date):
    """Fetch stock data from TWSE."""
    url = f"http://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date={date.strftime('%Y%m%d')}&type=ALL"
    r = requests.post(url)

    # Clean and process data
    ret = pd.read_csv(StringIO("\n".join([
        i.translate({ord(c): None for c in ' '})
        for i in r.text.split('\n')
        if len(i.split('",')) == 17 and i[0] != '='
    ])), header=0)

    ret = ret.set_index('證券代號')
    ret['成交金額'] = ret['成交金額'].str.replace(',', '', regex=True)
    ret['成交股數'] = ret['成交股數'].str.replace(',', '', regex=True)
    return ret


data = {}
n_days = 9
date = datetime.datetime.now()
fail_count = 0
allow_continuous_fail_count = 5

while len(data) < n_days:
    print(f'Fetching data for {date.date()}...')
    try:
        data[date.date()] = crawl_price(date)
        print(f'Success! Data for {date.date()} added.')
        fail_count = 0
    except Exception as e:
        print(f'Fail! Check if {date.date()} is a holiday. Error: {e}')
        fail_count += 1
        if fail_count == allow_continuous_fail_count:
            raise

    date -= datetime.timedelta(days=1)
    time.sleep(10)

# Define save location
save_path = "/Users/siashenjie/Desktop/"

if data:
    for stock_date, df in data.items():
        filename = os.path.join(save_path, f"twse_stock_{stock_date}.xlsx")
        df.to_excel(filename, engine='openpyxl')
        print(f"Data for {stock_date} saved to {filename}")
