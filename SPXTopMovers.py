import pandas as pd
import win32com.client as win32
from pandas_datareader import data
import matplotlib.pyplot as plt
import numpy as np
import urllib.request
from datetime import date
from datetime import timedelta
from bs4 import BeautifulSoup

tickers_page_address = 'https://en.wikipedia.org/wiki/S%26P_100'
tickers_page_html = urllib.request.urlopen(tickers_page_address)
soup = BeautifulSoup(tickers_page_html,'html.parser')
all_ticker_text = soup.find_all('td')

tickers = []
names = []
for x in range(1, 100):
    tickers.append((all_ticker_text[x*2+13].get_text())[:-1])
    names.append((all_ticker_text[x*2+14].get_text())[:-1])

today = str(date.today())
start_date = str(date.today()-timedelta(days=5))

price_data = []
valid_tickers = []
valid_names = []
for i in range(0, 99):
    try: 
        print(tickers[i])
        price_data.append(data.DataReader(tickers[i],'iex',start_date,today).close)
        valid_tickers.append(tickers[i])
        valid_names.append(names[i])
    except: 
        pass
print('Out of loop')

df_price_data = pd.DataFrame(price_data)
movements = pd.DataFrame()
for i in range(0, len(valid_tickers)-1):
    percent_move = np.log(float(df_price_data.iloc[i,-1])/float(df_price_data.iloc[i,-2]))
    movements = movements.append({'ticker': valid_tickers[i], 'name': valid_names[i], 'change': "{0:.2f}%".format(percent_move*100), 'absChange': np.abs(percent_move)}, ignore_index=True)

movements_top10 = movements.sort_values(by=['absChange'], ascending=False)[:10]
movements_top10 = movements_top10.reset_index(drop=True)
movements_top10.index = movements_top10.index + 1

del movements_top10['absChange']

movements_top10 = movements_top10[['ticker','name','change']]
movements_top10.columns = ['Ticker','name',list(df_price_data)[-1]]

import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'james9995@hotmail.co.uk'
mail.Subject = 'S&P 100 Top 10 Movers '+list(df_price_data)[-1]
mail.Body = 'S&P 100 Top 10 Movers'
html1 = movements_top10.to_html()
mail.HTMLBody = html1
mail.Send()
