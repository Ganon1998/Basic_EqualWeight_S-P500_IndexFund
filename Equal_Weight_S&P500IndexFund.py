# download these libraries since Colab doesn't come with these out of the box

import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math


def chunks(lst, n):
  for i in range (0, len(lst), n):
    yield lst[i:i + n]

# Python script that will accept the value of your portfolio and tell you how mahy shares of each S&P500 constituent you should purchase to get an equal-weight version of the index fund
# In other words, this program tells you which and how many shares of each S&P500 stock you should buy (the highests S&P 500 companies, like Apple, are given a lower weight and the lowests S&P 500 companies are given higher weight)

# NOTE: The S&P 500 calling cards used is static and is from DEC 2020. I tried to get live S&P 500 calling cards but sadly doing so would require a subscription :( (I can only do 5 requests for free a day and after that 5th request I'll be charged $0.10 per request)

stocks = pd.read_csv('sp_500_stocks.csv')

# test api calls just to make sure api calls work in reallife, it returns randomized data
# with that said, the Apple stock price is only off by $1.50 so it's pretty legit
from secrets import IEX_CLOUD_API_TOKEN
symbol = 'AAPL'
api_url = f"https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}"
data = requests.get(api_url).json()
price = data['latestPrice']
market_cap = data['marketCap']

# create dataframe column for stocks
mycolumns = ['Ticker', 'Stock Price', 'Market Cap', 'Number of Shares to Buy']

# use Batch API calls to speed up request times and reduce load on IEX_Cloud's end of things
# it does so by splitting our list of stocks into groups of 100 instead of asking all 500 in one go.
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []

# create string of calling cards seperated by commas
for i in range(0, len(symbol_groups)):
  symbol_strings.append(','.join(symbol_groups[i]))

# make batched api call for each calling card
final_dataframe = pd.DataFrame(columns = mycolumns)
for symbol_string in symbol_strings:
  batch_api_call_url = f"https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}"
  data = requests.get(batch_api_call_url).json()

  for symbol in symbol_string.split(','):
    if symbol == 'HFC' or symbol == 'VIAC' or symbol == 'WLTW' or symbol == 'DISCA':
       continue
    # we're gonna have to parse things differently now
    final_dataframe = final_dataframe.append(
        pd.Series([
            symbol,
            data[symbol]['quote']['latestPrice'],
            data[symbol]['quote']['marketCap'],
            'N/A'
        ], index = mycolumns
        ),
        ignore_index=True
    )

# ask user how much money they want to invest in each stock in the S&P500. Every stock in the portfolio will ahve the same position size (how much money you're gonna invest in each stock)
# Ex: if youre portfolio size is $1000, each stock will be $1.98
portfolio_size = input('Enter the value of your portfolio ')
val = 0.0
while True:
  try:
    val = float(portfolio_size)
    break
  except ValueError:
    print('Portfolio value is invalid. Please enter an integer.')
    portfolio_size = input('Enter the value of your portfolio ')


position_size = val/len(final_dataframe.index)

# how many shares of each stock need to be purchased to get to that position size
# Ex: I_want_Apple_Shares = position_size / price_of_apple_stock
# if I_want_Apple_Shares ends up being s decimal, we must round down. If we round up, we'll be buying more stocks than we can afford
for i in range(0, len(final_dataframe.index)):
  final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / final_dataframe.loc[i, 'Stock Price'])


# store portfolio in excel file
writer = pd.ExcelWriter('recommended_trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

# format columms
writer.sheets['Recommended Trades'].set_column('A:A', 18, string_format)
writer.sheets['Recommended Trades'].set_column('B:B', 18, string_format)
writer.sheets['Recommended Trades'].set_column('C:C', 18, string_format)
writer.sheets['Recommended Trades'].set_column('D:D', 18, string_format)


column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Cap', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer.save()
