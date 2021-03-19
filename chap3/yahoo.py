#야후 파이낸스, 판다스데이터리더 라이브러리 설치
#pip install yfinance
#pip install pandas-datareader

import pandas as pd
import yfinance as yf
from pandas_datareader import data as pdr
yf.pdr_override()

sec = pdr.get_data_yahoo('005930.KS',start='2020-03-04')
msft = pdr.get_data_yahoo('MSFT',start='2020-03-04')

print(sec.head(10))