# -*- coding: utf-8 -*-
"""
Created on Wed May 23 22:48:11 2018

@author: kim
"""

import time
import pandas as pd
import numpy as np
import requests
from io import BytesIO
import win32com.client
import datetime
stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")
stock_list=pd.read_excel('stock_list.xlsx')
stock_list.set_index('종목코드', inplace=True)

def get_datetime(x):
    logdate = str(x['Date'])
    year = int(logdate[:4])
    month = int(logdate[4:6])
    day = int(logdate[6:8])
    return datetime.datetime(year, month, day)


def get_datetime2(x):
    logdate = str(x['일자'])
    year = int('20'+logdate[:2])
    month = int(logdate[3:5])
    day = int(logdate[6:8])
    return datetime.datetime(year, month, day)

def process(start, end, code):
    stock_chart.SetInputValue(0, '%s'%code)
    stock_chart.SetInputValue(1, ord('1'))
    stock_chart.SetInputValue(2, start)
    stock_chart.SetInputValue(3, end)
    stock_chart.SetInputValue(5, [0, 2, 3, 4, 5, 8])
    stock_chart.SetInputValue(6, ord('D'))
    stock_chart.SetInputValue(9, ord('1'))
    stock_chart.BlockRequest()
    count = stock_chart.GetHeaderValue(3)
    for ii in range(count):
        caller_dates.append(stock_chart.GetDataValue(0, ii))
        caller_opens.append(stock_chart.GetDataValue(1, ii))
        caller_highs.append(stock_chart.GetDataValue(2, ii))
        caller_lows.append(stock_chart.GetDataValue(3, ii))
        caller_closes.append(stock_chart.GetDataValue(4, ii))
        caller_vols.append(stock_chart.GetDataValue(5, ii))
    return count


def daum(code, i):
    url = 'http://finance.daum.net/item/quote_yyyymmdd_sub.daum?page=%s&code=%s&modify=0' %(i,code)
    dfs = pd.read_html(url)
    dfs = dfs[0]
    dfs.columns = ['일자', '시가', '고가', '저가', '종가', '전일비', '등락률', '거래량']
    dfs = dfs[1:]
    dfs['시가'] = pd.to_numeric(dfs['시가'], errors='coerce')
    dfs['고가'] = pd.to_numeric(dfs['고가'], errors='coerce')
    dfs['저가'] = pd.to_numeric(dfs['저가'], errors='coerce')
    dfs['종가'] = pd.to_numeric(dfs['종가'], errors='coerce')
    dfs['거래량'] = pd.to_numeric(dfs['거래량'], errors='coerce')
    dfs = dfs[dfs['종가'] > 0]
    dfs.set_index('일자', inplace=True)
    del dfs['전일비']
    del dfs['등락률']
    return dfs
def daum2(code, i):
    url = 'http://finance.daum.net/item/quote_yyyymmdd_sub.daum?page=%s&code=%s&modify=1' %(i,code)
    dfs = pd.read_html(url)
    dfs = dfs[0]
    dfs.columns = ['일자', '시가', '고가', '저가', '종가', '전일비', '등락률', '거래량']
    dfs = dfs[1:]
    dfs['시가'] = pd.to_numeric(dfs['시가'], errors='coerce')
    dfs['고가'] = pd.to_numeric(dfs['고가'], errors='coerce')
    dfs['저가'] = pd.to_numeric(dfs['저가'], errors='coerce')
    dfs['종가'] = pd.to_numeric(dfs['종가'], errors='coerce')
    dfs['거래량'] = pd.to_numeric(dfs['거래량'], errors='coerce')
    dfs = dfs[dfs['종가'] > 0]
    dfs.set_index('일자', inplace=True)
    del dfs['전일비']
    del dfs['등락률']
    return dfs


stock_start = datetime.datetime.now()
stock_end = datetime.datetime(2010,4,1)


caller_dates = []
caller_opens = []
caller_highs = []
caller_lows = []
caller_closes = []
caller_vols = []
caller_df=[]
j =0
while j < len(stock_list.index):
    code = stock_list.index[j]
    if stock_list.loc[code, '상폐여부'] == '상장':
        print(j, code)
        try:
            df_non = daum(code[1:], 1)
            for i in range(2, 68):
                df2 = daum(code[1:], i)
                df_non = pd.concat([df_non, df2])
        except ValueError:
            print('오류 재시도')
            continue
        if len(df_non.index) >= 30:
            if df_non.index[29] == df_non.index[30]:
                df_non=pd.concat([df_non.iloc[:29], df_non.iloc[30:]])
        df_non= df_non.reset_index()
        df_non["Datetime"] = df_non.apply(get_datetime2, axis=1)
        df_non.set_index('Datetime',inplace=True)
#        df_non = df_non.sort_index(ascending=True)
                
        caller_dates = []
        caller_opens = []
        caller_highs = []
        caller_lows = []
        caller_closes = []
        caller_vols = []
        
        ## 제한 15초당 60건
        ## 1초당 4건.
        cnt = process(stock_start.strftime('%Y%m%d'), stock_end.strftime('%Y%m%d'), '%s'%stock_list.index[j])
        chartData = {'Date': caller_dates, 'Open': caller_opens, 'High': caller_highs, 'Low': caller_lows, 'Close': caller_closes, 'Vol': caller_vols }
        df_edit = pd.DataFrame(chartData, columns=['Date', 'Open', 'High', 'Low', 'Close', 'Vol'])
        df_edit["Datetime"] = df_edit.apply(get_datetime, axis=1)
        df_edit.sort_values(by="Datetime", ascending=True, inplace=True)
        df_edit.set_index(df_edit["Datetime"], inplace=True)
        del df_edit['Datetime']
        df_non['edit'] = df_edit['Close']
    elif stock_list.loc[code, '상폐여부'] == '폐지':
        print(j, code, stock_list.loc[code, '상폐여부'])
        try:
            df_non = daum(code[1:], 2)
            for i in range(3, 68):
                df2 = daum(code[1:], i)
                df_non = pd.concat([df_non, df2])  
            df_non= df_non.reset_index()
            df_non["Datetime"] = df_non.apply(get_datetime2, axis=1)
            df_non.set_index('Datetime',inplace=True)
        except ValueError:
            print('오류 재시도')
            continue
        try:
            df_edit = daum2(code[1:], 2)
            for i in range(3,68):
                df1 = daum2(code[1:], i)
                df_edit = pd.concat([df_edit, df1])   
            df_edit= df_edit.reset_index()
            df_edit["Datetime"] = df_edit.apply(get_datetime2, axis=1)
            df_edit.set_index('Datetime',inplace=True)
        except ValueError:
            print('오류 재시도')
            continue
        df_non['edit'] = df_edit['종가']

    
    df_non['비율1'] = round(df_non['edit']/df_non['종가'], 2)
    
    for i in range(0, len(df_non)-10):
        if df_non.iloc[i, -1] == df_non.iloc[i+5, -1]:
            df_non.iloc[i+1, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+2, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+3, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+4, -1] = df_non.iloc[i, -1]
        
        if df_non.iloc[i, -1] == df_non.iloc[i+10, -1]:
            df_non.iloc[i+1, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+2, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+3, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+4, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+5, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+6, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+7, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+8, -1] = df_non.iloc[i, -1]
            df_non.iloc[i+9, -1] = df_non.iloc[i, -1]
    df_non['비율2'] = 1
#    df_non= df_non.sort_index(ascending=True)
    for i in range(0, len(df_non)-1):
        day = df_non.index[i]
        day1 = df_non.index[i+1]
        df_non.loc[day, '비율2'] = round(df_non.loc[day1, '비율1']/df_non.loc[day, '비율1'],2)
        if df_non.loc[day, '비율2'] != 1:
            caller_df.append([code, day, df_non.loc[day, '비율2']])
        
    if j == 1300:
        df_result = df_non.iloc[:, 7:8]
        df_result[code] = df_result['비율1']
        del df_result['비율1']
    else:
        df_result[code] = df_non['비율1']
    j+=1
    if j%100 == 0:
        df_result.to_csv('D:/data/adjprice2_all.csv',encoding='euc-kr')
        df_adj = pd.DataFrame(caller_df, columns=['code','date','ratio'])
        df_adj.to_csv('D:/data/adjprice2.csv',encoding='euc-kr')

df_result.to_csv('D:/data/adjprice2_all.csv',encoding='euc-kr')
df_adj = pd.DataFrame(caller_df, columns=['code','date','ratio'])
df_adj.to_csv('D:/data/adjprice2.csv',encoding='euc-kr')
