# -*- coding: utf-8 -*-
"""
Created on Tue May  1 19:24:15 2018

@author: kim
"""

#종목마스터 구현하기

import win32com.client
import pandas as pd
import requests
from datetime import datetime, timedelta
from io import BytesIO
 
def getcodelist():
    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
    code = []
    name = []
    market = []
    kind = []
    for i, j in enumerate(codeList):
        if j[0] == 'A': #ETN 제외 ETN은 Q로 시작
            code.append(j)
            name.append(objCpCodeMgr.CodeToName(j)) #코드 입력 종목명 출력
            market.append('거래소') #시장 유형
            secondCode = objCpCodeMgr.GetStockSectionKind(j) #구분
            if secondCode == 1 or secondCode == 6 or secondCode == 13:
                kind.append('주식')
            elif secondCode == 2 or secondCode == 3 or secondCode == 4 or secondCode == 5:
                kind.append('주식(투자회사)')
            elif secondCode == 10:
                kind.append('ETF')
            else:
                kind.append('기타')
    
    for i, j in enumerate(codeList2):
        if j[0]== 'A': #ETN 제외 ETN은 Q로 시작
            code.append(j)
            name.append(objCpCodeMgr.CodeToName(j)) #코드 입력 종목명 출력
            market.append('코스닥') #시장 유형
            secondCode = objCpCodeMgr.GetStockSectionKind(j) #구분
            if secondCode == 1 or secondCode == 6 or secondCode == 13:
                kind.append('주식')
            elif secondCode == 2 or secondCode == 3 or secondCode == 4 or secondCode == 5:
                kind.append('주식(투자회사)')
            elif secondCode == 10:
                kind.append('ETF')
            else:
                kind.append('기타')
    
    df = pd.DataFrame({'종목코드':code, '종목명':name, '시장':market, '구분':kind}, columns = ['종목코드', '종목명','시장','구분','상폐여부','폐지일'])
    df.set_index('종목코드', inplace=True)
    return df

def delisting(market, date=None):
    #market 은 거래소 또는 코스닥 입력
    if market =='거래소':
        mkcode = 'STK'
    elif market == '코스닥':
        mkcode ='KSQ'
        
    if date == None:
        date = datetime.today().strftime('%Y%m%d')

    url1 = 'http://marketdata.krx.co.kr/contents/COM/GenerateOTP.jspx'
    data1 = {
        'name': 'fileDown',
        'filetype': 'xls',
        'url': 'MKD/04/0406/04060600/mkd04060600',
        'market_gubun': mkcode, #STK =거래소 KSQ =코스닥
        'isu_cdnm': '전체',
        'isu_cd': '',
        'isu_nm': '',
        'isu_srt_cd': '',
        'fromdate': '20110401',
        'todate': date,
        'pagePath': '/contents/MKD/04/0406/04060600/MKD04060600.jsp'
    }

    r = requests.post(url1, data1)
    code = r.content

    url2 = 'http://file.krx.co.kr/download.jspx'
    data2 = {
        'code': code,
    }
    r = requests.post(url2, data2)
    df = pd.read_excel(BytesIO(r.content), header=0, thousands=',')
    df.set_index('종목코드', inplace=True)
    df['시장']=market
    df['상폐여부']='폐지'
    df['종목명']=df['기업명']
    df['구분']='주식'
    return df

stock_list = getcodelist()
Delist1 = delisting('거래소')
Delist2 = delisting('코스닥')

stock_list = pd.concat([stock_list,Delist1])
stock_list = pd.concat([stock_list,Delist2])
stock_list = pd.DataFrame(stock_list, columns = ['종목명','시장','구분','상폐여부','폐지일'])

for i in range(0, len(stock_list)):
    if stock_list.index[i][-1] != '0':
        stock_list['구분'][i] = '주식(우선주)' #우선주 구분
    if stock_list['구분'][i] == '주식':
        for j in range(0, len(stock_list['종목명'][i])-1):
            if stock_list['종목명'][i][j:j+2] == '스팩':
                stock_list['구분'][i]='주식(스팩)' # 스팩 구분
    if stock_list['구분'][i] == '주식':
        for j in range(0, len(stock_list['종목명'][i])): 
            if stock_list['종목명'][i][j] == '0' or stock_list['종목명'][i][j] == '1' or stock_list['종목명'][i][j] == '2' or stock_list['종목명'][i][j] == '3' or stock_list['종목명'][i][j] == '4'or stock_list['종목명'][i][j] == '5'or stock_list['종목명'][i][j] == '6'or stock_list['종목명'][i][j] == '7'or stock_list['종목명'][i][j] == '8'or stock_list['종목명'][i][j] == '9':
                if stock_list['종목명'][i] != '한세예스24홀딩스' and stock_list['종목명'][i] != 'E1' and stock_list['종목명'][i] != '까페24' and stock_list['종목명'][i] != '예스24' and stock_list['종목명'][i] != '3S' and stock_list['종목명'][i] != '3노드디지탈':
                    stock_list['구분'][i] ='주식(투자회사)' #투자회사 구분
                    
stock_list.to_excel('stock_list.xlsx')
                    

