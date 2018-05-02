# -*- coding: utf-8 -*-
"""
Created on Wed May  2 14:04:42 2018

@author: kim
"""

#종목리스트 업데이트
import win32com.client
import pandas as pd
from datetime import datetime
import requests

def codelist_update():
    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥
    code = []
    name = []
    market = []
    kind = []
    for i, j in enumerate(codeList):
        if j[0]== 'A': #ETN 제외 ETN은 Q로 시작
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

today = codelist_update() # 현재 상장주식 목록 대신 API에서 수신
stock_list = pd.read_excel('stock_list.xlsx') #기존 stock_list 불러오기
stock_list.set_index('종목코드', inplace=True) # index설정 (불러올때마다 해주기)


yesterday = stock_list[(pd.isnull(stock_list['상폐여부']))] #비교를 위해 기존 stock_list에서 상장종목만 yesterday에 저장

today['비교']= yesterday['종목명'] #yasterday 종목명을 today '비교' 열에 저장
yesterday['비교'] =today['종목명'] #today 종목명을 yesterday '비교' 열에 저장
#today에는 있으나 yesterday에는 없는것 -> today['비교']값이 nan 인 경우 -> 신규상장
#yesterday에는 있으나, today에는 없는것 -> yesterday['비교']값이 nan 인 경우 -> 상장폐지

add = today[(pd.isnull(today['비교']))] #신규상장 종목
sub = yesterday[(pd.isnull(yesterday['비교']))] #상장폐지 종목

#상장폐지 종목 stock_list에서 값 변환해주기  
for i in range(0, len(sub)):
    stock_list.loc[sub.index[i]]['상폐여부'] = '폐지'
    stock_list.loc[sub.index[i]]['폐지일'] = datetime.today().strftime('%Y/%m/%d')
    
#신규상장 종목 stock_list에 추가
del add['비교']
stock_list = pd.concat([stock_list, add])

stock_list.to_excel('stock_list.xlsx')



