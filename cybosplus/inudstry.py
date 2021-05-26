#대신증권 API(CybosPlus)를 통해 회사명, 종목코드, 현재가, 시가총액, 코스피/코스닥 값을 불러와 db로 저장
import win32com.client
import ctypes
from datetime import datetime

from sqlalchemy import create_engine
import pymysql

engine = create_engine(" ", encoding='utf-8')
conn = engine.connect()

################################################
# PLUS 공통 OBJECT
from pandas import DataFrame

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
g_objStIdx = win32com.client.Dispatch('Dscbo1.StockIndexIR')
today = datetime.now().strftime('%Y%m%d%H%M')

##################################
codeList = g_objCodeMgr.GetStockListByMarket(1)

kospi = {}
for code in codeList:
    name = g_objCodeMgr.CodeToName(code)
    kospi[code] = name

#######################3
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True


class CpMarketEye:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0

    def Request(self, codes, dataInfo):
        # 0: 종목코드 4: 현재가 20: 상장주식수
        rqField = [0, 4, 20, 118, 120]  # 요청 필드

        self.objRq.SetInputValue(0, rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        print("통신상태", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(2)

        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            cur = self.objRq.GetDataValue(1, i)  # 종가
            listedStock = self.objRq.GetDataValue(2, i)  # 상장주식수
            foreigner = self.objRq.GetDataValue(3, i) # 당일 외국인 순매수
            agency = self.objRq.GetDataValue(4, i) # 당일 기관 순매수

            maketAmt = listedStock * cur
            if g_objCodeMgr.IsBigListingStock(code):
                maketAmt *= 1000
            #            print(code, maketAmt)

            # key(종목코드) = tuple(상장주식수, 시가총액)
            dataInfo[code] = (code, cur, maketAmt, foreigner, agency)

        return True


class CMarketTotal():
    def __init__(self):
        self.dataInfo = {}

    def GetAllMarketTotal(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
        allcodelist = codeList + codeList2
        print('전 종목 코드 %d, 거래소 %d, 코스닥 %d' % (len(allcodelist), len(codeList), len(codeList2)))


        objMarket = CpMarketEye()
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo)
                rqCodeList = []
                continue
        # end of for

        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo)

    def PrintMarketTotal(self):
        result = []
        # 시가총액 순으로 소팅
        data2 = sorted(self.dataInfo.items(), key=lambda x: x[1][2], reverse=True)

        print('전종목 시가총액 순 조회 (%d 종목)' % (len(data2)))
        for item in data2:
            name = g_objCodeMgr.CodeToName(item[0])
            code = item[1][0]
            cur = item[1][1]
            markettot = item[1][2]
            foreign = item[1][3]
            agen = item[1][4]
            market = g_objCodeMgr.GetStockMarketKind(item[0])

            print('%s, %s, 현재가 : %s, 시가총액 : %s  업종 : %s' % (code.replace("A","").replace("Q",""), name, cur, format(markettot, ','),  market))

            data3 = [name, code.replace("A","").replace("Q",""), cur, markettot, market]
            result.append(data3)
        df = DataFrame(result, columns=['company', 'code', 'price', 'siga', 'market'])
        return df.to_sql(name='company_siga', con=engine, if_exists='append', index=False)
        conn.close()

if __name__ == "__main__":
    objMarketTotal = CMarketTotal()
    objMarketTotal.GetAllMarketTotal()
    objMarketTotal.PrintMarketTotal()