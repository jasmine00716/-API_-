import win32com.client
import win32com
import pandas as pd
import sqlite3

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

index = 0
# 네이버, 삼성전자, 아모레퍼시픽, 현대차, LG화학, 카카오, POSCO, SK바이오사이언스, 엔씨소프트, 두산중공업, HMM, 아시아나항공, 한화솔루션, 신세계, 에스엠
# 유한양행, 카카오게임즈, sk하이닉스, 셀트리온, 현대모비스
code = ["A035420", "A005930", "A090430", "A005380", "A051910", "A035720", "A005490", "A302440",
        "A036570", "A034020", "A011200", "A020560", "A009830", "A004170", "A041510",
        "A000100", "A293490", "A000660", "A068270", "A012330"]
value = []
f = open('F:\\정수민\\dwu 3학년 1학기\\한이음 프로젝트 - 동학개미\\data\\stockprice_data.csv', 'w')
for k in code:
    # 인덱스 부여
    index = index + 1
    value.append(index)

    stock_name = instCpStockCode.CodeToName(k) # 종목이름
    value.append(stock_name)

    stock_code = instStockChart.SetInputValue(0, k)  # 종목코드
    value.append(k)
    f.write("%s,%s,%s," % (index, stock_name, k))

    instStockChart.SetInputValue(1, ord('1')) # 요청구분(기간으로 요청)
    instStockChart.SetInputValue(2, 20210901) # 요청종료일
    instStockChart.SetInputValue(3, 20210901) # 요청시작일

    # 0: 날짜, 1: 시간, 2: 시가, 3: 고가, 4: 저가, 5: 종가
    # 6: 전일대비, 8: 거래량, 9: 거래대금, 10: 누적체결매도수량
    instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 8))
    instStockChart.SetInputValue(6, ord('D')) # 차트구분(일)
    instStockChart.SetInputValue(9, ord('0')) # 무수정주가

    instStockChart.BlockRequest()

    numData = instStockChart.GetHeaderValue(3) # 수신개수
    numField = instStockChart.GetHeaderValue(1) # 필드개수

    for i in range(numData):
        for j in range(numField):
            data = instStockChart.GetDataValue(j, i)
            value.append(data)
            f.write("%s," % data)
        f.write("\n")
        value.clear() # 다른 종목을 저장하기 위해 리스트를 비운다
f.close()
