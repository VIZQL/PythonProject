import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import KiwoomAPI
# from KiwoomAPI import *
import pythoncom

output_list = {
    'OPT10017': ['종목코드',
                 '종목정보',
                 '종목명',
                 '현재가',
                 '전일대비기호',
                 '전일대비',
                 '등락률',
                 '거래량',
                 '전일거래량',
                 '매도잔량',
                 '매도호가',
                 '매수호가',
                 '매수잔량',
                 '횟수'
                 ],
    'OPT10032': ['종목코드',
                 '현재순위',
                 '전일순위',
                 '종목명',
                 '현재가',
                 '전일대비기호',
                 '전일대비',
                 '등락률',
                 '매도호가',
                 '매수호가',
                 '현재거래량',
                 '전일거래량',
                 '거래대금',
                 ],

    'OPT10035': ['종목코드',
                 '종목명',
                 '현재가',
                 '전일대비기호',
                 '전일대비',
                 'D-1',
                 'D-2',
                 'D-3',
                 '합계',
                 '한도소진율',
                 '전일대비1',
                 '전일대비2',
                 '전일대비3',
                 ],
    'OPT10001': ['종목코드',
                 '종목명',
                 '결산월',
                 '액면가',
                 '자본금',
                 '상장주식',
                 '신용비율',
                 '연중최고',
                 '시가총액',
                 '시가총액비중',
                 '외인소진률',
                 '대용가'
                 'PER',
                 'EPS',
                 'ROE',
                 'PBR',
                 'EV',
                 'BPS',
                 '매출액',
                 '영업이익',
                 '당기순이익',
                 '250최고',
                 '250최저',
                 '시가',
                 '고가',
                 '저가',
                 '상한가',
                 '하한가',
                 '기준가',
                 '예상체결가',
                 '예상체결수량',
                 '250최고가일',
                 '250최고가대비율',
                 '250최저가일',
                 '250최저가대비율',
                 '현재가',
                 '대비기호',
                 '전일대비',
                 '등락율',
                 '거래량',
                 '거래대비',
                 '액면가단위',
                 '유통주식',
                 '유통비율',
                 ],
    'OPT10081': ['종목코드',
                 '현재가',
                 '거래량',
                 '거래대금',
                 '일자',
                 '시가',
                 '고가',
                 '저가',
                 '수정주가구분',
                 '수정비율',
                 '대업종구분',
                 '소업종구분',
                 '종목정보',
                 '수정주가이벤트',
                 '전일종가',
                 ],
}


class KiwoonMain:
    def __init__(self):
        self.kiwoom = KiwoomAPI.KiwoomAPI()
        # self.kiwoom.CommConnect()

# ========== #
    def GetLoginInfo(self):
        pass

    # TR 목록
    def OPT10001(self):
        self.kiwoom.output_list = ['종목명']

        self.kiwoom.SetInputValue("종목코드", "005930")
        self.kiwoom.CommRqData("OPT10001", "OPT10001", 0, "0101")

        return self.kiwoom.ret_data

    def OPT10017(self): # 상한가 종목 요청
        self.kiwoom.output_list = output_list['OPT10017']

        self.kiwoom.SetInputValue("시장구분", "0")
        self.kiwoom.SetInputValue("상하한구분", "1")
        self.kiwoom.CommRqData("OPT10017", "OPT10017", 0, "0101")

        return self.kiwoom.ret_data['OPT10017']

    def OPT10032(self): # 거래대금 상위 종목 
        self.kiwoom.output_list = output_list['OPT10032']

        self.kiwoom.SetInputValue("시장구분", "0")

        self.kiwoom.CommRqData("OPT10032", "OPT10032", 0, "0101")

        return self.kiwoom.ret_data['OPT10032']

    def OPT10035(self): # 외국인 순매수 
        self.kiwoom.output_list = output_list['OPT10035']

        self.kiwoom.SetInputValue("시장구분", "0")
        self.kiwoom.CommRqData("OPT10035", "OPT10035", 0, "0101")

        return self.kiwoom.ret_data['OPT10035']
    
    def OPT10001(self, code): # 종목 기본정보 요청 
        self.kiwoom.output_list = output_list['OPT10001']

        self.kiwoom.SetInputValue("종목코드", code)
        self.kiwoom.CommRqData("OPT10001", "OPT10001", 0, "0101")

        return self.kiwoom.ret_data['OPT10001']
    
    def OPT10081(self, code, day): # 일봉데이터요청
        self.kiwoom.output_list = output_list['OPT10081']

        self.kiwoom.SetInputValue("종목코드", code) 
        self.kiwoom.SetInputValue("기준일자", day) #YYYYMMDD

        self.kiwoom.CommRqData("OPT10081", "OPT10081", 0, "0101")

        return self.kiwoom.ret_data['OPT10081']


'''
app = QApplication(sys.argv)
api_con = KiwoonMain()  

result = api_con.OPT10017() # 상한가 조회 
print(result['Data'][0])

result = api_con.OPT10032() # 거래대금 상위 조회 
print(result['Data'][0])

result = api_con.OPT10035() # 외국인 연속 순매수 상위 
print(result['Data'][0])
'''

