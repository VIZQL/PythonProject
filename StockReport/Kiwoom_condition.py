from pykiwoom.kiwoom import *
import re
import pandas as pd
# import KiwoomMain
# from kiwoom_pro import KiwoomMainT
# from KiwoomMainT import *
# import kiwoom_pro.KiwoomMainT
from KiwoomMain import *
import time
import datetime 

attention_name = []
attention_percent = []
attention_amount = []

chart20sun = []
chart8sun = []
chart45sun = []

sanghan_name = []
sanghan_tramount = []

trTop_name = []
trTop_amount = []
trTop_percent = []

except_stock = ['삼성전자','하이닉스', 'KODEX', '코덱스', 'ETF', '레버리지', '채권', 'TIGER']


# test_stock = ['삼성전자', '팜스토리', 'SK 하이닉스', 'KODEX 200선물인버스2X', 'KODEX 레버리지', 'KBSTAR KIS단기종합채권(AA-이상)액티브', '대한전선']

def getConditionKiwoom():
    
    count = 0 # 거래대금 상위종목중 필요없는 종목 제외 

    # 로그인
    kiwoom = Kiwoom()
    kiwoom.CommConnect(block=True)

    # app = QApplication(sys.argv)
    api_con = KiwoonMain()  

    # 조건식을 PC로 다운로드
    kiwoom.GetConditionLoad()
 

    condition_list = {'index':[], 'name':[]}
    temporary_condition_list = kiwoom.GetConditionNameList()

    for data in temporary_condition_list:
        try: 
            condition_list["index"].append(str(data[0]))
            condition_list["name"].append(str(data[1]))
        except IndexError:
            pass
    

    # 5-8 일선 
    condition_index = condition_list["index"][condition_list["name"].index("5-8일선")]
    codes = kiwoom.SendCondition("0101", "5-8일선", condition_index, 0)

    for code in codes:
        chart8sun.append(kiwoom.GetMasterCodeName(code))

    # 20 일선 
    condition_index = condition_list["index"][condition_list["name"].index("18일선")]
    codes = kiwoom.SendCondition("0101", "18일선", condition_index, 0)

    for code in codes:
        chart20sun.append(kiwoom.GetMasterCodeName(code))

    # 45 일선 
    condition_index = condition_list["index"][condition_list["name"].index("45일선")]
    codes = kiwoom.SendCondition("0101", "45일선", condition_index, 0)

    for code in codes:
        chart45sun.append(kiwoom.GetMasterCodeName(code))


    # 주목받은 종목 - 주식 기본조회로 추가 정보 얻어야함 
    condition_index = condition_list["index"][condition_list["name"].index("주목받은종목")]
    codes = kiwoom.SendCondition("0101", "주목받은종목", condition_index, 0)

    today_time = datetime.date.today().strftime("%Y%m%d")

    for code in codes:
        result = api_con.OPT10001(code)
        time.sleep(0.2)
        result_2 = api_con.OPT10081(code,today_time)
        attention_name.append(result['Data'][0]['종목명'])
        attention_percent.append(result['Data'][0]['등락율'])
        tmp = int(result_2['Data'][0]['거래대금'])//1000
        attention_amount.append(f'{int(tmp):,}')
        time.sleep(0.2)


    result = api_con.OPT10017() # 상한가 조회 

    for indiv in result['Data']:
        sanghan_name.append(indiv['종목명'])
        time.sleep(0.2)
        result_2 = api_con.OPT10081(indiv['종목코드'],today_time) 
        tmp = int(result_2['Data'][0]['거래대금'])//1000
        sanghan_tramount.append(f'{int(tmp):,}')


    result = api_con.OPT10032() # 거래대금 상위 조회 

    for index, indiv in enumerate(result['Data']):
        
        except_flag = 0

        for word in except_stock:
            if(indiv['종목명'].find(word) > -1):
                except_flag = 1

        if(except_flag == 0):
            trTop_name.append(indiv['종목명'])
            tmp = int(indiv['거래대금'])//1000
            trTop_amount.append(f'{int(tmp):,}')
            trTop_percent.append(indiv['등락률'])
            count += 1
        
        if(count == 10):
            break
    

    # print(trTop_name)
    # print(trTop_amount)
    # print(trTop_percent)

    # print(result['Data'][0])

    # result = api_con.OPT10035() # 외국인 연속 순매수 상위 
    # print(result['Data'][0])
    # result = api_con.OPT10001()
    # print(result['Data'][0])

    # print(attention_name)
    # print(attention_percent)
    # print(attention_amount)
    # print(chart8sun)
    # print(chart20sun)
    # print(chart45sun)

    # print(sanghan_name)     
    # print(sanghan_tramount)

if __name__ == "__main__":
    getConditionKiwoom()
    





    


        


