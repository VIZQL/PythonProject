import re
import enum
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

org_dic = {'name':[], 'price':[], 'quantity':[], 'rate':[] }

def create_soup(url):
    
    headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36" }
    res = requests.get(url, headers = headers)

    res.raise_for_status()
    # print("응답코드 :", res.status_code) #200 이면 정상 

    # 처음 태그 정보 확인을 위한 html 문서 확인 
    with open("stock_influ.html", "w", encoding='utf-8') as f:
        f.write(res.text)
    
    # #boxInfluentialInvestors > div.box_contents > div:nth-child(1) > table > tbody > tr.first > td:nth-child(1) > a

    # soup = BeautifulSoup(res.text, "lxml")
    soup = BeautifulSoup(res.text, "html5lib")

    return soup

def trading_trend():
    
    options = webdriver.ChromeOptions()
    
    # 1. headess chrome option 확인 
    # options.headless = True

    headers = {
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.60 Safari/537.36", 
        "Accept-Language": "ko-KR,ko"
        } # accept language 요청해야함 

    # 2. 페이지 이동 
    # url = "https://finance.daum.net/domestic/influential_investors?market=KOSPI"

    # options.add_argument("window-size=1920x1080")
    # options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36")

    # browser = webdriver.Chrome(options=options)
    # browser.get(url)

    # # 3. 기관 클릭
    # elem = browser.find_element(By.XPATH, "//*[@id='boxInfluentialInvestors']/div[1]/ul/li[2]/a")
    # elem.click()
    
    # soup = BeautifulSoup(browser.page_source, "lxml")

    # 임시 작업
    # 다음 금융 
    # url = "https://finance.daum.net/domestic/influential_investors"
    # soup = create_soup(url)

    # 네이버 금융
    url = "https://finance.naver.com/item/main.naver?code=005930"
    soup = create_soup(url)

    css_selector = "#content > div.section.cop_analysis > div.sub_section > table"
    result = soup.select(css_selector)
    # print(result)

    data = pd.read_html(str(result))


    print(data)

    # url = 'https://finance.naver.com/item/main.nhn?code=035720'
    # table_df_list = pd.read_html(url, encoding='euc-kr')    # 한글이 깨짐. utf-8도 깨짐. 그래서 'euc-kr'로 설정함 


    # url = "https://finance.naver.com/sise/sise_deal_rank.naver"
    # table_df_list = pd.read_html(url, encoding='euc-kr')    # 한글이 깨짐. utf-8도 깨짐. 그래서 'euc-kr'로 설정함 
    # # table_df = table_df_list[3]  # 첫번째 방법과 마찬가지로 리스트 중에서 원하는 데이터프레임 한개를 가져온다   

    # print(data)
    

    # for item in result:
    #      print(item)

   


 
    
    # # 4. 뉴스 더보기 클릭 
    # browser.find_element(By.XPATH, "//*[@id='main_pack']/section[1]/div/div[3]/a").click()
    # soup = BeautifulSoup(browser.page_source, "lxml")

    # # 5. 버튼 1 부터 3 click 하며 뉴스 가져오기 


if __name__ == "__main__":
    trading_trend() 