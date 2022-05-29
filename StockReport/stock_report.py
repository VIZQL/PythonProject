import re
import enum
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import sys
import io

# 현재 날짜 가져오기
from datetime import datetime
from datetime import date

list_day = []
list_report = []
list_target_price = []
list_opinion = []
list_url = []

def create_soup(url):
    
    headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36 Edg/100.0.1185.44" }
    res = requests.get(url, headers = headers)

    res.raise_for_status()
    # print("응답코드 :", res.status_code) #200 이면 정상 

    # 처음 태그 정보 확인을 위한 html 문서 확인 
    # with open("stock_influ.html", "w", encoding='utf-8') as f:
    #     f.write(res.text)

    # soup = BeautifulSoup(res.text, "lxml")
    # soup = BeautifulSoup(res.text, "html5lib",encoding = 'utf-8')

    # euc-kr 를 encoding 하기 위한 방법 
    soup = BeautifulSoup(res.content.decode('euc-kr','replace'), "lxml")

    return soup

def Stock_reports():


    # 1. 기업 페이지 이동 
    url = "http://hkconsensus.hankyung.com/apps.analysis/analysis.list?sdate={A}&edate={B}&now_page=1&search_value=&report_type=CO&pagenum=20&search_text=&business_code=".format(A= date.today() , B= date.today())
    
    soup = create_soup(url)
    
    table_tr = soup.select("#contents > div.table_style01 > table > tbody > tr")
    
    for index, i in enumerate(table_tr):
        
        # 딕셔너리 쌍 추가 
        # st_report.append({'name' : i.find("strong").text, 
        # 'link' : i.find("a")["href"])

        list_report.append(i.find("strong").text)
        list_url.append("http://consensus.hankyung.com" + i.find("a")["href"])
        list_target_price.append(i.find("td", class_='text_r txt_number').text)
        list_opinion.append(i.select_one('td:nth-of-type(4)').get_text(strip=True))
        list_day.append(i.find("td", class_ = "first txt_number").text) 
        

    # print(list_report)
    # 나중에 판다스 이용할지도 모르니까 

    # df = pd.DataFrame({'title':list_report, 'url':list_url, 'target_price':list_target_price, 'opinion':list_opinion})
    # df.to_csv("today_report.csv", encoding='utf-8')

    # data = pd.read_html(str(stocks_titles))


if __name__ == "__main__":
    Stock_reports() 