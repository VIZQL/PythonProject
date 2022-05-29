from numpy import append
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re 

name_list = []
day_rate_list = []
week_rate_list = []
month_rate_list = []
YTD_rate_list = []

def fs(soup_variable, k):
    
    temp = soup_variable

    for i in range(k):
        temp = temp.find_next_sibling("td")
    
    return temp


def scrape_rawM(): 

    url = "https://kr.investing.com/commodities/"

    headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36" }
    res = requests.get(url, headers = headers)
    res.raise_for_status()


    soup = BeautifulSoup(res.text, "lxml")



    table = soup.select_one("#dailyTab > tbody")
    table_row = table.select("tbody > tr")

    # table_row[0] re.findall('(?<=\>)(.*?)(?=\<)', str(contents.find("a")))

    for index, contents in enumerate(table_row):
        name_list.append(re.findall('(?<=\>)(.*?)(?=\<)', str(contents.find("a")))[0])
        day_rate_list.append(re.findall('(?<=\>)(.*?)(?=\<)', str(fs(contents.td,4)))[0])
        week_rate_list.append(re.findall('(?<=\>)(.*?)(?=\<)', str(fs(contents.td,5)))[0])
        month_rate_list.append(re.findall('(?<=\>)(.*?)(?=\<)', str(fs(contents.td,6)))[0])
        YTD_rate_list.append(re.findall('(?<=\>)(.*?)(?=\<)', str(fs(contents.td,7)))[0])

 
if __name__ == "__main__":
    scrape_rawM() # 헤드라인 뉴스 정보 가져오기 
else:
    pass
