from dataclasses import replace
import re
import enum
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import reduncheck


# 지수명 list
indice_name = ["코스피지수", "코스닥지수", "다우존스지수", "S&P500", "나스닥종합지수"]


# 기관 5거래일 순매수 상위 
dic = {''}

# 장 마감 후 주요 공시 
org_dic = {'name':[], 'contents':[]}
newsSecondlist = []


import docx
from docx.enum.dml import MSO_THEME_COLOR_INDEX

headers = {
    "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.60 Safari/537.36", 
    "Accept-Language": "ko-KR,ko"
    } # accept language 요청해야함 

def ChromeOn(headless):
    options = webdriver.ChromeOptions()

    # 1. headess chrome option 확인 
    options.headless = headless

    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36")

    browser = webdriver.Chrome(options=options)

    return browser


def create_soup(url):
    
    headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36" }
    res = requests.get(url, headers = headers)
    res.raise_for_status()
    # print("응답코드 :", res.status_code) #200 이면 정상 

    # 처음 태그 정보 확인을 위한 html 문서 확인 
    # with open("naver.html", "w", encoding='utf-8') as f:
    #     f.write(res.text)
    
    soup = BeautifulSoup(res.text, "lxml")
    return soup


# 네이버에 특징주 검색해서 가져오기 
def scrape_stocks_info():
   
    # 네이버 특징주 title 
    title_list = []

    # 네이버 특징주 url
    link_list = []

    browser = ChromeOn(True)
    url = "https://www.naver.com/"
    browser.get(url)

    # 3. 특징주 입력 
    elem = browser.find_element(By.CLASS_NAME, "green_window").find_element(By.ID, "query")
    elem.send_keys("특징주")
    elem.send_keys(Keys.ENTER)
    
    # 4. 뉴스 더보기 클릭 
    browser.find_element(By.XPATH, "//*[@id='main_pack']/section[1]/div/div[3]/a").click()
    soup = BeautifulSoup(browser.page_source, "lxml")

    # 5. 버튼 1 부터 3 click 하며 뉴스 가져오기 
    for i in range(1,4): 
        
        # element 에 page 1 ~3 클릭하기 
        elem = browser.find_element(By.CLASS_NAME, "api_sc_page_wrap").find_element(By.XPATH, "//*[@id='main_pack']/div[2]/div/div/a[{}]".format(i)) # X path 확인
        elem.click()
        
        # 페이지 브라우져 읽어오기 
        soup = BeautifulSoup(browser.page_source, "lxml")
        news_list = soup.find("ul", attrs = {"class" : "list_news"}).find_all("li", attrs = {"class" : "bx"})

        for index, news in enumerate(news_list):
            
            # title 이랑 link 읽어오기 
            title = news.div.div.find("a", attrs = {"class" : "news_tit"})["title"]
            link = news.div.div.find("a")["data-url"]
            
            # title list, link list append 하기 
            title_list.append(title)
            link_list.append(link)

    redun_result = reduncheck.redundency_check(3,title_list)
    # print(redun_result)
    title_list = reduncheck.remove_redunlist(redun_result,title_list)
    link_list =  reduncheck.remove_redunlist(redun_result,link_list)

    return title_list, link_list
    # print(title_list)
    # print(link_list)

# 2. 스탁 인베스터에서 주요 증시 가져오기 
def scrape_major_indice(): 

    # 지수 list 
    indice_number = []

    # 지수 증감 %
    indice_percent = []

    url = "https://kr.investing.com/indices/major-indices"
    res = requests.get(url, headers = headers)
    soup = BeautifulSoup(res.text, "lxml")

    indices = soup.find_all("tr", attrs = {"class":"datatable_row__2vgJl", "data-test" : "price-row"}, limit = 7)

    # 코스피 지수 
    for index, indice in enumerate(indices) :
        if index != 1 : 
            
            temp = indice.find("td", attrs = {"class" : "datatable_cell__3gwri datatable_cell--align-end__Wua8C table-browser_col-last__1ZaGj"}).get_text()
            indice_number.append(temp)

            try:
                temp = indice.find("td", attrs = {"class" : "datatable_cell__3gwri datatable_cell--align-end__Wua8C datatable_cell--up__2984w datatable_cell--bold__3e0BR table-browser_col-chg-pct__9p1T3"})
                indice_percent.append(temp.get_text())
            except:
                try:
                    temp = indice.find("td", attrs = {"class" : "datatable_cell__3gwri datatable_cell--align-end__Wua8C datatable_cell--down__2CL8n datatable_cell--bold__3e0BR table-browser_col-chg-pct__9p1T3"})
                    indice_percent.append(temp.get_text())
                except:
                    try: 
                        indice.find("td", attrs = {"class" : "datatable_cell__3gwri datatable_cell--align-end__Wua8C datatable_cell--down__2CL8n datatable_cell--bold__3e0BR table-browser_col-chg-pct__9p1T3"})
                        print(temp.get_text())
                    except:
                        continue
   
                    continue
    
        # print(indice_number)
        # print(indice_percent)
    return  indice_name, indice_number, indice_percent


# 장 마감 후 주요 공시 edaily
def main_announce_AftMarket(): 
    name = []
    content = []

    browser = ChromeOn(True)

    try:
        url = "https://www.edaily.co.kr/articles/stock/item"
        # url = "https://www.edaily.co.kr/articles/stock/stock"
        browser.get(url)

        # li tag 들 검색하여 elems 에 넣기 
        elems = browser.find_elements(By.TAG_NAME, "li")
        
        # li tag text 중에서 장 마감 포함한 elem 찾아서 클릭하기 
        for elem in elems:

            p = re.compile("장 마감")
            m = p.search(elem.text) # 주어진 문자열 중에 일치하는게 있는지 확인  
            
            if m:
                elem.click()
                break
            else:
                continue

        soup = BeautifulSoup(browser.page_source, "lxml")

        # 대체 구문 전체 텍스트 받아서 정규식으로 구분 
        news_text_all = soup.find('div', attrs = {"class": "news_body"}).get_text()
    except:
        # url = "https://www.edaily.co.kr/articles/stock/item"
        url = "https://www.edaily.co.kr/articles/stock/stock"
        browser.get(url)

        # li tag 들 검색하여 elems 에 넣기 
        elems = browser.find_elements(By.TAG_NAME, "li")
        
        # li tag text 중에서 장 마감 포함한 elem 찾아서 클릭하기 
        for elem in elems:

            p = re.compile("장 마감")
            m = p.search(elem.text) # 주어진 문자열 중에 일치하는게 있는지 확인  
            
            if m:
                elem.click()
                break
            else:
                continue

        soup = BeautifulSoup(browser.page_source, "lxml")

        # 대체 구문 전체 텍스트 받아서 정규식으로 구분 
        news_text_all = soup.find('div', attrs = {"class": "news_body"}).get_text()      

    
    # △ 로 split 처리 
    replace_list = re.split(r'[△]',news_text_all)
    replace_list.pop(0)
    replace_list = list(map(lambda x: x.strip(), replace_list))
    replace_list = list(filter(lambda x: x != '', replace_list))

    #  정규표현식 d{6} 써서 (xxxxx) 패턴끝의 index 알아내서 문자열 slicing 적용
    regex = re.compile(r'\(\d{6}\)')

    for inform in replace_list:
        try:
            matchobj = regex.search(inform)
            idx = matchobj.end()
            name.append(inform[:idx]) 
            content.append(inform[idx:].strip('=').strip(')=').strip())
        except:
            continue


    return name, content



# 상한가 종목 네이버 검색후 기사 가져오기 
def scrape_stocks_info_saghan(name_list):

    browser = ChromeOn(True)
    url = "https://www.naver.com/"
    browser.get(url)

    title_list = []
    link_list = []

    # 3. 상한가 종목 입력 

    for index, name in enumerate(name_list):

        if(index == 0):
            elem = browser.find_element(By.CLASS_NAME, "green_window").find_element(By.ID, "query")
            elem.send_keys(name)
            elem.send_keys(Keys.ENTER)
        else:
            elem = browser.find_element(By.CLASS_NAME, "greenbox").find_element(By.NAME, "query")
            elem.clear()
            elem.send_keys(name)
            elem.send_keys(Keys.ENTER)
    

        # 4. 뉴스 더보기 클릭 
        # browser.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.sp_nnews._prs_nws_all > div > div.api_more_wrap > a").click()
        if(index == 0):
            browser.find_element(By.LINK_TEXT, "뉴스 더보기").click() 
            time.sleep(2)
        
        soup = BeautifulSoup(browser.page_source, "lxml")
        news = soup.find("ul", attrs = {"class" : "list_news"}).find("li", attrs = {"class" : "bx"})
        title_list.append(news.div.div.find("a", attrs = {"class" : "news_tit"})["title"])
        link_list.append(news.div.div.find("a")["data-url"])

    return title_list, link_list 


def scrape_major_indice_money(): 

    doller_kr = []
    doller_index = []

    url = "https://kr.investing.com/currencies/"
    res = requests.get(url, headers = headers)
    soup = BeautifulSoup(res.text, "lxml")

    a = soup.select_one("#cr1 > tbody")
    b = a.find("a", string = "달러/원").parent.parent
    c = b.select_one("#pair_650 > td.pid-650-last")
    doller_kr.append(c.text)
    

    a = soup.select_one("#dailyTab > tbody")
    b = a.find("a", string = "달러/원").parent.parent

    for i in range(5,10):
        tmp = b.select_one(f"#pair_650 > td:nth-child({i})")
        doller_kr.append(tmp.text)

    url = "https://kr.investing.com/currencies/us-dollar-index"
    res = requests.get(url, headers = headers)
    soup = BeautifulSoup(res.text, "lxml")
    
    a = soup.select_one("#last_last")
    doller_index.append(a.text)
    try:
        b = soup.select_one("#quotes_summary_current_data > div.instrumentDataDetails > div.left.current-data > div.main-current-data > div.top.bold.inlineblock > span.arial_20.pid-8827-pcp.parentheses.greenFont")
        doller_index.append(b.text)
    except:
        b = soup.select_one("#quotes_summary_current_data > div.instrumentDataDetails > div.left.current-data > div.main-current-data > div.top.bold.inlineblock > span.arial_20.pid-8827-pcp.parentheses.redFont")
        doller_index.append(b.text)
    
    # print(doller_kr)
    # print(doller_index)
    
    return doller_kr, doller_index




if __name__ == "__main__":
    # scrape_stocks_info() # 네이버 특징주 
    name, content = main_announce_AftMarket() 
    print(name,content)
    # scrape_stocks_info_saghan("부산주공")
    # scrape_major_indice()
    # scrape_major_indice_money()

    # for index, i in enumerate(name):
    #     print(i)
    #     print(content[index])
