# -*- coding: utf-8 -*

from dataclasses import replace
import re
import enum
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# 현재 날짜 가져오기
import datetime 
from datetime import timedelta
import reduncheck

# # 헤드라인 뉴스 제목
# title_list = []

# # 헤드라인 뉴스 링크 
# link_list = []

# # 헤드라인 뉴스 링크 
# Day_list = []

# 뉴스 기사 제한
limit_news = 5

except_info_list = ['뉴시스','마이데일리','뉴스핌','위키트리','스타뉴스', 'OSEN', '노컷', '뉴스1', '뉴스워커', '중도일보', '스포츠', '데일리안', '머니투데이', '아시아투데이', '제주', '브레이크', '연예', '전민일보', '인천','경기','프레시안', '더팩트', '서울신문']
except_title_list = ['개관','개장','컴백','결혼','축제', '개봉', '기부', '박람회', '한정 판매', '공연', '음악회', '나눔', '티켓', '특별전', '문화제', '연주회', '캠페인', '콘서트', '숲길', '소식]', r'\S\S시', r'\S\S군']

def create_soup(url):
    
    headers = {"User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36" }
    res = requests.get(url, headers = headers)
    res.raise_for_status()
    # print("응답코드 :", res.status_code) #200 이면 정상 

    # 처음 태그 정보 확인을 위한 html 문서 확인 
    with open("naver.html", "w", encoding='utf-8') as f:
        f.write(res.text)
    

    soup = BeautifulSoup(res.text, "lxml")
    return soup


def scrape_schedule():

    # 헤드라인 뉴스 제목
    title_list = []

    # 헤드라인 뉴스 링크 
    link_list = []

    # 헤드라인 뉴스 링크 
    Day_list = []

    # 컴퍼니 리스트 
    company_list = []
   
    options = webdriver.ChromeOptions()
    
    # 1. headess chrome option 확인 
    options.headless = True

    headers = {
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.60 Safari/537.36", 
        "Accept-Language": "ko-KR,ko"
        } # accept language 요청해야함 

    # 2. 페이지 이동 
    url = "https://www.naver.com/"

    options.add_argument("window-size=1920x1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36")

    browser = webdriver.Chrome(options=options)
    browser.get(url)
    

    # 3. 오는 {}일 입력 
    for count in range(0,5):
        
        target_Day = datetime.date.today() + timedelta(days = count)

        if(count == 0):
            elem = browser.find_element(By.CLASS_NAME, "green_window").find_element(By.ID, "query")
            elem.send_keys("오는 {}일".format(target_Day.day))
            elem.send_keys(Keys.ENTER)
        else:
            elem = browser.find_element(By.CLASS_NAME, "greenbox").find_element(By.NAME, "query")
            elem.clear()
            if(pre_targetday.month == target_Day.month):
                elem.send_keys("오는 {}일".format(target_Day.day))
                elem.send_keys(Keys.ENTER)
            else:
                elem.send_keys("내달 {}일".format(target_Day.day))
                elem.send_keys(Keys.ENTER)
        

        # browser.execute_script("window.scrollTo(0, document.body.scrollHeight)")
        
        time.sleep(2)

        # 4. 뉴스 더보기 클릭 
        
        # 첫번째 수행할때 뉴스 더보기, 이후에는 필요 없음 
        if(count == 0):
    
            # browser.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.sp_nnews._prs_nws_all > div > div.api_more_wrap > a").click()
            # browser.find_element(By.XPATH, "//*[@id='main_pack']/section[5]/div/div[3]/a").click() 
            browser.find_element(By.LINK_TEXT, "뉴스 더보기").click() 
            
            # //*[@id="main_pack"]/section[4]/div/div[3]/a
            time.sleep(2)
            soup = BeautifulSoup(browser.page_source, "lxml")
            
        
        # 5. 버튼 1 부터 2 click 하며 뉴스 가져오기 
        for i in range(1,3): 
            
            # element 에 page 1 ~3 클릭하기 
            elem = browser.find_element(By.CLASS_NAME, "api_sc_page_wrap").find_element(By.XPATH, "//*[@id='main_pack']/div[2]/div/div/a[{}]".format(i)) # X path 확인
            elem.click()
            time.sleep(1)

            # 페이지 브라우져 읽어오기 
            soup = BeautifulSoup(browser.page_source, "lxml")
            news_list = soup.find("ul", attrs = {"class" : "list_news"}).find_all("li", attrs = {"class" : "bx"})

            for index, news in enumerate(news_list):
            
                skip_flag = 0

                # title 이랑 link 읽어오기 
                title = news.div.div.find("a", attrs = {"class" : "news_tit"})["title"]
                company = news.div.div.find("a", attrs = {"class" : "info press"}).get_text()
                link = news.div.div.find("a")["data-url"]
                
                # 컴퍼니가 skip list 에 있으면 skip 하기 
                for list in except_info_list:
                    if (company.find(list) > -1):
                        skip_flag = 1

                for list in except_title_list:
                    # if (title.find(list) > -1):
                    if (re.search(list, title)):
                        skip_flag = 1

                if (skip_flag == 0):
                    title_list.append(title)
                    link_list.append(link)
                    Day_list.append("{A}월 {B}일".format(A= target_Day.month, B = target_Day.day))


        if(count == 0):
            pre_targetday = target_Day

    # with open("redun.check.txt", "w", encoding='utf-8') as f:
    #     for a in title_list:
    #         f.write(a)
    #         f.write("\n")


    print(title_list)
    print(link_list)
    print(Day_list)
    print(company_list)
        
    redun_result = reduncheck.redundency_check(3,title_list)

    # print(redun_result)

    title_list = reduncheck.remove_redunlist(redun_result,title_list)
    link_list =  reduncheck.remove_redunlist(redun_result,link_list)
    Day_list =  reduncheck.remove_redunlist(redun_result,Day_list)

    return title_list, link_list, Day_list


if __name__ == "__main__":
    title, link, day = scrape_schedule() # 헤드라인 뉴스 정보 가져오기 
    print(title,link,day)




else:
    pass
