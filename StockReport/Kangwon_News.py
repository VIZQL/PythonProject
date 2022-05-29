import re
import enum
import requests
from bs4 import BeautifulSoup

# 뉴스 기사 제한
limit_news = 5

limit_news_bell = 5

limit_news_guru = 5


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


url_list = ["https://news.naver.com/main/main.naver?mode=LSD&mid=shm&sid1=100" # 정치
,"https://news.naver.com/main/main.naver?mode=LSD&mid=shm&sid1=101" # 경제
,"https://news.naver.com/main/main.naver?mode=LSD&mid=shm&sid1=104" # 세계
]

url_list_eco = [
"https://news.naver.com/main/list.naver?mode=LS2D&mid=shm&sid1=101&sid2=259" # 경제 - 금융
,"https://news.naver.com/main/list.naver?mode=LS2D&mid=shm&sid1=101&sid2=258" # 경제 - 증권
,"https://news.naver.com/main/list.naver?mode=LS2D&mid=shm&sid1=101&sid2=261" # 경제 - 산업
]

def scrape_headline_news_eco():

    title_eco_list = []
    link_eco_list = []

    for url in url_list_eco:

        soup = create_soup(url)

        for index in range(limit_news):
            try:
                a = soup.select_one("#main_content > div.list_body.newsflash_body > ul.type06_headline")
                b = a.select_one(f"li:nth-child({index+1}) > dl > dt:nth-child(2) > a")
                title_eco_list.append(b.get_text().strip())
                link_eco_list.append(b["href"])
            except:
                a = soup.select_one("#main_content > div.list_body.newsflash_body > ul.type06_headline")
                b = a.select_one(f"li:nth-child({index+1}) > dl > dt > a")
                title_eco_list.append(b.get_text().strip())  
                link_eco_list.append(b["href"])
                    
    # print(title_eco_list)
    # print(link_eco_list)

    return title_eco_list, link_eco_list


def scrape_headline_news(): 

    title_list = [] # 헤드라인 뉴스 제목
    link_list = [] # 헤드라인 뉴스 링크 


    for url in url_list:

        soup = create_soup(url)

        # 각 섹터별 헤드라인 뉴스 3개씩만 불러오기 
        news_list = soup.find_all("div", attrs = {"class" : "cluster_body"}, limit = limit_news)

        for index, news in enumerate(news_list):
            title = news.find("div", attrs = {"class" : "cluster_text"}).a.get_text().strip()
            link = news.find("a")["href"]

            # title, link  를 list 로 정렬 
            title_list.append(title.strip())
            link_list.append(link.strip())

    return title_list, link_list

def scrape_headline_news_thebell(): 

    title_list = [] # 헤드라인 뉴스 제목
    link_list = [] # 헤드라인 뉴스 링크 

    url = "http://www.thebell.co.kr/free/index.asp"

    soup = create_soup(url)

    a = soup.select_one("#contents > div.contentSection > div > div.content.R > div.bestclickWrap > ul")
    
    for i in range(limit_news_bell):
        b = a.select_one(f"li:nth-child({i+1}) > a")

        # title, link  를 list 로 정렬 
        title_list.append(b.get_text().strip())
        link_list.append("http://www.thebell.co.kr"+b["href"])

    print(title_list)
    print(link_list)

    return title_list, link_list

def scrape_headline_news_guru(): 

    title_list = [] # 헤드라인 뉴스 제목
    link_list = [] # 헤드라인 뉴스 링크 

    # url = "https://www.theguru.co.kr/news/article_list_all.html"
    url = "https://www.theguru.co.kr/news/review_list_all.html?rvw_no=32"
    
    soup = create_soup(url)

    a = soup.select_one("#container > div > div.column.sublay > div:nth-child(1) > div > div.ara_001 > ul")

    # print(c)
    for i in range(limit_news_guru):
        
        b = a.select_one(f"li:nth-child({i+1}) > a")
        c = b.h2

        # title, link  를 list 로 정렬 
        title_list.append(c.get_text().strip())
        link_list.append("https://www.theguru.co.kr"+b["href"])

    print(title_list)
    print(link_list)

    return title_list, link_list



# f = open("C:/PYTHONWORKSPACE/naver_news.txt", 'a', encoding='utf-8')
# f.write("\n")
# f.close()


if __name__ == "__main__":
    # scrape_headline_news_eco() # 헤드라인 뉴스 정보 가져오기
    title_list, link_list = scrape_headline_news() # 헤드라인 뉴스 정보 가져오기 
    print(title_list)
    # scrape_headline_news_thebell()
    # scrape_headline_news_guru()
else:
    pass
