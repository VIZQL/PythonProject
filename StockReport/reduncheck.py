# -*- coding: utf-8 -*

from numpy import result_type
# from Kangwon_News import *
# import Kangwon_News

def check_final(fist_result):
    
    if(len(fist_result) == 1):
        return fist_result
    else:
        for i in range(len(fist_result)):
            for j in range(len(fist_result)):
                if(i!=j):
                    if(set(fist_result[i])&set(fist_result[j])):
                        fist_result.append(list(set(fist_result[i])|set(fist_result[j])))
                        fist_result.pop(i) # i 요소 삭제 
                        fist_result.pop(j-1) # i 요소가 삭제됨으로 j 인덱스가 바뀌게됨 
                        check_final(fist_result)
                        return fist_result
        return fist_result



def redundency_check(th, title):
    
    set_list = []

    same_list = []

    redun = []

    same_result = []


    # print(title)

    # title 을 단어 단위로 구분 --> redun 에 split 된 list 가 저장
    for news in title:
        a = news.replace('“'," ").replace('”'," ").replace('…'," ").replace('"'," ").replace("'"," ").replace("["," ").replace("]"," ").replace("."," ").replace(','," ")
        b = a.split()
        redun.append(b)

    # print(redun)

    # redun 의 split 된 list 중 공통 요소를 set_list 에 저장 
    for index, a in enumerate(redun):
        set_list.append(set(a)) 
    
    # print(set_list)
    # i, j 돌아가면서 각 요소간 교집합이 있는 i,j 를 same_list 로 반환 (동일 기사 index 반환)
    for i in range(len(set_list)):
        for j in range(len(set_list)):
            if(i!=j):
                if(len(set_list[i] & set_list[j])>=th):
                    same_list.append([i,j])

    # print(same_list)
    # 1. 각 배열 요소를 순서대로 정렬 
    for content in same_list:
        content.sort()

    # 2. 동일값이 있으면 하나만 저장 
    for content in same_list:
        if content not in same_result:
            same_result.append(content)

    # print(same_result)
    # 3. 동일 기사가 3개 이상인 경우 하나의 list 로 처리해주는 함수 - 재귀 함수 사용
    final_reulst = check_final(same_result)
    
    # print(final_reulst)

    return final_reulst
    # print(final_reulst)


# 나머지 제거해주는 로직 설계 
def remove_redunlist(redunlist, remove_list):

    # print(remove_list)

    if(len(redunlist)>1):
        for sub_list in redunlist:
            for index, list in enumerate(sub_list):
                # print(remove_list)
                # print(sub_list)
                # print(index)
                # print(sub_list[index])
                if(index != 0):
                    remove_list[sub_list[index]] = "None"
    elif(redunlist == 1):
        for index, list in enumerate(redunlist):
            if(index != 0):
                remove_list[sub_list[index]] = "None"
    else:
        pass


    for i in range(len(remove_list)):
        try:remove_list.remove("None")
        except:pass
        


    return remove_list


if __name__ == "__main__":
    # title, link = Kangwon_News.scrape_headline_news()

    # print(len(title))
    # print(len(link_list))
    # print(len(day_list))
    check_list = redundency_check(3, title)
    print(check_list)
    remove_list = remove_redunlist(check_list, title)
    # print(check_list)
    # print(len(title))
    
    for index, i in enumerate(last_error):
        print(index)
        print(i)

