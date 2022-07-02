from calendar import c
# from vsave import *
import clipboard
import pyautogui as pa
import time, win32con, win32api, win32gui, ctypes
import re
import pywinauto 
from apscheduler.schedulers.background import BackgroundScheduler

from telethon.sync import TelegramClient
from telethon.tl.functions.messages import (GetHistoryRequest)
import telethon
import asyncio

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.jobstores.base import JobLookupError
from apscheduler.schedulers.background import BlockingScheduler

import datetime

########################### 모니터에 따른 설정값 #############################################
# FHD : 1920 X 1080
# UHD : 3840 X 2160 

UHD_SETTING = {"fix_data" :[450, 170, 840, 1000] }
FHD_SETIING = {"fix_data" :[450, 50, 840, 800] }

width, height = pa.size()

if (width, height) == (1920,1080):
    SETTING = [450, 50, 840, 800]
elif(width, height) == (3840,2160):
    SETTING = [450, 170, 840, 1000]
else:
    print("해당 사이즈를 찾을수 없습니다")
    SETTING = [450, 170, 840, 1000]
# print(width, height)


######################################### 카카오톡 오픈 채팅방 처리 함수 ###############################
kakao_opentalk_name = '영앤리치투자 주식정보방'
kakao_opentalk_name_temp = "영앤리치투자"  ## 너무길어서 인식을 잘 못해서 임시로 만듬
# kakao_opentalk_name = '메모장2'
# kakao_opentalk_name_temp = '메모장2'
# Master = ['YRStockLeader']
fix_data = [450, 170, 840, 1000]  ## UHD 모니터 기준 캘값임 
share_JJ_flag =0
share_JH_flag =0
share_Jing_flag =0
share_news_flg = 0
manager = ['[방장봇]', '[박강원전문가]', '[YRStockLeader]', '[부관리자]']

######################################### 텔레그램 관련 변수  ###############################
api_id = '14078381'
api_hash = '1c92f649600a09317f38c1df4f9e1ee4'
chat = "https://t.me/mainnews"

offset_id = 0
tele_message = 0
client = TelegramClient('KKW', api_id, api_hash)


## 송출 금지 단어 ## 
inhibitword = ['t.me/', '리차드', 'youtu']

######################################### 필요한 함수 선언  ###############################


## 채팅 방에 메세지 전송  
def kakao_sendtext(hwndEdit, text):
    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# # 메세지를 읽어오는 함수 
def read_message():
    pa.hotkey('ctrl', 'c')
    data = clipboard.paste()
    return data
    
# 메시지를 전송하는 함수 
def send_message(breaktime=0, send=''): 
    clipboard.copy(send)
    pa.hotkey('ctrl', 'v')
    time.sleep(breaktime)
    pa.press('enter')

# 메세지를 읽어오는 함수 
def find_message(): 
    pa.hotkey('ctrl', 'f')
    data = 'open'
    time.sleep(0.1)
    pa.press('enter')


# # 오픈채팅 링크 공유사람 강퇴시키는 기능 
def find_openchat_postion():
    sucess = 0
    Ox = 0
    Oy = 0
    try:
        # pa.click(button='left') # http 링크가 클릭되서 팝업되는 경우가 있음...
        Ox,Oy = pa.locateCenterOnScreen("find_http.png", confidence=0.7)
        sucess = 1
        return Ox, Oy, sucess
    except:
        print("스크린샷 캡쳐 실패")
        return Ox, Oy, sucess

# # 오픈채팅 링크 공유하면 가리는 기능 하는 기능 
def hide_open_chat(dx, dy):
    try:
        pa.click(dx, dy,button='right')
        time.sleep(0.2)
        pa.click(dx+73, dy+248, button='left')
        time.sleep(0.2)
        pa.click(dx+73+54, dy+248+12, button='left')
        time.sleep(0.2)
        kx,ky = pa.locateCenterOnScreen("delete_check.png", confidence=0.7)
        pa.click(kx, ky, button='left')
    except:
        print("가릴 메세지를 찾을수 없습니다")

# # 오픈채팅 링크 공유사람 강퇴시키는 기능 
def exit_user(dx, dy, line_cnt, user, contents, line_diff):
    try:
        a,b = dx-40, dy-line_cnt*17-line_diff*17 # 프로필 위치 계산 
        if a > chat_position[0] and b > chat_position[1]:
            pa.click(a, b, button='left') # 프로필 위치 계산 
            time.sleep(0.2)

            Ex,Ey = a-191, b+341 # 프로필 위치로부터 내보내기 좌표계산
            pa.click(Ex,Ey, button='left')
            time.sleep(0.2)

            ################## 그림인식이 잘 안먹힘 ###############
            Ex,Ey = pa.locateCenterOnScreen("exit_user.png", grayscale = True, confidence=0.5)
            time.sleep(0.2)
            pa.click(Ex, Ey, button='left')
            ###############################
            
            pa.click(Ex-40, Ey, button='left') ## 내보내기 / 내보내기&신고 중에 선택
            time.sleep(0.2)
            pa.click(Ex-40, Ey-165, button='left') ## 최종 내보내기 

            with open("강퇴list.txt", "a" ) as f:
                f.write("\n")
                f.write("-"*100 + "\n")
                f.write(str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M')) + " : ")
                f.write(f"강퇴ID : {user}, 채팅내용: {contents}")
        else:
            print("강퇴할 유저가 화면 밖에 있습니다")
            pa.click(dx-40,dy, button='left')

    except:
        print("강퇴할 유저를 찾을수 없습니다")

########################################### # 리포트공유
def share_report(postion, file_name):

    pa.hotkey('ctrl', 't') # 카카오톡 단축키 
    time.sleep(3)

    '''   
    x = postion[0] + (postion[2] - postion[0])/17
    y = postion[3]+ (postion[3] - postion[1])/9.5 
    # print(x,y)
    pa.moveTo(x, y, 0.1)
    time.sleep(1)
    pa.click(x, y, button='left') 
    pa.click(x, y, button='left') 
    time.sleep(3)
    '''

    # hwndMain = win32gui.FindWindow(None, "32770")
    hwndMain = win32gui.FindWindow(None, "열기")
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "ComboBoxEx32", None) 
    hwndEdit2 = win32gui.FindWindowEx( hwndEdit, None, "ComboBox", None) 
    hwndEdit3 = win32gui.FindWindowEx( hwndEdit2, None, "Edit", None) 

    kakao_sendtext(hwndEdit3, file_name)

    time.sleep(2)

    try:
        Ex,Ey = pa.locateCenterOnScreen("uploadNG.png", grayscale = True, confidence=0.5)
        time.sleep(0.2)
        pa.click(Ex, Ey, button='left')
    except:
        pass
    


# # 채팅방 열기
def open_chatroom(chatroom_name):
    # # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)    # ㄴ시작핸들을 첫번째 자식 핸들(친구목록) 을 줌(hwndkakao_edit2_1)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)

    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(1)   # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(1)

## 채팅쓰기 
def find_linecnt_until_user(data, cur_index, line_cnt):
    total_len = len(data)
    
    if(total_len >= -cur_index+3):
        
        if(data[cur_index-1].split(" ")[0] != data[cur_index-3].split(" ")[0]):
            final_index =  cur_index
        else:
            line_cnt = line_cnt + data[cur_index-2].count("\n")
            final_index, line_cnt = find_linecnt_until_user(data, cur_index-2, line_cnt)
    else:
        final_index = cur_index
    
    return final_index, line_cnt



########################## 링크 유저 강퇴, 글 가리기  ##############################
def Check_linkMsg(send_msg, cmd):

    global share_JJ_flag, share_JH_flag, share_Jing_flag
    offset = 0
    cnt = 0
    data = ""
    skip = 0

    ## 메세지창 중간부터 아래까지 드래그 
    ## 만일 읽을수 없으면 시작 y 값 좀 낮춰서 읽음 
    
    # 마우스 스크롤 다운 아래로
    pa.scroll(-100)

    ####################### 채팅창 읽는데 처음 클릭한값이 하이퍼링크 부분이면 복사가 안됨###################
    while (data.count('\n') <= 2):

        pa.moveTo(start_x, start_y+offset)
        pa.mouseDown(button='left')
        pa.moveTo(start_x, end_y, 0.1)
        pa.mouseUp()

        ## 메시지창 ctrl + c 
        data = read_message() 
        time.sleep(0.5)

        if(cnt < 8):
            offset = offset + (chat_position[3] - chat_position[1])/10
        else:
            offset = 0
            cnt = 0 

        cnt = cnt+1

    ################################## 누가 링크 공유했는지 확인하는 부분 #################

    if "http" in data : ## 오픈채팅방 공유했는지 확인하기 
        data_split = re.split(r'(\[\S*\s*\S*\s*\S*\s*\S*\s*\S*\s*\] \[\S*\s*\d+:\d+\])',data)
        data_split = [i for i in data_split if i != ""]

        for i in range(-1,-len(data_split),-2):
            if "http" in data_split[i] and "YRStockLeader" not in data_split[i-1]:
                print(f'강퇴 대상자는 {data_split[i-1].split(" ")[0]}')
                a = data_split[i].count('\n')

                # print(data_split[i])
                print(f"rstrip 전  line count 는 {a}")

                if(i==-1):
                    line_cnt = data_split[i].count('\n')-1
                else:
                    line_cnt = data_split[i].count('\n')
                
                print(f"처음 line count 는 {line_cnt}")
                line_until_user, line_cnt = find_linecnt_until_user(data_split, i, line_cnt)

                print(f"line_until_user 는 {line_until_user}")
                print(f"총 line_cnt 는 {line_cnt}")

                line_diff = ((-line_until_user)-(-i))//2  # 채팅 말풍선 count

                line_count_aft_http = data_split[i].split("http")[-1].count("\n") - 1
                print(f"http 이후 띄어쓰기", data_split[i].split("http")[-1].count("\n"))
    
                line_cnt = line_cnt - line_count_aft_http  # http 밑에 있는줄을 계산

                print(f"line_diff 는 {line_diff}")

                Ox,Oy,success = find_openchat_postion()
                user_name = data_split[i-1]
                contents = data_split[i]

                if success == 1:
                    exit_user(Ox, Oy,line_cnt, user_name, contents, line_diff)
                    hide_open_chat(Ox, Oy)
    else: 
        pass

    if send_msg != 0 and cmd == 0:
    # if send_msg != 0:
        for a in inhibitword:
            if a in send_msg:
                skip = 1
        if(skip == 0):
            send_msg_update = send_msg.replace("[","<").replace("]", ">")
            kakao_sendtext(hwndEdit, send_msg_update)

    
    # now = datetime.datetime.now()
    # today8am = now.replace(hour=8, minute=0, second=0, microsecond=0)
    # now < today8am

    today_time = datetime.date.today().strftime("%Y%m%d") 
    time_now = str(datetime.datetime.now().strftime('%H:%M'))

    ## 장전 리포트 공유 
    if  time_now == "08:10" and share_JJ_flag == 0:
        file_name = f"C:\\PYTHONWORKSPACE\\webscraping_basic\\webscraping_project\\2022\\Y&R_리포트_{today_time}_장시작전.pdf"
        kakao_sendtext(hwndEdit, f"금일자 장전 YR 리포트 공유드립니다~!")
        share_report(chat_position, file_name)
        share_JJ_flag = 1
        
        ## 장중 리포트 공유 
    if time_now == "12:00" and share_Jing_flag == 0:
        file_name = f"C:\\PYTHONWORKSPACE\\webscraping_basic\\webscraping_project\\2022\\Y&R_리포트_{today_time}_오전장.pdf"
        kakao_sendtext(hwndEdit, f"오전장 모두 수고 많으셨습니다")
        kakao_sendtext(hwndEdit, f"오전장 YR 리포트 공유드립니다!")
        share_report(chat_position, file_name)
        share_Jing_flag = 1

    ## 장후 리포트 공유 
    if time_now == "16:00" and share_JH_flag == 0:
        file_name = f"C:\\PYTHONWORKSPACE\\webscraping_basic\\webscraping_project\\2022\\Y&R_리포트_{today_time}_장마감.pdf"
        kakao_sendtext(hwndEdit, f"금일자 장후 YR 리포트 공유드립니다~!")
        share_report(chat_position, file_name)
        share_JH_flag = 1


########################################## 글 만 가리기 ######################################
def chat_delete():

    global manager
    offset = 0
    cnt = 0
    data = ""
    ishttp = 0
    ispicture = 0

    # 마우스 스크롤 다운 아래로
    pa.scroll(-100)

    posX = chat_position[0]+(chat_position[2] - chat_position[0])/10
    posY = chat_position[3] - 10 

    final_x = chat_position[0] + (chat_position[2] - chat_position[0])/2 - (chat_position[2] - chat_position[0])/10
    final_y = chat_position[1]+(chat_position[3] - chat_position[1])/2 + (chat_position[3] - chat_position[1])/8

    while (data.count('\n') <= 2):

        pa.moveTo(start_x, start_y+offset)
        pa.mouseDown(button='left')
        pa.moveTo(start_x, end_y, 0.1)
        pa.mouseUp()

        ## 메시지창 ctrl + c 
        data = read_message() 
        time.sleep(0.2)

        if(cnt < 8):
            offset = offset + (chat_position[3] - chat_position[1])/10
        else:
            offset = 0
            cnt = 0 

        cnt = cnt+1
    
    data_split = re.split(r'(\[\S*\s*\S*\s*\S*\s*\S*\s*\S*\s*\S*\s*\S*]\s\[\S*\s*\d+:\d+\])',data)
    data_split = [i for i in data_split if i != ""]
    # data_split.pop(0)

    
    # print(len(data_split))

    # print(data_split)

    for i in range(-1,-len(data_split),-2):
        if(data_split[i-1].split(" ")[0] not in manager) and "삭제된 메시지입니다." not in data_split[i]:
            text_line_cnt = data_split[i].rstrip().count("\n")
            extra_line_cnt = data_split[i].count("채팅방 관리자가 메시지를 가렸습니다.") + data_split[i].count("나갔습니다.") + data_split[i].count("들어왔습니다.")+ data_split[i].count("내보냈습니다.")
            print (f"index 는 {i} text line count 는 {text_line_cnt},  가려진 메세지 count 는 {extra_line_cnt}")
            print(data_split[i].rstrip())
            if 'http' in data_split[i]:
                ishttp = 1
            elif "사진" in data_split[i].rstrip():
                ispicture = 1
            break
            
    try:
        dx,dy = posX, posY - extra_line_cnt*42
        pa.click(posX -20, posY,button='left') # 전체선택한거 한번 해제 해줘야함
        pa.click(dx, dy ,button='right')
        time.sleep(0.2)

        if ishttp == 1: ## 링크는 가리기까지 9칸
            pa.click(dx+73, dy+248, button='left')
            time.sleep(0.2)
            pa.click(dx+73+54, dy+248+12, button='left')
        elif ispicture == 1: ## 그림은 가리기까지 7칸
            pa.click(dx+73, dy+248-20-20, button='left')
            time.sleep(0.2)
            pa.click(dx+73+54, dy+248-20-20, button='left')          
        else: ## text 는 가리기까지 8칸
            pa.click(dx+73, dy+248-20, button='left')
            time.sleep(0.2)
            pa.click(dx+73+54, dy+248-20, button='left')

        time.sleep(0.2)
        # kx,ky = pa.locateCenterOnScreen("delete_check.png", confidence=0.7)
        pa.click(final_x, final_y, button='left')
    except:
        pass
        print("가릴 메세지를 찾을수 없습니다")
    


########################## 텔레그램 메시지 기능 ##############################

async def telebot_message(chatname):

    # await client.start()

    history =  await client(GetHistoryRequest(
                peer=chatname,
                offset_id=offset_id,
                offset_date=None,
                add_offset=0,
                limit=5,
                max_id=0,
                min_id=0,
                hash=0
    ))

    message = history.messages[0].message

    return message


################################ while 문 내에서 sync 맞추기 위해 timer 설정#################
timer_count = 0

def check_timer(resol, target_sec):
    global timer_count
    trigger = 0

    timer_count = timer_count+resol

    if(timer_count >= target_sec):
        trigger = 1
        timer_count = 0

    return trigger


################################ 송출 메시지 결정#############################
## 이전 메세지와 동일하면 송출하지 않고 동일하지 않으면 송출 cmd 메세지이면 무시 

def determine_msg(pre_result, result):

    send_msg = ""
    cmd = 0
    new = 0

    if len(result)>=1 and len(pre_result)>=1:
        for i, contents in enumerate(result):
            if pre_result[i] != result[i]:
                send_msg = result[i]
                new = 1
                if i == 0:
                    cmd = 1

    return send_msg, cmd, new

# pa.FAILSAFE = False

News_result = []

def generate_news(msg1, msg2, cmd):

    global News_result
    global share_news_flg

    if (msg1 != "" and cmd == 0):
        News_result.append(msg1)
    
    if (msg2 != ""):
        News_result.append(msg2)

    if (msg1 == "show" and cmd == 1):
        print(News_result)

    today_time = datetime.date.today().strftime("%Y%m%d") 

    right_now = str(datetime.datetime.now().strftime('%H:%M'))

    if  (right_now == "12:00" or right_now == "19:40") and share_news_flg == 0:
        share_news_flg = 1
        with open(f"{today_time}뉴스_list.txt", "a", encoding= 'utf-8' ) as f:
            for news in News_result:
                f.write("\n")
                f.write(news)


############################################# 전체 창 닫기 ###################################
pa.hotkey('winleft', 'd')


# ########################## 1.오카방 열기 #################################
open_chatroom(kakao_opentalk_name_temp) # 오픈 채팅방 열기 

########################## 2. 채팅 치기 위해 핸들러 가져오기 #####################################
hwndMain = win32gui.FindWindow(None, kakao_opentalk_name)
hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None) # 채팅 치는부분 좌표 가져오기
open_chat = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None) # 채팅방 전체 좌표가져오기


# ######################### 3. 오픈채팅방 검색해서 활성화 , 비쥬얼 스튜디오도 열기######################

app = pywinauto.application.Application()  # 오카방 열기
app.connect(handle=hwndMain)
app_dialog = app.top_window_()
app_dialog.Restore()

pa.hotkey('ctrl', 'shift', 't') # ctrl + shift t 로  카톡방을 고정할수 있음 

app = pywinauto.application.Application()  # 비쥬얼 스튜디오 열기 
app.connect(title_re=".*Visual Studio Code.*")
app_dialog = app.top_window_()
app_dialog.Restore()

time.sleep(2)

# ############################## 3. 시작전 size fix 하기 #########################################
# win32gui.MoveWindow(hwndMain, fix_data[0], fix_data[1], fix_data[2], fix_data[3], True)
win32gui.MoveWindow(hwndMain, SETTING[0], SETTING[1], SETTING[2], SETTING[3], True)

chat_position = win32gui.GetWindowRect(open_chat)
start_x = (chat_position[0] + chat_position[2])/2
start_y = chat_position[1] + (chat_position[3] - chat_position[1])/10
end_y = chat_position[3] - (chat_position[3] - chat_position[1])/20

print(chat_position)

def init():
    ########################## 0. 시작전 다 창 아래로 깔기#################################
    pa.hotkey('winleft', 'd')


    ########################## 1.오카방 열기 #################################
    open_chatroom(kakao_opentalk_name_temp) # 오픈 채팅방 열기 

    ########################## 2. 채팅 치기 위해 핸들러 가져오기 #####################################
    hwndMain = win32gui.FindWindow(None, kakao_opentalk_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None) # 채팅 치는부분 좌표 가져오기
    open_chat = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None) # 채팅방 전체 좌표가져오기


    ######################### 3. 오픈채팅방 검색해서 활성화 , 비쥬얼 스튜디오도 열기######################

    app = pywinauto.application.Application()  # 오카방 열기
    app.connect(handle=open_chat)
    app_dialog = app.top_window_()
    app_dialog.Restore()

    # app = pywinauto.application.Application()  # 비쥬얼 스튜디오 열기 
    # app.connect(title_re=".*Visual Studio Code.*")
    # app_dialog = app.top_window_()
    # app_dialog.Restore()

    time.sleep(2)

    ############################## 3. 시작전 size fix 하기 #########################################
    win32gui.MoveWindow(hwndMain, fix_data[0], fix_data[1], fix_data[2], fix_data[3], True)

    chat_position = win32gui.GetWindowRect(open_chat)
    start_x = (chat_position[0] + chat_position[2])/2
    start_y = chat_position[1] + (chat_position[3] - chat_position[1])/10
    end_y = chat_position[3] - (chat_position[3] - chat_position[1])/10

    print(chat_position)


############################################# telegram 활성화하기 ################################


'''
"https://t.me/kikawoTest")) # 내 채널
"https://t.me/rassiro_channel")) #라시로 채널
"https://t.me/davidstocknew")) # 주식 소리통 네오 
https://t.me/SeeBullReal 주식 돋보기 
https://t.me/FastStockNews 주식 대장주 
https://t.me/YeouidoStory2 여의도 스토리 
https://t.me/corevalue 가치 투자 클럽
https://t.me/FastStockNews 급등일보 
https://t.me/gangsec 요약
'''


urls = ["https://t.me/kikawoTest","https://t.me/FastStockNews","https://t.me/gangsec" ]

urls2 = ["https://t.me/rassiro_channel", "https://t.me/davidstocknew", "https://t.me/corevalue", "https://t.me/mainnews"  ]



################################################## main 함수 ##############################################
msginitflg = 0
msginitflg2 = 0
pre_result = []
pre_result2 = []
result = []
result2 = []

async def main():
    
    global tele_message, msginitflg, msginitflg2, pre_result, pre_result2, result, result2


    My_phone = "+82-10-7764-0568"
    password = "!rlarkddnjs1"
    resol = 0.1

    total_news = []

    await client.start(phone = My_phone)

    while True:

        if check_timer(resol,3) == 1 or msginitflg == 0:
            futures = [asyncio.ensure_future(telebot_message(url)) for url in urls] # 태스크(퓨처) 객체를 리스트로 만듦
            result = await asyncio.gather(*futures)
            msginitflg = 1
        

        # if check_timer(resol,3) == 1 or msginitflg2 == 0:
        #     futures = [asyncio.ensure_future(telebot_message(url)) for url in urls2] # 태스크(퓨처) 객체를 리스트로 만듦
        #     result2 = await asyncio.gather(*futures)
        #     msginitflg2 = 1
        

        send_msg, cmd, new1 = determine_msg(pre_result, result)
        # send_msg2, cmd2, new2 = determine_msg(pre_result2, result2)

        # if new1 == 1:
        #     print(send_msg)
        # if new2 == 1:
        #     print(send_msg2)
        
        # generate_news(send_msg, send_msg2, cmd)
        

        pre_result = result 
        # pre_result2 = result2 

        if msginitflg == 1:
            if "pause" in result[0]: # pause 는 유지하고 있어야함 
                pass
            elif "채팅금지" in result[0]:   
                chat_delete()
            elif "re" in send_msg and cmd == 1:
                init()
            else:
                Check_linkMsg(send_msg, cmd) # 0.5 초마다 실행 
                pass

        time.sleep(resol)

if __name__ == "__main__":
    loop = asyncio.get_event_loop() 
    loop.run_until_complete(main())
    loop.close() 





