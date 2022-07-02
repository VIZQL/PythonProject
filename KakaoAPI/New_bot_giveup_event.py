from calendar import c
# from pyautogui import *
from vsave import *
import clipboard
import pyautogui as pa
import time, win32con, win32api, win32gui, ctypes
import re
import pywinauto 
from telethon import TelegramClient, events, sync
import asyncio

########################################### 텔레그램 API 사용을 위한 함수 ############################
# Remember to use your own values from my.telegram.org!
api_id = '14078381'
api_hash = '1c92f649600a09317f38c1df4f9e1ee4'
client = TelegramClient('KKW', api_id, api_hash)
News_Tele = ""




async def teleProc():
    @client.on(events.NewMessage(chats= "https://t.me/kikawoTest")) # 내 채널
    @client.on(events.NewMessage(chats= "https://t.me/mainnews")) # main news 채널
    @client.on(events.NewMessage(chats= "https://t.me/sedaily_worldwide_news")) # 국제뉴스 채널 서울경제
    @client.on(events.NewMessage(chats= "https://t.me/rassiro_channel")) # AI 찌라시로
    @client.on(events.NewMessage(chats= "https://t.me/davidstocknew")) # 주식소리통 NEO
    
    async def my_event_handler(event):
        print(event.raw_text)
        News_Tele = event.raw_text

    client.start()
    client.run_until_disconnected()


######################################### 카카오톡 오픈 채팅방 처리 함수 ###############################
kakao_opentalk_name = '메모장'
Master = ['YRStockLeader']
fix_data = [450, 170, 840, 1000]


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
        pa.click(button='left')
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
def exit_user(dx, dy, line_cnt):
    try:
        a,b = dx-40, dy-line_cnt*17 # 프로필 위치 계산 
        pa.click(a, b, button='left') # 프로필 위치 계산 
        time.sleep(0.2)

        # Ex,Ey = a-191, b+341 # 프로필 위치로부터 내보내기 좌표계산
        # pa.click(Ex,Ey, button='left')
        # time.sleep(0.2)

        ################## 그림인식이 잘 안먹힘 ###############
        Ex,Ey = pa.locateCenterOnScreen("exit_user.png", grayscale = True, confidence=0.5)
        time.sleep(0.2)
        pa.click(Ex, Ey, button='left')
        ###############################
        
        pa.click(Ex-40, Ey, button='left') ## 내보내기 / 내보내기&신고 중에 선택
        time.sleep(0.2)
        pa.click(Ex-40, Ey-165, button='left') ## 최종 내보내기 
    except:
        print("강퇴할 유저를 찾을수 없습니다")

    

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
def kakao_sendtext(hwndEdit, text):
    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

cnt = 0


## 마우스 현재 위치 확인 코드 
while 0:
    print(pa.position())
    time.sleep(1)
    cnt = cnt+1
    if(cnt > 10):
        break


# pa.FAILSAFE = False

########################## 0. 시작전 다 창 아래로 깔기#################################
# pa.hotkey('winleft', 'd')


########################## 1.오카방 열기 #################################
open_chatroom(kakao_opentalk_name) # 오픈 채팅방 열기 

########################## 2. 채팅 치기 위해 핸들러 가져오기 #####################################
hwndMain = win32gui.FindWindow(None, kakao_opentalk_name)
hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None) # 채팅 치는부분 좌표 가져오기
open_chat = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None) # 채팅방 전체 좌표가져오기


######################### 3. 오픈채팅방 검색해서 활성화 , 비쥬얼 스튜디오도 열기######################

# app = pywinauto.application.Application()  # 오카방 열기
# app.connect(handle=open_chat)
# app_dialog = app.top_window_()
# app_dialog.Restore()

# app = pywinauto.application.Application()  # 비쥬얼 스튜디오 열기 
# app.connect(title_re=".*PYTHON.*")
# app_dialog = app.top_window_()
# app_dialog.Restore()

# time.sleep(2)



############################## 3. 시작전 size fix 하기 #########################################
win32gui.MoveWindow(hwndMain, fix_data[0], fix_data[1], fix_data[2], fix_data[3], True)

chat_position = win32gui.GetWindowRect(open_chat)
start_x = (chat_position[0] + chat_position[2])/2
y = (chat_position[1] + chat_position[3])/2
start_y = chat_position[1] + (chat_position[3] - chat_position[1])/10
end_y = chat_position[3] - (chat_position[3] - chat_position[1])/10

# print(chat_position)


############################################# telegram 활성화하기 ######################
loop = asyncio.get_event_loop()
loop.run_until_complete(teleProc())

################################### 4. 메인 코드 ###########################################

while(1):

    ## 메세지창 중간부터 아래까지 드래그 
    pa.moveTo(start_x, start_y)
    pa.mouseDown(button='left')
    pa.moveTo(start_x, end_y, 0.1)

    ## 메시지창 ctrl + c 
    data = read_message() 
    time.sleep(0.1)

    if "http" in data : ## 오픈채팅방 공유했는지 확인하기 
        data_split = re.split(r'\[([\w]+)\]' , data)
        data_split.pop(0)

        for i in range(-1,-len(data_split),-2):
            if "http" in data_split[i]:
                line_cnt = data_split[i].rstrip().count('\n')
                Ox,Oy,success = find_openchat_postion()
                if success == 1:
                    exit_user(Ox, Oy,line_cnt)
                    hide_open_chat(Ox, Oy)
    else: 
        pass
        # kakao_sendtext(hwndEdit, "오픈링크 없습니다")

    time.sleep(1)
    cnt = cnt+1
    if (cnt>10):
        break

#########################################################################################################
'''
while True:
    image_me = locateCenterOnScreen('send_check.png')
    image_master = locateCenterOnScreen('send_mastercheck.png')
    image_user = locateCenterOnScreen('send_usercheck.png')
    if not image_me == None or (not image_user == None and not image_master == None): # 방장이나 부방장이 명령어를 사용했을때
        master = True if not image_me == None else False # 방장이 명령어를 사용했는가?
        x,y = image_me if master else image_master # 느낌표 !의 위치를 구하기 
        doubleClick(x+10, y-10)
        data = read_message() # 메시지 읽어오기 
        time.sleep(0.1)
        click(x-30, y+30)
        if data == "!안녕": # 일반명령어
            send_message(send="(bot) 안녕하세요!")
        elif data == "!마지막대화":
            click(x+50, y-50, clicks=2, button="right")
            copy_x, copy_y = position()
            click(copy_x+10, copy_y+140)
            click(x-30, y+30)
            read_last = clipboard.paste().replace('[','').split('] ')
            send_message(send=f'(bot) 마지막으로 {read_last[0]} 님이 {read_last[1]} 에 {read_last[2]}라고 전송하셨어요 ')
        elif data == "!추가":
            click(x+10, y-10, clicks=3)
            data = read_message().strip('!추가').split('/')
            if not len(data) ==3:
                click(x-30, y+30)
                send_message(send = "(bot)잘못된구문이에요")
            else:
                save = data[2]
                command[data[0]] = data[1] #명령어 추가 
                click(x-30, y+30)
                send_message(send = f'(bot) [{data[0]}] / [{data[1]}] 명령어가 추가되었어요 (save = {save})')
                if save == "True": upload('command', command)
        elif data == "!도배":
            click(x+10, y-10, clicks = 3)
            data = read_message().strip("!초기화")
            if data == "True":
                command.clear() #명령어초기화 
                click(x-30, y+30)
                send_message(send=f'(bot) 모든 명령어를 삭제했어요')
                upload('command', {})
            else:
                click(x-30,y+30)
                send_message(send=f'bot 초기화를 할 권한이 없어요')

'''

            




