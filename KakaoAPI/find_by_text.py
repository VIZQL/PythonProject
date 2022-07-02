from calendar import c
# from pyautogui import *
from vsave import *
import clipboard
import pyautogui as pa
import time, win32con, win32api, win32gui, ctypes
from PIL import ImageGrab  # screenshot
import pytesseract
from pytesseract import Output

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


# # 오픈채팅 링크 공유하면 hide 하는 기능 
def hide_open_chat(Ox, Oy):
    # pa.click(button='left')
    # Ox,Oy = pa.locateCenterOnScreen("open_chat.png", confidence=0.7)
    try:
        pa.click(Ox, Oy,button='right')
        time.sleep(0.2)
        pa.click(Ox+73, Oy+248, button='left')
        time.sleep(0.2)
        pa.click(Ox+73+54, Oy+248+12, button='left')
        time.sleep(0.2)
        dx,dy = pa.locateCenterOnScreen("delete_check.png", confidence=0.7)
        pa.click(dx, dy, button='left')
    except:
        print("가릴 메세지를 찾을수 없습니다")

# # 오픈채팅 링크 공유사람 강퇴시키는 기능 
def exit_user():
    try:
        pa.click(button='left')
        Ox,Oy = pa.locateCenterOnScreen("open_chat.png", confidence=0.7)
        pa.click(Ox-196, Oy-55, button='left')
        Ex,Ey = pa.locateCenterOnScreen("exit_user.png", grayscale = True, confidence=0.6)
        pa.click(Ex, Ey, button='left')
        time.sleep(0.5)
        pa.click(Ex-35, Ey, button='left')
        time.sleep(0.2)
        pa.click(Ex-35, Ey-165, button='left')
        return Ox, Oy
    except:
        print("강퇴할 유저를 찾을수 없습니다")
        return Ox, Oy
    

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


pytesseract.pytesseract.tesseract_cmd = (r"C:\Program Files\Tesseract-OCR\tesseract") # needed for Windows as OS

screen =  ImageGrab.grab()  # screenshot
cap = screen.convert('L')   # make grayscale

data=pytesseract.image_to_boxes(cap,output_type=Output.DICT)
print(data)

# ### 가리기 비율 확인 ###########
# click : Point(x=848, y=770)
# 가리기 : Point(x=921, y=1018)
# 가리기>가리기 : Point(x=975, y=1030)
# 선택메세지 가리겠습니까 : Point(x=817, y=692)

######### 강퇴 비율 확인 ##########
# Point(x=758, y=794)
# Point(x=562, y=739)

##################################

# 1. 오픈채팅방 검색해서 열기 
open_chatroom(kakao_opentalk_name)

# 2. 채팅 치기 위해 핸들러 가져오기  
hwndMain = win32gui.FindWindow(None, kakao_opentalk_name)
hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None) # 채팅 치는부분 좌표 가져오기
open_chat = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None) # 채팅방 전체 좌표가져오기

# 3. size fix 하기 
win32gui.MoveWindow(hwndMain, fix_data[0], fix_data[1], fix_data[2], fix_data[3], True)

chat_position = win32gui.GetWindowRect(open_chat)
start_x = (chat_position[0] + chat_position[2])/2
y = (chat_position[1] + chat_position[3])/2
start_y = chat_position[1] + (chat_position[3] - chat_position[1])/10
end_y = chat_position[3] - (chat_position[3] - chat_position[1])/10

print(chat_position)

# 2. 오픈채팅방 좌표 가져오기 

while(0):
    # x = 850 
    # y = 585

    ## 메세지창 중간부터 아래까지 드래그 
    pa.moveTo(start_x, start_y)
    pa.mouseDown(button='left')
    pa.moveTo(start_x, end_y, 0.1)

    ## 메시지창 ctrl c 
    data = read_message() 
    time.sleep(0.1)

    if "http:" in data : ## 오픈채팅방 공유했는지 확인하기 
        Ox,Oy = exit_user()
        hide_open_chat(Ox,Oy)
    else: 
        pass
        # kakao_sendtext(hwndEdit, "오픈링크 없습니다")

    time.sleep(1)
    cnt = cnt+1
    if (cnt>10):
        break


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

            




