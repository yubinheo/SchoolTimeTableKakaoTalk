# -- Copyright © 2021 Yubin Heo -- #

from time import sleep
import win32con
import win32api
import win32gui
import datetime
import time
import schedule

global name
global contents
global contents_temp
global TimeTable106_MON
global TimeTable106_TUE
global TimeTable106_WED
global TimeTable106_THU
global TimeTable106_FRI
global time
global typeValue
global repeat

name = "이준서"

contents_temp = ""

time = ["월", "화", "수", "목", "금", "토", "일"]
TimeTable106_MON = ["전기회로 이론", "전기회로 이론", "수학", "체육", "진로", "한국사", "한문"]
TimeTable106_TUE = ["체육", "음악", "한문", "수학", "과학", "국어", "영어"]
TimeTable106_WED = ["전기회로 실습", "전기회로 실습", "전기회로 실습", "전기회로 실습", "한문", "직업", "NULL"]
TimeTable106_THU = ["영어", "직업", "사회", "정보", "정보", "국어", "한국사"]
TimeTable106_FRI = ["프로그래밍 이론", "프로그래밍 실습", "프로그래밍 실습", "프로그래밍 실습", "창체", "창체", "NULL"]

r = datetime.datetime.today().weekday()

todayWeekdayName = time[r]
    
typeValue = 0
repeat = 1

kakao_opentalk_name = name


# 8:30분 모든 시간표 출력
def SchedulePrintAllTimeTable():
    open_chatroom(kakao_opentalk_name)
    
    contents = '''★ 오늘의 시간표 ★\n\n'''
    if todayWeekdayName == "월" :
        for i in range(len(TimeTable106_MON)):
            contents_temp = TimeTable106_MON[i] + "\n"
            contents += contents_temp

    elif todayWeekdayName == "화" :
        for i in range(len(TimeTable106_TUE)):
            contents_temp = TimeTable106_TUE[i] + "\n"
            contents += contents_temp

    elif todayWeekdayName == "수" :
        for i in range(len(TimeTable106_WED)):
            contents_temp = TimeTable106_WED[i] + "\n"
            contents += contents_temp

    elif todayWeekdayName == "목" :
        for i in range(len(TimeTable106_THU)):
            contents_temp = TimeTable106_THU[i] + "\n"
            contents += contents_temp

    elif todayWeekdayName == "금" :
        for i in range(len(TimeTable106_FRI)):
            contents_temp = TimeTable106_FRI[i] + "\n"
            contents += contents_temp
            
    contents += "\nDevelopment by Yubin.Heo \n(모든 권리 보유)"

    text = contents
    kakao_sendtext(kakao_opentalk_name, text)


# 특정 시간마다 시간 알림 (조례 알림)
def SchedulePrintMorning():
    open_chatroom(kakao_opentalk_name)
    
    contents = '''★ 아침 조례 안내 ★

(AM) 8:40분 ~ 8:45분
지금은 아침 조례 시간입니다
담임 선생님과 조례를 진행해 주세요

링크 : https://url.kr/3owbgh

Development by Yubin.Heo
(모든 권리 보유)'''
    
    text = contents
    kakao_sendtext(kakao_opentalk_name, text)

# 특정 시간마다 시간 알림 (점심 알림)
def SchedulePrintEat():
    open_chatroom(kakao_opentalk_name)
    
    contents = '''★ 점심 시간 안내 ★

(PM) 12:20분 ~ 13:40분
지금은 점심 시간 입니다
맛밥 하세요~!

Development by Yubin.Heo
(모든 권리 보유)'''
    
    text = contents
    kakao_sendtext(kakao_opentalk_name, text)

# 특정 시간마다 시간 알림 (종례 알림)
def SchedulePrintEnd():
    open_chatroom(kakao_opentalk_name)
    
    contents = '''★ 종례 시간 안내 ★

(PM) 16:15분 ~ 16:30분
지금은 오후 종례 입니다

링크 : https://url.kr/3owbgh

Development by Yubin.Heo
(모든 권리 보유)'''
    
    text = contents
    kakao_sendtext(kakao_opentalk_name, text)

# 특정 시간마다 시간 알림 (수업 종료 알림)
def SchedulePrintClassEnd():
    open_chatroom(kakao_opentalk_name)
    
    contents = '''★ 수업 종료 안내 ★

현재 수업이 종료 되었습니다.
10분간 쉬는시간 후
수업 안내 해드리겠습니다.

다음 수업을 준비해 주세요.

Development by Yubin.Heo
(모든 권리 보유)'''
    
    text = contents
    kakao_sendtext(kakao_opentalk_name, text)


# 특정 시간마다 시간 알림
def SchedulePrintTimeTable():

    nowTimeTableObject = ""

    now = datetime.datetime.now()

    nowTime = now.strftime('%H:%M')

    i = 0

    if nowTime == "8:50" or nowTime == "8:51" or nowTime == "8:52":
        i = 0
    elif nowTime == "9:45" or nowTime == "9:46" or nowTime == "9:47":
        i = 1
    elif nowTime == "10:40" or nowTime == "10:41" or nowTime == "10:42":
        i = 2
    elif nowTime == "11:35" or nowTime == "11:36" or nowTime == "11:37":
        i = 3
    elif nowTime == "13:40" or nowTime == "13:41" or nowTime == "13:42":
        i = 4
    elif nowTime == "14:35" or nowTime == "14:36" or nowTime == "14:37":
        i = 5
    elif nowTime == "15:30" or nowTime == "15:31" or nowTime == "15:32":
        i = 6
        

    if todayWeekdayName == "월" :
        nowTimeTableObject = TimeTable106_MON[i]

    elif todayWeekdayName == "화" :
        nowTimeTableObject = TimeTable106_TUE[i]

    elif todayWeekdayName == "수" :
        nowTimeTableObject = TimeTable106_WED[i]

    elif todayWeekdayName == "목" :
        nowTimeTableObject = TimeTable106_THU[i]

    elif todayWeekdayName == "금" :
        nowTimeTableObject = TimeTable106_FRI[i]

    contents = ""

    contents += "★ 수업 시작 안내 ★\n\n"
    contents += nowTimeTableObject + " 수업이 시작되었습니다.\n"
    contents += "https://url.kr/e32yni \n위 링크의 해당 수업 클래스로 접속해서 수업을 들어주세요.\n"
    contents += "\nDevelopment by Yubin.Heo \n(모든 권리 보유)"

            
    open_chatroom(kakao_opentalk_name)
    
    text = contents
    kakao_sendtext(kakao_opentalk_name, text)

schedule.every().day.at("08:30").do(SchedulePrintAllTimeTable)

schedule.every().day.at("08:40").do(SchedulePrintMorning)

schedule.every().day.at("08:50").do(SchedulePrintTimeTable)
schedule.every().day.at("09:35").do(SchedulePrintClassEnd)

schedule.every().day.at("09:45").do(SchedulePrintTimeTable)
schedule.every().day.at("10:30").do(SchedulePrintClassEnd)

schedule.every().day.at("10:40").do(SchedulePrintTimeTable)
schedule.every().day.at("11:25").do(SchedulePrintClassEnd)

schedule.every().day.at("11:35").do(SchedulePrintTimeTable)
schedule.every().day.at("12:20").do(SchedulePrintClassEnd)

schedule.every().day.at("12:21").do(SchedulePrintEat)

schedule.every().day.at("13:40").do(SchedulePrintTimeTable)
schedule.every().day.at("14:25").do(SchedulePrintClassEnd)

schedule.every().day.at("14:35").do(SchedulePrintTimeTable)
schedule.every().day.at("15:20").do(SchedulePrintClassEnd)

schedule.every().day.at("15:30").do(SchedulePrintTimeTable)
schedule.every().day.at("16:15").do(SchedulePrintClassEnd)

schedule.every().day.at("16:16").do(SchedulePrintEnd)


def kakao_sendtext(chatroom_name, text):

    hwndMain = win32gui.FindWindow( None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RICHEDIT50W", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)



def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def open_chatroom(chatroom_name):

    hwndkakao = win32gui.FindWindow(None, "KakaoTalk")
    hwndkakao_edit1 = win32gui.FindWindowEx( hwndkakao, None, "EVA_ChildWindow", None)
    hwndkakao_edit2_1 = win32gui.FindWindowEx( hwndkakao_edit1, None, "EVA_Window", None)
    hwndkakao_edit2_2 = win32gui.FindWindowEx( hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    hwndkakao_edit3 = win32gui.FindWindowEx( hwndkakao_edit2_2, None, "Edit", None)


    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    sleep(1)
    SendReturn(hwndkakao_edit3)
    sleep(1)


def main():
    while True:
        schedule.run_pending()
        sleep(1)

if __name__ == '__main__':
    main()
