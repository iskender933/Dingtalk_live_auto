# -*- coding:utf-8 -*-
# 作者：伊斯卡尼达尔
# 创建：2022-10-31
# 更新：2022-10-31
# 用意：钉钉直播签到自动化程序

"""
我提供你们的是python源码，运行程序前请下载相关的库，
签到部分是通过点击坐标实现的，每个人的电脑屏幕不同，所以你得改一下坐标！
请尊重成果，有疑问可以QQ：1457436639联系我.
"""

import os
import time

import pyautogui
import win32api
import win32con
import win32gui
from PIL import ImageGrab


hwnd_title = {}
def get_all_hwnd(hwnd, mouse):
    if (win32gui.IsWindow(hwnd) and
            win32gui.IsWindowEnabled(hwnd) and
            win32gui.IsWindowVisible(hwnd)):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})


def get_all_child_window(parent):
    if not parent:
        return
    hwndChildList = []
    win32gui.EnumChildWindows(
        parent, lambda hwnd, param: param.append(hwnd), hwndChildList)
    return hwndChildList


def setforeground_window(window_handle):
    while True:
        try:
            win32gui.SetForegroundWindow(window_handle)  # 强制在最前端显示
            #Iskender()
            return
        except:
            time.sleep(0.1)


def close_analyse_window():
    print("检测到直播结束...")
    # 关闭统计窗口
    print("尝试关闭统计窗口...")
    try:
        win32gui.EnumWindows(get_all_hwnd, 0)
        for h, t in hwnd_title.items():
            if win32gui.GetClassName(h) == "DingEAppWnd":
                setforeground_window(h)  # 使当前窗口在最前
                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                print("成功关闭统计窗口")
                break
    except:
        print("关闭失败，可能统计窗口已经被关闭了...")
    print("本轮检测结束，" + str(delay_time) + "s后进行下一轮检测")
    return


def get_live_window_isopened(live_window_handle):
    while True:
        time.sleep(delay_time)
        # 查找所有窗口标题和句柄 StandardFrame
        isOpened_temp = False
        win32gui.EnumWindows(get_all_hwnd, 0)
        try:
            for h, t in hwnd_title.items():
                if t == '钉钉' and win32gui.GetClassName(h) == "StandardFrame":
                    isOpened_temp = True
        except:
            close_analyse_window()
            return
        if not isOpened_temp:
            close_analyse_window()
            return


delay_time = 60

def Iskender():
    var = 1;s = 0
    while var == 1 :  # 表达式永远为 true
        time.sleep(5)
        s=s+1
        pyautogui.click(x=1082, y=615)# 单击  
        print("点击",s,"次")

# 查找所有窗口标题和句柄 StandardFrame_DingTalk
win32gui.EnumWindows(get_all_hwnd, 0)
for h, t in hwnd_title.items():
    if t != '钉钉':
        continue
    if win32gui.GetClassName(h) != "StandardFrame_DingTalk":  # 跳过其他窗口捕获主窗口
        continue
    print("获取到钉钉窗口句柄 " + str(h))
    ding_main_window_handle = h
    break

try:
    if ding_main_window_handle is None:
        exit(0)
except:
    print("请先打开钉钉窗口")
    os.system("pause")
    exit(0)

ding_child_list = get_all_child_window(ding_main_window_handle)

print(ding_child_list)

win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MAXIMIZE)  # 最大化
setforeground_window(ding_main_window_handle)  # 强制在最前端显示
time.sleep(2)
for c in ding_child_list:
    if win32gui.GetWindowText(c) == "Chrome Legacy Window":
        ding_chrome_window = c
print("获取到聊天窗口 " + str(ding_chrome_window))
win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MINIMIZE)  # 完成后最小化

print("准备就绪，3s后开始检测")

while True:
    win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MAXIMIZE)  # 最大化

    time.sleep(0.3)
    x_start, y_start, x_end, y_end = win32gui.GetWindowRect(ding_chrome_window)
    box = (x_start, y_start, x_end, y_end)
    image = ImageGrab.grab(box)
    win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MINIMIZE)  # 截图完成后最小化
    if image.getpixel((5, 5)) == (224, 237, 254):
        live_window_handle_temp = ''
        print("检测到有直播可进入，准备检测是否已启动直播页面...")
        # 查找所有窗口标题和句柄 StandardFrame
        win32gui.EnumWindows(get_all_hwnd, 0)
        isOpened = False
        for h, t in hwnd_title.items():
            if t != '钉钉':
                continue
            if win32gui.GetClassName(h) != "StandardFrame":  # 跳过其他窗口捕获直播窗口
                continue
            isOpened = True
            live_window_handle_temp = h
            break
        if isOpened:
            print("直播窗口已打开，开始监控直播窗口变化")
            Iskender()
            get_live_window_isopened(live_window_handle_temp)
        else:
            print("直播窗口未打开，尝试打开")
            win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MAXIMIZE)  # 最大化
            setforeground_window(ding_main_window_handle)  # 强制在最前端显示
            time.sleep(0.5)
            left, top, right, bottom = win32gui.GetWindowRect(ding_chrome_window)
            move_x = left + 5
            move_y = top + 5
            win32api.SetCursorPos((move_x, move_y))  # 鼠标挪到点击处
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)  # 鼠标左键按下
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)  # 鼠标左键抬起
            print("启动完成，等待直播进入...")
            time.sleep(8)
            print("开始获取直播窗口...")
            # 查找所有窗口标题和句柄 StandardFrame
            win32gui.EnumWindows(get_all_hwnd, 0)
            isOpened = False
            for h, t in hwnd_title.items():
                if t != '钉钉':
                    continue
                if win32gui.GetClassName(h) != "StandardFrame":  # 跳过其他窗口捕获直播窗口
                    continue
                isOpened = True
                live_window_handle_temp = h
                break
            if isOpened:
                print("直播窗口已打开，开始监控直播窗口变化")
                Iskender()
                get_live_window_isopened(live_window_handle_temp)
      
            else:
                print("打开失败，" + str(delay_time) + "s后再次尝试...")
    else:
        print("未检测到直播，本轮检测完毕")
    image.save('temp.jpg')
    time.sleep(delay_time)