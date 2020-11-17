# -*- coding: utf-8 -*-
from ctypes import *
from time import sleep
import os
import math
import random
import win32api
import win32con
import mp3play
import threading

advapi32 = windll.LoadLibrary('advapi32.dll')
user32 = windll.LoadLibrary('user32.dll')
size_correct=win32api.GetSystemMetrics (1)/864.0

VK_CODE = {
    'backspace': 0x08,
    'tab': 0x09,
    'clear': 0x0C,
    'enter': 0x0D,
    'shift': 0x10,
    'ctrl': 0x11,
    'alt': 0x12,
    'pause': 0x13,
    'caps_lock': 0x14,
    'esc': 0x1B,
    'spacebar': 0x20,
    'page_up': 0x21,
    'page_down': 0x22,
    'end': 0x23,
    'home': 0x24,
    'left_arrow': 0x25,
    'up_arrow': 0x26,
    'right_arrow': 0x27,
    'down_arrow': 0x28,
    'select': 0x29,
    'print': 0x2A,
    'execute': 0x2B,
    'print_screen': 0x2C,
    'ins': 0x2D,
    'del': 0x2E,
    'help': 0x2F,
    '0': 0x30,
    '1': 0x31,
    '2': 0x32,
    '3': 0x33,
    '4': 0x34,
    '5': 0x35,
    '6': 0x36,
    '7': 0x37,
    '8': 0x38,
    '9': 0x39,
    'a': 0x41,
    'b': 0x42,
    'c': 0x43,
    'd': 0x44,
    'e': 0x45,
    'f': 0x46,
    'g': 0x47,
    'h': 0x48,
    'i': 0x49,
    'j': 0x4A,
    'k': 0x4B,
    'l': 0x4C,
    'm': 0x4D,
    'n': 0x4E,
    'o': 0x4F,
    'p': 0x50,
    'q': 0x51,
    'r': 0x52,
    's': 0x53,
    't': 0x54,
    'u': 0x55,
    'v': 0x56,
    'w': 0x57,
    'x': 0x58,
    'y': 0x59,
    'z': 0x5A,
    'numpad_0': 0x60,
    'numpad_1': 0x61,
    'numpad_2': 0x62,
    'numpad_3': 0x63,
    'numpad_4': 0x64,
    'numpad_5': 0x65,
    'numpad_6': 0x66,
    'numpad_7': 0x67,
    'numpad_8': 0x68,
    'numpad_9': 0x69,
    'multiply_key': 0x6A,
    'add_key': 0x6B,
    'separator_key': 0x6C,
    'subtract_key': 0x6D,
    'decimal_key': 0x6E,
    'divide_key': 0x6F,
    'F1': 0x70,
    'F2': 0x71,
    'F3': 0x72,
    'F4': 0x73,
    'F5': 0x74,
    'F6': 0x75,
    'F7': 0x76,
    'F8': 0x77,
    'F9': 0x78,
    'F10': 0x79,
    'F11': 0x7A,
    'F12': 0x7B,
    'F13': 0x7C,
    'F14': 0x7D,
    'F15': 0x7E,
    'F16': 0x7F,
    'F17': 0x80,
    'F18': 0x81,
    'F19': 0x82,
    'F20': 0x83,
    'F21': 0x84,
    'F22': 0x85,
    'F23': 0x86,
    'F24': 0x87,
    'num_lock': 0x90,
    'scroll_lock': 0x91,
    'left_shift': 0xA0,
    'right_shift ': 0xA1,
    'left_control': 0xA2,
    'right_control': 0xA3,
    'left_menu': 0xA4,
    'right_menu': 0xA5,
    'browser_back': 0xA6,
    'browser_forward': 0xA7,
    'browser_refresh': 0xA8,
    'browser_stop': 0xA9,
    'browser_search': 0xAA,
    'browser_favorites': 0xAB,
    'browser_start_and_home': 0xAC,
    'volume_mute': 0xAD,
    'volume_Down': 0xAE,
    'volume_up': 0xAF,
    'next_track': 0xB0,
    'previous_track': 0xB1,
    'stop_media': 0xB2,
    'play/pause_media': 0xB3,
    'start_mail': 0xB4,
    'select_media': 0xB5,
    'start_application_1': 0xB6,
    'start_application_2': 0xB7,
    'attn_key': 0xF6,
    'crsel_key': 0xF7,
    'exsel_key': 0xF8,
    'play_key': 0xFA,
    'zoom_key': 0xFB,
    'clear_key': 0xFE,
    '+': 0xBB,
    ',': 0xBC,
    '-': 0xBD,
    '.': 0xBE,
    '/': 0xBF,
    '`': 0xC0,
    ';': 0xBA,
    '[': 0xDB,
    '\\': 0xDC,
    ']': 0xDD,
    "'": 0xDE,
    '`': 0xC0}


def MoveTo(x, y):
    user32.SetCursorPos(int((x)*size_correct), int(y*size_correct))


def DrawTo(x1, y1, x2, y2):
    MoveTo(int(x1), int(y1))
    user32.mouse_event(0x0002, 0, 0, 0, 0)
    MoveTo(int(x2), int(y2))
    user32.mouse_event(0x0004, 0, 0, 0, 0)
    sleep(0.005)



# http://mathworld.wolfram.com/HeartCurve.html


def DrawHeart(x, y, enlarge=1,theta=0, degreeStep=1):
    nx, ny = 0, 0
    theta = -math.pi * theta / 180
    n = [16, 13, 5, 2, 1, math.cos(theta), math.sin(theta)]
    for i in range(0, int(360 / degreeStep)):
        t = i * (math.pi * degreeStep / 180)
        xorig0 = n[0] * math.sin(t) ** 3
        yorig0 = n[1] * math.cos(t) - n[2] * math.cos(2 * t) - \
            n[3] * math.cos(3 * t) - n[4] * math.cos(4 * t)
        xorig=xorig0*n[5]-yorig0*n[6]
        yorig=xorig0*n[6]+yorig0*n[5]
        ox = nx
        oy = ny
        nx = x - enlarge * xorig
        ny = y - enlarge * yorig
        DrawTo(ox, oy, nx, ny)


def DrawEllipse(x, y, a, b, theta=0, degreeStep=1, randomStrength=0,  isMouth=False):
    nx, ny = 0, 0
    theta = -math.pi * theta / 180
    n0 = [a, b, math.cos(theta), math.sin(theta)]
    n = n0
    if isMouth:
        loopStart = int(210 / degreeStep)
        loopEnd = int(335 / degreeStep)
    else:
        loopStart = 0
        loopEnd = int(360 / degreeStep)
    step = math.pi * degreeStep / 180
    for i in range(loopStart, loopEnd):
        t = -i * step
        if randomStrength != 0:
            nd = [random.uniform(-a * randomStrength, randomStrength), random.uniform(-b * randomStrength,
                                                                                      b * randomStrength), random.uniform(-math.cos(theta) * randomStrength,
                                                                                                                          math.cos(theta) * randomStrength), random.uniform(-math.sin(theta) * randomStrength,
                                                                                                                                                                            math.sin(theta) * randomStrength)]
            n = [n0[j] + nd[j] for j in xrange(0, 4)]
        xorig = n[0] * math.cos(t) * n[2] - n[1] * math.sin(t) * n[3]
        yorig = n[0] * math.cos(t) * n[3] + n[1] * math.sin(t) * n[2]
        ox = nx
        oy = ny
        nx = x + xorig
        ny = y + yorig
        DrawTo(ox, oy, nx, ny)

def DrawXAM(x,y,fontSize):
    DrawTo(x,y-fontSize,x+fontSize,y)#X
    sleep(0.3)
    DrawTo(x+fontSize,y-fontSize,x,y)
    sleep(0.5)
    DrawTo(x+3*fontSize/2,y-fontSize,x+fontSize,y)#A
    sleep(0.3)
    DrawTo(x+3*fontSize/2,y-fontSize,x+2*fontSize,y)
    sleep(0.3)
    DrawTo(x+5*fontSize/4,y-fontSize/2,x+7*fontSize/4,y-fontSize/2)
    sleep(0.5)
    DrawTo(x+9*fontSize/4,y-fontSize,x+2*fontSize,y)#M
    sleep(0.3)
    DrawTo(x+9*fontSize/4,y-fontSize,x+5*fontSize/2,y)
    sleep(0.3)
    DrawTo(x+5*fontSize/2,y,x+11*fontSize/4,y-fontSize)
    sleep(0.3)
    DrawTo(x+11*fontSize/4,y-fontSize,x+3*fontSize,y)



def key_input(str=''):
    for c in str:
        win32api.keybd_event(VK_CODE[c], 0, 0, 0)
        win32api.keybd_event(VK_CODE[c], 0, win32con.KEYEVENTF_KEYUP, 0)
        sleep(0.01)

def play_mp3(file_address):
	if os.path.exists(file_address):
		clip = mp3play.load(file_address)
		clip.play()
		while True:
			sleep(10)

if __name__ == "__main__":
    os.startfile("mspaint")
    sleep(2)

    bgm_thread = threading.Thread(target=play_mp3,args=(u'bgm.mp3',))
    bgm_thread.setDaemon(True)
    bgm_thread.start()

    sleep(0.5)
    user32.ShowWindow(user32.GetForegroundWindow(), 3)
    sleep(0.5)
    tts = open('TempVoice.vbs', 'w')
    text = u'''
    Set objSpv=CreateObject("SAPI.SpVoice")
    objSpv.Speak"我正在画画。"
    '''
    textB=text.encode('gbk')
    tts.write(textB)
    tts.close()

    win32api.keybd_event(VK_CODE['ctrl'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['e'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['e'], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE['ctrl'], 0, win32con.KEYEVENTF_KEYUP, 0)
    key_input('%d' %int(1366*size_correct))
    key_input(['tab'])
    key_input('%d' %int(768*size_correct))
    key_input(['enter'])

    os.startfile("TempVoice.vbs")
    DrawEllipse(500, 400, 250, 200, 10, 0.25, 0.03)  # head
    DrawEllipse(280, 250, 35, 30, 48, 0.5, 0.8)  # ear
    DrawEllipse(740, 270, 35, 30, -27, 0.5, 0.7)
    DrawEllipse(400, 370, 50, 30, 40, 0.5, 0.1)  # eye frame
    DrawEllipse(600, 350, 50, 30, -25, 0.5, 0.1)
    DrawEllipse(400, 370, 10, 10, 40, 1, 0.5)  # eye ball
    DrawEllipse(600, 350, 10, 10, -25, 1, 0.5)
    DrawEllipse(500, 450, 6, 5, 0, 1, 0.3)  # nose
    DrawEllipse(500, 470, 100, 50, 5, 0.5, 0.05, True)  # mouth

    DrawXAM(700,600,50)
    for x in xrange(0,22):
        DrawHeart(random.randint(100,1000), random.randint(200,650), random.uniform(0.5,2.5),random.uniform(0,360),1)  # heart

    tts = open('TempVoice.vbs', 'w')
    text = u'''
    Set objSpv=CreateObject("SAPI.SpVoice")
    objSpv.Speak"画完了!"
    '''
    textB=text.encode('gbk')
    tts.write(textB)
    tts.close()
    os.startfile("TempVoice.vbs")

    key_input(['print_screen'])
    sleep(0.3)
if os.path.exists("TempVoice.vbs"):
    os.remove("TempVoice.vbs")

    win32api.keybd_event(VK_CODE['ctrl'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['alt'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['z'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['z'], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE['alt'], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE['ctrl'], 0, win32con.KEYEVENTF_KEYUP, 0)
    sleep(1)
    key_input('zz')
    key_input(['enter'])
    key_input(['enter'])
    sleep(1)
    win32api.keybd_event(VK_CODE['ctrl'], 0, 0, 0)
    win32api.keybd_event(VK_CODE['v'], 0, 0, 0)
    sleep(0.2)
    win32api.keybd_event(VK_CODE['v'], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE['ctrl'], 0, win32con.KEYEVENTF_KEYUP, 0)
    sleep(1)
    key_input(['enter'])
    sleep(1)
    key_input('zheshiwohuade')
    key_input(['spacebar', 'enter'])
    sleep(1)