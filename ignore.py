# -*- coding: utf-8 -*-
"""
Created on Mon Dec  3 11:52:24 2018

@author: Yang
"""

import win32gui
import win32con
import win32com
import win32com.client
import threading

flag = False
def terminate():
    global flag
    while (1):
        hwnd = win32gui.FindWindow(None, 'Password')
        if hwnd != 0:
            win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
            break
        if flag == True:
            break


filename = 'pw.pptx'


PPTApplication = win32com.client.Dispatch("PowerPoint.Application")
PPTApplication.DisplayAlerts = False

print('thread start')
t = threading.Thread(target=terminate)
t.start()

try:
    print('open ppt')
    PPTApplication.Presentations.Open(filename,ReadOnly=True,WithWindow=True)
except:
    print('exception error')
    t.join()
    None

if t.is_alive():
    print('thread alive')
    flag = True
    t.join()
    print('thread dead')

print('done')