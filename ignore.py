# -*- coding: utf-8 -*-
"""
Created on Mon Dec  3 11:52:24 2018

@author: Yang


MIT License

Copyright (c) [2018] [Yee Yang, Tan] (yeeyang.tan@live.com)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

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


if __name__ == "__main__":

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