# -*- coding: utf-8 -*-
__author__ = 'echeg'

import cv2
import numpy as np
from PIL import  ImageGrab
import time
import win32api, win32con
import random

#cv_im = cv2.CreateImageHeader(pi.size, cv2.IPL_DEPTH_8U, 1)
#cv2.SetData(cv_im, pi.tostring())

#print result.argmax()
#print result.shape
#print np.unravel_index(result.argmax(),result.shape)
def mouse_move (x,y, speed = 10):

    pos = win32api.GetCursorPos()
    step_x=abs(pos[0]-x)/speed
    step_y=abs(pos[1]-y)/speed
    if not step_x:
        step_x=1
    if not step_y:
        step_y=1
        #print (step_x, abs(pos[0]-result_x), speed)
    for i in range (speed):
        pos = win32api.GetCursorPos()
        if x>pos[0] and y>pos[1]:
            win32api.SetCursorPos((pos[0]+step_x, pos[1]+step_y))
            #print (pos[0]+step_x, pos[1]+step_y)
        elif x>pos[0] and y<pos[1]:
            win32api.SetCursorPos((pos[0]+step_x, pos[1]-step_y))
            #print (pos[0]+step_x, pos[1]+step_y)
        elif x<pos[0] and y>pos[1]:
            win32api.SetCursorPos((pos[0]-step_x, pos[1]+step_y))
            #print (pos[0]+step_x, pos[1]+step_y)
        elif x<pos[0] and y<pos[1]:
            win32api.SetCursorPos((pos[0]-step_x, pos[1]-step_y))
            #print (pos[0]+step_x, pos[1]+step_y)
        time.sleep(time_mouse_move_sleep)

    win32api.SetCursorPos((x, y))
    return True

def mouse_click (type):

    if type==1:# one left click
        time.sleep(0.10)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0) # нажали левую кнопку мыши
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0) # отжали левую кнопку мыши
        time.sleep(0.10)

    elif type==2:# one rigth click
        time.sleep(0.10)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0) # нажали левую кнопку мыши
        time.sleep(0.10)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0) # отжали левую кнопку мыши
        time.sleep(0.10)

    elif type==3:# double left click
        time.sleep(0.25)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0) # нажали левую кнопку мыши
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0) # отжали левую кнопку мыши
        time.sleep(0.25)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0) # нажали левую кнопку мыши
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0) # отжали левую кнопку мыши
        time.sleep(0.25)

    if type==4:# one left click not up
        time.sleep(0.10)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0) # нажали левую кнопку мыши

    if type==5:
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0) # отжали левую кнопку мыши
        time.sleep(0.10)

    return True

def mouse_wheel():
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0, -120)
    time.sleep(0.10)

    return True

def find_image_in_screen(image):
    im = ImageGrab.grab()
    #im.save("test.png")
    #cv_im = cv2.imread("test.png")
    cv_im = cv2.cvtColor(np.asarray(im), cv2.COLOR_RGB2BGR)

    result = cv2.matchTemplate(cv_im,image,cv2.TM_CCOEFF_NORMED)
    #print result
    confidence = 0.95 # 95% совпадения
    match_indices = np.arange(result.size)[(result>confidence).flatten()]
    #print ()
    a = np.unravel_index(match_indices,result.shape)
    #print (len(a[0]))
    #print image.shape[0]
    #print a

    if len(a[0])==0:
        # ничего не нашло возможно список закончился
        exit()

    for i in range (len(a[0])):
        #
        mouse_move(a[1][i]+random.randint(0,image.shape[1])/2,a[0][i]+random.randint(0,image.shape[0])/2)
        if click_or_not:
            mouse_click(1)
        time.sleep(time_after_click)
        #mouse_move(a[1][i],a[0][i])


time_after_click = 0
time_mouse_move_sleep = 0.03
click_or_not = False
people = 500

image = cv2.imread("template.png")

time.sleep(5)

time1 = time.time()
for i in range(people/6):
    find_image_in_screen(image)
    for i in range(10):
        mouse_wheel()

print ("time all %f" % (time.time()-time1,))
#for i in range(10):
#    mouse_wheel()