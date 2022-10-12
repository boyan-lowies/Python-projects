
from PIL import ImageGrab
import random
import cv2 as cv
import numpy as np
from xmlrpc.client import boolean
import openpyxl as op
import pyautogui as gui
import time
import pytesseract

L_type = "G"
L_count = "H"
L_instant = "I"

N = 2



item_list=[] 
item_count=[]
item_instant=[]
zeropoint=[]
location_TPB = [314, 0, 38, 37]

First_ofset = [392,81,780,129]
TP_buy_ofset = [403,412,526,438]
TP_sell_ofset = [744,413,879,434]



pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

def move_png(img, offsetx, offsety, randofset):
    location = gui.locateOnScreen(img, confidence=0.9)
    print(img,location)
    location = [location[0]+ random.randint(0+randofset,location[2]-randofset) + offsetx, location[1] + random.randint(0+randofset,location[3] - randofset) + offsety]
    gui.moveTo(location[0], location[1], 1, gui.easeInOutQuad) 
    gui.click(location[0] ,location[1]) 


def move_cords(location, offsetx, offsety, randofset):
    location = [location[0]+ random.randint(0+randofset,location[2]-randofset-location[0]) + offsetx, location[1] + random.randint(0+randofset,location[3] - randofset-location[1]) + offsety]
    gui.moveTo(location[0], location[1], 1, gui.easeInOutQuad) 
    gui.click(location[0] ,location[1]) 

def find_pos(count):
    search = [First_ofset[0] + zeropoint[0], First_ofset[1] + zeropoint[1], First_ofset[2] + zeropoint[0], First_ofset[3]+zeropoint[1]]
    img = ImageGrab.grab(search)
    result = pytesseract.image_to_string(img)
    result = result[:-1].replace("\n", " ") 
    if item_list[count] == result:
        return search
    return

def find_price():
    search = [TP_buy_ofset[0] + zeropoint[0], TP_buy_ofset[1] + zeropoint[1], TP_buy_ofset[2] + zeropoint[0], TP_buy_ofset[3]+zeropoint[1]]
    img = ImageGrab.grab(search)
    img.save('temp.png')
    img = cv.imread('test', cv.IMREAD_GRAYSCALE)
    img = cv.threshold(img,100,255,cv.THRESH_BINARY)
    img.save('test1.png')
    price_b = pytesseract.image_to_string(img, config=r'--oem 3')
    
    search = [TP_sell_ofset[0] + zeropoint[0], TP_sell_ofset[1] + zeropoint[1], TP_sell_ofset[2] + zeropoint[0], TP_sell_ofset[3]+zeropoint[1]]
    img = ImageGrab.grab(search)
    img = cv.threshold(img,100,255,cv.THRESH_BINARY)
    price_s = pytesseract.image_to_string(img, config=r'--oem 3'  )
    
    
    return price_b, price_s

#--------------------#
#     Main Code      #
#====================#

wb = op.load_workbook('C:\\Users\\lowie\\Desktop\\pytest.xlsx')
sheet = wb['Next']

while 1:
    adress= L_type + str(N)  
    if sheet[adress].value is None:
        break

    item_list.append(sheet[adress].value)
    adress= L_count + str(N)  
    item_count.append(sheet[adress].value)
    adress = L_instant + str(N)
    item_instant.append(bool(sheet[adress].value))
    
    N += 1
    

gui.click(location_TPB[0] + random.randint(0,location_TPB[2]) ,location_TPB[1]+ random.randint(0,location_TPB[3]))
time.sleep(0.5)

move_png('tp.png', 0, 46,5)

time.sleep(2.0)

zeropoint = gui.locateOnScreen('search.png', confidence=0.9)
move_png('search.png',0,-1, 5)
time.sleep(1.0)

N = 1
item_list[N] = "Lorekeeper Axe Skin"

gui.write(item_list[N], interval=0.1)

time.sleep(1.0)

cords = find_pos(N)
move_cords(cords,0,0,5)

time.sleep(2)

price = find_price()
print(price)


