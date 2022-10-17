import pytesseract
from pytesseract import Output
import pyautogui as gui
import random
import re
import time
from PIL import ImageGrab

pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

Searchbox = [67, -221, 341, -195]
Itembox = [67, -181, 367, 591]

items = [[19685], [3], [22], ['Fulgurite']]

def move_cords(location, offsetx, offsety, randofset):
    location = [location[0] + zeropoint[0],location[1] + zeropoint[1], location[2] - location[0], location[3] - location[1]]
    location = [location[0]+ random.randint(0+randofset,location[2]-randofset) + offsetx, location[1] + random.randint(0+randofset,location[3] - randofset) + offsety]
    gui.moveTo(location[0], location[1], 1, gui.easeInOutQuad) 
    gui.click(location[0] ,location[1]) 

def move_id(name, randofset):
    name = name.replace(" ","")
    search = [Itembox[0] + zeropoint[0], Itembox[1] + zeropoint[1], Itembox[2] + zeropoint[0], Itembox[3] + zeropoint[1]]
    img = ImageGrab.grab(search)
    img.save('test.png')
    d = pytesseract.image_to_data(img, output_type=Output.DICT,lang='eng', config='--psm 4 --oem 1 -c tessedit_char_whitelist=qwertyuiopQWERTYUIOPasdfghjklASDFGHJKLzxcvbnmZXCVBNM(1234567890)')
    for x in range(len(d['text'])):
        test = re.sub("\(.*?\)","", d['text'][x])
        if test == name:
            location = [d['left'][x]+ random.randint(0+randofset,d['width'][x]-randofset) + search[0], d['top'][x] + random.randint(0+randofset,d['height'][x] - randofset) + search[1]]
            gui.moveTo(location[0], location[1], 1, gui.easeInOutQuad) 
            gui.click(location[0] ,location[1]) 
            return
    return "Could not find"

zeropoint = gui.locateOnScreen('C:\\Users\\lowie\\wardrobe.jpg', confidence=0.9)
print(zeropoint)

move_cords(Searchbox,0,0,7)
time.sleep(0.5)

# gui.hotkey('ctrl','a')

time.sleep(0.5) 
gui.write(items[3][0], interval=0.1)
move_id(items[3][0],7)

