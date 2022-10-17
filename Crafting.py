import pytesseract
from pytesseract import Output
import pyautogui as gui
import random
import re
import time
from PIL import ImageGrab
from difflib import SequenceMatcher
import ctypes

pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

Searchbox = [67, -221, 341, -195]
Itembox = [67, -181, 367, 591]
Countbox = [589,392,598,403]
Craftbox = [627,384,767,406]
pixelsearch = [939, 390, 949, 400]



def move_cords(location, offsetx, offsety, randofset):
    location = [location[0] + zeropoint[0],location[1] + zeropoint[1], location[2] - location[0], location[3] - location[1]]
    location = [location[0]+ random.randint(0+randofset,location[2]-randofset) + offsetx, location[1] + random.randint(0+randofset,location[3] - randofset) + offsety]
    gui.moveTo(location[0], location[1], 1, gui.easeInOutQuad) 
    gui.click(location[0] ,location[1]) 

def delete_line():
    gui.moveTo(zeropoint[0] + 334, zeropoint[1] - 207, 1, gui.easeInOutQuad)
    time.sleep(0.5)
    gui.dragTo(zeropoint[0] + 51, zeropoint[1] - 208, 1, gui.easeInOutQuad)

def move_slider():
    gui.moveTo(zeropoint[0] + 373, zeropoint[1] - 150, 1, gui.easeInOutQuad)
    time.sleep(0.5)
    gui.dragTo(zeropoint[0] + 373, zeropoint[1] +555, 1, gui.easeInOutQuad)

def wait_done():
    search = [pixelsearch[0] + zeropoint[0], pixelsearch[1] + zeropoint[1], pixelsearch[2] + zeropoint[0], pixelsearch[3] + zeropoint[1]]
    while 1:
        time.sleep(1)
        img = ImageGrab.grab(search)
        img.save('plz.png')
        img = img.load()
        colour = img[1,1]
        print(colour[1])
        if colour[1] > 150:
            break
       


def move_id(name, randofset):
    slid = 0
    while 1:
        certainty =[]
        itemnames = []
        name = name.replace(" ","")
        name = name[:22]
        search = [Itembox[0] + zeropoint[0], Itembox[1] + zeropoint[1], Itembox[2] + zeropoint[0], Itembox[3] + zeropoint[1]]
        img = ImageGrab.grab(search)
        img.save('test.png')
        d = pytesseract.image_to_data(img, output_type=Output.DICT,lang='eng', config='--psm 4 --oem 1 -c tessedit_char_whitelist=qwertyuiopQWERTYUIOPasdfghjklASDFGHJKLzxcvbnmZXCVBNM(1234567890)')
        for y in range(len(d['text'])):
            test = re.sub("\(.*?\)","", d['text'][y])
            test = test[:22]
            itemnames.append(test)
            certainty.append(SequenceMatcher(None, name, test).ratio())
        count = max(certainty) 
        x = certainty.index(count)
        print(count)
        
        if count > 0.90:
            break
        if count > 0.80:
            string = "Are\n "+str(name)+" \nand\n "+str(itemnames[x])+" \nthe same item?"
            var = ctypes.windll.user32.MessageBoxW(0, string, "", 4)
            if var == 6:
                break
            print(var)

        if slid == 1:
            print("Not in list")
            return

        move_slider()
        slid = 1

        
    
    location = [d['left'][x]+ random.randint(0+randofset,d['width'][x]-randofset) + search[0], d['top'][x] + random.randint(0+randofset,d['height'][x] - randofset) + search[1]]
    gui.moveTo(location[0], location[1], 1, gui.easeInOutQuad) 
    gui.click(location[0] ,location[1]) 

    return 




zeropoint = gui.locateOnScreen('C:\\Users\\lowie\\wardrobe.jpg', confidence=0.9)
print(zeropoint)

for x in range(len(items[3])):
    delete_line()
    time.sleep(0.5) 
    gui.write(items[3][x], interval=0.1)
    move_id(items[3][x],7)
    time.sleep(0.5)
    move_cords(Countbox,0,0,1)
    time.sleep(0.5)
    gui.write(str(items[1][x]), interval=0.1)
    time.sleep(0.5)
    move_cords(Craftbox,0,0,5)
    wait_done()

print("Done")