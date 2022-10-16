from operator import itemgetter
from pyparsing import ZeroOrMore
import pytesseract
import pyautogui as gui
import random
import time

pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

Searchbox = [67, -221, 341, -195]
Itembox = [67, -141, 351, -127]

def move_cords(location, offsetx, offsety, randofset):
    location = [location[0] + zeropoint[0],location[1] + zeropoint[1], location[2] - location[0], location[3] - location[1]]
    location = [location[0]+ random.randint(0+randofset,location[2]-randofset) + offsetx, location[1] + random.randint(0+randofset,location[3] - randofset) + offsety]
    gui.moveTo(location[0], location[1], 1, gui.easeInOutQuad) 
    gui.click(location[0] ,location[1]) 


zeropoint = gui.locateOnScreen('C:\\Users\\lowie\\wardrobe.jpg', confidence=0.9)
print(zeropoint)

move_cords(Searchbox,0,0,5)
time.sleep(2)
gui.write("Zojja's Footwear", interval=0.1)
move_cords(Itembox,0,0,5)

