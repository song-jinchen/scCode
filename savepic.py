'''
Author: error: error: git config user.name & please set dead value or install git && error: git config user.email & please set dead value or install git & please set dead value or install git
Date: 2023-07-14 11:17:08
LastEditors: error: error: git config user.name & please set dead value or install git && error: git config user.email & please set dead value or install git & please set dead value or install git
LastEditTime: 2023-08-08 09:08:14
FilePath: \autowork\savepic.py
Description: 这是默认设置,请设置`customMade`, 打开koroFileHeader查看配置 进行设置: https://github.com/OBKoro1/koro1FileHeader/wiki/%E9%85%8D%E7%BD%AE
'''

import pyautogui

# pyautogui.sleep(1)
c = True
while c:
    if pyautogui.locateCenterOnScreen('wechatheader.png', confidence=0.9) != None:
        
        pyautogui.hotkey('ctrl', 's')
        pyautogui.sleep(1)
        pyautogui.hotkey('enter')
        pyautogui.sleep(1)
        pyautogui.hotkey('right')
        pyautogui.sleep(3)
    elif pyautogui.locateCenterOnScreen('whatsappheader.png', confidence=0.9) != None:
        
        pyautogui.hotkey('ctrl', 's')
        pyautogui.sleep(1)
        pyautogui.hotkey('enter')
        pyautogui.sleep(1)
        pyautogui.hotkey('right')
        pyautogui.sleep(1)
    # elif pyautogui.locateCenterOnScreen('excelClick.png', confidence=0.9) != None:
        
    #     # pyautogui.write('6:30')
    #     # pyautogui.sleep(1)
    #     # pyautogui.hotkey('right')
    #     # pyautogui.sleep(1)
    #     pyautogui.write('15:30')
    #     pyautogui.sleep(1)
        
        
    # elif pyautogui.locateCenterOnScreen('deletecol.png', confidence=0.9) != None:
    #     # pyautogui.hotkey('ctrl','v') 
    #     pyautogui.hotkey('delete')
    #     pyautogui.sleep(1)
    