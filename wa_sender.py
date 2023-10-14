"""
This is a simple project created to send Whatsapp messages using Python script to send large volume of texts.

Please note that this library is meant to work on 1-monitor and you will have to tinker with the mouse movement if you are working with multiple monitors.

To run the code:
- create & activate the environment
-- mkdir <dir> # Create directory
-- cd <dir> # Go to created directory
-- python<version> -m venv <virtual-environment-name> # Create env
-- source env/bin/activate # Activate env
- install dependencies / requirements (one-time only)
-- pip install -r requirements.txt
- update excel document
- update image filepath
- run script
- Deactivate env before exiting
-- ~ deactivate

Author: @ahijmomo
Date Created: 4 October 2023 
"""

########## Libraries ##########
import pandas as pd         # Reading excel
import pywhatkit as pwk     # Sending whatsapp via WA Web
import pyautogui            # Mouse control
import keyboard as k        # Keyboard control
import datetime
import time 


########## Import Sample files ##########
rec = pd.read_excel('./input/save_the_date_test.xlsx')
img = './input/KWY_save_the_date.png'

#print(rec.loc[0, 'number'])
########## Iterate through df and send messages via Chome Whatsapp ##########

for idx in range(len(rec)): 
    
    num = '+' + str(rec.loc[idx, 'number'])
    msg = rec.loc[idx, 'opening_msg'] + ' ' + rec.loc[idx, 'name'] + ' ' + rec.loc[idx, 'closing_msg']

    # Send message
    pwk.sendwhats_image(num, img, msg, datetime.datetime.now().hour, datetime.datetime.now().minute + 1, 30)

    # Wait and Press enter
    pyautogui.click(1700, 900)
    k.press_and_release('enter')
    time.sleep(10)
    print(f"{idx + 1} message is sent, {len(rec) - (idx + 1)} left to send...")
    
print('Sent successfully') 

#print(pyautogui.size())
 
 