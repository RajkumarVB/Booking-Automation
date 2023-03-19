import pandas as pd
import pyautogui

import sys

sheet_no = sys.argv[1]

df = pd.read_excel(r"300rs_ticket_details.xlsx", sheet_name=f"Sheet{sheet_no}")
user_details = df.to_dict(orient="records")

pyautogui.hotkey("alt", "tab")
for person_no, person in enumerate(user_details, start=1):

    pyautogui.write(str(person["Name"]))
    pyautogui.hotkey("tab")
    for alphabet in str(person["Age"]):
        pyautogui.hotkey(alphabet)
    pyautogui.hotkey("tab")
    each_person = int(person_no)
    if(str(person["Gender"]).casefold() == 'm'.casefold()):
        pyautogui.hotkey("enter")
        pyautogui.hotkey("tab")
        pyautogui.hotkey("tab")
        pyautogui.hotkey("enter")
    elif(str(person["Gender"]).casefold() == 'f'.casefold()):
        pyautogui.hotkey("enter")
        pyautogui.hotkey("tab")
        pyautogui.hotkey("enter")
    pyautogui.hotkey("tab")
    if(str(person["Proof"]).casefold() == 'a'.casefold()):
        pyautogui.hotkey("enter")
        pyautogui.hotkey("tab")
        pyautogui.hotkey("enter")
    elif(str(person["Proof"]).casefold() == 'p'.casefold()):
     pyautogui.hotkey("tab")
     pyautogui.hotkey("tab")
     pyautogui.hotkey("tab")
     pyautogui.hotkey("enter")
    elif(str(person["Proof"]).casefold() == 'v'.casefold()):
     pyautogui.hotkey("tab")
     pyautogui.hotkey("tab")
     pyautogui.hotkey("tab")
     pyautogui.hotkey("enter")
    pyautogui.hotkey("tab")
    pyautogui.write(str(person["Proof_no"]))
    pyautogui.hotkey('tab') 
pyautogui.hotkey('tab')   
pyautogui.hotkey("enter")

# pyautogui.hotkey('tab')
# pyautogui.hotkey('tab')
# pyautogui.write(str(user_details[0]["Phone"]).split('.')[0])
# pyautogui.hotkey('tab')
# pyautogui.hotkey('tab')
# pyautogui.hotkey("enter")
