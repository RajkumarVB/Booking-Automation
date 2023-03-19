import pandas as pd
import pyautogui
from time import sleep
import sys

sheet_no = sys.argv[1]

df = pd.read_excel(r"Tatkal_details.xlsx", sheet_name=f"Sheet{sheet_no}")
user_details = df.to_dict(orient="records")
pyautogui.hotkey("alt", "tab")
sleep(0.5)
for person_no, person in enumerate(user_details, start=1):
    if(person_no!=1):
        for i in range(5):
            pyautogui.hotkey("shift","tab")
    pyautogui.write(person["Name"])
    pyautogui.hotkey("tab")
    pyautogui.write(str(person["Age"]))
    pyautogui.hotkey("tab")
    pyautogui.hotkey(person["Gender"])
    pyautogui.hotkey("tab")
    pyautogui.hotkey("tab")
    for singlekey in person["Preference"]:
        pyautogui.hotkey(singlekey)
    pyautogui.hotkey("tab")
    for singlekey in person["Food"]:
        pyautogui.hotkey(singlekey)
    pyautogui.hotkey("tab")
    if person_no != len(user_details) :
        pyautogui.hotkey('enter')
for i in range(3):
    pyautogui.hotkey('tab')
pyautogui.hotkey('backspace')
pyautogui.write(str(user_details[0]["Phone"]).split('.')[0])
for i in range(3):
    pyautogui.hotkey('tab')