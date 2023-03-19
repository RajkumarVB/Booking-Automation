import pandas as pd
import pyautogui

import sys

sheet_no = sys.argv[1]
# print(sheet_no)


# pyautogui.TypingInterval = 0.1
# list_no = int(input("Enter list no : "))


df = pd.read_excel(r"Seva_details.xlsx", sheet_name=f"Sheet{sheet_no}")
user_details = df.to_dict(orient="records")
# print(str(user_details[0]["Phone"]).split('.')[0])
# print(user_details)


pyautogui.hotkey("alt", "tab")
for person_no, person in enumerate(user_details, start=1):

    pyautogui.write(person["Name"])
    pyautogui.hotkey("tab")
    for alphabet in str(person["Age"]):
        pyautogui.hotkey(alphabet)
    pyautogui.hotkey("tab")
    pyautogui.hotkey(person["Gender"])
    pyautogui.hotkey("tab")
    pyautogui.hotkey(person["Proof"])
    if person["Proof"].casefold == "p".casefold() :
        pyautogui.hotkey(person["Proof"])
    pyautogui.hotkey("tab")
    pyautogui.write(str(person["Proof_no"]))
    if person_no != len(user_details) :
        pyautogui.hotkey('tab')

pyautogui.hotkey('tab')
pyautogui.write(str(user_details[0]["Email"]))
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.write(str(user_details[0]["Phone"]).split('.')[0])
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
