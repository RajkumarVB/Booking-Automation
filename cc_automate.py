import pandas as pd
import pyautogui
import sys

row_no = int(sys.argv[1])
row_no = row_no - 1
# print(sheet_no)
# pyautogui.TypingInterval = 0.1
df = pd.read_excel(r"cc_details.xlsx", sheet_name="Sheet1")
login_details = df.to_dict(orient="records")
pyautogui.hotkey("alt", "tab")
pyautogui.write(str(login_details[row_no]["cc_no"]))
pyautogui.hotkey("tab")
pyautogui.hotkey("tab")
pyautogui.write(str(login_details[row_no]["name"]))
pyautogui.hotkey("tab")
for alphabet in str(login_details[row_no]["expiry_month"]):
        pyautogui.hotkey(alphabet)
pyautogui.hotkey("tab")
# pyautogui.write(str(login_details[row_no]["expiry_month"]))
# pyautogui.hotkey('tab')
for alphabet in str(login_details[row_no]["expiry_year"]):
        pyautogui.hotkey(alphabet)
pyautogui.hotkey("tab")
# pyautogui.write(str(login_details[row_no]["expiry_year"]))
# pyautogui.hotkey('tab')
pyautogui.write(str(login_details[row_no]["cvv"]))
