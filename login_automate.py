import pandas as pd
import pyautogui
import sys

row_no = int(sys.argv[1])
# print(sheet_no)
# pyautogui.TypingInterval = 0.1
df = pd.read_excel(r"Login_Details.xlsx", sheet_name="Sheet1")
login_details = df.to_dict(orient="records")
pyautogui.hotkey("alt", "tab")
pyautogui.write(str(login_details[row_no-1]["Login_id"]))
pyautogui.hotkey("tab")
pyautogui.write(str(login_details[row_no-1]["Password"]))
pyautogui.hotkey('tab')
