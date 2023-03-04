from openpyxl import Workbook,load_workbook
import pandas as pd
import json
import os
from flask import Flask, jsonify, request
  
app = Flask(__name__)
# class Taktal:
#     def __init__(self, age, name):
#         self.age = age
#         self.name = name
  
  
# class Tpty:
#     def __init__(self, proof, name, proof_no):
#         self.proof = proof
#         self.name = name
#         self.proof_no = proof_no
def sheetFinder(panel_name):
    switcher = {
        "taktal": "Taktal_Details.xlsx",
        "taktal_food": "Taktal_Details.xlsx",
        "seva": "Seva_details.xlsx",
        "300": "300rs_ticket_details.xlsx",
        "accomodation": "Accomodation_details.xlsx",
        "login": "Login_details.xlsx",
        "credit_card":"cc_details.xlsx",
    }
    return switcher.get(panel_name,"nothing")

@app.route('/taktal', methods = ['GET','POST'])
def getTaktalDetails():
    panel_name = "taktal"
    workbook = load_workbook(sheetFinder(panel_name))
    TaktalDict = {}
    sheets_list = workbook.sheetnames
    for sheet in sheets_list:
        df = pd.read_excel(sheetFinder(panel_name),sheet_name=sheet,usecols=['Name','Age'])
        if(df.empty == 1):
            continue
        df = df.dropna(axis='rows',subset=['Name'])
        TaktalDict[sheet] = df.to_dict(orient="records")
    return jsonify(TaktalDict)

@app.route('/tirupati/<panel_name>', methods = ['GET','POST'])
def getTirupatiDetails(panel_name):
    workbook = load_workbook(sheetFinder(panel_name))
    TptyDict = {}
    sheets_list = workbook.sheetnames
    for sheet in sheets_list:
        df = pd.read_excel(sheetFinder(panel_name),sheet_name=sheet,usecols=['Name','Proof','Proof_no'])
        if(df.empty == 1):
            continue
        df = df.dropna(axis='rows',subset=['Name'])
        TptyDict[sheet] = df.to_dict(orient="records")
    return jsonify(TptyDict)

@app.route('/login', methods = ['GET','POST'])
def getLoginDetails():
    panel_name = "login"
    df = pd.read_excel(sheetFinder(panel_name),sheet_name="Sheet1",usecols=['Login_id'])
    df = df.dropna(axis='rows',subset=['Login_id'])
    result = df.to_json(orient="index")
    parsed = json.loads(result)
    return jsonify(parsed)

@app.route('/cc', methods = ['GET','POST'])
def getCCDetails():
    panel_name = "credit_card"
    df = pd.read_excel(sheetFinder(panel_name),sheet_name="Sheet1",usecols=['cc_no'])
    df = df.dropna(axis='rows',subset=['cc_no'])
    result = df.to_json(orient="index")
    parsed = json.loads(result)
    return jsonify(parsed)

@app.route('/execute/<panel_name>/<int:sheet_no>', methods = ['GET','POST'])
def execute_script(panel_name,sheet_no):
    match panel_name:
        case "taktal":
            os.system(f"python taktal_automate.py {sheet_no}")
        case "seva":
            os.system(f"python seva_automate.py {sheet_no}")
        case "300":
            os.system(f"python 300_tpty.py {sheet_no}")
        case "accomodation":
            os.system(f"python accomodation_automate.py {sheet_no}")
        case "login":
            os.system(f"python login_automate.py {sheet_no}")
        case "credit_card":
            os.system(f"python cc_automate.py {sheet_no}")
        case "taktal_food":
            os.system(f"python taktal_food.py {sheet_no}")
        case other:
            print("error: panel not found")
#     return history(panel_name,sheet_no)

# def history(panel_name,sheet_no):

if __name__ == "__main__":
    # getTaktalDetails("taktal")
    # getTirupatiDetails("seva")
    app.run(debug=True)
    # getLoginDetails("login")
    # execute_script("taktal",1)


