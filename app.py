import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import pandas as pd
from dotenv import load_dotenv
import os
import sys

def convert_to_excel_column(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

load_dotenv()
url = os.environ.get("SHEET_URL")
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(os.environ.get("CREDENTIALS_FILE"), scope)
client = gspread.authorize(creds)

classes = ["A", "B", "C", "D", "E", "F", "IUP"]
modul = int(sys.argv[2])
type = sys.argv[1]
header = 2
side = 3

if type == "praktikum":
    contest = "PRAKTIKUM"
    n = 0
    r = 4
elif type == "remidi":
    contest = "REMIDI"
    n = 4
    r = 4
elif type == "revisi":
    contest = "REVISI"
    n = 8
    r = 3
else:
    print("Invalid type!")
    exit()

data = {}
for kelas in classes:
    api = f'https://www.its.ac.id/informatika/domjudge/api/v4/contests/{contest}-{kelas}-MODUL{modul}/scoreboard?strict=false'
    response = requests.get(api)
    if response.status_code == 200:
        data[kelas] = response.json()
    else:
        print('Error:', response.status_code)
        data[kelas] = None

sheet = client.open_by_url(url)
worksheet = sheet.worksheet("REKAP")
kelas_col = worksheet.col_values(3)
nrp_col = worksheet.col_values(1)

value = []
row = 0
for kelas in classes:
    class_data = data[kelas]
    if class_data is None:
        break
    
    df = pd.DataFrame(data[kelas]['rows'])
    
    for nrp_cell, kelas_cell in zip(nrp_col[header+row:], kelas_col[header+row:]):
        if kelas_cell != kelas:
            break

        team = df[df['team_id'] == nrp_cell]
        solved = [0] * r

        if not team.empty:
            problems = team.iloc[0]['problems']
            for i, problem in enumerate(problems):
                if problem['solved']:
                    solved[i] = 1
            value.append(solved)
        else:
            print("Praktikan tidak ditemukan!")
            value.append(solved)

        row += 1

if row > 0:
    range = f"{convert_to_excel_column(side + (1 + ((modul-1)*11)) + n)}{header+1}:{convert_to_excel_column(side + (1 + ((modul-1)*11)) + n + (r-1))}{header + row}"
    worksheet.update(value, range)