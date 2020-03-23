from gspread_formatting import CellFormat, format_cell_range, Color
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from datetime import datetime, date
from gspread import authorize
from pickle import load, dump
from os.path import exists
from pyodbc import connect

students = {}
groups = {}
con, cur, client, sheet = None, None, None, None

class credentials_wrapper(Credentials):
    def __init__(self, creds : Credentials):
        super().__init__(creds.token, creds.refresh_token, creds.id_token, creds.token_uri, creds.client_id, creds.client_secret, creds.scopes, creds.quota_project_id)
        self.access_token = creds.token

def get_credentials(scopes : list or tuple) -> Credentials: #note type safety
    credentials = None
    if exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = load(token)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', scopes)
            credentials = flow.run_local_server(port = 0)
            with open('token.pickle', 'wb') as token:
                dump(credentials, token)

    return credentials

def write_row(sheet, data, row):
    cells = sheet.range('A' + str(row) + ':E' + str(row))
    for i, j in zip(cells, data):
        i.value = j
    sheet.update_cells(cells)

def get_students(cursor, block):
    global groups
    return [i[0] for i in cursor.execute(f'SELECT EmployeeId FROM tblEmployeeGroup WHERE GroupId = {groups[block]}').fetchall()]

def get_recent_records(cursor, block):
    today_records = cursor.execute(f'SELECT EmployeeId, DateClockedIn FROM tblTimeEntry WHERE DateValue(DateClockedIn) = #{datetime(date.today().year, date.today().month, date.today().day)}#').fetchall()
    ids = set(get_students(cursor, block))
    today_records = filter(lambda i: i[0] in ids, today_records)
    return today_records

def init():
    global con, cur, client, sheet, groups
    con = connect(driver = 'Microsoft Access Driver (*.mdb, *.accdb)', dbq = r'C:\ProgramData\LotHill Solutions LLC\TimeDrop\timeclock.mdb', pwd = 'jz073314')
    cur = con.cursor()
    client = authorize(credentials_wrapper(get_credentials(['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'])))

    if exists('sheet_link.pickle'):
        with open('sheet_link.pickle', 'rb') as infile:
            link = load(infile)
        sheet = client.open_by_key(link).sheet1
        sheet.clear()
    else:
        sheet = client.create('Attendance')
        s = open('sheet_link.pickle', 'wb')
        s.close()
        with open('sheet_link.pickle', 'wb') as outfile:
            dump(sheet.id, outfile)
        sheet = sheet.sheet1
    
    for i, j in cur.execute('SELECT GroupName, GroupId FROM tblGroup').fetchall():
        groups[i] = j

init()
print(*get_recent_records(cur, '1st Block'), sep = '\n')

cur.close()
con.close()