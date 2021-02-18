#Thư viện dùng để sử dụng GC API
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
import google.oauth2.credentials
import google_auth_oauthlib.flow
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES,redirect_uri='urn:ietf:wg:oauth:2.0:oob')
        creds = flow.run_console();
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
# Tới đây là đã xong xuôi phần key rồi 

#Thư viện dùng để xử lý thời khóa biểu
import pandas as pd
from openpyxl import load_workbook
import math
from datetime import datetime, timedelta
'''============= Xử lý dữ liệu từ file excel =========================='''
folderEvent = 'Book1.xlsx'
folderName = 'Book2.xlsx'
dataEvent = pd.read_excel(folderEvent,sheet_name = 'Sheet1')
nameEvent = pd.read_excel(folderName,sheet_name  = 'Sheet1')

def processColumn(data): # hàm dùng để lọc những cột mà không chứa thông tin gì cả
    for column in data:
        if (column.find("Unnamed") != -1):
            del data[column]

processColumn(dataEvent)
processColumn(nameEvent)

def insertClassName(dataEvent, nameEvent):#hàm này sử dụng để chèn thêm tên lớp vào kèm cùng với mã lớp và loại lớp 
    className = [] 
    typeClass = []
    classCode = dataEvent['Lớp học']
    for code in classCode:
        if math.isnan(code):#kiểm tra xem có phải là NaN hay không, nếu phải thì next
            continue;
        else:
            x = nameEvent[nameEvent['Mã lớp'] == code]['Tên lớp']
            y = nameEvent[nameEvent['Mã lớp'] == code]['Loại lớp']
            if (x.empty):
                x = nameEvent[nameEvent['Mã lớp kèm'] == code]['Tên lớp']
                y = nameEvent[nameEvent['Mã lớp kèm'] == code]['Loại lớp']
                className.insert(len(className),y.values[0] + ':' +x.values[0])
            else:
                className.insert(len(className),y.values[0] + ':' +x.values[0])
    cN = pd.DataFrame(className) #convert về cùng là DataFrame mới có thể thêm vào được
    dataEvent.insert(1,"Tên lớp",cN)

insertClassName(dataEvent, nameEvent)
print(dataEvent)

dateBegin = datetime(2021,2,22) 
weekBegin = 28 
weekEnd = 48

def processLearningWeek(ls,week):
    if (ls.find('-') != -1):
        ls = ls.split('-')
        if (week >= int(ls[0]) and week <= int(ls[1])):
            return True
        else: 
            return False
    else:
        if (ls.find(str(week)) != -1):
            return True
        else: 
            return False


def processDateEvent(event,dateBegin,week):
    print(week)
    for count in range (0,len(event)):
        row = event.iloc[count]
        if (processLearningWeek(row['Tuần học'],week)):
            sumary = row['Tên lớp']
            time = row['Thời gian'].split('-')
            time[0] = dateBegin.strftime("%Y-%m-%dT") + time[0] + ":00+07:00"
            time[1] = dateBegin.strftime("%Y-%m-%dT") + time[1] + ":00+07:00"
            location = row['Phòng học']
            print(sumary)
            print(time[0])
            print(time[1])
            print(location)
            events = {
              'summary': sumary,
              'location': location,
              'start': {
                'dateTime': time[0],
                'timeZone': 'Asia/Ho_Chi_Minh',
              },
              'end': {
                'dateTime': time[1],
                'timeZone': 'Asia/Ho_Chi_Minh',
              },
              'reminders': {
                'useDefault': False,
                'overrides': [
                  {'method': 'email', 'minutes': 24 * 60},
                  {'method': 'popup', 'minutes': 10},
                ],
              },
            }
            events = service.events().insert(calendarId='primary', body=events).execute()

for weekBegin in range(25,44):
    for day in range(2,7):
        event = dataEvent[dataEvent['Thứ'] == day]
        if (not (event.empty)):
            processDateEvent(event,dateBegin,weekBegin)
        dateBegin+= timedelta(days=1)
    dateBegin+=timedelta(days=2)

