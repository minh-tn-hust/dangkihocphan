import pandas as pd
from openpyxl import load_workbook
import math
from datetime import datetime, timedelta
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

def insertClassName(dataEvent, nameEvent):#hàm này sử dụng để chèn thêm tên lớp vào kèm cùng với mã lớp
    className = [] 
    classCode = dataEvent['Lớp học']
    for code in classCode:
        if math.isnan(code):#kiểm tra xem có phải là NaN hay không, nếu phải thì next
            continue;
        else:
            x = nameEvent[nameEvent['Mã lớp'] == code]['Tên lớp']
            if (x.empty):
                x = nameEvent[nameEvent['Mã lớp kèm'] == code]['Tên lớp']
            className.insert(len(className),x.values[0])
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
            print(row['Tên lớp'],row['Thời gian'],row['Phòng học'],dateBegin)

for weekBegin in range(28,44):
    for day in range(2,7):
        event = dataEvent[dataEvent['Thứ'] == day]
        if (not (event.empty)):
            processDateEvent(event,dateBegin,weekBegin)
        dateBegin+= timedelta(days=1)
    dateBegin+=timedelta(days=2)


'''
cứ mỗi một dòng là một event trong lịch, việc của mình là viết một hàm
để có thể lấy được hết thông tin trong đấy
có thể sử dụng vòng lặp để có thể tạo được sự kiện, hoặc là mình sẽ sử dụng thuộc tính recurrence để có thể đỡ phải vòng lặp :v :
các thuộc tính cần quan tâm 
    summary: tên sẽ lấy từ mã lớp :)) toang rồi
    start và end: sẽ lấy được từ cột thời gian
    location sẽ lấy được từ cột phòng học
    lặp sự kiện sẽ sử dụng  tuần học để có thể lạp được
- công việc đầu tiên là phải sync được mã lớp và lấy được tên mã lớp
    đã xong việc link mã lớp và tên lớp
- công việc việc tiếp theo là phải tạo ra được những sự kiện từ để có thể up lên google calendar theo tuần học
    đã xong việc còn lại là link api và phần xử lý số liệu này nữa là được
'''




