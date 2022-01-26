# pip3 install pytz python-dateutil httpx
###
import json
from datetime import datetime
import pytz
from dateutil.parser import parse
import csv
import httpx
###

def getCalendarJSON(bearer_token, teams_cookie, url):
    headers = {
        'authorization': bearer_token,
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36',
        'sec-ch-ua-platform': 'Windows',
        'referer': 'https://teams.microsoft.com/_',
        'cookie': teams_cookie
    }

    teams_json_resp = httpx.get(url, headers=headers)
    return teams_json_resp.json()

def writeCSVFile(write_data):
    csv_header = ['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time']
    with open('./weeklyevents.csv', 'w', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(csv_header)
        writer.writerows(write_data)

def convertToPHXTime(time, format, pytz_timezone):
    datestr = time
    dt = parse(datestr)
    localtime = dt.astimezone (pytz.timezone(pytz_timezone))
    if format == 'date':
        return localtime.strftime("%m-%d-%Y")
    elif format == 'time':
        return localtime.strftime("%H:%M %p")
    else:
        print("No format found")

## Change variables to match your configuration
pytz_timezone = "America/Phoenix"
bearer_token = "Bearer XXXXX"
teams_cookie = "MC1=GUID=3140f995c5XXXX"
start_date = "2022-01-23"
end_date = "2022-01-30"
gmt_offset = "T07:00:00.000Z"
url = "https://teams.microsoft.com/api/mt/amer/beta/me/calendarEvents?StartDate={startDate}{gmtOffset}&EndDate={endDate}{gmtOffset}".format(startDate=start_date, endDate=end_date, gmtOffset=gmt_offset)

calendarJson = getCalendarJSON(bearer_token, teams_cookie, url)

events_array = []

for event in calendarJson['value']:
    temp_array = []
    temp_array.append(event['subject'])
    temp_array.append(convertToPHXTime(event['startTime'], 'date', pytz_timezone))
    temp_array.append(convertToPHXTime(event['startTime'], 'time', pytz_timezone))
    temp_array.append(convertToPHXTime(event['endTime'], 'date', pytz_timezone))
    temp_array.append(convertToPHXTime(event['endTime'], 'time', pytz_timezone))
    events_array.append(temp_array)

writeCSVFile(events_array)
