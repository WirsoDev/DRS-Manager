# send report to aqSMTPapi 
from reports import REPORTS
import requests
from collections import defaultdict
from datetime import datetime


def sendData():
    #get global data for global report!

    global_report = REPORTS(global_=True)
    global_data = global_report.getAllDrsData()
    dates = getAllDates(global_data)
    occurrences = count_occurrences_by_month(dates)



    reports = REPORTS(global_=False)
    data = reports.parseAllData()
    data['occurrences'] = occurrences
    data['total_year'] = calculate_total_year(occurrences)
    url = 'http://localhost:5000/drsreport'
    resp = requests.post(url, json=data)
    print(resp.text)


#helpers

def calculate_total_year(occurrences_list):
    # Calculate the total values of the months
    return sum(count for _, count in occurrences_list)

def getAllDates(data):
    dates = []
    for d in data:
        date = data[d]['data_registo']
        current_month = datetime.now().month
        if int(date.split('-')[1]) != int(current_month):
            dates.append(date)
    return dates


def count_occurrences_by_month(dates):
    occurrences = defaultdict(int)
    
    for date_str in dates:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        month_key = date_obj.strftime("%b")
        occurrences[month_key] += 1
    
    # Convert the dictionary to a list of tuples and return
    return [(month, count) for month, count in occurrences.items()]



sendData()
