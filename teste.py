from reports import REPORTS
from collections import defaultdict
from datetime import datetime

reports = REPORTS(global_=True)

total_drs = reports.getDrsTotalNum()
data = reports.getAllDrsData()


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
    return dict(occurrences)

if __name__ == '__main__':
    dates = getAllDates(data)
    occurrences = count_occurrences_by_month(dates)
    print(occurrences)