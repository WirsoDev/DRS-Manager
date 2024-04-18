# send report to aqSMTPapi 
from reports import REPORTS
import requests


def sendData():
    reports = REPORTS()
    data = reports.parseAllData()

    url = 'http://localhost:5000/drsreport'
    resp = requests.post(url, json=data)
    print(resp.text)

sendData()
