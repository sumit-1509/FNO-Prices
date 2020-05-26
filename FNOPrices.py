import requests
import json
import xlwings as xw

wb = xw.Book()
sht = wb.sheets[0]

url = 'https://www.nseindia.com/api/equity-stockIndices?index=SECURITIES%20IN%20F%26O'
headers = {'User-Agent': 'Mozilla/5.0'}
res = requests.get(url, headers= headers)
res.raise_for_status()
data = json.loads(res.text)
stocks_no = len(data["data"])

for i in range(0, stocks_no):
    stocks = data["data"][i]["symbol"]
    ltp = data["data"][i]["lastPrice"]
    year_high = data["data"][i]["yearHigh"]
    year_low = data["data"][i]["yearLow"]
    sht.range("a" + str(i + 2)).value = stocks
    sht.range("b" + str(i + 2)).value = ltp
    sht.range("c" + str(i + 2)).value = year_high
    sht.range("d" + str(i + 2)).value = year_low
