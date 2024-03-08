### 本脚本需要python3解释执行，请下载并安装python3.6及以上版本
### 安装完成后执行
### 	py AnalasysBTCbubbles.py
### 本脚本会自动下载所需数据并根据江卓尔先生的泡沫指数公式生成图表，并调用系统默认程序打开该图表。
###
### 如果运行时报缺少xlsxwriter,请找到pip.exe（python3安装后就有）并执行
### 	pip install xlsxwriter


from __future__ import print_function
import sys,os
from datetime import datetime
import urllib.request
import xlsxwriter, csv


price_file = "market_price.csv"
cap_file = "market_cap.csv"
addr_file = "unique_address.csv"
result_file = "btc_bubbles.xlsx"

def Schedule(blocks, blocksize, totalsize):
    sys.stdout.write(".")
    sys.stdout.flush()

def getfile(url, saveasfile):
    print("正在下载: ", url, " ==> ", saveasfile)
    urllib.request.urlretrieve(url, saveasfile, Schedule)
    print("")

def download_data():
    getfile("https://api.blockchain.info/charts/market-price?scale=1&timespan=all&daysAverageString=7&format=csv", price_file)
    #getfile("https://api.blockchain.info/charts/market-price?daysAverageString=7&scale=1&timespan=all&format=csv", price_file)

    getfile("https://api.blockchain.info/charts/market-cap?daysAverageString=7&timespan=all&format=csv", cap_file)
    getfile("https://api.blockchain.info/charts/n-unique-addresses?timespan=all&daysAverageString=7&format=csv",
            addr_file)

def print_version():
    print("")
    print("===================================")
    print("     比特币泡沫指数计算器 v1.0     ")
    print("===== By hanfuqishe@gmail.com =====")
    print("")
    print("根据江卓尔先生提出的泡沫指数计算公式自动下载最新数据并生成比特币泡沫指数")
    print("")


def main():
    print_version()
    download_data()

    workbook_result = xlsxwriter.Workbook(result_file)
    worksheet_result = workbook_result.add_worksheet()
    headings = ["date", "price", "price_log", "cap", "addresses", "bubbles"]
    worksheet_result.write_row(0, 0, headings)

    money = workbook_result.add_format({'num_format': '#,##0.00'})
    dateformat  = workbook_result.add_format({"num_format" : 'yyyy-mm-dd'})

    col = 0
    csv_reader = csv.reader(open(price_file, encoding="utf-8"))
    i=1
    for row in csv_reader:
        # worksheet_result.write_row(i, 0, row, money)
        date_time = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
        worksheet_result.write_datetime(i, col, date_time, dateformat)
        worksheet_result.write(i, col + 1, float(row[1]), money)
        worksheet_result.write_formula(i, col + 2, f"=log10(B{i+1:d} + 1)")
        i+=1

    col += 3
    csv_reader = csv.reader(open(cap_file, encoding="utf-8"))
    i=1
    for row in csv_reader:
        worksheet_result.write(i, col, float(row[1]), money)
        i+=1

    col += 1
    csv_reader = csv.reader(open(addr_file, encoding="utf-8"))
    i=1
    for row in csv_reader:
        worksheet_result.write(i, col, float(row[1]), money)
        worksheet_result.write_formula(i, col + 1, f"=D{i+1:d}/(E{i+1:d}^2-E{i+1:d})")
        i+=1

    rowcount=i
    # print("rowcount", i)

    chart_col = workbook_result.add_chart({"type": "line"})

    chart_col.add_series({
        "name": "=Sheet1!C1",
        "categories": f"=Sheet1!A2:A{rowcount:d}",
        "values": f"=Sheet1!C2:C{rowcount:d}",
        "line": {"color": "green"},
    })

    chart_col.add_series({
        "name": "=Sheet1!F1",
        "categories": f"=Sheet1!A2:A{rowcount:d}",
        "values": f"=Sheet1!F2:F{rowcount:d}",
        "line": {"color": "blue"},
        'y2_axis': True,
    })

    chart_col.set_title({"name": "BTC泡沫指数"})
    chart_col.set_x_axis({"name": "日期"})
    chart_col.set_y_axis({"name": "价格"})
    chart_col.set_y2_axis({"name": "泡沫"})

    #chart_col.set_style(1)

    # worksheet_result.insert_chart("A10", chart_col, {"x_offset": 25, "y_offset": 10})
    chartsheet = workbook_result.add_chartsheet()
    chartsheet.set_chart(chart_col)
    chartsheet.activate()

    workbook_result.close()


    print("打开图表...")
    os.system(result_file)

main()

