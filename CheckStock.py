# -*- coding: utf-8 -*-
"""
Created on Sun Feb 18 00:05:21 2024

@author: JIA
"""

import subprocess
import time

import win32ui
import dde

import pandas as pd

import talib
from talib import abstract

import requests

from pywinauto import Application
import pywinauto

# LINE Notify 權杖
token = 'nkss6NkZp619xiVjXtQo0psTWIrgimB1PT5gbQGl8Gh'

# 要發送的訊息
#message = '這是用 Python 發送的訊息'
def send_line_notify(message):
    headers = { "Authorization": "Bearer " + token }
    data = { 'message': message }
    response = requests.post('https://notify-api.line.me/api/notify', 
                             headers=headers, data=data)
    return response
    



# return data ('Open', 'High', 'Low', 'Close', 'Volume', 'K', 'D', 'BBAND_u', 'BBAND_l') of given stock
# conversation: Connection to the DDE server
# id: Stock ID
def StockData(conversation, id):
    # Request data from the server, returned data: string
    requested_data = conversation.Request(id + ".TW-Day-250")
    
    # Split into rows
    rows = requested_data.split(';')
    
    # Split each row into columns
    data = [row.split(',') for row in rows]
    
    # Convert the list of lists into a DataFrame
    df = pd.DataFrame(data)
    
    # Set the column names
    df.columns = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']
    
    # Convert from string to float
    df['Open'] = df['Open'].astype(float)
    df['High'] = df['High'].astype(float)
    df['Low'] = df['Low'].astype(float)
    df['Close'] = df['Close'].astype(float)
    df['Volume'] = df['Volume'].astype(float)
    
    # KD值
    RSV = calculate_rsv(df)
    df['K'] = calculate_k(df, RSV)
    df['D'] = calculate_d(df, df['K'])
    
    # 布林通道，參數與goodinfo一致
    df['BBAND_u'], df['BBAND_m'], df['BBAND_l'] = talib.abstract.BBANDS(df['Close'], timeperiod=21, nbdevup=2.0, nbdevdn=2.0, matype=0)
    
    # 均線，參數與goodinfo一致
    # 月均線:21日。季均線:62日。半年均線:124日。年均線:248日。
    
    o = float(df.iloc[-1]['Open'])
    h = float(df.iloc[-1]['High'])
    l = float(df.iloc[-1]['Low'])
    c = float(df.iloc[-1]['Close'])
    v = float(df.iloc[-1]['Volume'])
    k = float(df.iloc[-1]['K'])
    d = float(df.iloc[-1]['D'])
    bu = float(df.iloc[-1]['BBAND_u'])
    bm = float(df.iloc[-1]['BBAND_m'])
    bl = float(df.iloc[-1]['BBAND_l'])
    return (o, h, l, c, v, k, d, bu, bl, df)
    
def calculate_rsv(df, n=9):
    rsv = ((df['Close'] - df['Low'].rolling(window=n).min()) / (df['High'].rolling(window=n).max() - df['Low'].rolling(window=n).min())) * 100
    return rsv

def calculate_k(df, rsv, n=9):
    k = [None] * len(df)    
    k[n-2] = 50.  # Set the first value of K to 50
    for i in range(n-1, len(df)):
        k[i] = (2. * k[i-1] + rsv[i]) / 3.
    return k

def calculate_d(df, k, n=9):
    d = [None] * len(df)    
    d[n-2] = 50.  # Set the first value of D to 50
    for i in range(n-1, len(df)):
        d[i] = (2. * d[i-1] + k[i]) / 3.
    return d



# ht: High threshold compared by Close price. set as ZERO to disable.
# lt: Low threshold compared by Close price. set as ZERO to disable.
def CheckToday(id, name, conversation, ht, lt):
    # stock data ('Open', 'High', 'Low', 'Close', 'Volume', 'K', 'D', 'BBAND_u', 'BBAND_l')
    (o, h, l, c, v, k, d, bu, bl, _) = StockData(conversation, id)
    
    if ht != 0:
        if c >= ht:
            message = id + " " + name + "\n"
            message = message + "價格接近高點:" + "\n"
            message = message + "價格: " + str(c) + "\n"
            message = message + "高點: " + "{:.2f}".format(ht) + "\n"
            send_line_notify(message)
    
    if lt != 0:
        if c <= lt:
            message = id + " " + name + "\n"
            message = message + "價格接近低點:" + "\n"
            message = message + "價格: " + str(c) + "\n"
            message = message + "低點: " + "{:.2f}".format(lt) + "\n"
            send_line_notify(message)
        
    # if l < 1.01 * bl then send warning
    if l < 1.01 * bl:
        message = id + " " + name + "\n"
        message = message + "價格接近布林下軌:" + "\n"
        message = message + "最低價: " + str(l) + "\n"
        message = message + "布林下軌: " + "{:.2f}".format(bl) + "\n"
        send_line_notify(message)
    
    # if 15 <= k <= 25 then send warning
    if 15. <= k and k <= 25.:
        message = id + " " + name + "\n"
        message = message + "K值接近20:" + "\n"
        message = message + "K值: " + "{:.2f}".format(k) + "\n"
        send_line_notify(message)
    
    # if 75 <= k <= 85 then send warning
    if 75. <= k and k <= 85.:
        message = id + " " + name + "\n"
        message = message + "K值接近80:" + "\n"
        message = message + "K值: " + "{:.2f}".format(k) + "\n"
        send_line_notify(message)



#
# Initialize
#

# Launch XQLite
program_path = 'C:\SysJust\XQLite\daqxqlite.exe'
XQLite = subprocess.Popen(program_path)

# Wait for XQLite to open
time.sleep(60)

# Create a DDE server
server = dde.CreateServer()

# Give the server a name
server.Create("TestClient") 

# Start a conversation with the server
conversation = dde.CreateConversation(server)

# Connect to the DDE server
conversation.ConnectTo("XQLITE", "Kline")



#
# Stock ID, 0056 元大高股息
#
id = "0056"
name = "元大高股息"

# High threshold compared by Close price. set as ZERO to disable.
ht = 40.
# Low threshold compared by Close price. set as ZERO to disable.
lt = 36.

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2408 南亞科
#
id = "2408"
name = "南亞科"

# High threshold compared by Close price. set as ZERO to disable.
ht = 70.
# Low threshold compared by Close price. set as ZERO to disable.
lt = 64.

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2344 華邦電
#
id = "2344"
name = "華邦電"

# High threshold compared by Close price. set as ZERO to disable.
ht = 28.
# Low threshold compared by Close price. set as ZERO to disable.
lt = 25.5

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 1101 台泥
#
id = "1101"
name = "台泥"

# High threshold
ht = 0
# Low threshold
lt = 32.5

CheckToday(id, name, conversation, ht, lt)

#
# Stock ID, 1102 亞泥
#
id = "1102"
name = "亞泥"

# High threshold
ht = 43.5
# Low threshold
lt = 40

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2886 兆豐金
#
id = "2886"
name = "兆豐金"

# High threshold
ht = 39
# Low threshold
lt = 36

CheckToday(id, name, conversation, ht, lt)

#
# Stock ID, 2890 永豐金
#
id = "2890"
name = "永豐金"

# High threshold
ht = 0
# Low threshold
lt = 17

CheckToday(id, name, conversation, ht, lt)

#
# Stock ID, 2891 中信金
#
id = "2891"
name = "中信金"

# High threshold
ht = 0
# Low threshold
lt = 28

CheckToday(id, name, conversation, ht, lt)

#
# Stock ID, 2812 台中銀
#
id = "2812"
name = "台中銀"

# High threshold
ht = 0
# Low threshold
lt = 16

CheckToday(id, name, conversation, ht, lt)

#
# Stock ID, 2892 第一金
#
id = "2892"
name = "第一金"

# High threshold
ht = 27.5
# Low threshold
lt = 26.5

CheckToday(id, name, conversation, ht, lt)

#
# Stock ID, 2884 玉山金
#
id = "2884"
name = "玉山金"

# High threshold
ht = 0
# Low threshold
lt = 24.5

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 8271 宇瞻
#
id = "8271"
name = "宇瞻"

# High threshold
ht = 0
# Low threshold
lt = 60

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 00642U 期元大S&P石油
#
id = "00642U"
name = "期元大S&P石油"

# High threshold
ht = 19
# Low threshold
lt = 16.2

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 9921 巨大
#
id = "9921"
name = "巨大"

# High threshold
ht = 0
# Low threshold
lt = 170

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2301 光寶科
#
id = "2301"
name = "光寶科"

# High threshold
ht = 0
# Low threshold
lt = 110

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 3702 大聯大
#
id = "3702"
name = "大聯大"

# High threshold
ht = 0
# Low threshold
lt = 50

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2385 群光
#
id = "2385"
name = "群光"

# High threshold
ht = 0
# Low threshold
lt = 170

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 1605 華新
#
id = "1605"
name = "華新"

# High threshold
ht = 38.5
# Low threshold
lt = 35.5

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 1513 中興電
#
id = "1513"
name = "中興電"

# High threshold
ht = 200
# Low threshold
lt = 170

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2409 友達
#
id = "2409"
name = "友達"

# High threshold
ht = 18
# Low threshold
lt = 16

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 3481 群創
#
id = "3481"
name = "群創"

# High threshold
ht = 15.5
# Low threshold
lt = 12.5

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2303 聯電
#
id = "2303"
name = "聯電"

# High threshold
ht = 0
# Low threshold
lt = 49

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 2454 聯發科
#
id = "2454"
name = "聯發科"

# High threshold
ht = 0
# Low threshold
lt = 49

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 1303 南亞
#
id = "1303"
name = "南亞"

# High threshold
ht = 0
# Low threshold
lt = 49

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 1301 台塑
#
id = "1301"
name = "台塑"

# High threshold
ht = 0
# Low threshold
lt = 49

CheckToday(id, name, conversation, ht, lt)



#
# Stock ID, 1326 台化
#
id = "1326"
name = "台化"

# High threshold
ht = 0
# Low threshold
lt = 49

CheckToday(id, name, conversation, ht, lt)



#
# Finish
#

# Close XQLite
XQLite.terminate()

# Wait for XQLite to close
time.sleep(10)