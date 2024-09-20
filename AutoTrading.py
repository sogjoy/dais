import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
import time, calendar
import requests

# Slack webhook URL
SLACK_WEBHOOK_URL = "https://hooks.slack.com/services/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"

def post_message(webhook_url, text):
    response = requests.post(
        webhook_url,
        json={"text": text},
        headers={"Content-Type": "application/json"}
    )
    if response.status_code != 200:
        raise ValueError(
            f"Request to Slack returned an error {response.status_code}, the response is:\n{response.text}"
        )

def dbgout(message):
    """Prints the message to the Python shell and sends it to Slack."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    post_message(SLACK_WEBHOOK_URL, strbuf)

def printlog(message, *args):
    """Prints the message to the Python shell."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)

# Creon Plus common OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  

def check_creon_system():
    """Checks Creon Plus system connection status."""
    # Check if process is run with admin privileges
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        return False
 
    # Check connection status
    if (cpStatus.IsConnect == 0):
        printlog('check_creon_system() : connect to server -> FAILED')
        return False
 
    # Initialize trade - only use when account-related code is available
    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True

def get_current_price(code):
    """Returns the current price, ask price, and bid price for the given stock code."""
    cpStock.SetInputValue(0, code)  # Set the stock code to get price info
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # Current price
    item['ask'] =  cpStock.GetHeaderValue(16)        # Ask price
    item['bid'] =  cpStock.GetHeaderValue(17)        # Bid price    
    return item['cur_price'], item['ask'], item['bid']

def get_ohlc(code, qty):
    """Returns OHLC price information for the given stock code and quantity."""
    cpOhlc.SetInputValue(0, code)           # Stock code
    cpOhlc.SetInputValue(1, ord('2'))        # 1: Period, 2: Quantity
    cpOhlc.SetInputValue(4, qty)             # Number of requests
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5]) # 0: Date, 2~5: OHLC
    cpOhlc.SetInputValue(6, ord('D'))        # D: Daily
    cpOhlc.SetInputValue(9, ord('1'))        # 0: Unadjusted price, 1: Adjusted price
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)   # Number of received items
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count): 
        index.append(cpOhlc.GetDataValue(0, i)) 
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
            cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)]) 
    df = pd.DataFrame(rows, columns=columns, index=index) 
    return df

def get_stock_balance(code):
    """Returns the stock name and quantity for the given stock code."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # Account number
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1: All, 1: Stock, 2: Futures/Options
    cpBalance.SetInputValue(0, acc)         # Account number
    cpBalance.SetInputValue(1, accFlag[0])  # Product category - first stock product
    cpBalance.SetInputValue(2, 50)          # Number of requests (maximum 50)
    cpBalance.BlockRequest()     
    if code == 'ALL':
        dbgout('Account name: ' + str(cpBalance.GetHeaderValue(0)))
        dbgout('Settlement balance quantity: ' + str(cpBalance.GetHeaderValue(1)))
        dbgout('Valuation amount: ' + str(cpBalance.GetHeaderValue(3)))
        dbgout('Valuation profit/loss: ' + str(cpBalance.GetHeaderValue(4)))
        dbgout('Number of stocks: ' + str(cpBalance.GetHeaderValue(7)))
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # Stock code
        stock_name = cpBalance.GetDataValue(0, i)   # Stock name
        stock_qty = cpBalance.GetDataValue(15, i)   # Quantity
        if code == 'ALL':
            dbgout(str(i+1) + ' ' + stock_code + '(' + stock_name + ')' 
                + ':' + str(stock_qty))
            stocks.append({'code': stock_code, 'name': stock_name, 
                'qty': stock_qty})
        if stock_code == code:  
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0

def get_current_cash():
    """Returns the amount of cash available for 100% margin orders."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # Account number
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1: All, 1: Stock, 2: Futures/Options
    cpCash.SetInputValue(0, acc)              # Account number
    cpCash.SetInputValue(1, accFlag[0])      # Product category - first stock product
    cpCash.BlockRequest() 
    return cpCash.GetHeaderValue(9) # Cash available for 100% margin orders

def get_target_price(code):
    """Returns the target price for buying."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open 
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]                                      
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]
        target_price = today_open + (lastday_high - lastday_low) * 0.5
        return target_price
    except Exception as ex:
        dbgout("`get_target_price() -> exception! " + str(ex) + "`")
        return None
    
def get_movingaverage(code, window):
    """Returns the moving average price for the given stock code and window."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()         
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        dbgout('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None    

def buy_etf(code):
    """Buys the given stock using the most favorable FOK condition."""
    try:
        global bought_list      # Global to modify within the function
        if code in bought_list: # If the stock is already bought, do not buy again
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code) 
        target_price = get_target_price(code)    # Target price
        ma5_price = get_movingaverage(code, 5)   # 5-day moving average price
        ma10_price = get_movingaverage(code, 10) # 10-day moving average price
        buy_qty = 0        # Initialize buy quantity
        if ask_price > 0:  # If ask price exists   
            buy_qty = buy_amount // ask_price  
        stock_name, stock_qty = get_stock_balance(code)  # Get stock name and quantity
        if current_price > target_price and current_price > ma5_price \
            and current_price > ma10_price:  
            
            printlog(stock_name + '(' + str(code) + ') ' + str(buy_qty) +
                'EA : ' + str(current_price) + ' meets the buy condition!`')            
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]      # Account number
            accFlag = cpTradeUtil.GoodsList(acc, 1) # -1: All, 1: Stock, 2: Futures/Options                
            # Set FOK buy order with the most favorable condition
            cpOrder.SetInputValue(0, "2")        # 2: Buy
            cpOrder.SetInputValue(1, acc)        # Account number
            cpOrder.SetInputValue(2, accFlag[0]) # Product category - first stock product
            cpOrder.SetInputValue(3, code)       # Stock code
            cpOrder.SetInputValue(4, buy_qty)    # Quantity to buy
            cpOrder.SetInputValue(7, "2")        # Order condition 0: Basic, 1: IOC, 2: FOK
            cpOrder.SetInputValue(8, "12")       # Order price type 1: Normal, 3: Market price
                                                 # 5: Conditional, 12: Most favorable, 13: Priority 
            # Send buy order request
            ret = cpOrder.BlockRequest() 
            printlog('Most favorable FoK buy ->', stock_name, code, buy_qty, '->', ret)
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                printlog('Warning: Continuous order restriction. Waiting time:', remain_time/1000)
                time.sleep(remain_time/1000) 
                return False
            time.sleep(2)
            printlog('Cash available for order:', buy_amount)
            stock_name, bought_qty = get_stock_balance(code)
            printlog('get_stock_balance :', stock_name, stock_qty)
            if bought_qty > 0:
                bought_list.append(code)
                dbgout("`buy_etf("+ str(stock_name) + ' : ' + str(code) + 
                    ") -> " + str(bought_qty) + "EA bought!" + "`")
    except Exception as ex:
        dbgout("`buy_etf("+ str(code) + ") -> exception! " + str(ex) + "`")

def sell_all():
    """Sells all owned stocks with the most favorable IOC condition."""
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]       # Account number
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1: All, 1: Stock, 2: Futures/Options   
        while True:    
            stocks = get_stock_balance('ALL') 
            total_qty = 0 
            for s in stocks:
                total_qty += s['qty'] 
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:                  
                    cpOrder.SetInputValue(0, "1")         # 1: Sell, 2: Buy
                    cpOrder.SetInputValue(1, acc)         # Account number
                    cpOrder.SetInputValue(2, accFlag[0])  # First stock product
                    cpOrder.SetInputValue(3, s['code'])   # Stock code
                    cpOrder.SetInputValue(4, s['qty'])    # Quantity to sell
                    cpOrder.SetInputValue(7, "1")   # Condition 0: Basic, 1: IOC, 2: FOK
                    cpOrder.SetInputValue(8, "12")  # Price type 12: Most favorable, 13: Priority 
                    # Send most favorable IOC sell order request
                    ret = cpOrder.BlockRequest()
                    printlog('Most favorable IOC sell', s['code'], s['name'], s['qty'], 
                        '-> cpOrder.BlockRequest() -> returned', ret)
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        printlog('Warning: Continuous order restriction. Waiting time:', remain_time/1000)
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        dbgout("sell_all() -> exception! " + str(ex))

if __name__ == '__main__': 
    try:
        #symbol_list = ['A018310', 'A229000', 'A000880', 'A005930', 'A010140',
        #    'A028260', 'A042660', 'A452260']
        symbol_list = ['A069500', 'A091180', 'A091230', 'A102780', 'A104520', 'A117680', 'A117700', 'A138540', 'A005930']
        #    'A028260', 'A042660', 'A452260', 'A112610', 'A397030', 'A310210'
        #symbol_list = ['A005930']  
        #symbol_list = ['A228790', 'A261110', 'A266410', 'A291680', 'A334700', 'A346000', 'A360140', 'A373530', 'A422260', 'A430500', 'A432850']


        bought_list = []     # List of bought stocks
        target_buy_count = 5 # Number of stocks to buy
        buy_percent = 0.19   
        printlog('check_creon_system() :', check_creon_system())  # Check Creon connection
        stocks = get_stock_balance('ALL')      # Check all owned stocks
        total_cash = int(get_current_cash())   # Check cash available for 100% margin orders
        buy_amount = total_cash * buy_percent  # Calculate order amount per stock
        printlog('100% margin order available cash:', total_cash)
        printlog('Order percentage per stock:', buy_percent)
        printlog('Order amount per stock:', buy_amount)
        printlog('Start time:', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldout = False

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0,microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # Exit on Saturday or Sunday
                printlog('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                sys.exit(0)
            if t_9 < t_now < t_start and soldout == False:
                soldout = True
                sell_all()
            if t_start < t_now < t_sell :  # AM 09:05 ~ PM 03:15 : Buy stocks
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy_etf(sym)
                        time.sleep(1)
                if t_now.minute == 30 and 0 <= t_now.second <= 5: 
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : Sell all stocks
                if sell_all() == True:
                    dbgout('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ : Exit the program
                dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')
