import MetaTrader5 as mt5  # Import MetaTrader5 library
import time
import pandas as pd

# Define the path to the Excel file containing orders
OrderFilePathAndName = "C:/Users/Shahab Kabiri/PycharmProjects/HomeworkAI/Projects/Forex Vensim And Much More/Orders.xlsx"

# Function to read orders from the Excel file
def ReadingOrdersExcelFile(OrderFilePathAndName):
    Orders = pd.read_excel(OrderFilePathAndName)
    
    # Keep rows where the USED column is TRUE and select the last 100 rows
    Orders = Orders.loc[Orders.USED, :]
    Orders = Orders.tail(100)
    Orders.reset_index(drop=True, inplace=True)
    return Orders

# Function to execute limit orders
def ExecutingLimitOrders():
    Orders = ReadingOrdersExcelFile(OrderFilePathAndName)
    
    for i in range(Orders.shape[0]):
        NumberOfIntervals = int(Orders.loc[i, 'NUMBER_OF_INTERVALS'])
        OrderType = Orders.loc[i, 'ORDER_TYPE']
        Symbol = Orders.loc[i, 'SYMBOL']
        point = mt5.symbol_info(Symbol).point
        CrashedVolume = Orders.loc[i, 'VOLUME'] / Orders.loc[i, 'NUMBER_OF_INTERVALS']
        
        if OrderType == 'ORDER_TYPE_BUY_LIMIT':
            FirstOrderPrice = Orders.loc[i, 'UPPERBOUND_PRICE']
            PriceLevels = round((Orders.loc[i, 'UPPERBOUND_PRICE'] - Orders.loc[i, 'LOWERBOUND_PRICE']) / Orders.loc[i, 'NUMBER_OF_INTERVALS'], 7)
            
            for j in range(NumberOfIntervals):
                BUYrequest = {
                    "action": mt5.TRADE_ACTION_PENDING,
                    "symbol": Symbol,
                    "volume": CrashedVolume,
                    "type": mt5.ORDER_TYPE_BUY_LIMIT,
                    "price": FirstOrderPrice - (j - 1) * PriceLevels,
                    "sl": FirstOrderPrice - (j - 1) * PriceLevels - Orders.loc[i, 'STOPLOSS_IN_POINTS'] * point,
                    "tp": FirstOrderPrice - (j - 1) * PriceLevels + Orders.loc[i, 'TAKEPROFIT_IN_POINTS'] * point,
                    "deviation": 0,
                    "comment": str(int(Orders.loc[i, 'LEVEL_UNIQUE_ID'])) + " " + str(j),
                    "type_time": mt5.ORDER_TIME_GTC,
                }
                resultBUY = mt5.order_send(BUYrequest)
                print(resultBUY)
        
        if OrderType == 'ORDER_TYPE_SELL_LIMIT':
            FirstOrderPrice = Orders.loc[i, 'LOWERBOUND_PRICE']
            PriceLevels = round((Orders.loc[i, 'UPPERBOUND_PRICE'] - Orders.loc[i, 'LOWERBOUND_PRICE']) / Orders.loc[i, 'NUMBER_OF_INTERVALS'], 7)
            
            for j in range(NumberOfIntervals):
                SELLrequest = {
                    "action": mt5.TRADE_ACTION_PENDING,
                    "symbol": Symbol,
                    "volume": CrashedVolume,
                    "type": mt5.ORDER_TYPE_SELL_LIMIT,
                    "price": FirstOrderPrice + (j - 1) * PriceLevels,
                    "sl": FirstOrderPrice + (j - 1) * PriceLevels + Orders.loc[i, 'STOPLOSS_IN_POINTS'] * point,
                    "tp": FirstOrderPrice + (j - 1) * PriceLevels - Orders.loc[i, 'TAKEPROFIT_IN_POINTS'] * point,
                    "deviation": 0,
                    "comment": str(int(Orders.loc[i, 'LEVEL_UNIQUE_ID'])) + " " + str(j),
                    "type_time": mt5.ORDER_TIME_GTC,
                }
                resultSELL = mt5.order_send(SELLrequest)
                print(resultSELL)

# Execute limit orders
ExecutingLimitOrders()
