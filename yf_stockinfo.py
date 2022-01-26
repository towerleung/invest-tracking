import yfinance as yf
from apiclient import discovery
from google.oauth2 import service_account
from datetime import datetime

def get_yf_stockinfo(stock_symbol):    
    out_dict={}
    yf_Ticker = yf.Ticker(stock_symbol)
    yf_stock = yf_Ticker.info
    if stock_symbol=="2840.HK": #for some reason, YF api doesn't have currentPrice value for this stock
      out_dict["Price"] = yf_stock["dayLow"]  
    else:
      out_dict["Price"] = yf_stock["currentPrice"]
    out_dict["Dividend"] = yf_stock["dividendYield"]
    return out_dict

def main():
        #initialize googlesheet api and connect to asset worksheet
        scopes = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        #secret_file = os.path.join(os.getcwd(), '.credentials/gd_client.json')
        #secret_file = "/home/tower/stock_dividend/web-spider-invest/.credentials/gd_client.json"
        secret_file = "/home/tower/stock_dividend/.credentials/gd_client.json"
        credentials = service_account.Credentials.from_service_account_file(secret_file,scopes=scopes)
        service = discovery.build('sheets', 'v4',credentials=credentials)
        spreadsheet_id = '1DgZPlZDt3lvsJHgptWGkdX-9PG334LkEqDLtb0iU54E'

        #capture begin time
        print(datetime.now()) 

        #Get stock count
        Stock_Count_Cell = 'Stock Pivot!AZ1'
                
        #stock symbols and dividend% ranges
        No_of_Stock_Cell = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,range=Stock_Count_Cell).execute() 
        No_of_Stock = str(int(No_of_Stock_Cell['values'][0][0]))
        stock_symbol_range = 'Stock Pivot!A2:A' + No_of_Stock
        stock_yield_range = 'Stock Pivot!G2:G' + No_of_Stock
        stock_price_range = 'Stock Pivot!AG2:AG' + No_of_Stock

        
        #get list of stock symbols
        stock_list = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,range=stock_symbol_range).execute()    

        stock_info = []
        price_list =[]
        dividend_list=[]

        for stock_symbol in stock_list['values']:
            #convert stock symbol 
            stock_symbol_text = stock_symbol[0].split(':',4)[1] + ".HK"
            stock_info = get_yf_stockinfo(stock_symbol_text)
            price_list.append([stock_info['Price']])
            dividend_list.append([stock_info['Dividend']])
       
        price_data = {
            'values' : price_list 
        }
        dividend_data = {
            'values' : dividend_list 
        }
        #update the dividend data back to googlesheet
        service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, body=price_data,range=stock_price_range, valueInputOption='USER_ENTERED').execute()
        service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, body=dividend_data,range=stock_yield_range, valueInputOption='USER_ENTERED').execute()
        #capture end time
        print(datetime.now())
    

if __name__ == "__main__":
    main()
