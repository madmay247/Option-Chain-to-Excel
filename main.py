import requests
import time
import pandas as pd
import xlwings as xw


def export_option_chain_to_xl(option_chain):
    
        # Open the existing workbook
        wb = xw.Book('option_chain_data.xlsx')

        # Select the active sheet
        sheet = wb.sheets.active

        sheet.clear_contents()
        
        # Write the dataframe to the active sheet, starting at cell B2
        sheet.range('B2').options(index=False).value = option_chain

        # Set the color of the header row to light grey
        sheet.range('B2').expand('right').color = (217, 217, 217)

        # Merge cells for "Calls" and "Puts" labels
        sheet.range('B1:E1').api.Merge()
        sheet.range('G1:J1').api.Merge()

        # Write labels
        sheet.range('B1').value = 'Calls'
        sheet.range('G1').value = 'Puts'

        # Color the labels
        sheet.range('B1:E1').color = (255, 199, 206)  # Light red
        sheet.range('G1:J1').color = (198, 239, 206)  # Light green
        sheet.range('K3:K84').color = (255, 223, 153) #Light Yellow
        sheet.range('F3:F84').color = (255, 223, 153) #Light Yellow

    
        # Auto adjust the column widths1
        sheet.autofit('c')

        # Center align all cells vertically and horizontally
        sheet.cells.api.HorizontalAlignment = -4108
        sheet.cells.api.VerticalAlignment = -4108
        
        # ...
        
        for i in range(3, len(option_chain) + 3):
            # Change cell color based on 'CE_OI_Change' value
            ce_oi_change_cell = sheet.range(f'D{i}')
            ce_oi_change_value = ce_oi_change_cell.value
            if ce_oi_change_value is not None:
                ce_oi_change_value = float(ce_oi_change_value)
                if ce_oi_change_value < 0:
                    ce_oi_change_cell.color = (255, 199, 206)  # Light red (same as "Calls" header)
                elif ce_oi_change_value > 0:
                    ce_oi_change_cell.color = (198, 239, 206)  # Light green (same as "Puts" header)

            # Change cell color based on 'PE_OI_Change' value
            pe_oi_change_cell = sheet.range(f'H{i}')
            pe_oi_change_value = pe_oi_change_cell.value
            if pe_oi_change_value is not None:
                pe_oi_change_value = float(pe_oi_change_value)
                if pe_oi_change_value < 0:
                    pe_oi_change_cell.color = (255, 199, 206)  # Light red (same as "Calls" header)
                elif pe_oi_change_value > 0:
                    pe_oi_change_cell.color = (198, 239, 206)  # Light green (same as "Puts" header)

        # ...
        
        # Get the range of cells containing values
        value_range = sheet.range(f'B2:{chr(65 + option_chain.shape[1])}{len(option_chain) + 2}')

        # Add borders to the value range
        value_range.api.Borders.LineStyle = 1  # Add border
        
        # Save the workbook
        wb.save()

    
def fetch_option_chain(symbol, xl=False):
    symbols_mapping = {
        1: 'NIFTY',
        2: 'BANKNIFTY',
        3: 'FINNIFTY',
        4: 'MIDCPNIFTY'
    }
    
    if symbol not in symbols_mapping.values():
        print("Invalid symbol. Exiting...Sir")
        return None
    
    url = f'https://www.nseindia.com/api/option-chain-indices?symbol={symbol}'
<<<<<<< HEAD
    
    headers = { 'Connection': 'keep-alive',
               'Cache-Control': 'max-age=0',
               'DNT': '1', 'Upgrade-Insecure-Requests': '1',
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36',
               'Sec-Fetch-User': '?1',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
               'Sec-Fetch-Site': 'none',
               'Sec-Fetch-Mode': 'navigate',
               'Accept-Encoding': 'gzip, deflate, br',
               'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8'}
=======
    headers = {'User-Agent': 'Mozilla/5.0'}
>>>>>>> a33b3a0e5928cf7834ff48618cc44bee76552342
    
    # Retry mechanism for request
    max_retries = 3
    retries = 0
    success = False

    while retries < max_retries and not success:
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()  # Raise exception if response status is not successful
            data = response.json()
            records = data['records']['data']
            df = pd.json_normalize(records)
            success = True
        except (requests.RequestException, requests.HTTPError) as e:
            retries += 1
            print(f"Request failed. Retrying... (Attempt {retries}/{max_retries})")
            time.sleep(4)  # Delay for 1 second before the next retry

    if not success:
        print("Request failed after maximum retries. Exiting...Sir")
        return None

  
    # Convert expiryDate to datetime format
    df['expiryDate'] = pd.to_datetime(df['expiryDate'])
    # Filter rows with the closest expiry date
    closest_expiry = df['expiryDate'].min()
    df_closest_expiry = df[df['expiryDate'] == closest_expiry]

    # Create a new dataframe with the desired columns
    new_columns = ['CE_Volume', 'CE_OI', 'CE_OI_Change', 'CE_LTP', 'Strike', 'PE_LTP', 'PE_OI_Change', 'PE_OI', 'PE_Volume', 'Spot_Price']
    option_chain = pd.DataFrame(columns=new_columns)

    # Assign values from the original dataframe to the new dataframe
    option_chain['CE_Volume'] = df_closest_expiry['CE.totalTradedVolume']
    option_chain['CE_OI'] = df_closest_expiry['CE.openInterest']
    option_chain['CE_OI_Change'] = df_closest_expiry['CE.changeinOpenInterest']
    option_chain['CE_LTP'] = df_closest_expiry['CE.lastPrice']
    option_chain['Strike'] = df_closest_expiry['strikePrice']
    option_chain['PE_LTP'] = df_closest_expiry['PE.lastPrice']
    option_chain['PE_OI_Change'] = df_closest_expiry['PE.changeinOpenInterest']
    option_chain['PE_OI'] = df_closest_expiry['PE.openInterest']
    option_chain['PE_Volume'] = df_closest_expiry['PE.totalTradedVolume']
    option_chain['Spot_Price'] = df_closest_expiry['PE.underlyingValue']

    # Reset the index to start from 0
    option_chain.reset_index(drop=True, inplace=True)
    
    if xl == True:
        
        export_option_chain_to_xl(option_chain=option_chain)
        
    return option_chain

def main():
    print("Hello Sir. Please select a symbol:")
    print("1. NIFTY")
    print("2. BANKNIFTY")
    print("3. FINNIFTY")
    print("4. MIDCPNIFTY")
    
    user_input = int(input("Please enter the corresponding number, Sir: "))
    
    symbol = None
    if user_input in [1, 2, 3, 4]:
        symbols_mapping = {
            1: 'NIFTY',
            2: 'BANKNIFTY',
            3: 'FINNIFTY',
            4: 'MIDCPNIFTY'
        }
        symbol = symbols_mapping[user_input]
    else:
        print("Invalid input. Exiting...")
        return

    while True:
        option_chain = fetch_option_chain(symbol, xl=True)
        if option_chain is not None:
            current_time = time.strftime("%Y-%m-%d %H:%M:%S")
            print(f"Option chain for {symbol} fetched and exported to Excel at {current_time}. Refresh in 30 seconds, Sir")
        time.sleep(30)  # Wait for 30 seconds before the next iteration

if __name__ == "__main__":
    main()