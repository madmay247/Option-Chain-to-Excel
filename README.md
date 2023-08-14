# Bring Entire Option Chain into an Excel file

This script fetches option chain data for specific symbols from the NSE (National Stock Exchange) website and exports it to an Excel file. It provides a user-friendly way to choose the symbol and periodically fetches and updates the data. 

**NSE updates its data on website every 3 to 5 mins, please keep that in mind as you may not see changes on every refresh**

https://github.com/madmay247/Option-Chain-to-Excel/assets/132202725/d8d39d25-1c50-4182-bc10-067427aabf59

## Requirements

- Python 3.x
- Required libraries: `requests`, `time`, `pandas`, `xlwings`

You can install the required libraries using the following command:

`pip install -r requirements.txt`

## Usage

1. Run the script using the command:

   `main.py`

3. Select a symbol from the list (1. NIFTY, 2. BANKNIFTY, 3. FINNIFTY, 4. MIDCPNIFTY).

4. The script will fetch option chain data from the NSE website, create an Excel file named `option_chain_data.xlsx`, and export the data to it.

5. The script will continue to periodically update the data in the Excel file, refreshing every 30 seconds.

## Features

- Fetches option chain data from the NSE website using the `requests` library.
- Organizes the fetched data into a structured `pandas` DataFrame.
- Exports the data to an Excel file using `xlwings`.
- Provides a simple user interface to select the desired symbol.
- Periodically updates the data in the Excel file to keep it current (every 30 seconds).

## Notes

- Ensure that you have a reliable internet connection to fetch data from the NSE website.
- The script handles request retries in case of connectivity issues.
- The Excel file is created and updated in the same directory as the script.
- This script can be customized and extended further to include additional features.

## Disclaimer

This script is provided for educational and informational purposes only. Use it at your own risk and discretion. The author is not responsible for any financial or legal implications resulting from its use.
