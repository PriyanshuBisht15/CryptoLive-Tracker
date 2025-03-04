# Project Description
This project fetches real-time cryptocurrency data for the top 50 cryptocurrencies by market capitalization using the CoinGecko API. The fetched data is automatically updated in an Excel sheet every 5 minutes. Additionally, the program performs basic analysis, including:

- Identifying the top 5 cryptocurrencies by market capitalization.
- Calculating the average price of the top 50 cryptocurrencies.
- Determining the highest and lowest percentage price changes over 24 hours.
# Features
## ✅ Fetches live cryptocurrency data from CoinGecko API
## ✅ Updates an Excel file (Crypto.xlsx) every 5 minutes
## ✅ Stores key metrics:
   - Cryptocurrency Name
   - Symbol
   - Current Price (USD)
   - Market Capitalization
   - 24-hour Trading Volume
   - Price Change (24h, %)
   
 ## ✅ Clears old data before adding new updates
 ## ✅ Performs basic analysis on cryptocurrency trends

## Installation & Setup
- Prerequisites
Ensure you have Python installed (version 3.7 or later).

- Required Libraries
Install the necessary dependencies using pip:
pip install pycoingecko openpyxl

## How to Run
1. Clone or download the project.
2. Open a terminal in the project directory.
3. Run the script using:
   python crypto.py

The script will continuously update the Excel file (Crypto.xlsx) every 5 minutes.

## Output : 
- Before:
![image](https://github.com/user-attachments/assets/f964bf27-38bb-447e-b383-5b0b6b87c76c)


- After:
![image](https://github.com/user-attachments/assets/6e91d7e0-5c1d-4663-8918-15fb406d2de2)


