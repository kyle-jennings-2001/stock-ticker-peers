import pandas as pd
import numpy as np
import pycot.reports

# 1. Fetch the Traders in Financial Futures (TFF) report for a given year directly into a DataFrame
# You can also use 'traders_in_financial_futures_futures_only', 'legacy_futures_and_options', or 'disaggregated_futures_and_options'
# df = pycot.reports.CommitmentsOfTraders(report_type="traders_in_financial_futures_fut")
# print(df)

cot = pycot.reports.CommitmentsOfTraders("traders_in_financial_futures_futop")
# contract_names = ("FED FUNDS - CHICAGO BOARD OF TRADE", "30-DAY FEDERAL FUNDS - CHICAGO BOARD OF TRADE")
contract_names = (
    "S&P 500 Consolidated - CHICAGO MERCANTILE EXCHANGE",
    "E-MINI S&P 500 - CHICAGO MERCANTILE EXCHANGE",
    "MICRO E-MINI S&P 500 INDEX - CHICAGO MERCANTILE EXCHANGE",
    "NASDAQ-100 Consolidated - CHICAGO MERCANTILE EXCHANGE",
    "NASDAQ MINI - CHICAGO MERCANTILE EXCHANGE", 
    "MICRO E-MINI NASDAQ-100 INDEX - CHICAGO MERCANTILE EXCHANGE",
    "RUSSELL E-MINI - CHICAGO MERCANTILE EXCHANGE",
    "MICRO E-MINI RUSSELL 2000 INDX - CHICAGO MERCANTILE EXCHANGE",

)
df = cot.report(contract_names)
print(df)

df.to_excel('COTtest.xlsx')