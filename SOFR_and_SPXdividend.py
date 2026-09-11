import pandas as pd
import requests
import blpapi

# ------------------------------------------------------------------------------
# 1. Fetch Historical Dividend Yield from Bloomberg Desktop API (blpapi)
# ------------------------------------------------------------------------------
def get_bloomberg_div_yield(ticker="SPX Index", field="EQY_DVD_YLD_12M", start_date="19000101"):
    """
    Connects to the local Bloomberg Terminal API session and requests 
    historical daily dividend yield for the specified ticker.
    """
    # Set up Bloomberg session options
    session_options = blpapi.SessionOptions()
    session_options.setServerHost("localhost")
    session_options.setServerPort(8194)
    
    session = blpapi.Session(session_options)
    if not session.start():
        raise ConnectionError("Failed to start Bloomberg API session. Ensure Bloomberg Terminal is running.")
    
    try:
        if not session.openService("//blp/refdata"):
            raise ConnectionError("Failed to open //blp/refdata service.")
            
        refDataService = session.getService("//blp/refdata")
        request = refDataService.createRequest("HistoricalDataRequest")
        
        request.getElement("securities").appendValue(ticker)
        request.getElement("fields").appendValue(field)
        request.set("startDate", start_date)
        request.set("endDate", pd.Timestamp.now().strftime("%Y%m%d"))
        request.set("periodicitySelection", "DAILY")
        
        session.sendRequest(request)
        
        dates = []
        values = []
        
        while True:
            event = session.nextEvent(5000)
            for msg in event:
                if msg.hasElement("securityData"):
                    security_data = msg.getElement("securityData")
                    field_data_array = security_data.getElement("fieldData")
                    
                    for i in range(field_data_array.numValues()):
                        field_data = field_data_array.getValue(i)
                        date_str = field_data.getElementAsString("date")
                        if field_data.hasElement(field):
                            val = field_data.getElementAsFloat(field)
                            dates.append(pd.to_datetime(date_str))
                            values.append(val)
                            
            if event.eventType() == blpapi.Event.RESPONSE:
                break
                
        df = pd.DataFrame({"Date": dates, "SP500_Div_Yield": values})
        df.set_index("Date", inplace=True)
        return df

    finally:
        session.stop()


# ------------------------------------------------------------------------------
# 2. Fetch Historical SOFR Data from New York Fed API
# ------------------------------------------------------------------------------
def get_ny_fed_sofr():
    """
    Fetches the full daily historical SOFR series directly from 
    the Federal Reserve Bank of New York API.
    """
    # The NY Fed endpoint provides full rate search history
    url = "https://markets.newyorkfed.org/api/rates/all/search.json?rateType=SOFR"
    
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()
    
    rates = data.get("refRates", [])
    df = pd.DataFrame(rates)
    
    # Process date and rate columns
    df["Date"] = pd.to_datetime(df["effectiveDate"])
    df["SOFR_Rate"] = pd.to_numeric(df["percentRate"], errors="coerce")
    
    df = df[["Date", "SOFR_Rate"]].dropna().sort_values("Date").set_index("Date")
    return df


# ------------------------------------------------------------------------------
# 3. Combine Data and Display Output
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    print("Fetching Bloomberg S&P 500 Dividend Yield...")
    sp500_df = get_bloomberg_div_yield(
        ticker="SPX Index", 
        field="EQY_DVD_YLD_12M",  # Daily 12-month trailing dividend yield
        start_date="19000101"     # Pulls as far back as Bloomberg dataset allows
    )
    
    print("Fetching NY Fed SOFR history...")
    sofr_df = get_ny_fed_sofr()
    
    # Merge on Date (Outer join retains maximum historical extent for both)
    combined_df = pd.merge(sp500_df, sofr_df, left_index=True, right_index=True, how="outer")
    combined_df.sort_index(inplace=True)
    
    print("\nCombined DataFrame Preview:")
    print("=" * 45)
    print(combined_df.tail(20))
    print("\nDataset Info:")
    print(f"Start Date: {combined_df.index.min().strftime('%Y-%m-%d')}")
    print(f"End Date:   {combined_df.index.max().strftime('%Y-%m-%d')}")
    print(f"Total Rows: {len(combined_df)}")