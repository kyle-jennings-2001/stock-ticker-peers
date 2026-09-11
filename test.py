import blpapi
import pandas as pd
from datetime import datetime

def parse_bloomberg_response(response):
    """
    Parses the historical data response from the Bloomberg API into a list of dictionaries.
    """
    data_rows = []
    
    for msg in response:
        # Check if the message contains historical data
        if msg.hasElement("securityData"):
            sec_data = msg.getElement("securityData")
            field_data_array = sec_data.getElement("fieldData")
            
            # Iterate through each historical date entry
            for i in range(field_data_array.numValues()):
                field_data = field_data_array.getValueAsElement(i)
                date_str = field_data.getElementAsString("date")
                
                row = {"Date": date_str}
                
                # Extract all requested fields that returned data for this date
                for j in range(field_data.numElements()):
                    element = field_data.getElement(j)
                    name = str(element.name())
                    if name != "date":
                        # Fetch value safely based on data type
                        row[name] = element.getValue()
                        
                data_rows.append(row)
                
    return data_rows

def get_bloomberg_financials():
    # 1. Initialize Bloomberg Session Options
    session_options = blpapi.SessionOptions()
    session_options.setServerHost("localhost")
    session_options.setServerPort(8194)  # Default Bloomberg API port
    
    session = blpapi.Session(session_options)
    
    if not session.start():
        print("Failed to start Bloomberg session. Ensure your Terminal is open and logged in.")
        return None
        
    if not session.openService("//blp/refdata"):
        print("Failed to open remote service //blp/refdata")
        return None
        
    ref_data_service = session.getService("//blp/refdata")
    request = ref_data_service.createRequest("HistoricalDataRequest")
    
    # 2. Define Ticker and Fields
    # Adding REGN US Equity. (Quarterly data is requested via overrides below)
    request.getElement("securities").appendValue("REGN US Equity")
    
    # Core Financial Statement Field Mappings (Examples of standard GAAP items)
    # You can add or modify these fields based on the specific lines you need
    fields = [
        "SALES_REV_TURN",          # Total Revenue
        "IS_COMP_SALES",
        "IS_COGS_TO_FE_AND_PP_AND_G",
        "GROSS_PROFIT",              # Gross Profit
        "IS_OPER_INC",               # Operating Income (EBIT)
        "NET_INCOME",                # Net Income
        "BS_TOT_ASSET",              # Total Assets
        "BS_TOT_LIAB",               # Total Liabilities
        "TOTAL_EQUITY",              # Total Equity
        "CF_CASH_FROM_OPER",         # Cash from Operating Activities
        "CF_CASH_FROM_INV_ACT",      # Cash from Investing Activities
        "CF_CASH_FROM_FIN_ACT",      # Cash from Financing Activities
    ]
    
    for field in fields:
        request.getElement("fields").appendValue(field)
        
    # 3. Define Parameters & Overrides for Historical Quarterly Data
    request.set("startDate", "19000101")  # Set far back to capture early history (REGN IPO was 1991)
    request.set("endDate", datetime.today().strftime('%Y%m%d'))
    request.set("periodicitySelection", "QUARTERLY") 
    
    # Optional: Force the currency to USD and adjust for corporate actions if preferred
    request.set("currency", "USD")
    
    print("Sending request to Bloomberg Terminal...")
    session.sendRequest(request)
    
    # 4. Process the incoming data stream
    responses = []
    while True:
        event = session.nextEvent()
        for msg in event:
            if event.eventType() in [blpapi.Event.RESPONSE, blpapi.Event.PARTIAL_RESPONSE]:
                responses.append(msg)
        if event.eventType() == blpapi.Event.RESPONSE:
            break
            
    session.stop()
    
    # 5. Build DataFrame
    print("Processing financial data...")
    raw_data = parse_bloomberg_response(responses)
    
    if not raw_data:
        print("No data retrieved.")
        return None
        
    df = pd.DataFrame(raw_data)
    
    # Clean up DataFrame structure
    df["Date"] = pd.to_datetime(df["Date"])
    df = df.sort_values(by="Date", ascending=False).reset_index(drop=True)
    
    return df

if __name__ == "__main__":
    df_financials = get_bloomberg_financials()
    
    if df_financials is not None:
        output_file = "REGN_Quarterly_Financials.xlsx"
        # Export to Excel
        df_financials.to_excel(output_file, index=False, sheet_name="REGN_Financials")
        print(f"Success! Data exported to {output_file}")
        print(df_financials.head())