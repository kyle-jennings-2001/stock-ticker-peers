import blpapi
import pandas as pd

TICKER = "LRCX US Equity"

def get_geographic_revenue(ticker: str, period: str = "A") -> pd.DataFrame:
    """
    Queries Bloomberg API for geographic revenue breakdown of a security.
    
    Parameters:
        # ticker (str): Bloomberg ticker (e.g., 'AAPL US Equity').
        period (str): 'A' for Annual, 'Q' for Quarterly, 'FQ' for Fiscal Quarter.
    """
    # 1. Initialize session connection (Desktop API default: localhost:8194)
    session_options = blpapi.SessionOptions()
    session_options.setServerHost("localhost")
    session_options.setServerPort(8194)

    session = blpapi.Session(session_options)
    if not session.start():
        raise ConnectionError("Failed to start Bloomberg API session.")

    if not session.openService("//blp/refdata"):
        session.stop()
        raise ConnectionError("Failed to open Bloomberg refdata service.")

    refDataService = session.getService("//blp/refdata")
    request = refDataService.createRequest("ReferenceDataRequest")

    # 2. Add ticker and bulk data field
    request.getElement("securities").appendValue(ticker)
    request.getElement("fields").appendValue("PG_REVENUE")

    # 3. Add overrides for Geographic breakdown and periodicity
    overrides = request.getElement("overrides")
    
    # Set Geographic mode ('G' for Geographic, 'P' for Product)
    ovr_geo = overrides.appendElement()
    ovr_geo.setElement("fieldId", "PRODUCT_GEO_OVERRIDE")
    ovr_geo.setElement("value", "G")

    # Set Periodicity ('A', 'Q', or 'FQ')
    ovr_per = overrides.appendElement()
    ovr_per.setElement("fieldId", "FUND_PER")
    ovr_per.setElement("value", period)

    # Send asynchronous request
    session.sendRequest(request)

    data = []

    # 4. Parse incoming event response stream
    while True:
        event = session.nextEvent(5000)
        
        for msg in event:
            if msg.hasElement("securityData"):
                sec_data_array = msg.getElement("securityData")
                for i in range(sec_data_array.numValues()):
                    sec_data = sec_data_array.getValueAsElement(i)
                    field_data = sec_data.getElement("fieldData")
                    
                    if field_data.hasElement("PG_REVENUE"):
                        pg_rev = field_data.getElement("PG_REVENUE")
                        
                        # Loop through returned array elements
                        for j in range(pg_rev.numValues()):
                            row_elem = pg_rev.getValueAsElement(j)
                            row_dict = {}
                            for k in range(row_elem.numElements()):
                                elem = row_elem.getElement(k)
                                row_dict[str(elem.name())] = elem.getValue()
                            data.append(row_dict)

        # Stop listening when final RESPONSE event is received
        if event.eventType() == blpapi.Event.RESPONSE:
            break

    session.stop()

    # Convert to pandas DataFrame
    df = pd.DataFrame(data)
    return df


if __name__ == "__main__":
    ticker = TICKER
    output_file = f"{TICKER}_geographic_revenue.xlsx"

    print(f"Fetching geographic revenue data for {ticker}...")
    df = get_geographic_revenue(ticker, period="A")

    if not df.empty:
        # Export DataFrame to Excel
        df.to_excel(output_file, index=False, engine="openpyxl")
        print(f"Data successfully saved to {output_file}")
        print(df.head())
    else:
        print("No geographic revenue data returned.")