import blpapi
import pandas as pd

TICKER = "LRCX"
COUNTRY_CODE = "US"
FREQUENCY = "A" # A=Annual, Q=Quarter, FQ=Fiscal Quarter
SEGMENTATION_TYPE = "G" # G=Geographic, P=Product

TICKER_COMBINATION = f"{TICKER} {COUNTRY_CODE} Equity"


def get_geographic_revenue(ticker, frequency, segmentation_type):
    # Initialize session connection; Desktop API default: localhost=8194
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

    # Bulk data field for given ticker
    request.getElement("securities").appendValue(ticker)
    request.getElement("fields").appendValue("PG_REVENUE")


    # Add overrides for Geographic breakdown and periodicity
    overrides = request.getElement("overrides")
    # Set Geographic mode ('G' for Geographic, 'P' for Product)
    ovr_geo = overrides.appendElement()
    ovr_geo.setElement("fieldId", "PRODUCT_GEO_OVERRIDE")
    ovr_geo.setElement("value", segmentation_type)

    # Set Periodicity ('A', 'Q', or 'FQ')
    ovr_per = overrides.appendElement()
    ovr_per.setElement("fieldId", "FUND_PER")
    ovr_per.setElement("value", frequency)

    # Send asynchronous request
    session.sendRequest(request)

    data = []

    # Parse incoming event response stream
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
                        #### How to make adding more fields more flexible for the code? 
                        
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

df1 = get_geographic_revenue(ticker=TICKER_COMBINATION, frequency=FREQUENCY, segmentation_type=SEGMENTATION_TYPE)


if not df1.empty:
    output_file = f"{TICKER}_geographic_revenue.xlsx"
    df1.to_excel(output_file, index=False, engine="openpyxl")
    print(f"Data successfully saved to {output_file}")
    print(df1.head())
else:
    print("No geographic revenue data returned.")