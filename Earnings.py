from datetime import datetime, timedelta
import pandas as pd
from xbbg import blp

def get_sp500_tickers() -> list:
    """
    Fetches S&P 500 index members, handling variations in column headers returned by xbbg.
    """
    try:
        print("Pulling S&P 500 constituents from Bloomberg...")
        df_members = blp.bds('SPX Index', 'INDX_MEMBERS')
        
        # Take the first column returned regardless of exact column header name
        raw_tickers = df_members.iloc[:, 0].dropna().tolist()
        
        # Ensure ' Equity' suffix is attached
        formatted_tickers = [
            f"{t} Equity" if not str(t).endswith("Equity") else str(t) 
            for t in raw_tickers
        ]
        print(f"Retrieved {len(formatted_tickers)} constituents from S&P 500.")
        return formatted_tickers
    except Exception as e:
        print(f"Failed to fetch index members dynamically. Error: {e}")
        print("Falling back to sample ticker universe...")
        return [
            'AAPL US Equity', 'MSFT US Equity', 'NVDA US Equity', 
            'AMZN US Equity', 'GOOGL US Equity', 'META US Equity', 
            'TSLA US Equity', 'JPM US Equity', 'V US Equity'
        ]


def get_upcoming_earnings(tickers: list, days_ahead: int = 7) -> pd.DataFrame:
    """
    Pulls upcoming earnings announcement dates and estimates from Bloomberg 
    for a list of equity tickers within the specified day window.
    """
    today = datetime.today().date()
    end_date = today + timedelta(days=days_ahead)
    
    print(f"Fetching earnings scheduled between {today} and {end_date}...")

    # Relevant Bloomberg reference fields
    fields = [
        'SECURITY_NAME',
        'EARN_ANN_DT',
        'EARN_ANN_TM',
        'IS_EPS',
        'EST_EPS'
    ]

    # Pull data using Bloomberg BDP
    raw_res = blp.bdp(tickers=tickers, flds=fields)

    # Coerce response into standard pandas DataFrame safely
    if raw_res is None:
        print("No response from Bloomberg API.")
        return pd.DataFrame()
        
    df = pd.DataFrame(raw_res)

    if df.empty:
        print("Empty DataFrame returned from Bloomberg.")
        return pd.DataFrame()

    # Reset index to make Ticker a proper column
    df = df.reset_index().rename(columns={'index': 'Ticker'})
    df.columns = [str(c).lower() for c in df.columns]

    # Filter by Earnings Date if column exists
    if 'earn_ann_dt' in df.columns:
        # Convert to datetime and strip time
        df['earn_ann_dt'] = pd.to_datetime(df['earn_ann_dt'], errors='coerce').dt.date
        
        # Filter for the 7-day window
        mask = (df['earn_ann_dt'] >= today) & (df['earn_ann_dt'] <= end_date)
        df_filtered = df[mask].copy()
    else:
        print("Warning: 'EARN_ANN_DT' field missing or unavailable in response.")
        df_filtered = df.copy()

    if df_filtered.empty:
        print(f"No companies found reporting earnings between {today} and {end_date}.")
        return pd.DataFrame()

    # Sort chronologically by date and ticker name
    df_filtered = df_filtered.sort_values(by=['earn_ann_dt', 'ticker']).reset_index(drop=True)

    # Rename columns for presentation
    rename_dict = {
        'ticker': 'Ticker',
        'security_name': 'Company Name',
        'earn_ann_dt': 'Earnings Date',
        'earn_ann_tm': 'Announcement Time',
        'is_eps': 'Trailing EPS',
        'est_eps': 'Consensus EPS Est'
    }
    df_filtered = df_filtered.rename(columns=rename_dict)

    return df_filtered


def export_to_excel(df: pd.DataFrame, output_filename: str = "Upcoming_Earnings.xlsx"):
    """
    Exports the pandas DataFrame into a structured Excel file.
    """
    if df.empty:
        print("No data to export. Skipping Excel file generation.")
        return

    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        # Main detail tab
        df.to_excel(writer, sheet_name='Earnings Next 7 Days', index=False)
        
        # High-level summary count tab
        if 'Earnings Date' in df.columns:
            summary = df.groupby('Earnings Date')['Ticker'].count().reset_index()
            summary.columns = ['Earnings Date', 'Company Count']
            summary.to_excel(writer, sheet_name='Summary', index=False)

    print(f"Successfully saved {len(df)} records to '{output_filename}'.")


if __name__ == "__main__":
    # Get universe (S&P 500 or fallback list)
    tickers = get_sp500_tickers()

    # Pull and filter earnings
    earnings_df = get_upcoming_earnings(tickers=tickers, days_ahead=7)

    # Preview results
    if not earnings_df.empty:
        print("\n--- Upcoming Earnings Preview ---")
        print(earnings_df.head(10))

    # Export to Excel
    export_to_excel(earnings_df, output_filename="Earnings_Next_7_Days.xlsx")