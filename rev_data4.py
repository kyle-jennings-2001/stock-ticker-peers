import re
from contextlib import contextmanager

import blpapi
import pandas as pd

TICKER = "LRCX"
COUNTRY_CODE = "US"
FREQUENCY = "A" # A=Annual, Q=Quarter, FQ=Fiscal Quarter
SEGMENTATION_TYPE = "G" # G=Geographic, P=Product

START_YEAR = 1900 # oldest fiscal year to request
END_YEAR = 2026 # newest fiscal year to request (anchor of the first window)

TICKER_COMBINATION = f"{TICKER} {COUNTRY_CODE} Equity"

# PG_REVENUE always returns a fixed 5-period window anchored at EQY_FUND_YEAR.
PERIODS_PER_WINDOW = 5
# Period 1 is the anchor year and periods count backwards from there. Verify this
# once for your dataset (see the overlap check printed at the end) and flip if needed.
PERIODS_NEWEST_FIRST = True
# Step 4 rather than 5 so consecutive windows overlap by one year, which lets us
# confirm the period-to-year mapping is correct.
ANCHOR_STEP = 4
# Zero-revenue segments come back as float noise (e.g. -2.42454E-14); snap them to 0
ZERO_THRESHOLD = 1e-4
# Stop walking backwards after this many consecutive empty windows, so a very old
# START_YEAR doesn't keep requesting years that predate the company's history
MAX_EMPTY_WINDOWS = 2

PERIOD_RE = re.compile(r"Period\s*(\d+)\s*Value", re.IGNORECASE)

# Returns the fiscal period end date for whatever period the overrides point at
PERIOD_END_FIELD = "LATEST_PERIOD_END_DT_FULL_RECORD"


@contextmanager
def bloomberg_session(host="localhost", port=8194):
    # Initialize session connection; Desktop API default: localhost=8194
    session_options = blpapi.SessionOptions()
    session_options.setServerHost(host)
    session_options.setServerPort(port)

    session = blpapi.Session(session_options)
    if not session.start():
        raise ConnectionError("Failed to start Bloomberg API session.")

    if not session.openService("//blp/refdata"):
        session.stop()
        raise ConnectionError("Failed to open Bloomberg refdata service.")

    try:
        yield session
    finally:
        session.stop()


def fetch_window(session, ticker, frequency, segmentation_type, fiscal_year):
    """Request the 5-period PG_REVENUE window anchored at fiscal_year."""
    refDataService = session.getService("//blp/refdata")
    request = refDataService.createRequest("ReferenceDataRequest")

    # Bulk data field for given ticker
    request.getElement("securities").appendValue(ticker)
    request.getElement("fields").appendValue("PG_REVENUE")
    # Exact period end date for the anchor, so the fiscal year labels can be checked
    request.getElement("fields").appendValue(PERIOD_END_FIELD)

    # Add overrides for Geographic breakdown, periodicity and anchor year
    overrides = request.getElement("overrides")
    # Set Geographic mode ('G' for Geographic, 'P' for Product)
    ovr_geo = overrides.appendElement()
    ovr_geo.setElement("fieldId", "PRODUCT_GEO_OVERRIDE")
    ovr_geo.setElement("value", segmentation_type)

    # Set Periodicity ('A', 'Q', or 'FQ')
    ovr_per = overrides.appendElement()
    ovr_per.setElement("fieldId", "FUND_PER")
    ovr_per.setElement("value", frequency)

    # Anchor the 5-period window on a specific fiscal year
    ovr_yr = overrides.appendElement()
    ovr_yr.setElement("fieldId", "EQY_FUND_YEAR")
    ovr_yr.setElement("value", str(fiscal_year))

    # Send asynchronous request
    session.sendRequest(request)

    rows = []
    period_end_date = None

    # Parse incoming event response stream
    while True:
        event = session.nextEvent(5000)

        for msg in event:
            if msg.hasElement("securityData"):
                sec_data_array = msg.getElement("securityData")
                for i in range(sec_data_array.numValues()):
                    sec_data = sec_data_array.getValueAsElement(i)

                    if sec_data.hasElement("securityError"):
                        err = sec_data.getElement("securityError")
                        raise RuntimeError(f"Security error for {ticker}: {err}")

                    field_data = sec_data.getElement("fieldData")

                    if field_data.hasElement(PERIOD_END_FIELD):
                        period_end_date = field_data.getElementAsString(PERIOD_END_FIELD)

                    if field_data.hasElement("PG_REVENUE"):
                        pg_rev = field_data.getElement("PG_REVENUE")

                        # Loop through returned array elements
                        for j in range(pg_rev.numValues()):
                            row_elem = pg_rev.getValueAsElement(j)
                            row_dict = {}
                            for k in range(row_elem.numElements()):
                                elem = row_elem.getElement(k)
                                row_dict[str(elem.name())] = (
                                    None if elem.isNull() else elem.getValue()
                                )
                            rows.append(row_dict)

        # Stop listening when final RESPONSE event is received
        if event.eventType() == blpapi.Event.RESPONSE:
            break

    return rows, period_end_date


def fetch_period_end_date(session, ticker, frequency, fiscal_year):
    """Fiscal period end date for a single fiscal year, or None if unavailable."""
    refDataService = session.getService("//blp/refdata")
    request = refDataService.createRequest("ReferenceDataRequest")

    request.getElement("securities").appendValue(ticker)
    request.getElement("fields").appendValue(PERIOD_END_FIELD)

    overrides = request.getElement("overrides")
    ovr_per = overrides.appendElement()
    ovr_per.setElement("fieldId", "FUND_PER")
    ovr_per.setElement("value", frequency)

    ovr_yr = overrides.appendElement()
    ovr_yr.setElement("fieldId", "EQY_FUND_YEAR")
    ovr_yr.setElement("value", str(fiscal_year))

    session.sendRequest(request)

    value = None
    while True:
        event = session.nextEvent(5000)

        for msg in event:
            if msg.hasElement("securityData"):
                sec_data_array = msg.getElement("securityData")
                for i in range(sec_data_array.numValues()):
                    sec_data = sec_data_array.getValueAsElement(i)
                    if sec_data.hasElement("securityError"):
                        continue

                    field_data = sec_data.getElement("fieldData")
                    if field_data.hasElement(PERIOD_END_FIELD):
                        elem = field_data.getElement(PERIOD_END_FIELD)
                        if not elem.isNull():
                            value = field_data.getElementAsString(PERIOD_END_FIELD)

        if event.eventType() == blpapi.Event.RESPONSE:
            break

    return value


def fetch_period_end_dates(session, ticker, frequency, fiscal_years):
    """Map each fiscal year to its period end date. One request per year."""
    dates = {}
    for year in fiscal_years:
        dates[year] = fetch_period_end_date(session, ticker, frequency, year)
    missing = [y for y, d in dates.items() if not d]
    if missing:
        print(f"  no period end date for: {', '.join(str(y) for y in sorted(missing))}")
    return dates


def rows_to_long(rows, anchor_year, period_end_date=None):
    """Explode the 'Period N Value' columns into one record per (segment, year)."""
    records = []
    for row in rows:
        name = next((v for k, v in row.items() if "Metric Name" in k), None)
        level = next((v for k, v in row.items() if "Hierarchy Level" in k), None)

        for key, val in row.items():
            match = PERIOD_RE.search(key)
            if not match:
                continue
            n = int(match.group(1))
            offset = (n - 1) if PERIODS_NEWEST_FIRST else (PERIODS_PER_WINDOW - n)
            records.append({
                "segment": str(name).strip(),
                "indent": name, # raw name; leading spaces encode the nesting
                "level": level,
                "fiscal_year": anchor_year - offset,
                "value": val,
                "anchor_year": anchor_year,
                "anchor_period_end": period_end_date,
            })
    return records


def anchor_years(start_year, end_year, step=ANCHOR_STEP):
    y = end_year
    while y >= start_year:
        yield y
        y -= step


def check_overlaps(long_df):
    """Overlapping windows should agree on the years they share."""
    dupes = long_df.dropna(subset=["value"]).groupby(
        ["segment", "level", "fiscal_year"]
    )["value"].nunique()
    return dupes[dupes > 1]


def add_period_end_row(wide, period_ends):
    """Prepend a 'Period End' row giving the period end date for each FY column."""
    header = {}
    for col in wide.columns:
        if col == "segment":
            header[col] = "Period End"
        elif isinstance(col, str) and col.startswith("FY"):
            header[col] = period_ends.get(int(col[2:])) or ""
        else:
            header[col] = ""
    return pd.concat([pd.DataFrame([header]), wide], ignore_index=True)


def get_segment_revenue(ticker, frequency, segmentation_type, start_year, end_year):
    records = []
    consecutive_empty = 0
    with bloomberg_session() as session:
        for anchor in anchor_years(start_year, end_year):
            rows, period_end = fetch_window(
                session, ticker, frequency, segmentation_type, anchor
            )
            window = rows_to_long(rows, anchor, period_end)
            # Bloomberg can return segment rows whose period values are all null,
            # so an empty window means "no usable values", not just "no rows"
            if not any(r["value"] is not None for r in window):
                consecutive_empty += 1
                print(
                    f"  anchor {anchor}: no data "
                    f"({consecutive_empty}/{MAX_EMPTY_WINDOWS} consecutive)"
                )
                if consecutive_empty >= MAX_EMPTY_WINDOWS:
                    print(f"  reached start of history, stopping at anchor {anchor}")
                    break
                continue

            consecutive_empty = 0
            print(f"  anchor {anchor}: {len(rows)} segment rows (period end {period_end})")
            records.extend(window)

        long_df = pd.DataFrame(records)
        if long_df.empty:
            return long_df, long_df, long_df

        long_df = long_df[long_df["fiscal_year"].between(start_year, end_year)].copy()
        long_df["value"] = pd.to_numeric(long_df["value"], errors="coerce")
        long_df = long_df.dropna(subset=["value"])

        # Snap floating-point zeros before the overlap check, so the same segment
        # reported as 0 and as -2.4e-14 in two windows doesn't look like a conflict
        noise = long_df["value"].abs() < ZERO_THRESHOLD
        if noise.any():
            print(f"  rounding {int(noise.sum())} near-zero values to 0")
        long_df.loc[noise, "value"] = 0.0

        # Only look up dates for years that actually came back with data
        years = sorted(long_df["fiscal_year"].unique(), reverse=True)
        print(f"\nFetching period end dates for {len(years)} fiscal years")
        period_ends = fetch_period_end_dates(session, ticker, frequency, years)

    long_df["period_end"] = long_df["fiscal_year"].map(period_ends)

    conflicts = check_overlaps(long_df)
    if not conflicts.empty:
        print("\nWARNING: overlapping windows disagree on these cells - the")
        print("period-to-year mapping may be reversed. Try flipping PERIODS_NEWEST_FIRST.")
        print(conflicts.head(10))

    # Overlapping windows repeat cells; prefer the value from the newest anchor
    deduped = long_df.sort_values("anchor_year", ascending=False).drop_duplicates(
        subset=["segment", "level", "fiscal_year"], keep="first"
    )

    wide = deduped.pivot_table(
        index=["segment", "level"],
        columns="fiscal_year",
        values="value",
        aggfunc="first",
    )
    wide = wide.sort_index(axis=1, ascending=False).reset_index()
    wide.columns = [
        f"FY{c}" if isinstance(c, (int, float)) else c for c in wide.columns
    ]
    wide = add_period_end_row(wide, period_ends)

    periods_df = pd.DataFrame(
        {"fiscal_year": years, "period_end": [period_ends.get(y) for y in years]}
    )
    return wide, long_df, periods_df


print(f"Requesting {TICKER_COMBINATION} segment revenue, FY{START_YEAR}-FY{END_YEAR}")
df1, df_long, df_periods = get_segment_revenue(
    ticker=TICKER_COMBINATION,
    frequency=FREQUENCY,
    segmentation_type=SEGMENTATION_TYPE,
    start_year=START_YEAR,
    end_year=END_YEAR,
)

if not df1.empty:
    output_file = f"{TICKER}_segment_revenue.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="wide", index=False)
        df_long.to_excel(writer, sheet_name="long", index=False)
        df_periods.to_excel(writer, sheet_name="periods", index=False)
    print(f"\nData successfully saved to {output_file}")
    print(df1.head(10))
else:
    print("No segment revenue data returned.")
