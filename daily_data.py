"""
Pull daily historical price data + dividend history for a ticker from the
Bloomberg Desktop API (BLPAPI, //blp/refdata) and combine them into one
DataFrame.

Requires a running Bloomberg Terminal on this machine (bbcomm.exe listening on
localhost:8194) and the blpapi Python package:

    pip install --index-url https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi

Output goes to a new Excel workbook (sheets: daily_data, prices, dividends,
summary, column_coverage). Existing files are never overwritten -- the path is
auto-suffixed _1, _2, ... unless you pass --overwrite.

Usage
-----
    # full available history -> AAPL_US_Equity_daily_data.xlsx
    python daily_data.py AAPL

    python daily_data.py AAPL --start 2015-01-01 --end 2025-12-31
    python daily_data.py AAPL --out aapl.xlsx --csv aapl.csv
    python daily_data.py AAPL --start 2024-01-01 --print-only

    # or from Python
    from daily_data import get_daily_data, export_excel
    df, prices, divs = get_daily_data("AAPL US Equity", "2015-01-01",
                                      return_parts=True)
    export_excel(df, ticker="AAPL US Equity", prices=prices, dividends=divs)

Needs `openpyxl` for the Excel writer:  pip install openpyxl


ON "ADJUSTED CLOSE"
-------------------
Bloomberg has no `adjclose` field. Price adjustment is a property of the
*request*, not a separate mnemonic -- you set adjustmentSplit /
adjustmentNormal / adjustmentAbnormal and PX_LAST comes back adjusted or not.
So this module issues TWO HistoricalDataRequests (one raw, one fully adjusted)
and merges them, giving you both `close` and `adjClose`. That is the only way
to get both in one frame.

`adjustmentFollowDPDF=True` makes Bloomberg honour your terminal's DPDF<GO>
defaults, which can silently override the explicit flags. This module sets it
False on both requests so the behaviour is deterministic and not dependent on
whoever last touched DPDF on this terminal.


ON VWAP
-------
EQY_WEIGHTED_AVG_PX is the full-day volume-weighted average price. It is not
populated for every security or every history depth -- if it comes back empty,
check the field in FLDS<GO> for your security type. VWAP_START_TIME /
VWAP_END_TIME overrides can narrow it to an intraday window.
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
import time
from typing import Iterable, Optional, Sequence

import pandas as pd

try:
    import blpapi
except ImportError:  # pragma: no cover
    blpapi = None


__all__ = [
    "DEFAULT_PRICE_FIELDS",
    "PRICE_VOLUME_FIELDS",
    "SIZE_FIELDS",
    "VALUATION_FIELDS",
    "FUNDAMENTAL_RATIO_FIELDS",
    "MILLIONS_FIELDS",
    "PRICE_FIELD_LABELS",
    "BloombergSession",
    "fetch_daily_prices",
    "fetch_dividend_history",
    "combine_price_and_dividends",
    "get_daily_data",
    "export_excel",
]


HOST = "localhost"
PORT = 8194
REFDATA_SVC = "//blp/refdata"

# Default start date: earlier than any listed equity. Bloomberg clamps to
# whatever history it actually has, so this just means "everything".
EARLIEST_START = "1900-01-01"


# ---------------------------------------------------------------------------
# Fields
# ---------------------------------------------------------------------------
# Daily price/volume fields available on HistoricalDataRequest. Trim or extend
# as needed -- any field returning #N/A for your security type just comes back
# as an all-NaN column (see the column_coverage sheet in the workbook).
PRICE_VOLUME_FIELDS: tuple[str, ...] = (
    "PX_OPEN",                        # open
    "PX_HIGH",                        # high
    "PX_LOW",                         # low
    "PX_LAST",                        # close (raw / unadjusted)
    "PX_VOLUME",                      # volume
    "EQY_WEIGHTED_AVG_PX",            # VWAP (full day)
    "TURNOVER",                       # traded value
    "NUM_TRADES",                     # trade count
    "VOLUME_AVG_30D",                 # 30-day avg volume
)

# Size fields. Bloomberg returns these in MILLIONS -- see MILLIONS_FIELDS.
SIZE_FIELDS: tuple[str, ...] = (
    "EQY_SH_OUT",                     # shares outstanding
    "CUR_MKT_CAP",                    # market cap
    "CURRENT_ENTERPRISE_VALUE",       # enterprise value          # VERIFY
)

# Valuation ratios. These are price-driven, so they genuinely move every day.
VALUATION_FIELDS: tuple[str, ...] = (
    "PE_RATIO",                       # trailing P/E
    "BEST_PE_RATIO",                  # forward P/E (consensus)   # VERIFY
    "PX_TO_BOOK_RATIO",               # P/B
    "PX_TO_SALES_RATIO",              # P/S                       # VERIFY
    # "PX_TO_CASH_FLOW",                # P/CF                      # VERIFY
    "PX_TO_FREE_CASH_FLOW",           # P/FCF                     # VERIFY
    "PX_TO_TANG_BV_PER_SH",           # P/tangible book           # VERIFY
    "CURRENT_EV_TO_T12M_EBITDA",      # EV/EBITDA (trailing)      # VERIFY
    "CURRENT_EV_TO_T12M_SALES",       # EV/Sales (trailing)       # VERIFY
    "EQY_DVD_YLD_IND",                # indicated dividend yield
    "EQY_DVD_YLD_12M",                # trailing 12m yield        # VERIFY
)

# Fundamentals that only change on a filing, so they step rather than move
# daily. Still useful joined to a daily series.
FUNDAMENTAL_RATIO_FIELDS: tuple[str, ...] = (
    "T12M_DIL_EPS_CONT_OPS",          # trailing 12m diluted EPS  # VERIFY
    # "RETURN_COM_EQY",                 # ROE                       # VERIFY
    # "RETURN_ON_ASSET",                # ROA                       # VERIFY
    # "GROSS_MARGIN",                   # VERIFY
    # "OPER_MARGIN",                    # VERIFY
    # "PROF_MARGIN",                    # net margin                # VERIFY
    # "TOT_DEBT_TO_TOT_EQY",            # D/E                       # VERIFY
    # "CUR_RATIO",                      # current ratio             # VERIFY
)

DEFAULT_PRICE_FIELDS: tuple[str, ...] = (
    PRICE_VOLUME_FIELDS
    + SIZE_FIELDS
    + VALUATION_FIELDS
    + FUNDAMENTAL_RATIO_FIELDS
)

# Bloomberg reports these in millions of shares / millions of currency. Left
# alone, 3.4bn shares arrives as 3416.9, which is easy to misread. Multiplied
# out to whole units by fetch_daily_prices(scale_to_units=True).
#
# VERIFY the scale against the terminal once -- if Bloomberg ever hands these
# back already in units, this would inflate them by 1e6. Sanity check: LRCX
# market cap should land around 1.0e11, not 1.0e17.
MILLIONS_FIELDS: dict[str, float] = {
    "EQY_SH_OUT": 1e6,
    "CUR_MKT_CAP": 1e6,
    "CURRENT_ENTERPRISE_VALUE": 1e6,
}

# Written to Excel with a thousands-separated integer format so they display
# in full rather than as 1.23E+11.
WHOLE_NUMBER_COLUMNS: frozenset[str] = frozenset({
    "volume", "PX_VOLUME", "TURNOVER", "NUM_TRADES", "VOLUME_AVG_30D",
    "EQY_SH_OUT", "CUR_MKT_CAP", "CURRENT_ENTERPRISE_VALUE",
})
WHOLE_NUMBER_FORMAT = "#,##0"

# Friendly names for the OHLCV core. Anything not listed keeps its mnemonic --
# the ratio fields deliberately stay as mnemonics so they match FLDS<GO>.
PRICE_FIELD_LABELS: dict[str, str] = {
    "PX_OPEN": "open",
    "PX_HIGH": "high",
    "PX_LOW": "low",
    "PX_LAST": "close",
    "PX_VOLUME": "volume",
    "EQY_WEIGHTED_AVG_PX": "vwap",
}

# Bulk-data field holding the dividend/corporate-action history.
DIVIDEND_FIELD = "DVD_HIST"


# ---------------------------------------------------------------------------
# Session management
# ---------------------------------------------------------------------------
class BloombergSession:
    """Context manager around a blpapi Session on //blp/refdata."""

    def __init__(self, host: str = HOST, port: int = PORT) -> None:
        if blpapi is None:
            raise ImportError(
                "blpapi is not installed. Install it with:\n"
                "  pip install --index-url "
                "https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi"
            )
        self.host = host
        self.port = port
        self.session: Optional["blpapi.Session"] = None

    def __enter__(self) -> "BloombergSession":
        opts = blpapi.SessionOptions()
        opts.setServerHost(self.host)
        opts.setServerPort(self.port)

        self.session = blpapi.Session(opts)
        if not self.session.start():
            raise ConnectionError(
                f"Could not start blpapi session on {self.host}:{self.port}. "
                "Is the Bloomberg Terminal running and logged in?"
            )
        if not self.session.openService(REFDATA_SVC):
            self.session.stop()
            raise ConnectionError(f"Could not open {REFDATA_SVC}")
        return self

    def __exit__(self, *exc) -> None:
        if self.session is not None:
            self.session.stop()
            self.session = None

    @property
    def service(self) -> "blpapi.Service":
        assert self.session is not None, "session not started"
        return self.session.getService(REFDATA_SVC)

    def send(
        self,
        request: "blpapi.Request",
        *,
        timeout_seconds: float = 300.0,
        poll_millis: int = 1000,
    ) -> list["blpapi.Message"]:
        """
        Send a request and drain the response messages.

        A TIMEOUT event just means nothing arrived in this poll window -- it is
        not an error. Long histories routinely go quiet for several seconds
        between partial responses, so we keep polling until either the final
        RESPONSE lands or `timeout_seconds` elapses with no traffic at all.
        """
        assert self.session is not None, "session not started"
        self.session.sendRequest(request)

        messages: list["blpapi.Message"] = []
        deadline = time.monotonic() + timeout_seconds

        while True:
            # NB: nextEvent takes the timeout POSITIONALLY -- the blpapi
            # Python binding has no `timeoutMillis` keyword.
            ev = self.session.nextEvent(poll_millis)
            etype = ev.eventType()

            if etype in (blpapi.Event.PARTIAL_RESPONSE, blpapi.Event.RESPONSE):
                for msg in ev:
                    if msg.hasElement("responseError"):
                        raise RuntimeError(
                            f"Bloomberg responseError: "
                            f"{msg.getElement('responseError')}"
                        )
                    messages.append(msg)
                if etype == blpapi.Event.RESPONSE:
                    break
                # more to come -- reset the idle clock
                deadline = time.monotonic() + timeout_seconds

            elif etype == blpapi.Event.TIMEOUT:
                if time.monotonic() > deadline:
                    raise TimeoutError(
                        f"No response from Bloomberg in {timeout_seconds:.0f}s. "
                        "Is the Terminal still logged in?"
                    )
                # otherwise keep waiting

        return messages


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _bbg_date(value) -> str:
    """Coerce anything date-like to Bloomberg's YYYYMMDD string."""
    if isinstance(value, str):
        value = pd.Timestamp(value)
    if isinstance(value, (dt.datetime, pd.Timestamp)):
        value = value.date()
    if not isinstance(value, dt.date):
        raise TypeError(f"cannot interpret {value!r} as a date")
    return value.strftime("%Y%m%d")


def _normalise_ticker(ticker: str) -> str:
    """Accept 'AAPL' and turn it into a valid Bloomberg security string."""
    t = ticker.strip()
    upper = t.upper()
    known_suffixes = (
        "EQUITY", "INDEX", "CURNCY", "COMDTY", "CORP", "GOVT", "MTGE", "PFD",
    )
    if upper.endswith(known_suffixes):
        return t
    # bare 'AAPL' -> 'AAPL US Equity'
    if " " not in t:
        return f"{t} US Equity"
    return f"{t} Equity"


def _element_to_py(element: "blpapi.Element"):
    """Convert a scalar blpapi Element to a Python value."""
    if element.isNull():
        return None
    return element.getValue()


# ---------------------------------------------------------------------------
# Price history
# ---------------------------------------------------------------------------
def _historical_request(
    bb: BloombergSession,
    security: str,
    fields: Sequence[str],
    start: str,
    end: str,
    *,
    adjusted: bool,
    currency: Optional[str] = None,
    periodicity: str = "DAILY",
    non_trading_days: str = "ACTIVE_DAYS_ONLY",
) -> pd.DataFrame:
    """One HistoricalDataRequest -> DataFrame indexed by date."""
    req = bb.service.createRequest("HistoricalDataRequest")
    req.getElement("securities").appendValue(security)
    for f in fields:
        req.getElement("fields").appendValue(f)

    req.set("startDate", start)
    req.set("endDate", end)
    req.set("periodicitySelection", periodicity)
    req.set("periodicityAdjustment", "ACTUAL")
    req.set("nonTradingDayFillOption", non_trading_days)
    req.set("nonTradingDayFillMethod", "NIL_VALUE")

    # Deterministic adjustment behaviour -- do NOT inherit DPDF<GO> defaults.
    req.set("adjustmentFollowDPDF", False)
    req.set("adjustmentSplit", adjusted)
    req.set("adjustmentNormal", adjusted)
    req.set("adjustmentAbnormal", adjusted)

    if currency:
        req.set("currency", currency)

    records: list[dict] = []
    for msg in bb.send(req):
        if not msg.hasElement("securityData"):
            continue
        sec_data = msg.getElement("securityData")

        if sec_data.hasElement("securityError"):
            raise ValueError(
                f"Bloomberg securityError for {security!r}: "
                f"{sec_data.getElement('securityError')}"
            )
        if sec_data.hasElement("fieldExceptions"):
            fx = sec_data.getElement("fieldExceptions")
            for i in range(fx.numValues()):
                ex = fx.getValueAsElement(i)
                bad = ex.getElementAsString("fieldId")
                print(f"  [warn] field not available for {security}: {bad}",
                      file=sys.stderr)

        field_data = sec_data.getElement("fieldData")
        for i in range(field_data.numValues()):
            row_el = field_data.getValueAsElement(i)
            row: dict = {}
            for j in range(row_el.numElements()):
                el = row_el.getElement(j)
                row[str(el.name())] = _element_to_py(el)
            records.append(row)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame.from_records(records)
    df["date"] = pd.to_datetime(df["date"])
    return df.set_index("date").sort_index()


def fetch_daily_prices(
    bb: BloombergSession,
    ticker: str,
    start,
    end=None,
    *,
    fields: Iterable[str] = DEFAULT_PRICE_FIELDS,
    include_adjusted_close: bool = True,
    currency: Optional[str] = None,
    rename: bool = True,
    scale_to_units: bool = True,
) -> pd.DataFrame:
    """
    Daily OHLCV + extended price data for one security.

    Issues an unadjusted request for everything in `fields`, then (if
    include_adjusted_close) a second split/dividend-adjusted request for
    PX_LAST, merged in as `adjClose`.

    scale_to_units multiplies the MILLIONS_FIELDS out to whole units, so
    shares outstanding and market cap read as real numbers.
    """
    security = _normalise_ticker(ticker)
    start_s = _bbg_date(start)
    end_s = _bbg_date(end or dt.date.today())

    prices = _historical_request(
        bb, security, list(fields), start_s, end_s,
        adjusted=False, currency=currency,
    )
    if prices.empty:
        return prices

    if scale_to_units:
        for col, factor in MILLIONS_FIELDS.items():
            if col in prices.columns:
                prices[col] = pd.to_numeric(prices[col], errors="coerce") * factor

    if include_adjusted_close:
        adj = _historical_request(
            bb, security, ["PX_LAST"], start_s, end_s,
            adjusted=True, currency=currency,
        )
        if not adj.empty:
            prices = prices.join(
                adj["PX_LAST"].rename("__ADJ_PX_LAST"), how="left"
            )

    if rename:
        prices = prices.rename(columns=PRICE_FIELD_LABELS)
    prices = prices.rename(columns={"__ADJ_PX_LAST": "adjClose"})

    # Put the canonical OHLCV block first, in a sensible order.
    lead = [c for c in ("open", "high", "low", "close", "adjClose",
                        "volume", "vwap") if c in prices.columns]
    rest = [c for c in prices.columns if c not in lead]
    prices = prices[lead + rest]

    prices.insert(0, "ticker", security)
    return prices


# ---------------------------------------------------------------------------
# Dividend history (DVD_HIST bulk field)
# ---------------------------------------------------------------------------
def fetch_dividend_history(
    bb: BloombergSession,
    ticker: str,
    start=None,
    end=None,
) -> pd.DataFrame:
    """
    DVD_HIST as a DataFrame. Columns come straight from Bloomberg and are left
    exactly as returned -- typically:

        Declared Date, Ex-Date, Record Date, Payable Date,
        Dividend Amount, Dividend Frequency, Dividend Type

    Column names vary by security and Bloomberg release, so nothing here
    hard-codes them.
    """
    security = _normalise_ticker(ticker)

    req = bb.service.createRequest("ReferenceDataRequest")
    req.getElement("securities").appendValue(security)
    req.getElement("fields").appendValue(DIVIDEND_FIELD)

    if start is not None or end is not None:
        overrides = req.getElement("overrides")
        if start is not None:
            ov = overrides.appendElement()
            ov.setElement("fieldId", "DVD_START_DT")
            ov.setElement("value", _bbg_date(start))
        if end is not None:
            ov = overrides.appendElement()
            ov.setElement("fieldId", "DVD_END_DT")
            ov.setElement("value", _bbg_date(end))

    records: list[dict] = []
    for msg in bb.send(req):
        if not msg.hasElement("securityData"):
            continue
        sec_array = msg.getElement("securityData")
        for s in range(sec_array.numValues()):
            sec_data = sec_array.getValueAsElement(s)

            if sec_data.hasElement("securityError"):
                raise ValueError(
                    f"Bloomberg securityError for {security!r}: "
                    f"{sec_data.getElement('securityError')}"
                )
            field_data = sec_data.getElement("fieldData")
            if not field_data.hasElement(DIVIDEND_FIELD):
                continue

            bulk = field_data.getElement(DIVIDEND_FIELD)
            for i in range(bulk.numValues()):
                row_el = bulk.getValueAsElement(i)
                row = {}
                for j in range(row_el.numElements()):
                    el = row_el.getElement(j)
                    row[str(el.name())] = _element_to_py(el)
                records.append(row)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame.from_records(records)

    # Parse anything that looks like a date column, leave names untouched.
    for col in df.columns:
        if "date" in col.lower():
            df[col] = pd.to_datetime(df[col], errors="coerce")

    return df


# ---------------------------------------------------------------------------
# Combine
# ---------------------------------------------------------------------------
def _find_ex_date_column(dividends: pd.DataFrame) -> str:
    """Locate the ex-date column without assuming its exact spelling."""
    for col in dividends.columns:
        squashed = col.lower().replace("-", "").replace("_", "").replace(" ", "")
        if squashed in ("exdate", "exdt", "exdividenddate"):
            return col
    # fall back to the first datetime column
    for col in dividends.columns:
        if pd.api.types.is_datetime64_any_dtype(dividends[col]):
            return col
    raise KeyError(
        f"Could not find an ex-date column in DVD_HIST output. "
        f"Columns were: {list(dividends.columns)}"
    )


def combine_price_and_dividends(
    prices: pd.DataFrame,
    dividends: pd.DataFrame,
    *,
    on: Optional[str] = None,
) -> pd.DataFrame:
    """
    Left-join the dividend frame onto the daily price frame by ex-date.

    Every price row is kept; dividend columns are NaN on non-ex-dates. The
    dividend columns keep their original Bloomberg names -- nothing is renamed
    or dropped. If a name collides with a price column, the dividend one gets a
    `_dvd` suffix so both survive.
    """
    if prices.empty or dividends.empty:
        return prices.copy()

    key = on or _find_ex_date_column(dividends)
    div = dividends.copy()
    div[key] = pd.to_datetime(div[key], errors="coerce")
    div = div.dropna(subset=[key])
    if div.empty:
        return prices.copy()

    # Multiple dividends can share one ex-date (e.g. regular + special). A
    # plain merge would duplicate the price row for that day; flag it loudly
    # rather than silently changing the row count.
    dupes = int(div[key].duplicated().sum())
    if dupes:
        print(
            f"  [warn] {dupes} dividend record(s) share an ex-date with "
            f"another; those price rows will be duplicated in the output.",
            file=sys.stderr,
        )

    overlap = (set(div.columns) & set(prices.columns)) - {key}
    if overlap:
        div = div.rename(columns={c: f"{c}_dvd" for c in overlap})

    # Merge on a scratch key rather than left_index/right_on, which would
    # backfill the ex-date column with the price date on every non-ex row.
    date_col = prices.index.name or "date"
    left = prices.reset_index().rename(columns={prices.index.name or "index": date_col})
    left["__merge_dt"] = pd.to_datetime(left[date_col]).dt.normalize()
    div["__merge_dt"] = div[key].dt.normalize()

    combined = left.merge(div, how="left", on="__merge_dt")
    combined = combined.drop(columns="__merge_dt").set_index(date_col)
    return combined


# ---------------------------------------------------------------------------
# Top-level convenience
# ---------------------------------------------------------------------------
def get_daily_data(
    ticker: str,
    start,
    end=None,
    *,
    fields: Iterable[str] = DEFAULT_PRICE_FIELDS,
    include_adjusted_close: bool = True,
    currency: Optional[str] = None,
    dividends: bool = True,
    scale_to_units: bool = True,
    host: str = HOST,
    port: int = PORT,
    return_parts: bool = False,
):
    """
    Open a session, pull prices + dividends, return one combined frame.

    With return_parts=True returns (combined, prices, dividends) instead, so
    callers can write each to its own Excel sheet.
    """
    divs = pd.DataFrame()
    with BloombergSession(host, port) as bb:
        prices = fetch_daily_prices(
            bb, ticker, start, end,
            fields=fields,
            include_adjusted_close=include_adjusted_close,
            currency=currency,
            scale_to_units=scale_to_units,
        )
        if prices.empty:
            print(f"  [warn] no price data returned for {ticker!r}",
                  file=sys.stderr)
            return (prices, prices, divs) if return_parts else prices

        if dividends:
            divs = fetch_dividend_history(bb, ticker, start, end)
            if divs.empty:
                print(f"  [warn] no DVD_HIST records returned for {ticker!r}",
                      file=sys.stderr)

    combined = (combine_price_and_dividends(prices, divs)
                if not divs.empty else prices)
    return (combined, prices, divs) if return_parts else combined


# ---------------------------------------------------------------------------
# Excel export
#
# NOTE: _unique_path / _write_sheet are duplicated in data.py rather than
# shared, so each file stays a standalone program you can move on its own.
# ---------------------------------------------------------------------------
def _unique_path(path: str) -> str:
    """Never overwrite: 'x.xlsx' -> 'x_1.xlsx' -> 'x_2.xlsx' if taken."""
    import os

    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    n = 1
    while os.path.exists(f"{stem}_{n}{ext}"):
        n += 1
    return f"{stem}_{n}{ext}"


def _write_sheet(writer, df: pd.DataFrame, sheet_name: str,
                 index: bool = False,
                 number_formats: Optional[dict] = None) -> None:
    """Write one DataFrame to a sheet with frozen header, autofilter and
    reasonable column widths. Strips tz-awareness, which Excel cannot store.

    number_formats maps a column name to an Excel format string, e.g.
    {'CUR_MKT_CAP': '#,##0'} to force full digits instead of 1.23E+11.
    """
    df = df.copy()
    for col in df.columns:
        if isinstance(df[col].dtype, pd.DatetimeTZDtype):
            df[col] = df[col].dt.tz_localize(None)
    if index and isinstance(df.index, pd.DatetimeIndex) and df.index.tz is not None:
        df.index = df.index.tz_localize(None)

    sheet_name = sheet_name[:31]  # Excel hard limit
    df.to_excel(writer, sheet_name=sheet_name, index=index)

    ws = writer.sheets[sheet_name]
    ws.freeze_panes = "B2" if index else "A2"
    n_cols = len(df.columns) + (1 if index else 0)
    n_rows = len(df) + 1
    if n_cols and n_rows:
        from openpyxl.utils import get_column_letter

        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{n_rows}"
        headers = ([df.index.name or "index"] if index else []) + \
                  [str(c) for c in df.columns]
        series = ([df.index.to_series()] if index else []) + \
                 [df[c] for c in df.columns]
        for i, (head, s) in enumerate(zip(headers, series), start=1):
            longest = s.astype(str).str.len().max() if len(s) else 0
            # .max() is NaN on an empty/all-null column, and NaN is truthy --
            # `longest or 0` would let it through to int() and blow up.
            if pd.isna(longest):
                longest = 0
            width = max(len(head), int(longest)) + 2
            ws.column_dimensions[get_column_letter(i)].width = min(width, 40)

            fmt = (number_formats or {}).get(head)
            if fmt:
                letter = get_column_letter(i)
                for row in range(2, n_rows + 1):
                    ws[f"{letter}{row}"].number_format = fmt


def _safe_name(ticker: str) -> str:
    """'AAPL US Equity' -> 'AAPL_US_Equity', usable in a filename."""
    keep = [c if (c.isalnum() or c in "-_") else "_" for c in ticker.strip()]
    return "".join(keep).strip("_") or "output"


def export_excel(
    combined: pd.DataFrame,
    path: Optional[str] = None,
    *,
    ticker: str = "",
    prices: Optional[pd.DataFrame] = None,
    dividends: Optional[pd.DataFrame] = None,
    overwrite: bool = False,
) -> str:
    """
    Write the daily data to a new Excel workbook.

    Sheets
    ------
    daily_data   combined prices + dividend columns (the main output)
    prices       price frame alone, if passed
    dividends    raw DVD_HIST records, if passed
    summary      row/column counts, date range, non-null coverage per column

    Returns the path actually written (auto-suffixed unless overwrite=True).
    """
    if path is None:
        path = f"{_safe_name(ticker) or 'daily'}_daily_data.xlsx"
    if not overwrite:
        path = _unique_path(path)

    summary_rows = [
        {"item": "ticker", "value": ticker},
        {"item": "rows", "value": len(combined)},
        {"item": "columns", "value": len(combined.columns)},
    ]
    if len(combined):
        summary_rows += [
            {"item": "start", "value": str(combined.index.min())},
            {"item": "end", "value": str(combined.index.max())},
        ]
    if dividends is not None:
        summary_rows.append({"item": "dividend records",
                             "value": len(dividends)})
    summary = pd.DataFrame(summary_rows)

    # Which columns actually came back populated -- an all-empty column means
    # the mnemonic isn't supported for this security type.
    coverage = pd.DataFrame(
        [
            {
                "column": str(c),
                "non_null": int(combined[c].notna().sum()),
                "pct_populated": (round(100 * combined[c].notna().mean(), 1)
                                  if len(combined) else 0.0),
            }
            for c in combined.columns
        ]
    )

    # Force full digits on the big integer columns -- Excel would otherwise
    # show 1.23E+11 for a market cap.
    num_fmt = {c: WHOLE_NUMBER_FORMAT for c in WHOLE_NUMBER_COLUMNS}

    with pd.ExcelWriter(path, engine="openpyxl",
                        datetime_format="yyyy-mm-dd") as writer:
        _write_sheet(writer, combined, "daily_data", index=True,
                     number_formats=num_fmt)
        if prices is not None and not prices.empty:
            _write_sheet(writer, prices, "prices", index=True,
                         number_formats=num_fmt)
        if dividends is not None and not dividends.empty:
            _write_sheet(writer, dividends, "dividends")
        _write_sheet(writer, summary, "summary")
        _write_sheet(writer, coverage, "column_coverage")

    return path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main(argv: Optional[Sequence[str]] = None) -> int:
    p = argparse.ArgumentParser(
        description="Pull daily prices + dividend history from the Bloomberg "
                    "Desktop API and combine them into one DataFrame."
    )
    p.add_argument("ticker",
                   help="e.g. 'AAPL' or 'AAPL US Equity'")
    p.add_argument("--start", default=EARLIEST_START,
                   help="YYYY-MM-DD (default: %(default)s -- i.e. as far back "
                        "as Bloomberg has data; it simply returns history from "
                        "the security's inception)")
    p.add_argument("--end", default=None, help="YYYY-MM-DD (default: today)")
    p.add_argument("--currency", default=None, help="e.g. USD")
    p.add_argument("--no-dividends", action="store_true")
    p.add_argument("--no-adj-close", action="store_true",
                   help="skip the second adjusted-price request")
    p.add_argument("--no-ratios", action="store_true",
                   help="price/volume and size only -- skip the valuation and "
                        "fundamental ratio fields")
    p.add_argument("--raw-millions", action="store_true",
                   help="leave EQY_SH_OUT / CUR_MKT_CAP / EV in millions as "
                        "Bloomberg returns them, instead of whole units")
    p.add_argument("--out", default=None,
                   help="workbook path (default: <TICKER>_daily_data.xlsx)")
    p.add_argument("--overwrite", action="store_true",
                   help="overwrite --out instead of auto-suffixing _1, _2, ...")
    p.add_argument("--csv", metavar="PATH", default=None,
                   help="also write a flat CSV of the combined frame")
    p.add_argument("--print-only", action="store_true",
                   help="print a preview instead of writing a workbook")
    p.add_argument("--host", default=HOST)
    p.add_argument("--port", type=int, default=PORT)
    args = p.parse_args(argv)

    fields = (PRICE_VOLUME_FIELDS + SIZE_FIELDS) if args.no_ratios \
        else DEFAULT_PRICE_FIELDS

    df, prices, divs = get_daily_data(
        args.ticker,
        args.start,
        args.end,
        fields=fields,
        include_adjusted_close=not args.no_adj_close,
        currency=args.currency,
        dividends=not args.no_dividends,
        scale_to_units=not args.raw_millions,
        host=args.host,
        port=args.port,
        return_parts=True,
    )

    if df.empty:
        print("No data returned.", file=sys.stderr)
        return 1

    if args.print_only:
        with pd.option_context("display.max_columns", None,
                               "display.width", 250):
            print(df.head(10))
            print(f"\n{len(df):,} rows x {len(df.columns)} columns")
            print(f"columns: {list(df.columns)}")
        return 0

    written = export_excel(
        df, args.out,
        ticker=_normalise_ticker(args.ticker),
        prices=prices,
        dividends=divs,
        overwrite=args.overwrite,
    )
    print(f"Wrote {len(df):,} rows x {len(df.columns)} cols -> {written}")

    if args.csv:
        df.to_csv(args.csv)
        print(f"Wrote CSV -> {args.csv}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
