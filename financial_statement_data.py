"""
FMP (Financial Modeling Prep) -> Bloomberg Desktop API (BLPAPI) field map.

Three columns per row:
    fmp_field       the FMP JSON key / column name
    bbg_field       the Bloomberg field mnemonic (None = no direct equivalent)
    datatable       'income_statement' | 'balance_sheet' | 'cash_flow_statement'
                    | 'employee_count'

HOW TO USE
----------
Bloomberg has no REST-style "endpoints"; the Desktop API exposes fundamentals
as field mnemonics on the //blp/refdata service:

    - Point-in-time / latest:  BDP(security, field)      (xbbg: blp.bdp)
    - Time series by period:   BDH(security, field, ...) (xbbg: blp.bdh)

Set periodicity / period with overrides:
    FUND_PER ('A' annual, 'Q' quarterly, 'S' semi)
    EQY_FUND_YEAR, EQY_FUND_RELATIVE_PERIOD (e.g. '-1FY', '-2FQ')
    FILING_STATUS ('OR' original, 'MR' most recent, 'PR' preliminary)
    FUNDAMENTAL_PUBLIC_DATE

Example:
    from xbbg import blp
    from data import bbg_fields
    df = blp.bdh('AAPL US Equity', bbg_fields('income_statement'),
                 '2015-01-01', '2025-12-31', FUND_PER='A', FILING_STATUS='MR')

IMPORTANT CAVEAT
----------------
Bloomberg mnemonics drift between releases and many do not map 1:1 to FMP's
taxonomy. Rows whose `note` contains 'VERIFY' are best-effort and MUST be
confirmed in the terminal with FLDS<GO> against a real security before you
rely on them. Rows with bbg_field=None have no standard Bloomberg field and
need to be derived -- see DERIVED_FIELDS at the bottom.

This module is primarily a static field map and imports nothing heavier than
the stdlib. The one live-data helper it carries -- fetch_employee_history()
for the 'employee_count' table -- imports pandas/blpapi lazily, so `import
data` still works on a machine with no Terminal.
"""

from __future__ import annotations

import csv
import re
from typing import TYPE_CHECKING, Iterable, NamedTuple, Optional

if TYPE_CHECKING:  # pragma: no cover
    import pandas as pd

__all__ = [
    "FieldMap",
    "FIELD_MAP",
    "INCOME_STATEMENT",
    "BALANCE_SHEET",
    "CASH_FLOW_STATEMENT",
    "EMPLOYEE_COUNT",
    "DATATABLES",
    "DERIVED_FIELDS",
    "rows",
    "mapping",
    "bbg_fields",
    "fmp_to_bbg",
    "bbg_to_fmp",
    "mapped",
    "unmapped",
    "needs_verification",
    "status",
    "export_csv",
    "export_txt",
    "export_excel",
    "fetch_employee_history",
    "fetch_financials",
    "fetch_all_financials",
]


class FieldMap(NamedTuple):
    """One FMP field and its Bloomberg counterpart."""

    fmp_field: str
    bbg_field: Optional[str]
    datatable: str
    note: str = ""


INCOME_STATEMENT = "income_statement"
BALANCE_SHEET = "balance_sheet"
CASH_FLOW_STATEMENT = "cash_flow_statement"
EMPLOYEE_COUNT = "employee_count"

DATATABLES = (INCOME_STATEMENT, BALANCE_SHEET, CASH_FLOW_STATEMENT,
              EMPLOYEE_COUNT)

_IS = INCOME_STATEMENT
_BS = BALANCE_SHEET
_CF = CASH_FLOW_STATEMENT
_EC = EMPLOYEE_COUNT


# =============================================================================
# INCOME STATEMENT
# =============================================================================
_INCOME_STATEMENT_ROWS = [
    # ---- Identifiers & metadata --------------------------------------------
    FieldMap("isin",             "ID_ISIN",                    _IS),
    FieldMap("symbol",           "TICKER",                     _IS, "alt: PARSEKYABLE_DES / ID_BB_SEC_NUM_DES / ID_EXCH_SYMBOL"),
    FieldMap("cik",              "CENTRAL__INDEX_KEY_NUMBER",  _IS),
    FieldMap("reportedCurrency", "EQY_FUND_CRNCY",             _IS, "fundamentals reporting currency"),
    FieldMap("fillingDate",      "FUNDAMENTAL_PUBLIC_DATE",    _IS, "VERIFY - date financials were filed/published"),
    FieldMap("acceptedDate",     "LATEST_ANNOUNCEMENT_DT",     _IS, "VERIFY - closest analogue to SEC accepted timestamp"),
    FieldMap("fiscalYear",       "EQY_FUND_YEAR",              _IS, "alt: FISCAL__YEAR_PERIOD"),
    FieldMap("period",           "FUND_PER",                   _IS, "VERIFY - 'A' annual / 'Q' quarter; normally an override, not a field"),

    # ---- Revenue & cost -----------------------------------------------------
    FieldMap("revenue",       "SALES_REV_TURN",           _IS, "alt: IS_COMP_SALES"),
    FieldMap("costOfRevenue", "IS_COG_AND_SERVICES_SOLD", _IS, "Cost of Goods & Services Sold"),
    FieldMap("grossProfit",   "GROSS_PROFIT",             _IS),

    # ---- Operating expenses -------------------------------------------------
    FieldMap("researchAndDevelopmentExpenses",          "IS_RD_EXPEND",                 _IS, "VERIFY"),
    FieldMap("generalAndAdministrativeExpenses",        "AR_GENL_&_ADMIN_EXPN_SUM",     _IS, "VERIFY"),
    FieldMap("sellingAndMarketingExpenses",             "IS_SELLING_EXPENSES", _IS, "VERIFY"),
    FieldMap("sellingGeneralAndAdministrativeExpenses", "IS_SGA_EXPENSE",               _IS, "VERIFY"),
    FieldMap("otherExpenses",                           "ARD_OTHER_EXPENSES",            _IS, "VERIFY - sign and scope differ from FMP"),
    FieldMap("operatingExpenses",                       "IS_OPERATING_EXPN",            _IS),
    FieldMap("costAndExpenses",                         "IS_TOT_OPER_COST_AND_EXP",     _IS, "VERIFY - total cost & expenses"),

    # ---- Interest -----------------------------------------------------------
    FieldMap("netInterestIncome", "NET_INT_INC",     _IS, "VERIFY - primarily a bank/financials field"),
    FieldMap("interestIncome",    "IS_INT_INC",      _IS, "VERIFY"),
    FieldMap("interestExpense",   "IS_INT_EXPENSE",  _IS),

    # ---- D&A, EBITDA, EBIT --------------------------------------------------
    FieldMap("depreciationAndAmortization", "CF_DEPR_AMORT", _IS, "alt: ARDR_DEPRECIATION_AMORTIZATION"),
    FieldMap("ebitda",                      "EBITDA",        _IS),
    FieldMap("ebit",                        "EBIT",          _IS),

    # ---- Non-operating & operating income -----------------------------------
    FieldMap("nonOperatingIncomeExcludingInterest", None,                        _IS, "NO DIRECT EQUIVALENT - derive"),
    FieldMap("operatingIncome",                     "IS_OPER_INC",               _IS),
    FieldMap("otherIncome",                         "ARD_OTHER_INC",               _IS),
    # FieldMap("totalOtherIncomeExpensesNet",         "IS_TOT_NON_OPER_INC_LOSS",  _IS, "VERIFY"),

    # ---- Pre-tax, tax, net income -------------------------------------------
    FieldMap("incomeBeforeTax",                     "PRETAX_INC",                _IS),
    FieldMap("incomeTaxExpense",                    "IS_INC_TAX_EXP",            _IS),
    FieldMap("netIncomeFromContinuingOperations",   "IS_INC_BEF_XO_ITEM",        _IS, "VERIFY"),
    FieldMap("netIncomeFromDiscontinuedOperations", "IS_NET_INC_DISCONT_OPER",   _IS, "VERIFY"),
    # FieldMap("otherAdjustmentsToNetIncome",         None,                        _IS, "NO DIRECT EQUIVALENT - derive"),
    FieldMap("netIncome",                           "NET_INCOME",                _IS),
    # FieldMap("netIncomeDeductions",                 None,                        _IS, "NO DIRECT EQUIVALENT - pfd divs / minority int; see BS_PFD_DVD"),
    # FieldMap("bottomLineNetIncome",                 "IS_NET_INCOME_TO_COMMON",   _IS, "VERIFY - net income attributable to common"),

    # ---- Per-share & share counts -------------------------------------------
    FieldMap("eps",                      "IS_EPS",                 _IS, "basic reported EPS"),
    FieldMap("epsDiluted",               "IS_DILUTED_EPS",         _IS),
    FieldMap("weightedAverageShsOut",    "IS_AVG_NUM_SH_FOR_EPS",  _IS, "basic weighted-average shares"),
    FieldMap("weightedAverageShsOutDil", "IS_SH_FOR_DILUTED_EPS",  _IS, "diluted weighted-average shares"),
]


# =============================================================================
# BALANCE SHEET
#
# Column list as queried from FMP's `balance-sheet-statement` endpoint.
# NOTE: this list has no isin/symbol/cik/reportedCurrency columns (unlike the
# income statement above), and spells the filing date `filingDate` with one 'l'
# where the income statement uses `fillingDate`. Both reproduced as given.
# =============================================================================
_BALANCE_SHEET_ROWS = [
    # ---- Metadata -----------------------------------------------------------
    FieldMap("filingDate",   "FUNDAMENTAL_PUBLIC_DATE", _BS, "VERIFY"),
    FieldMap("acceptedDate", "LATEST_ANNOUNCEMENT_DT",  _BS, "VERIFY"),
    FieldMap("fiscalYear",   "EQY_FUND_YEAR",           _BS, "alt: FISCAL__YEAR_PERIOD"),
    FieldMap("period",       "FUND_PER",                _BS, "VERIFY - normally an override"),

    # ---- Current assets -----------------------------------------------------
    FieldMap("cashAndCashEquivalents",     "BS_CASH_NEAR_CASH_ITEM",      _BS),
    FieldMap("shortTermInvestments",       "BS_MKT_SEC_OTHER_ST_INVEST",  _BS),
    FieldMap("cashAndShortTermInvestments", "C&CE_AND_STI_DETAILED",      _BS, "VERIFY - else sum of the two above"),
    FieldMap("netReceivables",             "BS_ACCT_NOTE_RCV",            _BS),
    FieldMap("accountsReceivables",        "BS_ACCTS_REC_EXCL_NOTES_REC", _BS, "VERIFY"),
    FieldMap("otherReceivables",           None,                          _BS, "NO DIRECT EQUIVALENT - derive: BS_ACCT_NOTE_RCV - accountsReceivables"),
    FieldMap("inventory",                  "BS_INVENTORIES",              _BS),
    FieldMap("prepaids",                   "BS_PREPAY",                   _BS, "VERIFY"),
    FieldMap("otherCurrentAssets",         "BS_OTHER_CUR_ASSET",          _BS),
    FieldMap("totalCurrentAssets",         "BS_CUR_ASSET_REPORT",         _BS),

    # ---- Non-current assets -------------------------------------------------
    FieldMap("propertyPlantEquipmentNet",   "BS_NET_FIX_ASSET",                 _BS),
    FieldMap("goodwill",                    "BS_GOODWILL",                      _BS),
    FieldMap("intangibleAssets",            "BS_DISCLOSED_INTANGIBLES",         _BS, "VERIFY - excludes goodwill"),
    FieldMap("goodwillAndIntangibleAssets", None,                               _BS, "NO DIRECT EQUIVALENT - derive: goodwill + intangibleAssets"),
    FieldMap("longTermInvestments",         "BS_LT_INVEST",                     _BS, "VERIFY"),
    FieldMap("taxAssets",                   "BS_DEFERRED_TAX_ASSETS",           _BS, "VERIFY"),
    FieldMap("otherNonCurrentAssets",       "BS_OTHER_ASSETS_DEF_CHRG_OTHER",   _BS, "VERIFY"),
    FieldMap("totalNonCurrentAssets",       None,                               _BS, "NO DIRECT EQUIVALENT - derive: BS_TOT_ASSET - BS_CUR_ASSET_REPORT"),
    FieldMap("otherAssets",                 "BS_OTHER_ASSETS_DEF_CHRG_OTHER",   _BS, "VERIFY - overlaps otherNonCurrentAssets"),
    FieldMap("totalAssets",                 "BS_TOT_ASSET",                     _BS),

    # ---- Current liabilities ------------------------------------------------
    FieldMap("totalPayables",                  None,                        _BS, "NO DIRECT EQUIVALENT - derive: accountPayables + otherPayables"),
    FieldMap("accountPayables",                "BS_ACCT_PAYABLE",           _BS),
    FieldMap("otherPayables",                  None,                        _BS, "NO DIRECT EQUIVALENT - derive"),
    FieldMap("accruedExpenses",                "BS_ACCRUAL",                _BS, "VERIFY"),
    FieldMap("shortTermDebt",                  "BS_ST_BORROW",              _BS),
    FieldMap("capitalLeaseObligationsCurrent", None,                        _BS, "NO DIRECT EQUIVALENT - VERIFY, often inside BS_ST_BORROW"),
    FieldMap("taxPayables",                    "BS_INCOME_TAX_PAYABLE",     _BS, "VERIFY"),
    FieldMap("deferredRevenue",                "BS_DEFERRED_REVENUE",       _BS, "VERIFY"),
    FieldMap("otherCurrentLiabilities",        "BS_OTHER_CUR_LIAB",         _BS, "VERIFY"),
    FieldMap("totalCurrentLiabilities",        "BS_CUR_LIAB",               _BS),

    # ---- Non-current liabilities --------------------------------------------
    FieldMap("longTermDebt",                       "BS_LONG_TERM_BORROWINGS",       _BS, "VERIFY"),
    FieldMap("capitalLeaseObligationsNonCurrent",  None,                            _BS, "NO DIRECT EQUIVALENT - VERIFY"),
    FieldMap("deferredRevenueNonCurrent",          None,                            _BS, "NO DIRECT EQUIVALENT - VERIFY"),
    FieldMap("deferredTaxLiabilitiesNonCurrent",   "BS_DEFERRED_TAX_LIABILITIES",   _BS, "VERIFY"),
    FieldMap("otherNonCurrentLiabilities",         "OTHER_NONCUR_LIAB",             _BS, "VERIFY"),
    FieldMap("totalNonCurrentLiabilities",         "BS_TOT_NON_CUR_LIAB",           _BS, "VERIFY - else BS_TOT_LIAB2 - BS_CUR_LIAB"),
    FieldMap("otherLiabilities",                   None,                            _BS, "NO DIRECT EQUIVALENT - derive"),
    FieldMap("capitalLeaseObligations",            "BS_CAPITAL_LEASE_OBLIGATIONS",  _BS, "VERIFY"),
    FieldMap("totalLiabilities",                   "BS_TOT_LIAB2",                  _BS),

    # ---- Equity -------------------------------------------------------------
    FieldMap("treasuryStock",                            "BS_AMT_OF_TSY_STOCK",             _BS),
    FieldMap("preferredStock",                           "BS_PFD_EQY",                      _BS, "VERIFY"),
    FieldMap("commonStock",                              "BS_COMMON_STOCK",                 _BS, "VERIFY"),
    FieldMap("retainedEarnings",                         "BS_RETAIN_EARN",                  _BS, "VERIFY"),
    FieldMap("additionalPaidInCapital",                  "BS_ADD_PAID_IN_CAP",              _BS, "VERIFY"),
    FieldMap("accumulatedOtherComprehensiveIncomeLoss",  "BS_ACCUM_OTHER_COMPR_INC",        _BS, "VERIFY"),
    FieldMap("otherTotalStockholdersEquity",             None,                              _BS, "NO DIRECT EQUIVALENT - plug/derive"),
    FieldMap("totalStockholdersEquity",                  "TOT_COMMON_EQY",                  _BS, "common equity, excludes minority interest"),
    FieldMap("totalEquity",                              "TOTAL_EQUITY",                    _BS, "includes minority interest"),
    FieldMap("minorityInterest",                         "MINORITY_NONCONTROLLING_INTEREST", _BS, "VERIFY"),
    FieldMap("totalLiabilitiesAndTotalEquity",           "BS_TOT_LIAB_AND_EQY",             _BS),

    # ---- Derived / summary --------------------------------------------------
    FieldMap("totalInvestments", None,                      _BS, "NO DIRECT EQUIVALENT - derive: shortTermInvestments + longTermInvestments"),
    FieldMap("totalDebt",        "SHORT_AND_LONG_TERM_DEBT", _BS),
    FieldMap("netDebt",          "NET_DEBT",                 _BS),
]


# =============================================================================
# CASH FLOW STATEMENT
# =============================================================================
_CASH_FLOW_ROWS = [
    # ---- Identifiers & metadata --------------------------------------------
    FieldMap("isin",             "ID_ISIN",                    _CF),
    FieldMap("symbol",           "TICKER",                     _CF, "alt: PARSEKYABLE_DES / ID_EXCH_SYMBOL"),
    FieldMap("cik",              "CENTRAL__INDEX_KEY_NUMBER",  _CF),
    FieldMap("reportedCurrency", "EQY_FUND_CRNCY",             _CF),
    FieldMap("fillingDate",      "FUNDAMENTAL_PUBLIC_DATE",    _CF, "VERIFY"),
    FieldMap("acceptedDate",     "LATEST_ANNOUNCEMENT_DT",     _CF, "VERIFY"),
    FieldMap("fiscalYear",       "EQY_FUND_YEAR",              _CF, "alt: FISCAL__YEAR_PERIOD"),
    FieldMap("period",           "FUND_PER",                   _CF, "VERIFY - normally an override"),

    # ---- Operating activities -----------------------------------------------
    FieldMap("netIncome",                    "NET_INCOME",                    _CF, "alt: CF_NET_INC"),
    FieldMap("depreciationAndAmortization",  "CF_DEPR_AMORT",                 _CF),
    FieldMap("deferredIncomeTax",            "CF_DEF_INC_TAX",                _CF, "VERIFY"),
    FieldMap("stockBasedCompensation",       "CF_STOCK_BASED_COMPENSATION",   _CF, "VERIFY - alt: ARDR_STOCK_BASED_COMPENSATION"),
    FieldMap("changeInWorkingCapital",       "CF_CHNG_NON_CASH_WORK_CAP",     _CF, "VERIFY"),
    FieldMap("accountsReceivables",          "CF_ACCT_RCV_UNBILLED_REV",      _CF, "VERIFY - change in AR"),
    FieldMap("inventory",                    "CF_CHANGE_IN_INVENTORIES",      _CF, "VERIFY - change in inventory"),
    FieldMap("accountsPayables",             "CF_CHANGE_IN_ACCOUNTS_PAYABLE", _CF, "VERIFY - change in AP"),
    FieldMap("otherWorkingCapital",          None,                            _CF, "NO DIRECT EQUIVALENT - derive"),
    FieldMap("otherNonCashItems",            "CF_OTHER_NON_CASH_ADJUST",      _CF, "VERIFY"),
    FieldMap("netCashProvidedByOperatingActivities", "CF_CASH_FROM_OPER",     _CF),

    # ---- Investing activities -----------------------------------------------
    FieldMap("investmentsInPropertyPlantAndEquipment", "CF_CAP_EXPEND_PRPTY_ADD",  _CF, "VERIFY - alt: ACQUIS_OF_FIXED_PROD_ASSETS; sign differs"),
    FieldMap("acquisitionsNet",                        "CF_CASH_PAID_FOR_ACQUIS",  _CF, "VERIFY - alt: CF_NET_CASH_PAID_FOR_ACQUIS"),
    FieldMap("purchasesOfInvestments",                 "CF_INCR_INVEST",           _CF, "VERIFY"),
    FieldMap("salesMaturitiesOfInvestments",           "CF_DECR_INVEST",           _CF, "VERIFY"),
    FieldMap("otherInvestingActivities",               "CF_OTHER_INV_ACT",         _CF, "VERIFY"),
    FieldMap("netCashProvidedByInvestingActivities",   "CF_CASH_FROM_INV_ACT",     _CF),

    # ---- Financing activities: debt -----------------------------------------
    FieldMap("netDebtIssuance",         "CF_NET_CHNG_IN_DEBT",    _CF, "VERIFY"),
    FieldMap("longTermNetDebtIssuance", "CF_LT_DEBT_CASH_FLOW",   _CF, "VERIFY"),
    FieldMap("shortTermNetDebtIssuance", "CF_ST_DEBT_CASH_FLOW",  _CF, "VERIFY"),

    # ---- Financing activities: equity ---------------------------------------
    FieldMap("netStockIssuance",           "PROC_FR_REPURCH_EQTY_DETAILED", _CF, "VERIFY - net equity issuance/repurchase"),
    FieldMap("netCommonStockIssuance",     None,                            _CF, "NO DIRECT EQUIVALENT - derive: commonStockIssuance + commonStockRepurchased"),
    FieldMap("commonStockIssuance",        "PROCEEDS_FROM_ISSUANCE_OF_COMMON", _CF, "VERIFY"),
    FieldMap("commonStockRepurchased",     "PURCHASES_OF_COMMON_STOCK",     _CF, "VERIFY"),
    FieldMap("netPreferredStockIssuance",  None,                            _CF, "NO DIRECT EQUIVALENT - VERIFY"),

    # ---- Financing activities: dividends & other ----------------------------
    FieldMap("netDividendsPaid",       "CF_DVD_PAID",         _CF),
    FieldMap("commonDividendsPaid",    "CF_COMMON_DVD_PAID",  _CF, "VERIFY"),
    FieldMap("preferredDividendsPaid", "CF_PFD_DVD_PAID",     _CF, "VERIFY - alt: BS_PFD_DVD"),
    FieldMap("otherFinancingActivities", "CF_OTHER_FNC_ACT",  _CF, "VERIFY"),
    FieldMap("netCashProvidedByFinancingActivities", "CF_CASH_FROM_FNC_ACT", _CF),

    # ---- Reconciliation -----------------------------------------------------
    FieldMap("effectOfForexChangesOnCash", "CF_EFFECT_FOREIGN_EXCHANGES", _CF, "VERIFY"),
    FieldMap("netChangeInCash",            "CF_NET_CHNG_CASH",            _CF),
    FieldMap("cashAtEndOfPeriod",          "CF_CASH_AT_END_OF_PERIOD",    _CF, "VERIFY - alt: BS_CASH_NEAR_CASH_ITEM"),
    FieldMap("cashAtBeginningOfPeriod",    "CF_CASH_AT_BEG_OF_PERIOD",    _CF, "VERIFY - else prior-period cashAtEndOfPeriod"),

    # ---- Summary / supplemental ---------------------------------------------
    FieldMap("operatingCashFlow",  "CF_CASH_FROM_OPER",       _CF, "same as netCashProvidedByOperatingActivities"),
    FieldMap("capitalExpenditure", "CAPITAL_EXPEND",          _CF, "alt: CF_CAP_EXPEND_PRPTY_ADD; sign differs"),
    FieldMap("freeCashFlow",       "CF_FREE_CASH_FLOW",       _CF),
    FieldMap("incomeTaxesPaid",    "IS_CASH_PAID_FOR_INC_TAX", _CF, "VERIFY - supplemental disclosure"),
    FieldMap("interestPaid",       "IS_CASH_PAID_FOR_INT",    _CF, "VERIFY - supplemental disclosure"),
]


# =============================================================================
# EMPLOYEE COUNT
#
# Mirrors FMP's `employee-count` endpoint (which sources SEC filing cover
# pages). Bloomberg's equivalent is the single fundamental field
# NUM_OF_EMPLOYEES -- there is no per-filing record with form type / source,
# so most of FMP's metadata columns have no counterpart here.
#
# PERIODICITY: annual in practice. Headcount is not a required quarterly
# disclosure, so Bloomberg carries one point per fiscal year for the large
# majority of issuers. Requesting -1FQ/-2FQ generally returns nulls or the
# prior annual figure repeated. fetch_employee_history() below can attempt
# quarterly and fall back to annual.
# =============================================================================
_EMPLOYEE_COUNT_ROWS = [
    FieldMap("symbol",         "TICKER",                          _EC, "alt: PARSEKYABLE_DES / ID_EXCH_SYMBOL"),
    FieldMap("cik",            "CENTRAL__INDEX_KEY_NUMBER",       _EC),
    FieldMap("companyName",    "NAME",                            _EC, "alt: LONG_COMP_NAME"),
    FieldMap("employeeCount",  "NUM_OF_EMPLOYEES",                _EC, "VERIFY - annual; point-in-time at fiscal period end"),
    FieldMap("periodOfReport", "LATEST_PERIOD_END_DT_FULL_RECORD", _EC, "VERIFY - fiscal period the headcount describes"),
    FieldMap("filingDate",     "FUNDAMENTAL_PUBLIC_DATE",         _EC, "VERIFY"),
    FieldMap("acceptanceTime", "LATEST_ANNOUNCEMENT_DT",          _EC, "VERIFY - no true SEC accepted-timestamp equivalent"),
    FieldMap("formType",       None,                              _EC, "NO EQUIVALENT - Bloomberg does not expose the source form type"),
    FieldMap("source",         None,                              _EC, "NO EQUIVALENT - no filing-URL field"),
]

# Related per-employee fields, useful alongside a headcount series.
# Not part of FMP's endpoint, so kept out of FIELD_MAP.
EMPLOYEE_RELATED_FIELDS: dict[str, str] = {
    "SALES_PER_EMPLOYEE":       "revenue / headcount",           # VERIFY
    "NET_INC_PER_EMPLOYEE":     "net income / headcount",        # VERIFY
    "TOT_EMPLOYEE_COMPENSATION": "total employee comp expense",  # VERIFY
}


FIELD_MAP: tuple[FieldMap, ...] = tuple(
    _INCOME_STATEMENT_ROWS
    + _BALANCE_SHEET_ROWS
    + _CASH_FLOW_ROWS
    + _EMPLOYEE_COUNT_ROWS
)


# =============================================================================
# Fields with no clean Bloomberg mnemonic -- compute these yourself.
# =============================================================================
DERIVED_FIELDS = {
    # income statement
    "nonOperatingIncomeExcludingInterest":
        "totalOtherIncomeExpensesNet - (interestIncome - interestExpense)",
    "otherAdjustmentsToNetIncome":
        "netIncome - netIncomeFromContinuingOperations - netIncomeFromDiscontinuedOperations",
    "netIncomeDeductions":
        "preferred dividends + minority interest (see BS_PFD_DVD, MINORITY_NONCONTROLLING_INTEREST)",
    # balance sheet
    "otherReceivables":            "BS_ACCT_NOTE_RCV - accountsReceivables",
    "goodwillAndIntangibleAssets": "goodwill + intangibleAssets",
    "totalNonCurrentAssets":       "BS_TOT_ASSET - BS_CUR_ASSET_REPORT",
    "totalPayables":               "accountPayables + otherPayables",
    "otherPayables":               "totalPayables - accountPayables (if totalPayables available)",
    "otherLiabilities":            "BS_TOT_LIAB2 - (current + non-current itemised liabilities)",
    "otherTotalStockholdersEquity":
        "TOT_COMMON_EQY - (commonStock + additionalPaidInCapital + retainedEarnings "
        "+ accumulatedOtherComprehensiveIncomeLoss - treasuryStock)",
    "totalInvestments":            "shortTermInvestments + longTermInvestments",
    # cash flow
    "otherWorkingCapital":
        "changeInWorkingCapital - (accountsReceivables + inventory + accountsPayables)",
    "netCommonStockIssuance":      "commonStockIssuance + commonStockRepurchased",
    "netPreferredStockIssuance":   "netStockIssuance - netCommonStockIssuance",
}


# =============================================================================
# Helpers
# =============================================================================
def rows(datatable: Optional[str] = None) -> tuple[FieldMap, ...]:
    """All FieldMap rows, optionally filtered to one datatable."""
    if datatable is None:
        return FIELD_MAP
    _check(datatable)
    return tuple(r for r in FIELD_MAP if r.datatable == datatable)


def mapping(datatable: Optional[str] = None) -> dict[str, Optional[str]]:
    """{fmp_field: bbg_field} for one datatable (or all -- names collide across
    tables, so pass a datatable unless you know what you want)."""
    return {r.fmp_field: r.bbg_field for r in rows(datatable)}


def bbg_fields(datatable: Optional[str] = None, include_meta: bool = False) -> list[str]:
    """Bloomberg mnemonics ready to hand to BDP/BDH, de-duplicated, order
    preserved. Identifier/metadata fields are dropped unless include_meta."""
    meta = {"ID_ISIN", "TICKER", "CENTRAL__INDEX_KEY_NUMBER", "EQY_FUND_CRNCY",
            "FUNDAMENTAL_PUBLIC_DATE", "LATEST_ANNOUNCEMENT_DT",
            "EQY_FUND_YEAR", "FUND_PER"}
    out: list[str] = []
    for r in rows(datatable):
        if r.bbg_field is None:
            continue
        if not include_meta and r.bbg_field in meta:
            continue
        if r.bbg_field not in out:
            out.append(r.bbg_field)
    return out


def fmp_to_bbg(fmp_field: str, datatable: Optional[str] = None) -> Optional[str]:
    """Look up the Bloomberg mnemonic for an FMP field."""
    for r in rows(datatable):
        if r.fmp_field == fmp_field:
            return r.bbg_field
    raise KeyError(f"unknown FMP field: {fmp_field!r}")


def bbg_to_fmp(bbg_field: str, datatable: Optional[str] = None) -> list[str]:
    """Reverse lookup -- a mnemonic can back more than one FMP column."""
    return [r.fmp_field for r in rows(datatable) if r.bbg_field == bbg_field]


def unmapped(datatable: Optional[str] = None) -> list[str]:
    """FMP fields with no Bloomberg equivalent (see DERIVED_FIELDS)."""
    return [r.fmp_field for r in rows(datatable) if r.bbg_field is None]


def needs_verification(datatable: Optional[str] = None) -> list[FieldMap]:
    """Rows with a Bloomberg mnemonic that is a best guess and should be
    checked in FLDS<GO>. Excludes unmapped rows (bbg_field=None) even if
    their note happens to mention 'VERIFY' -- those belong in unmapped()."""
    return [r for r in rows(datatable)
            if r.bbg_field is not None and "VERIFY" in r.note.upper()]


def mapped(datatable: Optional[str] = None) -> list[FieldMap]:
    """Rows with a confirmed (non-VERIFY) Bloomberg mnemonic."""
    return [r for r in rows(datatable)
            if r.bbg_field is not None and "VERIFY" not in r.note.upper()]


def status(r: FieldMap) -> str:
    """One of 'mapped' / 'needs verification' / 'unmapped' for a row --
    the same three, mutually-exclusive buckets used by mapped()/
    needs_verification()/unmapped() and by the CLI and Excel export."""
    if r.bbg_field is None:
        return "unmapped"
    if "VERIFY" in r.note.upper():
        return "needs verification"
    return "mapped"


def _check(datatable: str) -> None:
    if datatable not in DATATABLES:
        raise ValueError(f"datatable must be one of {DATATABLES}, got {datatable!r}")


def export_csv(path: str = "fields.csv") -> str:
    """Write the three-column map (plus note) to CSV."""
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["fmp_field", "bbg_field", "datatable", "note"])
        for r in FIELD_MAP:
            w.writerow([r.fmp_field, r.bbg_field or "", r.datatable, r.note])
    return path


def export_txt(path: str = "fields_table.txt") -> str:
    """Write the three-column map as a fixed-width text table."""
    w1 = max(len(r.fmp_field) for r in FIELD_MAP) + 2
    w2 = max(len(r.bbg_field or "-") for r in FIELD_MAP) + 2
    w3 = max(len(r.datatable) for r in FIELD_MAP) + 2
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(f"{'fmp_field':<{w1}}{'bbg_field':<{w2}}{'datatable':<{w3}}note\n")
        fh.write("-" * (w1 + w2 + w3 + 40) + "\n")
        for r in FIELD_MAP:
            fh.write(f"{r.fmp_field:<{w1}}{(r.bbg_field or '-'):<{w2}}{r.datatable:<{w3}}{r.note}\n")
    return path


# ---------------------------------------------------------------------------
# Excel export
#
# NOTE: _unique_path / _write_sheet are duplicated in daily_data.py rather than
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


# Row fill colors keyed by the 'status' column value -- mirrors the
# mapped/needs verification/unmapped buckets used everywhere else in this
# module (see status()).
_STATUS_FILL_COLORS = {
    "unmapped":            "FFC7CE",  # red   -- no Bloomberg equivalent
    "needs verification":  "FFEB9C",  # amber -- confirm in FLDS<GO>
    "mapped":              "C6EFCE",  # green -- confirmed mnemonic
}


def _write_sheet(writer, df, sheet_name: str, index: bool = False,
                  status_col: Optional[str] = None) -> None:
    """Write one DataFrame to a sheet with frozen header, autofilter and
    reasonable column widths. Strips tz-awareness, which Excel cannot store.
    If status_col names a column of 'mapped' / 'needs verification' /
    'unmapped' values, each data row is shaded accordingly."""
    import pandas as pd

    df = df.copy()
    for col in df.columns:
        s = df[col]
        if isinstance(s.dtype, pd.DatetimeTZDtype):
            df[col] = s.dt.tz_localize(None)
    if index and isinstance(df.index, pd.DatetimeIndex) and df.index.tz is not None:
        df.index = df.index.tz_localize(None)

    sheet_name = sheet_name[:31]  # Excel hard limit
    df.to_excel(writer, sheet_name=sheet_name, index=index)

    ws = writer.sheets[sheet_name]
    ws.freeze_panes = "A2"
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
            ws.column_dimensions[get_column_letter(i)].width = min(width, 60)

    if status_col and status_col in df.columns and n_cols:
        from openpyxl.styles import PatternFill

        fills = {k: PatternFill(start_color=c, end_color=c, fill_type="solid")
                  for k, c in _STATUS_FILL_COLORS.items()}
        for row_idx, value in enumerate(df[status_col].tolist(), start=2):
            fill = fills.get(value)
            if fill is None:
                continue
            for col_idx in range(1, n_cols + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill


def export_excel(
    path: str = "fmp_bloomberg_field_map.xlsx",
    *,
    overwrite: bool = False,
    employee_history: "Optional[pd.DataFrame]" = None,
    ticker: str = "",
    ticker_data: "Optional[dict[str, pd.DataFrame]]" = None,
    include_field_map: bool = True,
) -> str:
    """
    Write the field map -- and, if ticker_data is given, live pulled figures
    for one security -- to a new Excel workbook.

    Sheets
    ------
    summary                per-datatable counts (mapped / needs verification /
                            unmapped) -- only if include_field_map
    all_fields              every row, all datatables, color-coded status --
                            only if include_field_map
    <datatable>             one sheet per datatable, same status coding --
                            only if include_field_map
    derived_fields          fields with no mnemonic + how to compute them --
                            only if include_field_map
    needs_verification      rows to confirm in FLDS<GO> -- only if include_field_map
    unmapped                rows with no Bloomberg equivalent -- only if include_field_map
    employee_history        only if a fetch_employee_history() frame is passed
    <datatable>             (or <datatable>_data if include_field_map, to
                            avoid a name clash) -- only if ticker_data is
                            given: the actual pulled figures for `ticker`,
                            one column per fiscal period, unmapped rows
                            filled with the literal 'UNMAPPED'

    Every sheet with a 'status' column is row-shaded: green = mapped, amber =
    needs verification, red = unmapped.

    include_field_map=False skips every static-mapping sheet and writes only
    the ticker_data sheets -- use this for a data-only export.

    Returns the path actually written (auto-suffixed unless overwrite=True).
    """
    import pandas as pd

    if not overwrite:
        path = _unique_path(path)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        if include_field_map:
            def frame(datatable=None) -> "pd.DataFrame":
                return pd.DataFrame(
                    [
                        {
                            "fmp_field": r.fmp_field,
                            "bbg_field": r.bbg_field if r.bbg_field else "",
                            "datatable": r.datatable,
                            "note": r.note,
                            "status": status(r),
                        }
                        for r in rows(datatable)
                    ]
                )

            summary = pd.DataFrame(
                [
                    {
                        "datatable": dt,
                        "fields": len(rows(dt)),
                        "mapped": len(mapped(dt)),
                        "needs_verification": len(needs_verification(dt)),
                        "unmapped": len(unmapped(dt)),
                    }
                    for dt in DATATABLES
                ]
                + [
                    {
                        "datatable": "TOTAL",
                        "fields": len(FIELD_MAP),
                        "mapped": len(mapped()),
                        "needs_verification": len(needs_verification()),
                        "unmapped": len(unmapped()),
                    }
                ]
            )
            if ticker_data:
                summary_rows = summary.to_dict("records")
                # Blank spacer keeps this visually separate from the
                # per-datatable counts above it.
                summary_rows += [{"datatable": ""},
                                 {"datatable": "ticker", "fields": ticker}]
                summary = pd.DataFrame(summary_rows)

            derived = pd.DataFrame(
                [{"fmp_field": k, "derivation": v} for k, v in DERIVED_FIELDS.items()]
            )
            all_fields = frame()
            verify = all_fields[all_fields["status"] == "needs verification"].reset_index(drop=True)
            unmapped_sheet = all_fields[all_fields["status"] == "unmapped"].reset_index(drop=True)

            _write_sheet(writer, summary, "summary")
            _write_sheet(writer, all_fields, "all_fields", status_col="status")
            for dt in DATATABLES:
                _write_sheet(writer, frame(dt), dt, status_col="status")
            _write_sheet(writer, derived, "derived_fields")
            _write_sheet(writer, verify, "needs_verification", status_col="status")
            _write_sheet(writer, unmapped_sheet, "unmapped", status_col="status")

        if employee_history is not None and not employee_history.empty:
            _write_sheet(writer, employee_history, "employee_history")

        if ticker_data:
            for dt in DATATABLES:
                df = ticker_data.get(dt)
                if df is not None and not df.empty:
                    name = f"{dt}_data" if include_field_map else dt
                    _write_sheet(writer, df, name[:31], status_col="status")

    return path


# =============================================================================
# Live data: employee-count history
#
# The only fetch helper in this module. Everything above is static mapping.
# blpapi/pandas are imported lazily so this file stays importable without a
# Terminal; the session plumbing is reused from daily_data.py.
# =============================================================================
EMPLOYEE_FIELD = "NUM_OF_EMPLOYEES"
PERIOD_END_FIELD = "LATEST_PERIOD_END_DT_FULL_RECORD"


def fetch_employee_history(
    ticker: str,
    periods_back: int = 20,
    *,
    period: str = "FY",
    filing_status: str = "MR",
    fallback_to_annual: bool = True,
    session: object = None,
) -> "pd.DataFrame":
    """
    Historical headcount from NUM_OF_EMPLOYEES, one row per fiscal period.

    Walks EQY_FUND_RELATIVE_PERIOD backwards ('-0FY', '-1FY', ...) with one
    ReferenceDataRequest per period. That costs N requests instead of one, but
    unlike a HistoricalDataRequest it returns the fiscal period each figure
    actually describes rather than the date Bloomberg loaded the filing -- so
    the series aligns to fiscal years, not to load dates.

    Parameters
    ----------
    ticker          'AAPL' or 'AAPL US Equity'
    periods_back    how many fiscal periods to walk back
    period          'FY' annual (default) or 'FQ' quarterly. Quarterly
                    headcount is rarely disclosed -- expect mostly nulls.
    filing_status   'MR' most-recent/restated, 'OR' as-originally-reported.
                    Headcount gets restated more often than you'd expect;
                    use 'OR' for point-in-time backtests.
    fallback_to_annual
                    if period='FQ' yields nothing, retry annually.
    session         an open daily_data.BloombergSession to reuse. If None, one
                    is opened and closed for the call.

    Returns a DataFrame [ticker, relative_period, period_end, employees],
    sorted by period_end, or an empty frame if nothing came back.
    """
    import pandas as pd  # lazy: keeps this module stdlib-only on import

    from daily_data import BloombergSession, _normalise_ticker

    period = period.upper()
    if period not in ("FY", "FQ"):
        raise ValueError(f"period must be 'FY' or 'FQ', got {period!r}")

    def _pull(bb, per: str) -> list[dict]:
        security = _normalise_ticker(ticker)
        out: list[dict] = []
        for i in range(periods_back):
            req = bb.service.createRequest("ReferenceDataRequest")
            req.getElement("securities").appendValue(security)
            for f in (EMPLOYEE_FIELD, PERIOD_END_FIELD):
                req.getElement("fields").appendValue(f)

            ov = req.getElement("overrides")
            for fid, val in (("EQY_FUND_RELATIVE_PERIOD", f"-{i}{per}"),
                             ("FILING_STATUS", filing_status)):
                e = ov.appendElement()
                e.setElement("fieldId", fid)
                e.setElement("value", val)

            for msg in bb.send(req):
                if not msg.hasElement("securityData"):
                    continue
                arr = msg.getElement("securityData")
                for s in range(arr.numValues()):
                    sd = arr.getValueAsElement(s)
                    if sd.hasElement("securityError"):
                        continue
                    fd = sd.getElement("fieldData")
                    if not fd.hasElement(EMPLOYEE_FIELD):
                        continue
                    if fd.getElement(EMPLOYEE_FIELD).isNull():
                        continue
                    period_end = None
                    if fd.hasElement(PERIOD_END_FIELD) and not fd.getElement(PERIOD_END_FIELD).isNull():
                        period_end = fd.getElementAsDatetime(PERIOD_END_FIELD)
                    out.append({
                        "ticker": security,
                        "relative_period": f"-{i}{per}",
                        "period_end": period_end,
                        "employees": fd.getElementAsFloat(EMPLOYEE_FIELD),
                    })
        return out

    def _run(bb) -> list[dict]:
        recs = _pull(bb, period)
        if not recs and period == "FQ" and fallback_to_annual:
            recs = _pull(bb, "FY")
        return recs

    if session is not None:
        records = _run(session)
    else:
        with BloombergSession() as bb:
            records = _run(bb)

    if not records:
        return pd.DataFrame(
            columns=["ticker", "relative_period", "period_end", "employees"]
        )

    df = pd.DataFrame.from_records(records)
    df["period_end"] = pd.to_datetime(df["period_end"], errors="coerce")

    # Bloomberg repeats the prior annual figure across quarters for issuers
    # that don't disclose quarterly -- collapse identical period_end rows.
    df = df.drop_duplicates(subset=["period_end", "employees"])
    return df.sort_values("period_end").reset_index(drop=True)


# =============================================================================
# Live data: full field-map pull for one security
#
# Generic counterpart to fetch_employee_history() above -- works for any of
# the four datatables by walking every mapped mnemonic in FIELD_MAP through
# one ReferenceDataRequest per fiscal period.
# =============================================================================
#  Safety backstop for periods_back=None (auto) -- stops a runaway loop if
#  Bloomberg never signals "no more history" for some reason. 120 annual
#  periods / 480 quarters covers any listed company with room to spare.
_MAX_AUTO_PERIODS = 120

#  Consecutive periods with no period-end date before auto-mode concludes
#  history has run out. Mirrors rev_data4.py's MAX_EMPTY_WINDOWS.
_MAX_EMPTY_PERIODS = 2


def fetch_financials(
    ticker: str,
    datatable: str,
    periods_back: "Optional[int]" = None,
    *,
    period: str = "FY",
    filing_status: str = "MR",
    session: object = None,
) -> "pd.DataFrame":
    """
    Pull live figures for one datatable and lay them out next to the static
    field map: one row per FMP field (same order as rows(datatable)), one
    column per fiscal period (most recent first).

    periods_back=None (default) walks backward one fiscal period at a time
    and stops once Bloomberg has returned no period-end date for
    _MAX_EMPTY_PERIODS consecutive periods -- i.e. it pulls as far back as
    the security actually has history, the same "walk until it goes dry"
    approach rev_data4.py uses for segment revenue. Pass an int to fetch
    exactly that many periods instead.

    Rows with bbg_field=None (see unmapped()) have no mnemonic to request --
    every period column reads the literal 'UNMAPPED' for those rows rather
    than being left blank, so a missing mapping is never confused with a
    field that Bloomberg simply returned null for.

    Column headers are the actual fiscal year for that period (e.g.
    'FY2024'), derived from the fiscal period-end date Bloomberg returns for
    that EQY_FUND_RELATIVE_PERIOD override -- not the raw override string.
    EQY_FUND_YEAR itself does not reliably echo back the resolved year when
    combined with EQY_FUND_RELATIVE_PERIOD, so LATEST_PERIOD_END_DT_FULL_RECORD
    (proven to vary correctly per period -- see fetch_employee_history) is
    requested alongside the mapped fields purely to label columns, doubles
    as the "does this period exist at all" signal for auto mode, and is used
    as the primary label source; EQY_FUND_YEAR is a fallback, and the raw
    override (e.g. '-0FY') is the last resort if neither comes back.

    period          'FY' annual (default) or 'FQ' quarterly.
    filing_status   'MR' most-recent/restated (default) or 'OR'
                    as-originally-reported. See FILING_STATUS in FLDS<GO>.
    session         an open daily_data.BloombergSession to reuse across
                    datatables. If None, one is opened and closed for the call.
    """
    import pandas as pd

    from daily_data import BloombergSession, _element_to_py, _normalise_ticker

    _check(datatable)
    period = period.upper()
    if period not in ("FY", "FQ"):
        raise ValueError(f"period must be 'FY' or 'FQ', got {period!r}")

    security = _normalise_ticker(ticker)
    dt_rows = rows(datatable)
    mnemonics = bbg_fields(datatable, include_meta=True)
    # PERIOD_END_FIELD is requested purely for column labeling (and, in auto
    # mode, to detect when history has run out) -- it is dropped from the
    # output unless the datatable already maps it itself (employee_count's
    # periodOfReport).
    request_fields = list(mnemonics)
    if PERIOD_END_FIELD not in request_fields:
        request_fields.append(PERIOD_END_FIELD)

    def _pull_period(bb, rel_period: str) -> dict:
        req = bb.service.createRequest("ReferenceDataRequest")
        req.getElement("securities").appendValue(security)
        for f in request_fields:
            req.getElement("fields").appendValue(f)

        ov = req.getElement("overrides")
        for fid, val in (("EQY_FUND_RELATIVE_PERIOD", rel_period),
                         ("FILING_STATUS", filing_status)):
            e = ov.appendElement()
            e.setElement("fieldId", fid)
            e.setElement("value", val)

        values: dict = {}
        for msg in bb.send(req):
            if not msg.hasElement("securityData"):
                continue
            arr = msg.getElement("securityData")
            for s in range(arr.numValues()):
                sd = arr.getValueAsElement(s)
                if sd.hasElement("securityError"):
                    continue
                fd = sd.getElement("fieldData")
                for m in request_fields:
                    if fd.hasElement(m) and not fd.getElement(m).isNull():
                        values[m] = _element_to_py(fd.getElement(m))
        return values

    def _run(bb) -> dict:
        per_period: dict = {}
        order: list = []
        empty_streak = 0
        i = 0
        while True:
            if periods_back is not None:
                if i >= periods_back:
                    break
            elif i >= _MAX_AUTO_PERIODS:
                break

            rel = f"-{i}{period}"
            vals = _pull_period(bb, rel)
            per_period[rel] = vals
            order.append(rel)

            if vals.get(PERIOD_END_FIELD) is None:
                empty_streak += 1
            else:
                empty_streak = 0
            i += 1

            if periods_back is None and empty_streak >= _MAX_EMPTY_PERIODS:
                break

        if periods_back is None and empty_streak:
            # Drop the trailing empty periods that triggered the stop --
            # they carry no data and would just be blank columns.
            for rel in order[-empty_streak:]:
                del per_period[rel]

        return per_period

    if session is not None:
        per_period = _run(session)
    else:
        with BloombergSession() as bb:
            per_period = _run(bb)

    def _year_of(value) -> Optional[str]:
        if value is None or value == "":
            return None
        if hasattr(value, "year"):
            return str(value.year)
        match = re.match(r"(\d{4})", str(value))
        return match.group(1) if match else None

    def _label(rel: str) -> str:
        vals = per_period[rel]
        year = _year_of(vals.get(PERIOD_END_FIELD)) or _year_of(vals.get("EQY_FUND_YEAR"))
        return f"FY{year}" if year else rel

    # Disambiguate in the (rare) case two periods land on the same label,
    # e.g. two relative periods both come back with no EQY_FUND_YEAR.
    seen: dict[str, int] = {}
    labels: dict[str, str] = {}
    for rel in per_period:
        lbl = _label(rel)
        seen[lbl] = seen.get(lbl, 0) + 1
        labels[rel] = lbl if seen[lbl] == 1 else f"{lbl} ({seen[lbl]})"

    out_rows = []
    for r in dt_rows:
        row = {"fmp_field": r.fmp_field, "bbg_field": r.bbg_field or "",
               "status": status(r)}
        for rel, label in labels.items():
            row[label] = ("UNMAPPED" if r.bbg_field is None
                          else per_period[rel].get(r.bbg_field))
        out_rows.append(row)

    return pd.DataFrame(out_rows)


def fetch_all_financials(
    ticker: str,
    periods_back: "Optional[int]" = None,
    *,
    period: str = "FY",
    filing_status: str = "MR",
) -> "dict[str, pd.DataFrame]":
    """fetch_financials() for every datatable, reusing one Bloomberg session
    so logging in only happens once. Returns {datatable: DataFrame}.
    periods_back=None (default) pulls as far back as each datatable has
    history -- see fetch_financials()."""
    from daily_data import BloombergSession

    with BloombergSession() as bb:
        return {
            dt: fetch_financials(ticker, dt, periods_back, period=period,
                                 filing_status=filing_status, session=bb)
            for dt in DATATABLES
        }


def _safe_name(ticker: str) -> str:
    """'PLTR US Equity' -> 'PLTR_US_Equity', usable in a filename."""
    keep = [c if (c.isalnum() or c in "-_") else "_" for c in ticker.strip()]
    return "".join(keep).strip("_") or "output"


# =============================================================================
# CLI
# =============================================================================
def main(argv: "Optional[list[str]]" = None) -> int:
    import argparse

    p = argparse.ArgumentParser(
        description="FMP -> Bloomberg field map. Prints a summary, or exports "
                    "the map to Excel/CSV/text."
    )
    p.add_argument("--excel", action="store_true",
                   help="write the field map to a new Excel workbook")
    p.add_argument("--out", default=None,
                   help="workbook path (default: fmp_bloomberg_field_map.xlsx, "
                        "or <TICKER>_financial_statement_data.xlsx if --ticker "
                        "is given)")
    p.add_argument("--overwrite", action="store_true",
                   help="overwrite --out instead of auto-suffixing _1, _2, ...")
    p.add_argument("--csv", metavar="PATH", nargs="?", const="fields.csv",
                   help="also write a flat CSV")
    p.add_argument("--txt", metavar="PATH", nargs="?", const="fields_table.txt",
                   help="also write a fixed-width text table")
    p.add_argument("--employees", metavar="TICKER", default=None,
                   help="also pull headcount history for TICKER and add it as "
                        "a sheet (requires a running Terminal)")
    p.add_argument("--employee-periods", type=int, default=20,
                   help="fiscal periods back for --employees (default: %(default)s)")
    p.add_argument("--ticker", metavar="TICKER", default=None,
                   help="pull live figures for TICKER (e.g. 'PLTR US Equity' "
                        "or just 'PLTR') for every datatable and write them "
                        "to a data-only Excel workbook -- no field-map sheets, "
                        "no console field-map dump (requires a running "
                        "Terminal). Implies --excel.")
    p.add_argument("--periods", type=int, default=None,
                   help="fiscal periods back to pull for --ticker "
                        "(default: as far back as Bloomberg has history)")
    p.add_argument("--period", choices=["FY", "FQ"], default="FY",
                   help="annual or quarterly periods for --ticker "
                        "(default: %(default)s)")
    p.add_argument("--filing-status", default="MR",
                   help="'MR' most-recent/restated or 'OR' "
                        "as-originally-reported, for --ticker "
                        "(default: %(default)s)")
    p.add_argument("--summary", action="store_true",
                   help="print terse per-datatable counts only, without "
                        "listing individual fields")
    args = p.parse_args(argv)

    # --ticker means "just get me the data" -- skip the static field-map
    # report entirely rather than printing 152 rows of mapping detail first.
    if not args.ticker:
        if args.summary:
            for dt in DATATABLES:
                print(f"{dt:<22} {len(rows(dt)):>3} fields | "
                      f"{len(unmapped(dt)):>2} unmapped | "
                      f"{len(needs_verification(dt)):>2} need FLDS<GO> verification")
            print(f"{'TOTAL':<22} {len(FIELD_MAP):>3} fields")
        else:
            name_w = max(len(r.fmp_field) for r in FIELD_MAP) + 2
            for dt in DATATABLES:
                dt_mapped = mapped(dt)
                dt_verify = needs_verification(dt)
                dt_unmapped_rows = [r for r in rows(dt) if r.bbg_field is None]
                print(f"\n=== {dt} -- {len(rows(dt))} fields "
                      f"({len(dt_mapped)} mapped, {len(dt_verify)} need verification, "
                      f"{len(dt_unmapped_rows)} unmapped) ===")

                print(f"  MAPPED ({len(dt_mapped)}):")
                for r in dt_mapped:
                    print(f"    {r.fmp_field:<{name_w}} -> {r.bbg_field}")

                print(f"  NEEDS VERIFICATION -- confirm in FLDS<GO> ({len(dt_verify)}):")
                for r in dt_verify:
                    note = f"  [{r.note}]" if r.note else ""
                    print(f"    {r.fmp_field:<{name_w}} -> {r.bbg_field}{note}")

                print(f"  UNMAPPED -- no Bloomberg equivalent ({len(dt_unmapped_rows)}):")
                for r in dt_unmapped_rows:
                    note = f"  [{r.note}]" if r.note else ""
                    print(f"    {r.fmp_field:<{name_w}}{note}")

            print(f"\nTOTAL: {len(FIELD_MAP)} fields | "
                  f"{len(mapped())} mapped | "
                  f"{len(needs_verification())} need verification | "
                  f"{len(unmapped())} unmapped")

    emp = None
    if args.employees:
        print(f"\nFetching headcount history for {args.employees} ...")
        emp = fetch_employee_history(args.employees,
                                     periods_back=args.employee_periods)
        print(f"  {len(emp)} period(s) returned")

    ticker_norm = None
    ticker_data = None
    if args.ticker:
        from daily_data import _normalise_ticker

        ticker_norm = _normalise_ticker(args.ticker)
        periods_desc = (f"{args.periods} period(s) back" if args.periods is not None
                        else "as far back as available")
        print(f"\nFetching {args.period} fundamentals for {ticker_norm} "
              f"({periods_desc}, {args.filing_status}) ...")
        ticker_data = fetch_all_financials(
            ticker_norm, periods_back=args.periods,
            period=args.period, filing_status=args.filing_status,
        )
        for dt, df in ticker_data.items():
            n_periods = max(len(df.columns) - 3, 0)  # minus fmp_field/bbg_field/status
            print(f"  {dt:<22} {len(df):>3} fields x {n_periods} period(s)")

    if args.excel or args.employees or args.ticker:
        out = args.out
        if out is None:
            out = (f"{_safe_name(ticker_norm)}_financial_statement_data.xlsx"
                   if ticker_norm else "fmp_bloomberg_field_map.xlsx")
        written = export_excel(out, overwrite=args.overwrite,
                               employee_history=emp,
                               ticker=ticker_norm or "",
                               ticker_data=ticker_data,
                               include_field_map=not args.ticker)
        print(f"\nWrote workbook -> {written}")
    if args.csv:
        print(f"Wrote CSV      -> {export_csv(args.csv)}")
    if args.txt:
        print(f"Wrote text     -> {export_txt(args.txt)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
