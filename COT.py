"""
Simple field prober: test a list of candidate field mnemonics against one
ticker and report which ones actually return data, plus the real Bloomberg
error for the ones that don't. Edit TICKER and CANDIDATE_FIELDS below and
rerun as many times as needed while we hunt for the right field(s).
"""

from contextlib import contextmanager

import blpapi

TICKER = "QNAA Index"

CANDIDATE_FIELDS = [
    "COT_TRADER_TYPE",
    "CFTC_ALL",
    "CFTC_ALL_DISAGG",
    "CFTC_POSITIONS",
    "CFTC_LONG_ALL",
    "CFTC_SHORT_ALL",
    "FUT_CFTC_LONG_ALL",
    "FUT_CFTC_SHORT_ALL",
    "FUT_CFTC_OPEN_INT_ALL",
]


@contextmanager
def bloomberg_session(host="localhost", port=8194):
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


def probe_fields(session, ticker, fields):
    refdata = session.getService("//blp/refdata")
    request = refdata.createRequest("ReferenceDataRequest")
    request.getElement("securities").appendValue(ticker)
    for f in fields:
        request.getElement("fields").appendValue(f)

    session.sendRequest(request)

    while True:
        event = session.nextEvent(5000)
        for msg in event:
            if not msg.hasElement("securityData"):
                continue

            sec_data_array = msg.getElement("securityData")
            for i in range(sec_data_array.numValues()):
                sec_data = sec_data_array.getValueAsElement(i)

                if sec_data.hasElement("securityError"):
                    print("SECURITY ERROR:", sec_data.getElement("securityError"))
                    continue

                # Per-field failures live here, as siblings of fieldData -
                # not inside it. This is what our first attempt missed.
                errors_by_field = {}
                if sec_data.hasElement("fieldExceptions"):
                    exceptions = sec_data.getElement("fieldExceptions")
                    for j in range(exceptions.numValues()):
                        exc = exceptions.getValueAsElement(j)
                        field_id = exc.getElementAsString("fieldId")
                        error_info = exc.getElement("errorInfo")
                        errors_by_field[field_id] = error_info.getElementAsString("message")

                field_data = sec_data.getElement("fieldData")

                for f in fields:
                    if field_data.hasElement(f):
                        elem = field_data.getElement(f)
                        n = elem.numValues() if elem.isArray() else 1
                        print(f"[OK]    {f}  ({n} value(s))")
                        print(f"        {elem}".rstrip())
                    elif f in errors_by_field:
                        print(f"[ERROR] {f}: {errors_by_field[f]}")
                    else:
                        print(f"[EMPTY] {f}: no error, but no data either")

        if event.eventType() == blpapi.Event.RESPONSE:
            break


def main():
    print(f"Probing {len(CANDIDATE_FIELDS)} candidate fields on {TICKER}\n")
    with bloomberg_session() as session:
        probe_fields(session, TICKER, CANDIDATE_FIELDS)


if __name__ == "__main__":
    main()
