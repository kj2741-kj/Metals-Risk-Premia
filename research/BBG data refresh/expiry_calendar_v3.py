"""
Futures Expiry Calendar Builder (Interactive)
==============================================
Step 1: Pick asset class (Energy, NGLs, Metals, Equity, Fixed Income, FX, Ags)
Step 2: Pick commodities within that class
Step 3: Pick start year
Step 4: Pulls from Bloomberg, saves to one Excel file with separate sheets

Run in PyCharm: Shift+F10
"""

import blpapi
import pandas as pd
from datetime import datetime
import os
import sys
import time

OUTPUT_DIR = r"C:\Users\Bloomberg User\research_data"
os.makedirs(OUTPUT_DIR, exist_ok=True)

MONTH_CODES = {
    1: "F", 2: "G", 3: "H", 4: "J", 5: "K", 6: "M",
    7: "N", 8: "Q", 9: "U", 10: "V", 11: "X", 12: "Z"
}

MONTH_NAMES = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}

EXPIRY_FIELDS = [
    "LAST_TRADEABLE_DT",
    "FUT_NOTICE_FIRST",
    "FUT_FIRST_TRADE_DT",
    "FUT_DLV_DT_FIRST",
    "FUT_DLV_DT_LAST",
]

# ═══════════════════════════════════════════════
# ASSET CLASS → COMMODITY CATALOG
# ═══════════════════════════════════════════════

ASSET_CLASSES = {
    "1": {
        "name": "Energy",
        "commodities": {
            "CL":  "WTI Crude Oil (NYMEX)",
            "CO":  "Brent Crude Oil (ICE)",
            "XB":  "RBOB Gasoline (NYMEX)",
            "HO":  "ULSD / Heating Oil (NYMEX)",
            "QS":  "Gasoil (ICE)",
            "NG":  "Natural Gas Henry Hub (NYMEX)",
        }
    },
    "2": {
        "name": "NGLs & Petrochemicals",
        "commodities": {
            "CAP": "Ethane — Mt Belvieu Swap (NYMEX)",
            "BAP": "Propane — Mt Belvieu Swap (NYMEX)",
            "DAE": "Butane — Mt Belvieu Swap (NYMEX)",
            "IBD": "Isobutane — Mt Belvieu Swap Future (NYMEX)",
            "PCW": "Ethylene — Mt Belvieu Futures (NYMEX)",
            "PGP": "Propylene — Polymer Grade Futures (NYMEX)",
            "NMB": "Natural Gasoline C5+ (NYMEX)",
        }
    },
    "3": {
        "name": "Base Metals",
        "commodities": {
            "LP":  "Copper (LME)",
            "LA":  "Aluminium (LME)",
            "LX":  "Zinc (LME)",
            "LN":  "Nickel (LME)",
            "LL":  "Lead (LME)",
            "LT":  "Tin (LME)",
        }
    },
    "4": {
        "name": "Precious Metals",
        "commodities": {
            "GC":  "Gold (COMEX)",
            "SI":  "Silver (COMEX)",
            "PL":  "Platinum (NYMEX)",
            "PA":  "Palladium (NYMEX)",
        }
    },
    "5": {
        "name": "Equity Index Futures",
        "commodities": {
            "ES":  "S&P 500 E-mini (CME)",
            "NQ":  "Nasdaq 100 E-mini (CME)",
            "YM":  "Dow Jones E-mini (CBOT)",
            "RTY": "Russell 2000 E-mini (CME)",
            "VG":  "Euro Stoxx 50 (EUREX)",
            "Z ":  "FTSE 100 (ICE)",
            "NK":  "Nikkei 225 (CME)",
            "HI":  "Hang Seng (HKFE)",
            "IF":  "CSI 300 (CFFEX)",
        }
    },
    "6": {
        "name": "Fixed Income / Rates",
        "commodities": {
            "TY":  "10-Year US Treasury Note (CBOT)",
            "FV":  "5-Year US Treasury Note (CBOT)",
            "TU":  "2-Year US Treasury Note (CBOT)",
            "US":  "US Treasury Bond 30Y (CBOT)",
            "WN":  "Ultra Bond (CBOT)",
            "RX":  "Euro Bund 10Y (EUREX)",
            "OE":  "Euro Bobl 5Y (EUREX)",
            "DU":  "Euro Schatz 2Y (EUREX)",
            "G ":  "UK Long Gilt (ICE)",
            "JB":  "Japan 10Y JGB (OSE)",
            "ED":  "Eurodollar 3M (CME)",
            "SR":  "SOFR 3M (CME)",
        }
    },
    "7": {
        "name": "FX / Currency Futures",
        "commodities": {
            "EC":  "EUR/USD (CME)",
            "JY":  "JPY/USD (CME)",
            "BP":  "GBP/USD (CME)",
            "SF":  "CHF/USD (CME)",
            "AD":  "AUD/USD (CME)",
            "CD":  "CAD/USD (CME)",
            "NE":  "NZD/USD (CME)",
            "MP":  "MXN/USD (CME)",
            "DX":  "US Dollar Index (ICE)",
        }
    },
    "8": {
        "name": "Agriculture & Softs",
        "commodities": {
            "C ":  "Corn (CBOT)",
            "S ":  "Soybeans (CBOT)",
            "W ":  "Wheat (CBOT)",
            "SM":  "Soybean Meal (CBOT)",
            "BO":  "Soybean Oil (CBOT)",
            "CT":  "Cotton (ICE)",
            "KC":  "Coffee Arabica (ICE)",
            "SB":  "Sugar #11 (ICE)",
            "CC":  "Cocoa (ICE)",
            "LC":  "Live Cattle (CME)",
            "LH":  "Lean Hogs (CME)",
        }
    },
}


# ═══════════════════════════════════════════════
# BLOOMBERG SESSION
# ═══════════════════════════════════════════════

class BloombergSession:
    def __init__(self):
        self.session = None
        self.svc = None

    def start(self):
        opts = blpapi.SessionOptions()
        opts.setServerHost("localhost")
        opts.setServerPort(8194)
        self.session = blpapi.Session(opts)
        if not self.session.start():
            print("\n  ✗ Cannot connect. Is Bloomberg Terminal running?")
            sys.exit(1)
        if not self.session.openService("//blp/refdata"):
            print("  ✗ Cannot open refdata service.")
            self.session.stop()
            sys.exit(1)
        self.svc = self.session.getService("//blp/refdata")
        print("  ✓ Connected to Bloomberg\n")

    def stop(self):
        if self.session:
            self.session.stop()

    def bdp_batch(self, tickers, fields):
        all_results = {}
        chunk_size = 40
        for i in range(0, len(tickers), chunk_size):
            chunk = tickers[i:i + chunk_size]
            try:
                req = self.svc.createRequest("ReferenceDataRequest")
                for t in chunk:
                    req.getElement("securities").appendValue(t)
                for f in fields:
                    req.getElement("fields").appendValue(f)
                self.session.sendRequest(req)

                while True:
                    event = self.session.nextEvent(timeout=30000)
                    for msg in event:
                        if msg.hasElement("securityData"):
                            arr = msg.getElement("securityData")
                            for j in range(arr.numValues()):
                                sec = arr.getValueAsElement(j)
                                ticker = sec.getElementAsString("security")
                                if sec.hasElement("securityError"):
                                    all_results[ticker] = {f: None for f in fields}
                                    continue
                                fd = sec.getElement("fieldData")
                                row = {}
                                for f in fields:
                                    try:
                                        el = fd.getElement(f)
                                        if el.datatype() == blpapi.DataType.DATE:
                                            dt = el.getValueAsDatetime()
                                            row[f] = datetime(dt.year, dt.month, dt.day).strftime("%Y-%m-%d")
                                        else:
                                            row[f] = el.getValueAsString()
                                    except Exception:
                                        row[f] = None
                                all_results[ticker] = row
                    if event.eventType() == blpapi.Event.RESPONSE:
                        break
            except Exception as e:
                print(f"    Error: {e}")
                for t in chunk:
                    if t not in all_results:
                        all_results[t] = {f: None for f in fields}

            done = min(i + chunk_size, len(tickers))
            print(f"    {done}/{len(tickers)} contracts...", end="\r")
            time.sleep(0.3)

        print(f"    {len(tickers)}/{len(tickers)} contracts done   ")
        return all_results


# ═══════════════════════════════════════════════
# CORE FUNCTIONS
# ═══════════════════════════════════════════════

def build_tickers(prefix, start_year, end_year):
    prefix_clean = prefix.strip()
    tickers = []
    meta = []
    for year in range(start_year, end_year + 1):
        yy = str(year)[-2:]
        for month in range(1, 13):
            code = MONTH_CODES[month]
            ticker = f"{prefix_clean}{code}{yy} Comdty"
            label = f"{prefix_clean}{code}{yy}"
            tickers.append(ticker)
            meta.append({
                "Contract": label,
                "Ticker": ticker,
                "Year": year,
                "Month": month,
                "Month_Name": MONTH_NAMES[month],
                "Delivery": f"{MONTH_NAMES[month]} {year}",
            })
    return tickers, meta


def pull_one_commodity(bbg, prefix, name, start_year, end_year):
    n_contracts = (end_year - start_year + 1) * 12
    print(f"\n  ── {name} ({prefix.strip()}) | {start_year}–{end_year} | {n_contracts} contracts ──")

    tickers, meta = build_tickers(prefix, start_year, end_year)
    results = bbg.bdp_batch(tickers, EXPIRY_FIELDS)

    rows = []
    for m in meta:
        ticker = m["Ticker"]
        if ticker in results:
            row = m.copy()
            row.update(results[ticker])
            rows.append(row)

    df = pd.DataFrame(rows)
    valid = df["LAST_TRADEABLE_DT"].notna().sum()
    missing = df["LAST_TRADEABLE_DT"].isna().sum()
    print(f"    ✓ {valid} with expiry dates, {missing} missing")

    has_data = df[df["LAST_TRADEABLE_DT"].notna()]
    if not has_data.empty:
        first = has_data.iloc[0]
        last = has_data.iloc[-1]
        print(f"    First: {first['Contract']} ({first['Delivery']}) → {first['LAST_TRADEABLE_DT']}")
        print(f"    Last:  {last['Contract']} ({last['Delivery']}) → {last['LAST_TRADEABLE_DT']}")

    return df


# ═══════════════════════════════════════════════
# USER INTERFACE
# ═══════════════════════════════════════════════

def show_asset_classes():
    print("\n" + "=" * 60)
    print("  FUTURES EXPIRY CALENDAR BUILDER")
    print("=" * 60)
    print("\n  Select asset class(es):\n")

    for key, ac in ASSET_CLASSES.items():
        count = len(ac["commodities"])
        print(f"    [{key}]  {ac['name']}  ({count} instruments)")

    print(f"\n    [A]  ALL asset classes")
    print(f"    [C]  Custom — enter your own tickers")


def select_asset_classes():
    show_asset_classes()
    print("\n" + "-" * 60)
    raw = input("  Enter choice(s) (comma separated, e.g. 1,3,4 or A): ").strip().upper()

    if not raw:
        print("  Nothing entered. Exiting.")
        sys.exit(0)

    if raw == "A":
        return list(ASSET_CLASSES.keys())
    elif raw == "C":
        return ["CUSTOM"]
    else:
        choices = [c.strip() for c in raw.split(",")]
        valid = [c for c in choices if c in ASSET_CLASSES]
        if not valid:
            print("  No valid choices. Exiting.")
            sys.exit(0)
        return valid


def select_commodities(asset_class_keys):
    """Let user pick specific commodities from chosen asset classes."""
    all_commodities = {}  # prefix -> name

    if asset_class_keys == ["CUSTOM"]:
        print("\n" + "-" * 60)
        raw = input("  Enter Bloomberg ticker prefixes (comma separated, e.g. CL,ES,GC): ").strip()
        if not raw:
            print("  Nothing entered. Exiting.")
            sys.exit(0)
        for t in raw.split(","):
            t = t.strip().upper()
            if t:
                all_commodities[t] = f"Custom ({t})"
        return all_commodities

    # Gather all commodities from selected asset classes
    available = {}
    for key in asset_class_keys:
        ac = ASSET_CLASSES[key]
        available.update(ac["commodities"])

    # Show them numbered
    print(f"\n  Available instruments in selected class(es):\n")
    indexed = list(available.items())
    for i, (prefix, name) in enumerate(indexed, 1):
        print(f"    [{i:2d}]  {prefix.strip():5s} — {name}")

    print(f"\n    [A]   ALL of the above")

    print("\n" + "-" * 60)
    raw = input("  Enter number(s) (comma separated, e.g. 1,3,5 or A): ").strip().upper()

    if not raw:
        print("  Nothing entered. Exiting.")
        sys.exit(0)

    if raw == "A":
        return available

    selected = {}
    for choice in raw.split(","):
        choice = choice.strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(indexed):
                prefix, name = indexed[idx]
                selected[prefix] = name
            else:
                print(f"    ⚠ {choice} out of range, skipping")
        except ValueError:
            # Maybe they typed the ticker directly
            choice_upper = choice.upper()
            if choice_upper in available:
                selected[choice_upper] = available[choice_upper]
            else:
                print(f"    ⚠ '{choice}' not recognized, skipping")

    if not selected:
        print("  No valid selections. Exiting.")
        sys.exit(0)

    return selected


def select_year_range():
    print()
    year_input = input("  Start year (default 2005): ").strip()
    if year_input:
        try:
            start_year = int(year_input)
            if start_year < 1980:
                print("  ⚠ Too early, using 1980")
                start_year = 1980
            elif start_year > datetime.today().year:
                print(f"  ⚠ Future year, using {datetime.today().year}")
                start_year = datetime.today().year
        except ValueError:
            print("  ⚠ Not a number, using 2005")
            start_year = 2005
    else:
        start_year = 2005

    end_year = datetime.today().year + 1
    return start_year, end_year


def confirm_selections(commodities, start_year, end_year):
    total_contracts = 0
    print("\n" + "=" * 60)
    print("  CONFIRMATION")
    print("=" * 60)
    print(f"\n  Period: {start_year} – {end_year}\n")
    print(f"  {'Ticker':<8} {'Commodity':<40} {'Contracts'}")
    print(f"  {'─'*8} {'─'*40} {'─'*10}")

    for prefix, name in commodities.items():
        n = (end_year - start_year + 1) * 12
        total_contracts += n
        print(f"  {prefix.strip():<8} {name:<40} {n}")

    print(f"\n  Total: {len(commodities)} instruments, {total_contracts} contracts")

    print()
    confirm = input("  Proceed? (y/n): ").strip().lower()
    return confirm == "y"


# ═══════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════

def main():
    # Step 1: Asset class
    ac_keys = select_asset_classes()

    # Step 2: Commodities
    commodities = select_commodities(ac_keys)

    # Step 3: Year range
    start_year, end_year = select_year_range()

    # Step 4: Confirm
    if not confirm_selections(commodities, start_year, end_year):
        print("  Cancelled.")
        return

    # Connect
    print("\n" + "=" * 60)
    print("  CONNECTING TO BLOOMBERG")
    print("=" * 60)
    bbg = BloombergSession()
    bbg.start()

    # Pull data
    print("=" * 60)
    print("  PULLING EXPIRY DATA")
    print("=" * 60)

    all_dfs = {}
    for prefix, name in commodities.items():
        df = pull_one_commodity(bbg, prefix, name, start_year, end_year)
        all_dfs[prefix] = (name, df)

    bbg.stop()

    # Save to Excel
    print("\n" + "=" * 60)
    print("  SAVING TO EXCEL")
    print("=" * 60)

    timestamp = datetime.today().strftime("%Y%m%d")
    filename = f"expiry_calendars_{timestamp}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:

        # Summary sheet first
        summary_rows = []
        for prefix, (name, df) in all_dfs.items():
            valid = df["LAST_TRADEABLE_DT"].notna().sum()
            missing = df["LAST_TRADEABLE_DT"].isna().sum()
            has_data = df[df["LAST_TRADEABLE_DT"].notna()]
            summary_rows.append({
                "Ticker": prefix.strip(),
                "Commodity": name,
                "Start_Year": start_year,
                "End_Year": end_year,
                "Total_Contracts": len(df),
                "With_Expiry": valid,
                "Missing": missing,
                "Earliest_Expiry": has_data["LAST_TRADEABLE_DT"].min() if not has_data.empty else None,
                "Latest_Expiry": has_data["LAST_TRADEABLE_DT"].max() if not has_data.empty else None,
            })

        df_summary = pd.DataFrame(summary_rows)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)

        # Auto-width for summary
        ws = writer.sheets["Summary"]
        for col_idx, col_name in enumerate(df_summary.columns):
            max_len = max(
                df_summary[col_name].astype(str).str.len().max(),
                len(col_name)
            ) + 2
            col_letter = chr(65 + col_idx) if col_idx < 26 else chr(64 + col_idx // 26) + chr(65 + col_idx % 26)
            ws.column_dimensions[col_letter].width = min(max_len, 30)

        # Individual sheets
        for prefix, (name, df) in all_dfs.items():
            sheet_name = f"{prefix.strip()} - {name}"[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(df.columns):
                max_len = max(
                    df[col_name].astype(str).str.len().max(),
                    len(col_name)
                ) + 2
                col_letter = chr(65 + col_idx) if col_idx < 26 else chr(64 + col_idx // 26) + chr(65 + col_idx % 26)
                ws.column_dimensions[col_letter].width = min(max_len, 25)

    print(f"\n  ✓ Saved: {filepath}")
    print(f"\n  Sheets:")
    print(f"    • Summary")
    for prefix, (name, _) in all_dfs.items():
        print(f"    • {prefix.strip()} - {name}")

    print("\n" + "=" * 60)
    print("  DONE")
    print(f"  Open in Excel: {filepath}")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
