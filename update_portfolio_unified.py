import os
import json
import time
import argparse
import requests
import pandas as pd
import yfinance as yf
from datetime import datetime, timezone

API_KEY = os.getenv("OPENFIGI_API_KEY", "").strip()
OPENFIGI_URL = "https://api.openfigi.com/v3/mapping"

EXCH_SUFFIX = {
    "XETR": ".DE", "ETR": ".DE", "FRA": ".F",
    "NYS": "", "NAS": "", "NMQ": "", "NGM": "", "ASE": "",
    "TSE": ".T", "TYO": ".T", "JPX": ".T",
    "LSE": ".L", "IOB": ".IL",
    "MCE": ".MC", "MIL": ".MI", "PAR": ".PA", "AMS": ".AS",
    "VIE": ".VI", "SWX": ".SW", "BRU": ".BR", "CPH": ".CO", "STO": ".ST", "HEL": ".HE",
    "ASX": ".AX", "TSX": ".TO", "TSV": ".V", "HKG": ".HK", "SGX": ".SI",
}

def chunked(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i+n]

def openfigi_map_isins(isins):
    if not API_KEY:
        print("INFO: OPENFIGI_API_KEY not set. Skipping auto-mapping; will rely on provided Ticker.")
        return pd.DataFrame([{"ISIN": i} for i in isins])
    headers = {"Content-Type": "application/json", "X-OPENFIGI-APIKEY": API_KEY}
    rows = []
    for batch in chunked(list(isins), 50):
        payload = [{"idType": "ID_ISIN", "idValue": i} for i in batch]
        try:
            r = requests.post(OPENFIGI_URL, headers=headers, data=json.dumps(payload), timeout=20)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            for i in batch:
                rows.append({"ISIN": i})
            continue
        for i, item in zip(batch, data):
            best = None
            if item and isinstance(item.get("data"), list) and item["data"]:
                for cand in item["data"]:
                    if cand.get("exchCode") in EXCH_SUFFIX:
                        best = cand; break
                if not best:
                    best = item["data"][0]
            row = {"ISIN": i}
            if best:
                row.update({
                    "exchCode": best.get("exchCode"),
                    "ticker_raw": best.get("ticker"),
                    "name": best.get("name"),
                    "securityType2": best.get("securityType2"),
                    "marketSector": best.get("marketSecDes") or best.get("marketSector"),
                    "currency_base": best.get("crncy") or best.get("currency"),
                })
            rows.append(row)
        time.sleep(0.3)
    return pd.DataFrame(rows)

def build_yahoo_symbol(ticker_raw, exch_code, fallback_suffix=".DE"):
    if not isinstance(ticker_raw, str) or not ticker_raw:
        return None
    if "." in ticker_raw or "=" in ticker_raw:
        return ticker_raw
    suff = EXCH_SUFFIX.get(str(exch_code).upper(), None)
    if suff is None:
        suff = fallback_suffix
    return f"{ticker_raw}{suff}"

def _interval_plan(mode: str):
    if mode == "off":
        return [("60m", "7d"), ("1d", "10d")]
    if mode == "aggressive":
        return [("1m","5d"), ("2m","10d"), ("5m","10d"), ("15m","30d"), ("30m","30d"), ("60m","60d"), ("1d","10d")]
    return [("1m","5d"), ("5m","10d"), ("60m","60d"), ("1d","10d")]

def fetch_last_prices(tickers, mode="normal"):
    from datetime import datetime, timezone
    plan = _interval_plan(mode)
    out = {}
    uniq = [t for t in sorted(set(tickers)) if t]
    for t in uniq:
        price = ts_local = ts_utc = src = None
        try:
            tk = yf.Ticker(t)

            fi = getattr(tk, "fast_info", None)
            if fi and "last_price" in fi and fi["last_price"]:
                price = float(fi["last_price"])
                ts_utc = datetime.now(timezone.utc)
                ts_local = ts_utc
                src = "fast"
            if price is None or mode != "off":
                mstate = (fi.get("market_state") if isinstance(fi, dict) else None) or ""
                allow_short = (mode != "off") and (mstate.upper() in ("REGULAR","PRE","POST","EXTENDED",""))
                for interval, period in plan:
                    if price is not None and interval != "1d":
                        if mode == "normal" and not allow_short:
                            continue
                    try:
                        df = tk.history(period=period, interval=interval, auto_adjust=True, prepost=False)
                        if df is not None and not df.empty:
                            close = df["Close"].dropna()
                            if not close.empty:
                                price = float(close.iloc[-1])
                                idx = close.index[-1]
                                ts_local = idx.to_pydatetime()
                                ts_utc = idx.tz_convert("UTC").to_pydatetime() if hasattr(idx, "tz_convert") else ts_local
                                src = interval
                                break
                    except Exception:
                        continue
        except Exception:
            pass
        out[t] = (price, ts_local, ts_utc, src)
    return out



def fetch_fx_to_eur(currencies, mode="normal"):
    need = [c for c in currencies if (c and str(c).upper() != "EUR")]
    if not need:
        return {"EUR": 1.0}
    plan = _interval_plan(mode)
    out = {"EUR": 1.0}
    for c in sorted(set(str(ccy).upper() for ccy in need)):
        pair = f"EUR{c}=X"
        px = None
        for interval, period in plan:
            try:
                df = yf.Ticker(pair).history(period=period, interval=interval)
                if df is not None and not df.empty:
                    px = float(df["Close"].dropna().iloc[-1])
                    break
            except Exception:
                continue
        out[c] = (1.0/px) if px else None
    return out


def resolve_map(base_df, eur_only):
    for col in ["Ticker","Exchange","Currency","Name"]:
        if col not in base_df.columns: base_df[col] = None
    need = base_df[base_df["Ticker"].isna() | (base_df["Ticker"].astype(str).str.strip()=="")]
    if not need.empty:
        of = openfigi_map_isins(need["ISIN"].tolist())
        base_df = base_df.merge(of, on="ISIN", how="left", suffixes=("","_OF"))
        def pick(row):
            t = str(row.get("Ticker") or "").strip()
            if t:
                return t
            return build_yahoo_symbol(row.get("ticker_raw"), row.get("exchCode") or row.get("exchCode_OF"), fallback_suffix=".DE" if eur_only else "")
        base_df["Ticker"] = base_df.apply(pick, axis=1)
        base_df["Exchange"] = base_df["Exchange"].fillna(base_df.get("exchCode")).fillna(base_df.get("exchCode_OF"))
        base_df["Name"] = base_df["Name"].fillna(base_df.get("name"))
        if eur_only:
            base_df["Currency"] = "EUR"
        else:
            base_df["Currency"] = base_df["Currency"].fillna(base_df.get("currency_base")).fillna("EUR")
    else:
        if eur_only:
            base_df["Currency"] = "EUR"
    base_df["Ticker"] = base_df["Ticker"].astype(str).str.strip()
    base_df["Currency"] = base_df["Currency"].astype(str).str.upper().str.strip()
    return base_df

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel", help="D:\\Investments\\Portfolio.xlsx")
    ap.add_argument("--eur-only", action="store_true", help="Force FX=1 and treat all outputs as EUR (you buy only in EUR)")
    ap.add_argument(
        "--intraday",
        choices=["off", "normal", "aggressive"],
        default="normal",
        help="Intraday depth: off=60m/1d only, normal=1m→5m→60m→1d, aggressive=1m→2m→5m→15m→30m→60m→1d"
    )
    args = ap.parse_args()

    xls = pd.ExcelFile(args.excel)
    try:
        base = pd.read_excel(xls, sheet_name="Map")
    except Exception:
        raise SystemExit("ERROR: 'Map' sheet not found. It must at least have an ISIN column.")
    if "ISIN" not in base.columns:
        for c in base.columns:
            if c.lower()=="isin":
                base = base.rename(columns={c:"ISIN"}); break
    if "ISIN" not in base.columns:
        raise SystemExit("ERROR: Map sheet must contain column 'ISIN'.")

    base = base.dropna(subset=["ISIN"]).copy()
    base["ISIN"] = base["ISIN"].astype(str).str.strip()

    base = resolve_map(base, eur_only=args.eur_only)

    # Fetch prices
    tickers = [t for t in base["Ticker"].tolist() if t]
    price_map = fetch_last_prices(tickers, mode=args.intraday)

    if args.eur_only:
        fx_map = {"EUR": 1.0}
    else:
        fx_map = fetch_fx_to_eur(base["Currency"].tolist(), mode=args.intraday)

    rows = []
    for _, r in base.iterrows():
        isin = r["ISIN"]
        tkr  = r["Ticker"]
        ccy  = r["Currency"] if r["Currency"] else "EUR"
        p, ts_local, ts_utc, src = price_map.get(tkr, (None, None, None, None)) if tkr else (None, None, None, None)
        fx_to_eur = 1.0 if ccy == "EUR" else fx_map.get(ccy)
        price_eur = (p * fx_to_eur) if (p is not None and fx_to_eur not in (None, 0)) else None
        rows.append({
            "ISIN": isin,
            "Ticker": tkr,
            "Currency": ccy if not args.eur_only else "EUR",
            "Last Timestamp (Local)": ts_local.isoformat() if ts_local else "",
            "Last Timestamp (UTC)": ts_utc.isoformat() if ts_utc else "",
            "Last Close (Local)": p,
            "FX to EUR": fx_to_eur,
            "Last Close (EUR)": price_eur,
            "Source": src
        })

    prices = pd.DataFrame(rows)
    with pd.ExcelWriter(args.excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        base[["ISIN","Ticker","Exchange","Currency","Name"]].to_excel(w, sheet_name="Map", index=False)
        prices.to_excel(w, sheet_name="Prices", index=False)

    print(f"OK: Updated Map({len(base)}) and Prices({len(prices)}) at {datetime.now(timezone.utc).isoformat()}Z. eur_only={args.eur_only}")

if __name__ == "__main__":
    main()
