"""Daily Telegram alert: which stocks from your three lists moved UP today.

Lists read from data/:
  Stock_Research_Model.xlsx      (Reviewed_Summary + Multibagger_List)
  Long_Term_Stock_Watchlist.xlsx (Priority, Secondary, Speculative, Special situations)
  PP_Analysis.xlsx               (Final Watchlist)

Env vars:
  TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID  - required to send (else prints only)
  MIN_GAIN_PCT   - alert threshold, default 0 (every stock that is up). e.g. 3 = only +3% or more
  FORCE=1        - send even if the market was closed today
  DRY_RUN=1      - print, never send
"""
import csv
import datetime as dt
import difflib
import os
import re
import sys
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
import requests
import yfinance as yf

BASE = Path(__file__).parent
DATA = BASE / "data"
IST = ZoneInfo("Asia/Kolkata")

# (file, sheet, name column header, symbol column header or None, tag)
SOURCES = [
    ("Stock_Research_Model.xlsx", "Reviewed_Summary", "Stock Name", "Symbol", "SRM"),
    ("Stock_Research_Model.xlsx", "Multibagger_List", "Stock", None, "SRM"),
    ("Long_Term_Stock_Watchlist.xlsx", "Priority", "Company", None, "LTW"),
    ("Long_Term_Stock_Watchlist.xlsx", "Secondary", "Company", None, "LTW"),
    ("Long_Term_Stock_Watchlist.xlsx", "Speculative", "Company", None, "LTW"),
    ("Long_Term_Stock_Watchlist.xlsx", "Special situations", "Company", None, "LTW"),
    ("PP_Analysis.xlsx", "Final Watchlist", "Company", None, "PP"),
]

NOISE = re.compile(r"\b(ltd|limited|pvt|private|india|international|industries|the)\b")


def norm(name: str) -> str:
    s = name.lower().replace("&", " and ")
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = NOISE.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()


def read_overrides() -> dict:
    out = {}
    p = BASE / "symbols.csv"
    if p.exists():
        with p.open(newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("name") and row.get("symbol"):
                    out[norm(row["name"])] = row["symbol"].strip()
    return out


def read_lists():
    """Return list of (display name, explicit symbol or '', tag)."""
    rows = []
    for fname, sheet, ncol, scol, tag in SOURCES:
        path = DATA / fname
        if not path.exists():
            print(f"WARN missing {path}")
            continue
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        if sheet not in wb.sheetnames:
            print(f"WARN sheet {sheet} not in {fname}")
            continue
        it = wb[sheet].iter_rows(values_only=True)
        header = None
        for r in it:
            if r and ncol in r:
                header = list(r)
                break
        if not header:
            print(f"WARN header {ncol} not found in {fname}/{sheet}")
            continue
        ni = header.index(ncol)
        si = header.index(scol) if scol and scol in header else None
        for r in it:
            if not r or ni >= len(r) or not r[ni]:
                continue
            name = str(r[ni]).strip()
            if name.lower() in ("stock", "company") or name.lower().startswith("source:"):
                continue  # repeated headers / footnotes
            sym = str(r[si]).strip() if si is not None and si < len(r) and r[si] else ""
            rows.append((name, sym, tag))
    return rows


def to_yahoo(sym: str) -> str:
    sym = sym.strip().upper()
    if "." in sym:
        return sym
    return f"{sym}.BO" if sym.isdigit() else f"{sym}.NS"


def search_symbol(name: str):
    """Look a company name up on Yahoo; return NSE/BSE symbol or None."""
    try:
        quotes = yf.Search(name, max_results=8).quotes
    except Exception as e:  # network / API hiccup
        print(f"WARN search failed for {name}: {e}")
        return None
    target = norm(name)
    best = None
    for q in quotes:
        s = q.get("symbol", "")
        if not (s.endswith(".NS") or s.endswith(".BO")) or q.get("quoteType") != "EQUITY":
            continue
        label = norm(q.get("longname") or q.get("shortname") or "")
        score = difflib.SequenceMatcher(None, target, label).ratio()
        if target and (target in label or label in target):
            score = max(score, 0.9)
        rank = (score, s.endswith(".NS"))
        if score >= 0.6 and (best is None or rank > best[0]):
            best = (rank, s)
    return best[1] if best else None


def build_universe():
    overrides = read_overrides()
    universe = {}  # yahoo symbol -> {"name", "tags"}
    unresolved = []
    cache = {}
    for name, sym, tag in read_lists():
        key = norm(name)
        if key in cache:
            ysym = cache[key]
        else:
            if key in overrides:
                ysym = to_yahoo(overrides[key])
            elif sym:
                ysym = to_yahoo(sym)
            else:
                ysym = search_symbol(name)
            cache[key] = ysym
            if ysym:
                print(f"map  {name!r:45} -> {ysym}")
            else:
                print(f"MISS {name!r}")
                unresolved.append(name)
        if ysym:
            e = universe.setdefault(ysym, {"name": name, "tags": set()})
            e["tags"].add(tag)
    return universe, sorted(set(unresolved))


def fetch_moves(symbols):
    df = yf.download(
        symbols, period="7d", interval="1d", group_by="ticker",
        auto_adjust=False, progress=False, threads=False,
    )
    moves = {}
    for s in symbols:
        try:
            close = (df[s]["Close"] if len(symbols) > 1 else df["Close"]).dropna()
            if len(close) < 2:
                continue
            last, prev = float(close.iloc[-1]), float(close.iloc[-2])
            moves[s] = (close.index[-1].date(), last, (last / prev - 1) * 100)
        except Exception:
            continue
    return moves


def send_telegram(text: str):
    token, chat = os.getenv("TELEGRAM_BOT_TOKEN"), os.getenv("TELEGRAM_CHAT_ID")
    if os.getenv("DRY_RUN") or not (token and chat):
        print("---- (not sent: DRY_RUN or Telegram secrets missing) ----\n" + text)
        return
    chunks, cur = [], ""
    for line in text.split("\n"):
        if len(cur) + len(line) + 1 > 3800:
            chunks.append(cur)
            cur = ""
        cur += line + "\n"
    chunks.append(cur)
    for c in chunks:
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            data={"chat_id": chat, "text": c, "disable_web_page_preview": True},
            timeout=30,
        )
        r.raise_for_status()


def main():
    min_gain = float(os.getenv("MIN_GAIN_PCT", "0"))
    universe, unresolved = build_universe()
    if not universe:
        sys.exit("No stocks resolved - check data/ files.")
    moves = fetch_moves(list(universe))
    if not moves:
        sys.exit("No price data returned.")

    session = max(m[0] for m in moves.values())
    today = dt.datetime.now(IST).date()
    if session != today and not os.getenv("FORCE"):
        print(f"Latest session is {session}, today is {today} (market closed). Nothing sent.")
        return

    fresh = {s: m for s, m in moves.items() if m[0] == session}
    ups = sorted(((s, m) for s, m in fresh.items() if m[2] >= max(min_gain, 0.05)),
                 key=lambda x: -x[1][2])
    downs = sum(1 for m in fresh.values() if m[2] < 0)

    lines = [f"\U0001F4C8 Watchlist movers - {session:%d %b %Y}",
             f"{len(ups)} up{f' (>= {min_gain:g}%)' if min_gain else ''}, "
             f"{downs} down, {len(fresh)} tracked of {len(universe)}", ""]
    for s, (_, price, pct) in ups:
        e = universe[s]
        tags = "/".join(sorted(e["tags"]))
        lines.append(f"▲ {pct:+.1f}%  {e['name']} ({s.split('.')[0]})  Rs {price:,.2f}  [{tags}]")
    if not ups:
        lines.append("No stock met the threshold today.")
    stale = len(universe) - len(fresh)
    if stale:
        lines.append(f"\n{stale} stock(s) had no price for {session:%d %b} (suspended/illiquid/not traded).")
    if unresolved:
        lines.append(f"\nNo ticker found for: {', '.join(unresolved)} (add to symbols.csv)")
    lines.append("\nSRM=Stock_Research_Model LTW=Long_Term_Watchlist PP=PP_Analysis")
    send_telegram("\n".join(lines))


if __name__ == "__main__":
    main()

