"""Fetch current quotes for invested tickers via yfinance and update data/quotes.json.

Usage:
    python -m scripts.refresh_quotes

Reads tickers from data/investments.json (operations + deposits-derived tickers),
queries yfinance for last close + 1-day change, converts USD quotes to EUR using
the EURUSD=X rate, and writes the result to data/quotes.json.

Errors per-ticker are logged but don't abort the run. If yfinance has no data for
a ticker, that ticker is skipped and any existing quote for it is preserved.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

import yfinance as yf

from scripts.investments_store import load as load_investments


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INVESTMENTS_PATH = PROJECT_ROOT / "data" / "investments.json"
DEFAULT_QUOTES_PATH = PROJECT_ROOT / "data" / "quotes.json"
# Cross-project: universo de swing trading + trades abiertos (Proyecto Finanzas Claude)
DEFAULT_WATCHLIST_SWING_PATH = Path("c:/Users/migas/Documents/Proyecto Finanzas Claude/cartera/watchlist_swing.json")
DEFAULT_TRADES_PATH = Path("c:/Users/migas/Documents/Proyecto Finanzas Claude/cartera/trades.json")

FALLBACK_EUR_PER_USD = 0.92


def fetch_eur_per_usd() -> float:
    """Return EUR per 1 USD using the EURUSD=X ticker. Fallback if unavailable."""
    try:
        t = yf.Ticker("EURUSD=X")
        hist = t.history(period="5d", auto_adjust=False)
        if hist.empty:
            return FALLBACK_EUR_PER_USD
        # EURUSD=X close is USD per 1 EUR. We want EUR per 1 USD: 1 / that.
        last_close = float(hist["Close"].dropna().iloc[-1])
        if last_close <= 0:
            return FALLBACK_EUR_PER_USD
        return 1.0 / last_close
    except Exception as e:
        print(f"  (warn: EURUSD fetch failed: {e}, using fallback {FALLBACK_EUR_PER_USD})", file=sys.stderr)
        return FALLBACK_EUR_PER_USD


def fetch_quote(ticker: str, eur_per_usd: float) -> dict | None:
    """Fetch a single ticker. Returns dict with precio (EUR), moneda, cambio_pct_1d
    or None if data not available."""
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period="5d", auto_adjust=False)
        if hist.empty:
            return None
        closes = hist["Close"].dropna()
        if len(closes) == 0:
            return None
        price = float(closes.iloc[-1])
        prev_close = float(closes.iloc[-2]) if len(closes) >= 2 else price
        change_pct = (price - prev_close) / prev_close if prev_close > 0 else 0.0

        currency = "USD"
        try:
            fi = t.fast_info
            currency = (fi.get("currency") if isinstance(fi, dict) else getattr(fi, "currency", None)) or "USD"
        except Exception:
            pass

        if currency.upper() == "USD":
            price_eur = price * eur_per_usd
        elif currency.upper() == "EUR":
            price_eur = price
        else:
            # Unhandled currency — store raw and let user fix
            price_eur = price
            print(f"  (warn: {ticker} en {currency}, sin conversión)", file=sys.stderr)

        return {
            "precio": round(price_eur, 4),
            "moneda": "EUR",
            "cambio_pct_1d": round(change_pct, 6),
            "raw_currency": currency,
            "raw_price": round(price, 4),
        }
    except Exception as e:
        print(f"  (error: {ticker} -> {e})", file=sys.stderr)
        return None


def main() -> int:
    if not DEFAULT_INVESTMENTS_PATH.exists():
        print(f"ERROR: no existe {DEFAULT_INVESTMENTS_PATH}. Ejecuta init_data primero.", file=sys.stderr)
        return 2

    store = load_investments(DEFAULT_INVESTMENTS_PATH)
    ops = store.get("operations", [])
    tickers_set = {op["ticker"] for op in ops}

    # Añadir tickers del universo de swing trading (Proyecto Finanzas Claude)
    if DEFAULT_WATCHLIST_SWING_PATH.exists():
        try:
            with open(DEFAULT_WATCHLIST_SWING_PATH, encoding="utf-8") as f:
                wl = json.load(f)
            for w in wl.get("tickers", []):
                tickers_set.add(w.get("ticker_yfinance", w.get("ticker")))
        except Exception as e:
            print(f"  (warn: no pude leer watchlist_swing.json: {e})", file=sys.stderr)
    # Añadir tickers de los trades abiertos
    if DEFAULT_TRADES_PATH.exists():
        try:
            with open(DEFAULT_TRADES_PATH, encoding="utf-8") as f:
                td = json.load(f)
            for tr in td.get("trades_abiertos", []):
                tickers_set.add(tr["ticker"])
        except Exception as e:
            print(f"  (warn: no pude leer trades.json: {e})", file=sys.stderr)

    tickers = sorted(tickers_set)
    if not tickers:
        print("Sin tickers que actualizar.")
        return 0

    print(f"Tickers a refrescar: {', '.join(tickers)}")
    eur_per_usd = fetch_eur_per_usd()
    print(f"EUR/USD: {eur_per_usd:.4f}")

    # Preserve existing quotes (so unfetched tickers keep their last known value)
    existing = {}
    if DEFAULT_QUOTES_PATH.exists():
        try:
            existing = json.loads(DEFAULT_QUOTES_PATH.read_text(encoding="utf-8")).get("quotes", {})
        except Exception:
            existing = {}

    quotes = dict(existing)
    for ticker in tickers:
        q = fetch_quote(ticker, eur_per_usd)
        if q is None:
            kept = "(mantengo valor previo)" if ticker in existing else "(sin datos)"
            print(f"  {ticker}: SKIPPED {kept}")
            continue
        quotes[ticker] = q
        print(f"  {ticker}: {q['precio']:.4f} EUR ({q['raw_currency']} {q['raw_price']:.4f}, 1d {q['cambio_pct_1d'] * 100:+.2f}%)")

    out = {
        "version": 1,
        "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "eur_per_usd": round(eur_per_usd, 6),
        "quotes": quotes,
    }
    DEFAULT_QUOTES_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: quotes.json actualizado con {len(quotes)} tickers.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
