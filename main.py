#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fetch active markets from Polymarket and export a simple 3-column Excel:
Market Title | Option A | Option B

- Focuses on binary (Yes/No) markets by default.
- Can optionally include multi-outcome markets (we'll take the first two outcomes).
- Usage:
    python main.py --output markets.xlsx
    python main.py --output markets.xlsx --include-multi
"""
import argparse
import datetime
import sys
from typing import List, Dict, Any, Tuple

import requests
import pandas as pd

GAMMA_ENDPOINTS = [
    # Primary
    "https://gamma-api.polymarket.com/markets?limit=1000&active=true",
    # Fallbacks (just in case the primary query params change)
    "https://gamma-api.polymarket.com/markets?limit=1000&closed=false",
    "https://gamma-api.polymarket.com/markets?limit=1000",
]

# A very permissive default headers set
DEFAULT_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "User-Agent": "Mozilla/5.0 (compatible; PolymarketExcelBot/1.0)",
}


def fetch_markets() -> List[Dict[str, Any]]:
    """
    Try several endpoints under gamma-api.polymarket.com and return the 'markets' list.
    """
    last_err = None
    for url in GAMMA_ENDPOINTS:
        try:
            resp = requests.get(url, headers=DEFAULT_HEADERS, timeout=20)
            resp.raise_for_status()
            data = resp.json()
            # Some responses return {"markets":[...]} others may directly be a list
            if isinstance(data, dict) and "markets" in data and isinstance(data["markets"], list):
                return data["markets"]
            if isinstance(data, list):
                return data
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"Failed to fetch markets from Polymarket. Last error: {last_err}")


def extract_title_and_outcomes(m: Dict[str, Any]) -> Tuple[str, List[str]]:
    """
    Extract a readable title and outcomes array from a Polymarket market object.
    Tries common keys with graceful fallback.
    """
    title = (
        m.get("question")
        or m.get("title")
        or m.get("name")
        or m.get("slug")
        or "Untitled Market"
    )

    # Outcomes: usually ["Yes","No"] for binary; for categorical could be more
    outcomes = None
    # gamma markets often have 'outcomes'
    if isinstance(m.get("outcomes"), list) and m["outcomes"]:
        outcomes = [str(o) for o in m["outcomes"]]

    # Some older shapes may put outcomes under 'outcomeNames' or similar
    if not outcomes and isinstance(m.get("outcomeNames"), list):
        outcomes = [str(o) for o in m["outcomeNames"]]

    # Fallback to Yes/No if it's clearly binary-type by 'type' or 'conditionType'
    if not outcomes:
        cond_type = (m.get("conditionType") or m.get("type") or "").lower()
        if cond_type in ("binary", "scalar", "range"):
            outcomes = ["Yes", "No"]

    if not outcomes:
        outcomes = ["—", "—"]

    return str(title).strip(), outcomes


def build_rows(markets: List[Dict[str, Any]], include_multi: bool) -> List[Dict[str, Any]]:
    rows = []
    for m in markets:
        title, outcomes = extract_title_and_outcomes(m)
        if not outcomes:
            continue

        # Binary-only by default (two outcomes exactly).
        if not include_multi and len(outcomes) != 2:
            continue

        # Ensure exactly two columns for the sheet: take first two outcomes if more exist
        opt_a = outcomes[0] if len(outcomes) >= 1 else "—"
        opt_b = outcomes[1] if len(outcomes) >= 2 else "—"

        rows.append(
            {
                "Название рынка": title,
                "Параметр A": opt_a,
                "Параметр B": opt_b,
            }
        )
    return rows


def write_excel(rows: List[Dict[str, Any]], output_path: str) -> None:
    if not rows:
        raise RuntimeError("Нет данных для записи в Excel (список пуст).")

    df = pd.DataFrame(rows, columns=["Название рынка", "Параметр A", "Параметр B"])
    # Sort by title for readability
    df = df.sort_values(by="Название рынка").reset_index(drop=True)

    # Write to Excel
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Polymarket")
        # Add a generated-at note in a second sheet
        meta = pd.DataFrame(
            {
                "Ключ": ["Источник", "Сгенерировано"],
                "Значение": [
                    "gamma-api.polymarket.com",
                    datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                ],
            }
        )
        meta.to_excel(writer, index=False, sheet_name="Meta")


def main():
    parser = argparse.ArgumentParser(description="Export Polymarket markets to Excel (3 columns).")
    parser.add_argument("--output", "-o", default="polymarket_markets.xlsx", help="Путь к Excel файлу вывода")
    parser.add_argument("--include-multi", action="store_true", help="Включить рынки с >2 исходами (берём первые два)")
    args = parser.parse_args()

    try:
        markets = fetch_markets()
    except Exception as e:
        print(f"[Ошибка] Не удалось загрузить рынки Polymarket: {e}", file=sys.stderr)
        sys.exit(2)

    rows = build_rows(markets, include_multi=args.include_multi)
    if not rows:
        print("[Предупреждение] Не найдено подходящих рынков для экспорта.", file=sys.stderr)

    try:
        write_excel(rows, args.output)
    except Exception as e:
        print(f"[Ошибка] Не удалось записать Excel: {e}", file=sys.stderr)
        sys.exit(3)

    print(f"Готово. Записано: {args.output} (строк: {len(rows)})")


if __name__ == "__main__":
    main()
