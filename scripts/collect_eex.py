"""
Collecte des settlement prices EEX pour Power FR et Natural Gas PEG.
Écrit dans data/eex_prix.xlsx avec gestion des prix reportés (cellule jaune).
"""

import os
import sys
import time
import requests
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import openpyxl
from openpyxl.styles import PatternFill

# Force UTF-8 sur Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

BASE_URL = "https://api.eex-group.com/pub/market-data/price-ticker"
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "eex_prix.xlsx")

REQUEST_HEADERS = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Origin": "https://www.eex.com",
    "Referer": "https://www.eex.com/",
}

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

HEADERS = ["Date", "Contrat", "Commodite", "Type", "Settlement_Price", "Source_Prix"]

# ---------------------------------------------------------------------------
# Définition des contrats à collecter
# ---------------------------------------------------------------------------

def build_contracts():
    today = date.today()
    contracts = []

    # Power FR Baseload ── CAL 2027-2030
    for year in range(2027, 2031):
        contracts.append({
            "label": f"ELEC-CAL-{year}",
            "commodite": "Power FR Baseload",
            "type": "CAL",
            "params": {
                "shortCode": "F7BY",
                "area": "FR",
                "product": "Base",
                "commodity": "POWER",
                "pricing": "F",
                "maturity": f"{year}01",
            },
        })

    # Power FR Baseload ── QTR 2027-2030
    for year in range(2027, 2031):
        for q, month in enumerate(["01", "04", "07", "10"], start=1):
            contracts.append({
                "label": f"ELEC-Q{q}-{year}",
                "commodite": "Power FR Baseload",
                "type": "QTR",
                "params": {
                    "shortCode": "F7BQ",
                    "area": "FR",
                    "product": "Base",
                    "commodity": "POWER",
                    "pricing": "F",
                    "maturity": f"{year}{month}",
                },
            })

    # Power FR Baseload ── MTH : mois courant + 17 mois suivants
    for i in range(18):
        m = today + relativedelta(months=i)
        contracts.append({
            "label": f"ELEC-{m.strftime('%b')}-{m.year}",
            "commodite": "Power FR Baseload",
            "type": "MTH",
            "params": {
                "shortCode": "F7BM",
                "area": "FR",
                "product": "Base",
                "commodity": "POWER",
                "pricing": "F",
                "maturity": m.strftime("%Y%m"),
            },
        })

    # Power FR Peakload ── CAL 2027-2030
    for year in range(2027, 2031):
        contracts.append({
            "label": f"PEAK-CAL-{year}",
            "commodite": "Power FR Peakload",
            "type": "CAL",
            "params": {
                "shortCode": "F7PY",
                "area": "FR",
                "product": "Peak",
                "commodity": "POWER",
                "pricing": "F",
                "maturity": f"{year}01",
            },
        })

    # Natural Gas TTF ── CAL 2027-2030
    for year in range(2027, 2031):
        contracts.append({
            "label": f"GAZ-CAL-{year}",
            "commodite": "Natural Gas TTF",
            "type": "CAL",
            "params": {
                "shortCode": "G3BY",
                "area": "TTF",
                "product": "Physical",
                "commodity": "NATGAS",
                "pricing": "F",
                "maturity": f"{year}01",
            },
        })

    # Natural Gas TTF ── QTR 2027-2030
    for year in range(2027, 2031):
        for q, month in enumerate(["01", "04", "07", "10"], start=1):
            contracts.append({
                "label": f"GAZ-Q{q}-{year}",
                "commodite": "Natural Gas TTF",
                "type": "QTR",
                "params": {
                    "shortCode": "G3BQ",
                    "area": "TTF",
                    "product": "Physical",
                    "commodity": "NATGAS",
                    "pricing": "F",
                    "maturity": f"{year}{month}",
                },
            })

    # Natural Gas TTF ── MTH : mois courant + 17 mois suivants
    for i in range(18):
        m = today + relativedelta(months=i)
        contracts.append({
            "label": f"GAZ-{m.strftime('%b')}-{m.year}",
            "commodite": "Natural Gas TTF",
            "type": "MTH",
            "params": {
                "shortCode": "G3BM",
                "area": "TTF",
                "product": "Physical",
                "commodity": "NATGAS",
                "pricing": "F",
                "maturity": m.strftime("%Y%m"),
            },
        })

    return contracts


# ---------------------------------------------------------------------------
# Appel API
# ---------------------------------------------------------------------------

def fetch_settlement(params: dict) -> float | None:
    try:
        r = requests.get(BASE_URL, params=params, headers=REQUEST_HEADERS, timeout=15)
        r.raise_for_status()
        data = r.json()
        # Format: {"header": ["lastUpdatedAt", "settlPx", ...], "data": [[...val...]]}
        rows = data.get("data")
        if not rows:
            return None
        headers = data.get("header", [])
        idx = headers.index("settlPx") if "settlPx" in headers else 1
        price = rows[0][idx]
        if price is None or price == "":
            return None
        return float(price)
    except Exception:
        return None
    finally:
        time.sleep(0.4)  # évite le rate limiting (max ~150 req/min)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def load_or_create_workbook(path: str):
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Prices"
        ws.append(HEADERS)
        # Format header
        from openpyxl.styles import Font, PatternFill as PF
        header_fill = PF(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
    return wb, ws


def already_has_today(ws, today_str: str) -> bool:
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and str(row[0]) == today_str:
            return True
    return False


def get_last_known_price(ws, label: str) -> float | None:
    """Retourne le dernier prix connu pour ce contrat (hors reporté)."""
    last_price = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[1] == label and row[4] is not None:
            last_price = row[4]
    return last_price


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    today = date.today()
    today_str = today.strftime("%Y-%m-%d")

    print(f"\n=== Collecte EEX — {today_str} ===")

    contracts = build_contracts()

    os.makedirs(os.path.dirname(os.path.abspath(EXCEL_PATH)), exist_ok=True)
    wb, ws = load_or_create_workbook(EXCEL_PATH)

    if already_has_today(ws, today_str):
        print("  Données du jour déjà présentes — rien à faire.")
        sys.exit(0)

    added = reported = ignored = 0

    for c in contracts:
        price = fetch_settlement(c["params"])

        if price is not None:
            ws.append([today_str, c["label"], c["commodite"], c["type"], price, "Settlement"])
            print(f"  {c['label']}: {price} €/MWh")
            added += 1
        else:
            last = get_last_known_price(ws, c["label"])
            if last is not None:
                row_idx = ws.max_row + 1
                ws.append([today_str, c["label"], c["commodite"], c["type"], last, "Reporté J-1"])
                # Cellule prix en jaune
                ws.cell(row=row_idx, column=5).fill = YELLOW_FILL
                print(f"  {c['label']}: {last} €/MWh ⚠️ reporté")
                reported += 1
            else:
                print(f"  {c['label']}: ignoré (pas de prix)")
                ignored += 1

    wb.save(EXCEL_PATH)
    print(f"\n✓ Terminé — {added} lignes ajoutées ({reported} reportées, {ignored} ignorées)")


if __name__ == "__main__":
    main()
