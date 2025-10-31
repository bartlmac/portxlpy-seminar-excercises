import xlwings as xw
import pandas as pd
from pathlib import Path
import csv
import os

EXCEL_PATH = Path("TARIFRECHNER_KLV.xlsm")
OUT_ZELL = EXCEL_PATH.parent / "excelzell.csv"
OUT_BER = EXCEL_PATH.parent / "excelber.csv"

def safe_formula(cell):
    try:
        return cell.formula
    except Exception:
        return ""

def safe_value(cell):
    try:
        return cell.value
    except Exception:
        return ""

def export_zellen(wb, out_path):
    rows = []
    for sht in wb.sheets:
        used_range = sht.used_range
        if used_range is None or used_range.address == "$A$1" and not sht.range("A1").value:
            continue
        for row in used_range.rows:
            for cell in row:
                val = safe_value(cell)
                form = safe_formula(cell)
                if val is None and (not form or form == ""):
                    continue
                rows.append({
                    "Blatt": sht.name,
                    "Adresse": cell.address,
                    "Formel": form,
                    "Wert": val,
                })
    df = pd.DataFrame(rows)
    df.to_csv(out_path, index=False, quoting=csv.QUOTE_NONNUMERIC)

def export_bereiche(wb, out_path):
    rows = []
    for sht in wb.sheets:
        # Defined Names / Named Ranges
        for name in wb.names:
            try:
                refers_to = name.refers_to_range
            except Exception:
                continue
            if refers_to and refers_to.sheet.name == sht.name:
                rows.append({
                    "Blatt": sht.name,
                    "Name": name.name,
                    "Adresse": refers_to.address,
                })
        # Excel Table ranges (listobjects)
        for tbl in sht.api.ListObjects:
            rows.append({
                "Blatt": sht.name,
                "Name": tbl.Name,
                "Adresse": tbl.Range.Address,
            })
    df = pd.DataFrame(rows).drop_duplicates()
    df.to_csv(out_path, index=False, quoting=csv.QUOTE_NONNUMERIC)

def main():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"{EXCEL_PATH} nicht gefunden.")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(EXCEL_PATH))
        export_zellen(wb, OUT_ZELL)
        export_bereiche(wb, OUT_BER)
        wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    main()
