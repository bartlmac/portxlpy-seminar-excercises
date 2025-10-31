# run_calc.py (korrigierte Version mit robustem Casting)

import os
import argparse
import pandas as pd
import json
from importlib import import_module
from pathlib import Path

THIS_DIR = Path(__file__).resolve().parent   # …/Bartek/output
os.chdir(THIS_DIR)                           # CSV-Dateien immer gefunden

ALL_FUNCS = ["Bxt", "BJB", "VSt", "Rnt", "VBar", "REBar"]

def load_input(filepath):
    df = pd.read_csv(filepath)
    df.columns = [col.strip().lower() for col in df.columns]
    df["name"] = df.iloc[:, 0].str.lower()
    df["wert"] = df.iloc[:, 1]
    return dict(zip(df["name"], df["wert"]))

def parse_numeric(val):
    try:
        f = float(val)
        return int(f) if f.is_integer() else f
    except Exception:
        return val  # fallback: string

def main():
    parser = argparse.ArgumentParser(
        description="Berechnet Versicherungsausgaben auf Basis von CSV-Variablen",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument("--var-file", default="var.csv", help="Pfad zur Variablen-Datei")
    parser.add_argument("--tarif-file", default="tarif.csv", help="Pfad zur Tarif-Parameter-Datei")
    parser.add_argument("--funcs", help="Kommagetrennte Funktionsliste, z.B. Bxt,BJB")
    parser.add_argument("--all", action="store_true", help="Alle Funktionen ausführen (default)")

    args = parser.parse_args()

    try:
        var_data_raw = load_input(args.var_file)
        tarif_data = load_input(args.tarif_file)
        var_data = {k: parse_numeric(v) for k, v in var_data_raw.items()}
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        return

    funcs_to_run = ALL_FUNCS if args.all or not args.funcs else args.funcs.split(",")

    aus = import_module("ausfunct")
    results = {}

    for name in funcs_to_run:
        func = getattr(aus, name, None)

        if func is None or func.__doc__ == "PLACEHOLDER":
            results[name] = "not yet implemented"
            continue

        try:
            input_args = {
                "vs": var_data["vs"],
                "age": var_data["x"],
                "sex": str(var_data["sex"]),
                "n": var_data["n"],
                "t": var_data["t"],
                "zw": var_data["zw"],
                "tarif": str(var_data.get("tarif", "KLV"))
            }
            results[name] = func(**input_args)
        except Exception as e:
            results[name] = f"error: {str(e)}"

    print(json.dumps(results, ensure_ascii=False, indent=None))

if __name__ == "__main__":
    main()
