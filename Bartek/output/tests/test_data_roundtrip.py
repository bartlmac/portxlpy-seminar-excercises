import os
import pandas as pd

def test_csv_columns(temp_csv_dir):
    # Erwarten: var=2, tarif=2, grenzen=2, tafeln=5 Spalten (laut Header)
    expected = {
        "var.csv": 2,
        "tarif.csv": 2,
        "grenzen.csv": 2,
        "tafeln.csv": 5,
    }
    for fname, ncols in expected.items():
        fpath = temp_csv_dir / fname
        df    = pd.read_csv(fpath)

        # Print für den XML-Report
        print(f"{fname}: erwartet {ncols} Spalten | gefunden {df.shape[1]}")

        assert df.shape[1] == ncols, f"{fname}: erwartet {ncols} Spalten, gefunden {df.shape[1]}"