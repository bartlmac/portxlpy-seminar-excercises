import os
from pathlib import Path
import pytest

# ---------------------------------------------------------------------------
# 1) Arbeitsverzeichnis **einmalig und dauerhaft** auf .../Bartek/output setzen
#    (kein Zurückwechseln nach Testende – genau so gewünscht)
# ---------------------------------------------------------------------------
os.chdir(Path(__file__).resolve().parents[1])   # → .../Bartek/output

# ---------------------------------------------------------------------------
# 2) Optional: Fixture mit Minimal-CSVs für einzelne Tests
# ---------------------------------------------------------------------------
@pytest.fixture(scope="session")
def temp_csv_dir(tmp_path_factory):
    """
    Liefert ein temporäres Verzeichnis mit Minimal-CSVs,
    falls Tests bewusst mit Dummy-Daten arbeiten wollen.
    """
    tmpdir = tmp_path_factory.mktemp("csv_data")

    (tmpdir / "var.csv").write_text("Name,Wert\nAlter,42\n", encoding="utf-8")
    (tmpdir / "tarif.csv").write_text("Tarif,Wert\nStandard,1.5\n", encoding="utf-8")
    (tmpdir / "grenzen.csv").write_text("Grenze,Wert\nMax,1000\n", encoding="utf-8")
    (tmpdir / "tafeln.csv").write_text(
        "x/y,DAV1994_T_M,DAV1994_T_F,DAV2008_T_M,DAV2008_T_F\n"
        "0,0.01,0.02,0.03,0.04\n",
        encoding="utf-8",
    )

    return tmpdir
