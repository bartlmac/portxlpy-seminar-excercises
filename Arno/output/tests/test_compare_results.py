import sys
import pathlib

#  …/Arno/output in sys.path legen, damit vergleich importierbar ist
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))

import vergleich

def test_compare_results_no_diff():
    """Excel- und Python-Werte müssen identisch sein."""
    assert vergleich.main() == 0
