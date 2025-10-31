# barwerte.py
"""
Modul mBarwerte – Übersetzung des gleichnamigen VBA-Moduls in Python.
"""

from constants import vba_round
from gwerte import Act_Dx, Act_Nx, Act_Mx

# ===========================================================
# Hilfsfunktion: Abzugsglied
# ===========================================================

def Act_Abzugsglied(k: int, zins: float) -> float:

    if k <= 0:
        return 0.0

    abzug = 0.0
    for l in range(k):
        abzug += (l / k) / (1 + (l / k) * zins)

    abzug = abzug * (1 + zins) / k
    return abzug


# ===========================================================
# Barwertfunktionen
# ===========================================================

def Act_ax_k(alter: int, sex: str, tafel: str, zins: float, k: int,
             geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:

    if k > 0:
        wert = (Act_Nx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
                / Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
                - Act_Abzugsglied(k, zins))
    else:
        wert = 0.0

    return vba_round(wert, 16)


def Act_axn_k(alter: int, n: int, sex: str, tafel: str, zins: float, k: int,
              geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:

    if k > 0:
        dx_alt = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
        dx_alt_n = Act_Dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
        nx_alt = Act_Nx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
        nx_alt_n = Act_Nx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

        abzug = Act_Abzugsglied(k, zins)

        wert = ((nx_alt - nx_alt_n) / dx_alt
                - abzug * (1 - dx_alt_n / dx_alt))
    else:
        wert = 0.0

    return vba_round(wert, 16)


def Act_nax_k(alter: int, n: int, sex: str, tafel: str, zins: float, k: int,
              geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:

    if k > 0:
        dx_alt = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
        dx_alt_n = Act_Dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

        wert = (dx_alt_n / dx_alt) * Act_ax_k(alter + n, sex, tafel, zins, k,
                                              geb_jahr, rentenbeginnalter, schicht)
    else:
        wert = 0.0

    return vba_round(wert, 16)


def Act_nGrAx(alter: int, n: int, sex: str, tafel: str, zins: float,
              geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:

    mx_alt = Act_Mx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    mx_alt_n = Act_Mx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    dx_alt = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

    wert = (mx_alt - mx_alt_n) / dx_alt
    return vba_round(wert, 16)


def Act_nGrEx(alter: int, n: int, sex: str, tafel: str, zins: float,
              geb_jahr: int = None, rentenbeginnalter: int = None, schicht: int = 1) -> float:

    dx_alt = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    dx_alt_n = Act_Dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)

    wert = dx_alt_n / dx_alt
    return vba_round(wert, 16)


def Act_ag_k(g: int, zins: float, k: int) -> float:

    if k <= 0:
        return 0.0

    v = 1 / (1 + zins)

    if zins > 0:
        wert = ((1 - v ** g) / (1 - v)
                - Act_Abzugsglied(k, zins) * (1 - v ** g))
    else:
        wert = float(g)

    return vba_round(wert, 16)
