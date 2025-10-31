# basfunct.py

import pandas as pd
import math
from datetime import date

# Konstanten
rund_lx = 16
rund_tx = 16
rund_Dx = 16
rund_Cx = 16
rund_Nx = 16
rund_Mx = 16
rund_Rx = 16
max_Alter = 123

# Cache
_cache = {}

def _round(val, digits):
    return round(val, digits)

def _key(*args):
    return "_".join(str(a) for a in args)

def _tafel_col(tafel, sex):
    return f"{tafel.upper()}_{'M' if sex.upper() == 'M' else 'F'}"

def Act_qx(alter, sex, tafel, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    tafeln = pd.read_csv("tafeln.csv")
    col = _tafel_col(tafel, sex)
    if col not in tafeln.columns:
        return 1.0
    return tafeln.at[alter, col]

def v_lx(endalter, sex, tafel, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    grenze = max_Alter if endalter == -1 else endalter
    vek = [1_000_000.0]
    for i in range(1, grenze + 1):
        qx = Act_qx(i - 1, sex, tafel, geb_jahr, rentenbeginnalter, schicht)
        vek.append(_round(vek[i - 1] * (1 - qx), rund_lx))
    return vek

def Act_lx(alter, sex, tafel, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    return v_lx(alter, sex, tafel, geb_jahr, rentenbeginnalter, schicht)[alter]

def v_tx(endalter, sex, tafel, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    grenze = max_Alter if endalter == -1 else endalter
    lx = v_lx(grenze + 1, sex, tafel, geb_jahr, rentenbeginnalter, schicht)
    vek = [_round(lx[i] - lx[i + 1], rund_tx) for i in range(grenze)]
    vek.append(0.0)  # Index max_Alter auffüllen
    return vek

def Act_tx(alter, sex, tafel, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    return v_tx(alter + 1, sex, tafel, geb_jahr, rentenbeginnalter, schicht)[alter]

def v_Dx(endalter, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    grenze = max_Alter if endalter == -1 else endalter
    lx = v_lx(grenze, sex, tafel, geb_jahr, rentenbeginnalter, schicht)
    v = 1 / (1 + zins)
    return [_round(lx[i] * v ** i, rund_Dx) for i in range(grenze + 1)]

def Act_Dx(alter, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    k = _key("Dx", alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    if k not in _cache:
        _cache[k] = v_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)[alter]
    return _cache[k]

def v_Cx(endalter, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    grenze = max_Alter if endalter == -1 else endalter
    tx = v_tx(grenze, sex, tafel, geb_jahr, rentenbeginnalter, schicht)
    v = 1 / (1 + zins)
    vek = [_round(tx[i] * v ** (i + 1), rund_Cx) for i in range(grenze)]
    vek.append(0.0)  # Stelle [max_Alter] explizit auffüllen
    return vek

def Act_Cx(alter, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    k = _key("Cx", alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    if k not in _cache:
        _cache[k] = v_Cx(alter + 1, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)[alter]
    return _cache[k]

def v_Nx(sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    Dx = v_Dx(-1, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    vek = [0.0] * (max_Alter + 1)
    vek[max_Alter] = Dx[max_Alter]
    for i in reversed(range(max_Alter)):
        vek[i] = _round(vek[i + 1] + Dx[i], rund_Dx)
    return vek

def Act_Nx(alter, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    k = _key("Nx", alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    if k not in _cache:
        _cache[k] = v_Nx(sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)[alter]
    return _cache[k]

def v_Mx(sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    Cx = v_Cx(-1, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    vek = [0.0] * (max_Alter + 1)
    vek[max_Alter] = Cx[max_Alter]
    for i in reversed(range(max_Alter)):
        vek[i] = _round(vek[i + 1] + Cx[i], rund_Mx)
    return vek

def Act_Mx(alter, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    k = _key("Mx", alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    if k not in _cache:
        _cache[k] = v_Mx(sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)[alter]
    return _cache[k]

def v_Rx(sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    Mx = v_Mx(sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    vek = [0.0] * (max_Alter + 1)
    vek[max_Alter] = Mx[max_Alter]
    for i in reversed(range(max_Alter)):
        vek[i] = _round(vek[i + 1] + Mx[i], rund_Rx)
    return vek

def Act_Rx(alter, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    k = _key("Rx", alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    if k not in _cache:
        _cache[k] = v_Rx(sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)[alter]
    return _cache[k]

def Act_Abzugsglied(k, zins):
    if k <= 0:
        return 0.0
    summ = sum(l / k / (1 + l / k * zins) for l in range(k))
    return summ * (1 + zins) / k

def Act_ag_k(g, zins, k):
    if k <= 0:
        return 0.0
    if zins > 0:
        v = 1 / (1 + zins)
        return (1 - v ** g) / (1 - v) - Act_Abzugsglied(k, zins) * (1 - v ** g)
    return g

def Act_ax_k(alter, sex, tafel, zins, k, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    if k <= 0:
        return 0.0
    return Act_Nx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht) / \
           Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht) - \
           Act_Abzugsglied(k, zins)

def Act_axn_k(alter, n, sex, tafel, zins, k, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    if k <= 0:
        return 0.0
    dx = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    dxn = Act_Dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    nx = Act_Nx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    nxn = Act_Nx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    return (nx - nxn) / dx - Act_Abzugsglied(k, zins) * (1 - dxn / dx)

def Act_nax_k(alter, n, sex, tafel, zins, k, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    if k <= 0:
        return 0.0
    dx = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    dxn = Act_Dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    return dxn / dx * Act_ax_k(alter + n, sex, tafel, zins, k, geb_jahr, rentenbeginnalter, schicht)

def Act_nGrAx(alter, n, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    mx = Act_Mx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    mxn = Act_Mx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    dx = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    return (mx - mxn) / dx

def Act_nGrEx(alter, n, sex, tafel, zins, geb_jahr=None, rentenbeginnalter=None, schicht=1):
    dx = Act_Dx(alter, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    dxn = Act_Dx(alter + n, sex, tafel, zins, geb_jahr, rentenbeginnalter, schicht)
    return dxn / dx

def Act_Altersberechnung(gebdat: date, berdat: date, methode: str):
    if methode != "K":
        return int((berdat.year - gebdat.year) + (1 / 12) * (berdat.month - gebdat.month + 5))
    else:
        return berdat.year - gebdat.year
