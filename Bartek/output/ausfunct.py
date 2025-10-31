# ausfunct.py

import pandas as pd
from basfunct import (
    Act_nGrAx, Act_Dx, Act_axn_k
)

def _load_var(name: str):
    df = pd.read_csv("var.csv")
    row = df[df["Name"] == name]
    if row.empty:
        raise ValueError(f"Variable {name} nicht in var.csv gefunden")
    return row["Wert"].values[0]

def _load_tarif(name: str):
    df = pd.read_csv("tarif.csv")
    row = df[df.columns[[0, 1]]]  # Spaltenname, Wert
    val = row.loc[row.iloc[:, 0] == name]
    if val.empty:
        raise ValueError(f"Tarifwert {name} nicht gefunden")
    return val.iloc[0, 1]

def Bxt(vs, age, sex, n, t, zw, tarif):
    # Tarifparameter laden
    alpha = float(_load_tarif("alpha"))
    beta1 = float(_load_tarif("beta1"))
    gamma1 = float(_load_tarif("gamma1"))
    gamma2 = float(_load_tarif("gamma2"))
    tafel = str(_load_tarif("Tafel"))
    zins = float(_load_tarif("Zins"))

    # Formel umsetzen
    z1 = Act_nGrAx(age, n, sex, tafel, zins)
    z2 = Act_Dx(age + n, sex, tafel, zins) / Act_Dx(age, sex, tafel, zins)
    z3 = gamma1 * Act_axn_k(age, t, sex, tafel, zins, 1)
    z4 = gamma2 * (Act_axn_k(age, n, sex, tafel, zins, 1) - Act_axn_k(age, t, sex, tafel, zins, 1))
    nenner = (1 - beta1) * Act_axn_k(age, t, sex, tafel, zins, 1) - alpha * t

    return (z1 + z2 + z3 + z4) / nenner
