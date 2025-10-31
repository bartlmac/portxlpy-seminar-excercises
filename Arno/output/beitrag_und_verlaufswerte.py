# beitrag_und_verlaufswerte.py
# -*- coding: utf-8 -*-
"""
Berechnet Beitragswerte und Verlaufswerte für einen KLV-Tarif
gemäß den im Excel-Tarifrechner verwendeten Formeln.

Voraussetzungen:
- barwerte.Act_nGrAx, barwerte.Act_axn_k
- gwerte.Act_Dx

Autor: (du)
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import List, Dict, Any

# externe versicherungsmathematische Funktionen
from barwerte import Act_nGrAx, Act_axn_k
from gwerte import Act_Dx


@dataclass
class TarifInput:
    # Vertragsdaten
    x: int                  # Eintrittsalter
    sex: str                # 'M' / 'F' (oder wie in Deinen Funktionen erwartet)
    n: int                  # Vertragslaufzeit
    t: int                  # Beitragszahlungsdauer
    VS: float               # Versicherungssumme
    zw: int                 # Zahlungsweise (z.B. 12 = monatlich)

    # Tarifdaten
    zins: float             # Rechnungszins (dezimal, z.B. 0.0175)
    tafel: str              # Sterbetafel-Key, z.B. 'DAV1994_T'
    alpha: float            # Abschlusskosten (dezimal, z.B. 0.025)
    beta1: float            # Verwaltungskosten-Typ 1 (dezimal, z.B. 0.025)
    gamma1: float           # Überschuss-Parameter 1 (dezimal)
    gamma2: float           # Überschuss-Parameter 2 (dezimal)
    gamma3: float           # Überschuss-Parameter 3 (dezimal)
    k: float                # laufender Kostenfaktor (absolute Einheit)
    ratzu: float            # Ratenzuschlag (dezimal, z.B. 0.05)

    # Grenzen
    MinAlterFlex: int       # Mindestalter für Flexphase
    MinRLZFlex: int         # Mindest-Restlaufzeit für Flexphase

    # Sonstiges / Optionen
    rkw_min_storno: float = 50.0   # Untergrenze Stornoabzug
    rkw_max_storno: float = 150.0  # Obergrenze Stornoabzug
    rkw_storno_quote: float = 0.01 # 1% * (VS - kDRx_bpfl)


# ---------- Hilfsfunktionen ----------

def _safe_div(num: float, den: float, fallback: float = 0.0) -> float:
    """Division mit Fallback bei Division durch 0."""
    return num / den if den != 0 else fallback


# ---------- Beitragsberechnung ----------

def beitragsberechnung(inp: TarifInput) -> Dict[str, float]:
    """
    Repliziert die Excel-Formeln im Block 'Beitragsberechnung'.

    Rückgabe:
        dict mit Schlüsseln: Bxt, BJB, BZB, Pxt
    """

    x = inp.x
    n = inp.n
    t = inp.t
    sex = inp.sex
    tafel = inp.tafel
    zins = inp.zins

    # Excel: act_nGrAx(...) heißt in Python laut Vorgabe Act_nGrAx
    term1 = Act_nGrAx(x, n, sex, tafel, zins)

    # Act_Dx(x+n) / Act_Dx(x)
    term2 = _safe_div(
        Act_Dx(x + n, sex, tafel, zins),
        Act_Dx(x, sex, tafel, zins)
    )

    # gamma1*Act_axn_k(x; t; ...; 1)
    term3 = inp.gamma1 * Act_axn_k(x, t, sex, tafel, zins, 1)

    # gamma2*(Act_axn_k(x; n; ...; 1) - Act_axn_k(x; t; ...; 1))
    term4 = inp.gamma2 * (
        Act_axn_k(x, n, sex, tafel, zins, 1) - Act_axn_k(x, t, sex, tafel, zins, 1)
    )

    # Nenner: ((1 - beta1) * Act_axn_k(x; t; ...; 1) - alpha * t)
    denom = (1.0 - inp.beta1) * Act_axn_k(x, t, sex, tafel, zins, 1) - inp.alpha * t

    Bxt = _safe_div(term1 + term2 + term3 + term4, denom)

    # BJB = VS * K5; in Excel ist K5 = Bxt
    BJB = inp.VS * Bxt

    # BZB = (1 + ratzu)/zw * (K6 + k); in Excel ist K6 = BJB
    BZB = (1.0 + inp.ratzu) / inp.zw * (BJB + inp.k)

    # Pxt = (act_nGrAx(...) + Dx(x+n)/Dx(x) + t*alpha*B_xt) / Act_axn_k(x; t; ...; 1)
    # 'B_xt' ist hier Bxt (Beitragsfaktor)
    Pxt = _safe_div(
        Act_nGrAx(x, n, sex, tafel, zins)
        + _safe_div(Act_Dx(x + n, sex, tafel, zins), Act_Dx(x, sex, tafel, zins))
        + t * inp.alpha * Bxt,
        Act_axn_k(x, t, sex, tafel, zins, 1)
    )

    return {"Bxt": Bxt, "BJB": BJB, "BZB": BZB, "Pxt": Pxt}


# ---------- Verlaufswerte (Tabellenblock) ----------

def verlaufswerte(inp: TarifInput, k_max: int | None = None) -> List[Dict[str, Any]]:
    """
    Repliziert die Excel-Zeilen im Block 'Verlaufswerte'.

    Parameter:
        k_max: bis zu welcher Periode k gerechnet wird (inklusive).
               Standard: max(n, t) (du kannst das bei Bedarf höher setzen).

    Rückgabe:
        Liste von Zeilen-Dicts mit Schlüsseln:
        k, Axn, axn, axt, kVx_bpfl, kDRx_bpfl, kVx_bfr, kVx_MRV,
        flex_phase, StoAb, RKW, VS_bfr
    """

    if k_max is None:
        k_max = max(inp.n, inp.t)

    # Vorberechnung, die in mehreren Formeln gebraucht wird
    Axt_x_t = Act_axn_k(inp.x, inp.t, inp.sex, inp.tafel, inp.zins, 1)
    Axn_x_n = Act_axn_k(inp.x, inp.n, inp.sex, inp.tafel, inp.zins, 1)

    beitr = beitragsberechnung(inp)
    Pxt = beitr["Pxt"]
    BJB = beitr["BJB"]

    rows: List[Dict[str, Any]] = []

    for k in range(0, k_max + 1):

        # Excel: Axn = WENN(k <= n; Act_nGrAx(x+k; max(0; n-k)) + Dx(x+n)/Dx(x+k); 0)
        if k <= inp.n:
            Axn_val = (
                Act_nGrAx(inp.x + k, max(0, inp.n - k), inp.sex, inp.tafel, inp.zins)
                + _safe_div(
                    Act_Dx(inp.x + inp.n, inp.sex, inp.tafel, inp.zins),
                    Act_Dx(inp.x + k, inp.sex, inp.tafel, inp.zins),
                )
            )
        else:
            Axn_val = 0.0

        # axn = Act_axn_k(x+k; max(0; n-k); ...; 1)
        axn_val = Act_axn_k(inp.x + k, max(0, inp.n - k), inp.sex, inp.tafel, inp.zins, 1)

        # axt = Act_axn_k(x+k; max(0; t-k); ...; 1)
        axt_val = Act_axn_k(inp.x + k, max(0, inp.t - k), inp.sex, inp.tafel, inp.zins, 1)

        # kVx_bpfl = B - Pxt * D + gamma2 * (C - (Axn_x_n / Axt_x_t) * D)
        kVx_bpfl_val = (
            Axn_val
            - Pxt * axt_val
            + inp.gamma2 * (axn_val - _safe_div(Axn_x_n, Axt_x_t) * axt_val)
        )

        # kDRx_bpfl = VS * E
        kDRx_bpfl_val = inp.VS * kVx_bpfl_val

        # kVx_bfr = B + gamma3 * C
        kVx_bfr_val = Axn_val + inp.gamma3 * axn_val

        # kVx_MRV = F + alpha * t * BJB * Act_axn_k(x+k; max(5-k; 0); ...; 1) / Act_axn_k(x; 5; ...; 1)
        numerator = Act_axn_k(inp.x + k, max(5 - k, 0), inp.sex, inp.tafel, inp.zins, 1)
        denominator = Act_axn_k(inp.x, 5, inp.sex, inp.tafel, inp.zins, 1)
        kVx_MRV_val = kDRx_bpfl_val + inp.alpha * inp.t * BJB * _safe_div(numerator, denominator)

        # flex. Phase = WENN(UND(x+k >= MinAlterFlex; k >= n - MinRLZFlex); 1; 0)
        flex_phase_val = 1 if ((inp.x + k) >= inp.MinAlterFlex and k >= (inp.n - inp.MinRLZFlex)) else 0

        # StoAb = WENN(ODER(k > n; flex_phase); 0; MIN(150; MAX(50; 1%*(VS - F))))
        if (k > inp.n) or (flex_phase_val == 1):
            StoAb_val = 0.0
        else:
            raw = inp.rkw_storno_quote * (inp.VS - kDRx_bpfl_val)
            StoAb_val = max(inp.rkw_min_storno, min(inp.rkw_max_storno, raw))

        # RKW = MAX(0; H - J)
        RKW_val = max(0.0, kVx_MRV_val - StoAb_val)

        # VS_bfr = WENNFEHLER(WENN(k > n; 0; WENN(k < t; H/G; VS)); 0)
        if k > inp.n:
            VS_bfr_val = 0.0
        else:
            if k < inp.t:
                VS_bfr_val = _safe_div(kVx_MRV_val, kVx_bfr_val, 0.0)
            else:
                VS_bfr_val = inp.VS

        rows.append(
            {
                "k": k,
                "Axn": Axn_val,
                "axn": axn_val,
                "axt": axt_val,
                "kVx_bpfl": kVx_bpfl_val,
                "kDRx_bpfl": kDRx_bpfl_val,
                "kVx_bfr": kVx_bfr_val,
                "kVx_MRV": kVx_MRV_val,
                "flex_phase": flex_phase_val,
                "StoAb": StoAb_val,
                "RKW": RKW_val,
                "VS_bfr": VS_bfr_val,
            }
        )

    return rows


# ---------- Komfortfunktion: Alles auf einmal ----------

def berechne_alle(inp: TarifInput, k_max: int | None = None) -> Dict[str, Any]:
    """Beitrag und Verlaufswerte in einem Rutsch."""
    beitr = beitragsberechnung(inp)
    verlauf = verlaufswerte(inp, k_max=k_max)
    return {"beitrag": beitr, "verlauf": verlauf}
