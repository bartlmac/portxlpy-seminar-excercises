# constants.py
"""
Modul: constants
Beschreibung:
    Enthält globale Konstanten und zentrale Rundungsfunktion für den Tarifrechner.
    Entspricht dem VBA-Modul 'mConstants'.
"""

# --- Globale Konstanten ---
RUND_LX: int = 16
RUND_TX: int = 16
RUND_DX: int = 16
RUND_CX: int = 16
RUND_NX: int = 16
RUND_MX: int = 16
RUND_RX: int = 16
MAX_ALTER: int = 123


# --- Rundungsfunktion ---
def vba_round(value: float, digits: int = 16) -> float:
    """
    Rundet eine Zahl nach dem VBA-Prinzip (Banker's Rounding).

    Parameter:
        value : float  – zu rundender Wert
        digits: int    – Anzahl Dezimalstellen (Standard = 16)

    Rückgabe:
        float – gerundeter Wert
    """
    if value is None:
        return None
    return round(value, digits)
