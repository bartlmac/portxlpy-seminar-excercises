# gwerte.py
"""
Python-Umsetzung des VBA-Moduls mGWerte.

Voraussetzungen:
- Eine Datei "Tafeln.xml" liegt im selben Verzeichnis wie dieses Modul.
- Ein Modul `constants.py` existiert mit:
    * RUND_LX, RUND_TX, RUND_DX, RUND_CX, RUND_NX, RUND_MX, RUND_RX, MAX_ALTER
    * vba_round(value, digits)  -> führt Banker's Rounding durch (wie VBA)

Hinweis:
- Die Funktionen sind eng an die VBA-Signaturen angelehnt.
- Optional-Parameter, welche in VBA default sind, sind in Python als default None (oder 1 für Schicht) realisiert.
"""

from typing import Optional, Dict, List, Any
import xml.etree.ElementTree as ET
import os
import datetime

import constants as const  # dein bereits erstelltes constants.py

# --- Globaler Cache (entspricht VBA `cache` Dictionary) ---
_cache: Dict[str, float] = None

# --- Tafeln-Daten (geladen bei erstem Zugriff) ---
# Struktur: tafeln_keys = list von Spaltennamen (z.B. "DAV1994_T_M", ...)
#           tafeln_values: dict mapping key -> list_of_qx indexed by age (0..max_Alter)
_tafeln_loaded = False
_tafeln_keys: List[str] = []
_tafeln_values: Dict[str, List[float]] = {}
_max_age_in_file = 0


def InitializeCache():
    """Initialisiert das globale Cache-Dictionary (wie VBA CreateObject("Scripting.Dictionary"))."""
    global _cache
    _cache = {}


def _ensure_cache():
    global _cache
    if _cache is None:
        InitializeCache()


# --- XML einlesen ---
def _load_tafeln(xml_path: str = None):
    """
    Lädt Tafeln.xml und befüllt _tafeln_keys und _tafeln_values.
    Erwartet Datei im selben Verzeichnis wie dieses Modul, falls xml_path None.
    Die XML benutzt Komma als Dezimaltrennzeichen (z.B. '0,01168700') — das wird konvertiert.
    """
    global _tafeln_loaded, _tafeln_keys, _tafeln_values, _max_age_in_file

    if _tafeln_loaded:
        return

    if xml_path is None:
        # Datei im gleichen Ordner wie dieses Modul
        base = os.path.dirname(__file__)
        xml_path = os.path.join(base, "Tafeln.xml")

    if not os.path.exists(xml_path):
        raise FileNotFoundError(f"Tafeln-Datei nicht gefunden: {xml_path}")

    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Annahme: <dataset><record> ... </record> ... </dataset>
    records = root.findall(".//record")
    # Wenn keine records: Datei fehlerhaft
    if not records:
        raise ValueError("Tafeln.xml enthält keine <record>-Einträge.")

    # Bestimme Spaltennamen aus dem ersten Record (alles außer 'xy')
    first = records[0]
    cols = [child.tag for child in first if child.tag.lower() != "xy"]
    _tafeln_keys = cols.copy()
    # Init lists for each key, size at least const.MAX_ALTER+1 (fill later)
    _tafeln_values = {k: [] for k in _tafeln_keys}
    _max_age_in_file = 0

    # temporäre dict mapping age -> dict of key->value
    temp: Dict[int, Dict[str, float]] = {}

    for rec in records:
        # find xy
        xy_elem = rec.find("xy")
        if xy_elem is None or xy_elem.text is None:
            continue
        try:
            age = int(xy_elem.text.strip())
        except Exception:
            continue

        if age > _max_age_in_file:
            _max_age_in_file = age

        temp[age] = {}
        for child in rec:
            tag = child.tag
            if tag.lower() == "xy":
                continue
            text = child.text.strip() if child.text is not None else ""
            # replace comma decimal with dot
            text = text.replace(".", "").replace(",", ".") if text != "" else "0"
            try:
                val = float(text)
            except Exception:
                val = 0.0
            temp[age][tag] = val

    # Now build lists for all keys up to max between file and const.MAX_ALTER
    max_fill = max(_max_age_in_file, const.MAX_ALTER)
    for key in _tafeln_keys:
        lst = []
        for age in range(0, max_fill + 1):
            if age in temp and key in temp[age]:
                lst.append(temp[age][key])
            else:
                # missing age: fill with 0.0 (sauberer Default)
                lst.append(0.0)
        _tafeln_values[key] = lst

    _tafeln_loaded = True


def _ensure_tafeln_loaded():
    if not _tafeln_loaded:
        _load_tafeln()


# --- Hilfsfunktion: CreateCacheKey (wie VBA) ---
def CreateCacheKey(Art: str, Alter: int, Sex: str, Tafel: str, Zins: float,
                   GebJahr: Optional[int], Rentenbeginnalter: Optional[int], Schicht: int) -> str:
    # In VBA: Art & "_" & Alter & "_" & Sex & "_" & Tafel & "_" & Zins & "_" & GebJahr & "_" & Rentenbeginnalter & "_" & Schicht
    # Wir wandeln None in 0 (oder leeren String) ähnlich wie VBA leerer Wert. Um konsistent zu bleiben, nutzen wir '' für None.
    def none_to_str(x):
        return "" if x is None else str(x)
    return f"{Art}_{Alter}_{Sex}_{Tafel}_{Zins}_{none_to_str(GebJahr)}_{none_to_str(Rentenbeginnalter)}_{Schicht}"


# --- Act_qx ---
def Act_qx(Alter: int, Sex: str, Tafel: str, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    """
    Liefert q_x aus der Tafeldatei.
    Entspricht VBA-Index für Range("m_Tafeln") und Match in v_Tafeln.
    """
    _ensure_tafeln_loaded()

    sex = Sex.upper()
    if sex != "M":
        sex = "F"

    taf = Tafel.upper()
    # nur bestimmte Tafeln implementiert wie im VBA (hier Beispiel: DAV1994_T, DAV2008_T)
    # In VBA wurden in Select Case nur "DAV1994_T", "DAV2008_T" aufgeführt.
    if taf not in ("DAV1994_T", "DAV2008_T"):
        # VBA: Act_qx = 1# : Error (1)
        # Wir setzen Python-konformes Verhalten: ValueError mit Verhalten wie VBA-Error.
        raise ValueError(f"Tafel '{Tafel}' nicht implementiert.")

    sTafelvektor = f"{taf}_{sex}"
    # Prüfen ob key existiert
    if sTafelvektor not in _tafeln_values:
        # sollte eigentlich nicht passieren, es sei denn Tafeln.xml hat andere Tags
        raise KeyError(f"Tafelvektor '{sTafelvektor}' nicht gefunden in Tafeln.xml.")

    # Sicherstellen, dass Alter innerhalb Bereich ist
    if Alter < 0:
        raise ValueError("Alter darf nicht negativ sein.")
    # wenn Alter größer als gelieferte Liste, return 0.0 (oder raise?) - VBA würde Fehler vermeiden, weil Tabelle ggf. bis max_Alter reicht
    values = _tafeln_values[sTafelvektor]
    if Alter >= len(values):
        # außerhalb: 0.0 zurückgeben
        return 0.0
    return values[Alter]


# --- v_lx ---
def v_lx(Endalter: int, Sex: str, Tafel: str, GebJahr: Optional[int] = None,
         Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> List[float]:
    """
    Erzeugt Vektor der lx (Anfangspopulation bei jedem Alter).
    Wenn Endalter = -1 wird bis const.MAX_ALTER erzeugt.
    lx(0) = 1_000_000 (wie in VBA)
    Rundung: const.RUND_LX Stellen (via const.vba_round)
    """
    _ensure_tafeln_loaded()

    if Endalter == -1:
        Grenze = const.MAX_ALTER
    else:
        Grenze = Endalter

    vek = [0.0] * (Grenze + 1)
    vek[0] = 1000000.0
    for i in range(1, Grenze + 1):
        q_prev = Act_qx(i - 1, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
        vek[i] = vek[i - 1] * (1.0 - q_prev)
        # Rundung
        vek[i] = const.vba_round(vek[i], const.RUND_LX)
    return vek


def Act_lx(Alter: int, Sex: str, Tafel: str, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    vek = v_lx(Alter, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    return vek[Alter]


# --- v_tx (#Tote) ---
def v_tx(Endalter: int, Sex: str, Tafel: str, GebJahr: Optional[int] = None,
         Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> List[float]:
    if Endalter == -1:
        Grenze = const.MAX_ALTER
    else:
        Grenze = Endalter

    v_Temp_lx = v_lx(Grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    vek = [0.0] * (Grenze + 1)
    # in VBA: For i = 0 To Grenze - 1: vek(i) = v_Temp_lx(i) - v_Temp_lx(i + 1)
    for i in range(0, Grenze):
        val = v_Temp_lx[i] - v_Temp_lx[i + 1]
        val = const.vba_round(val, const.RUND_TX)
        vek[i] = val
    # last element (index Grenze) bleibt 0.0 (wie VBA ReDim vek(Gr) und loop 0..Gr-1)
    return vek


def Act_tx(Alter: int, Sex: str, Tafel: str, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    vek = v_tx(Alter, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    return vek[Alter]


# --- v_Dx ---
def v_Dx(Endalter: int, Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
         Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> List[float]:
    if Endalter == -1:
        Grenze = const.MAX_ALTER
    else:
        Grenze = Endalter

    vek = [0.0] * (Grenze + 1)
    v_factor = 1.0 / (1.0 + Zins)

    v_Temp_lx = v_lx(Grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    for i in range(0, Grenze + 1):
        val = v_Temp_lx[i] * (v_factor ** i)
        val = const.vba_round(val, const.RUND_DX)
        vek[i] = val
    return vek


def Act_Dx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    _ensure_cache()
    sKey = CreateCacheKey("Dx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    if sKey in _cache:
        return _cache[sKey]
    vek = v_Dx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    result = vek[Alter]
    _cache[sKey] = result
    return result


# --- v_Cx ---
def v_Cx(Endalter: int, Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
         Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> List[float]:
    if Endalter == -1:
        Grenze = const.MAX_ALTER
    else:
        Grenze = Endalter

    vek = [0.0] * (Grenze + 1)
    v_factor = 1.0 / (1.0 + Zins)

    v_Temp_tx = v_tx(Grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    for i in range(0, Grenze):
        val = v_Temp_tx[i] * (v_factor ** (i + 1))
        val = const.vba_round(val, const.RUND_CX)
        vek[i] = val
    # last index remains 0.0 (VBA loop 0..Gr-1)
    return vek


def Act_Cx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    _ensure_cache()
    sKey = CreateCacheKey("Cx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    if sKey in _cache:
        return _cache[sKey]
    vek = v_Cx(Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    result = vek[Alter]
    _cache[sKey] = result
    return result


# --- v_Nx ---
def v_Nx(Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
         Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> List[float]:
    # erzeugt Vektor der Nx
    Grenze = const.MAX_ALTER
    v_Temp_Dx = v_Dx(-1, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    vek = [0.0] * (Grenze + 1)
    vek[Grenze] = v_Temp_Dx[Grenze]
    for i in range(Grenze - 1, -1, -1):
        val = vek[i + 1] + v_Temp_Dx[i]
        val = const.vba_round(val, const.RUND_DX)
        vek[i] = val
    return vek


def Act_Nx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    _ensure_cache()
    sKey = CreateCacheKey("Nx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    if sKey in _cache:
        return _cache[sKey]
    vek = v_Nx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    result = vek[Alter]
    _cache[sKey] = result
    return result


# --- v_Mx ---
def v_Mx(Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
         Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> List[float]:
    Grenze = const.MAX_ALTER
    v_Temp_Cx = v_Cx(-1, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    vek = [0.0] * (Grenze + 1)
    vek[Grenze] = v_Temp_Cx[Grenze]
    for i in range(Grenze - 1, -1, -1):
        val = vek[i + 1] + v_Temp_Cx[i]
        val = const.vba_round(val, const.RUND_MX)
        vek[i] = val
    return vek


def Act_Mx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    _ensure_cache()
    sKey = CreateCacheKey("Mx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    if sKey in _cache:
        return _cache[sKey]
    vek = v_Mx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    result = vek[Alter]
    _cache[sKey] = result
    return result


# --- v_Rx ---
def v_Rx(Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
         Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> List[float]:
    Grenze = const.MAX_ALTER
    v_Temp_Mx = v_Mx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    vek = [0.0] * (Grenze + 1)
    vek[Grenze] = v_Temp_Mx[Grenze]
    for i in range(Grenze - 1, -1, -1):
        val = vek[i + 1] + v_Temp_Mx[i]
        val = const.vba_round(val, const.RUND_RX)
        vek[i] = val
    return vek


def Act_Rx(Alter: int, Sex: str, Tafel: str, Zins: float, GebJahr: Optional[int] = None,
           Rentenbeginnalter: Optional[int] = None, Schicht: int = 1) -> float:
    _ensure_cache()
    sKey = CreateCacheKey("Rx", Alter, Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    if sKey in _cache:
        return _cache[sKey]
    vek = v_Rx(Sex, Tafel, Zins, GebJahr, Rentenbeginnalter, Schicht)
    result = vek[Alter]
    _cache[sKey] = result
    return result


# --- Altersberechnung ---
def Act_Altersberechnung(GebDat: Any, BerDat: Any, Methode: str) -> int:
    """
    Altersberechnung nach Kalenderjahresmethode (K) bzw. Halbjahresmethode (H).

    GebDat, BerDat können entweder datetime.date/datetime.datetime-Objekte oder
    ISO-Datumsstrings ('YYYY-MM-DD') sein.
    """
    # parse dates falls strings
    def to_date(d):
        if isinstance(d, (datetime.date, datetime.datetime)):
            if isinstance(d, datetime.datetime):
                return d.date()
            return d
        if isinstance(d, str):
            # akzeptiere 'YYYY-MM-DD' oder 'YYYY-MM-DDTHH:MM:SS'
            try:
                return datetime.date.fromisoformat(d)
            except Exception:
                raise ValueError("Datumformat nicht erkannt. Erwarte ISO-Format 'YYYY-MM-DD'.")
        raise TypeError("GebDat/BerDat müssen date/time oder ISO-String sein.")

    Geb = to_date(GebDat)
    Ber = to_date(BerDat)

    Methode_local = Methode
    if Methode_local != "K":
        Methode_local = "H"

    J_GD = Geb.year
    J_BD = Ber.year
    M_GD = Geb.month
    M_BD = Ber.month

    if Methode_local == "K":
        return J_BD - J_GD
    else:
        # Int(J_BD - J_GD + 1# / 12# * (M_BD - M_GD + 5))
        val = int((J_BD - J_GD) + (1.0 / 12.0) * (M_BD - M_GD + 5))
        return val
