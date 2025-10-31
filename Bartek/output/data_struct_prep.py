import pandas as pd
import re
import pickle
from pathlib import Path
from typing import Dict, List

CELL_CSV = Path('excelzell.csv')
RANGE_CSV = Path('excelber.csv')
PKL_OUT = Path('tafeln.pkl')

def load_dataframes():
    cell_df = pd.read_csv(CELL_CSV, dtype=str).fillna("")
    range_df = pd.read_csv(RANGE_CSV, dtype=str).fillna("")
    return cell_df, range_df

def extract_references(formula: str, sheetnames: List[str]) -> List[str]:
    """
    Extrahiert Zell- und Bereichs-Referenzen (inklusive Blattbezug) aus einer Formel.
    Liefert Liste in der Schreibweise: Blatt!Adresse bzw. Blatt!Name.
    """
    if not formula or "=" not in formula:
        return []
    # Blätter: Kalkulation!A1, 'Kalkulation 2'!A1 etc.
    sheet_pat = r"(?:'[^']+'|[A-Za-z0-9_]+)?"
    # Zellen: A1, B2, Z$4, $A$1:$D$10 usw.
    addr_pat = r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?"
    # Bereichsnamen: keine Ziffer am Anfang, kein !, kein .xls, kein Operator
    name_pat = r"\b[A-Za-z_][A-Za-z0-9_\.]*\b"

    ref_pat = (
        # Blatt!Adresse
        rf"({sheet_pat})!({addr_pat})"
        "|"  # Bereichsnamen ohne !, aber keine Funktionen
        rf"\b({name_pat})\b"
    )
    refs = []
    # Fundstellen extrahieren
    for m in re.finditer(ref_pat, formula):
        if m.group(1) and m.group(2):  # Blatt!Adresse
            sheet = m.group(1).strip("'") if m.group(1) else ""
            addr = m.group(2)
            refs.append(f"{sheet}!{addr}" if sheet else addr)
        elif m.group(3):  # Bereichsname (potenziell)
            # Keine Excel-Funktion und kein bool/Operator
            token = m.group(3)
            if token.upper() in sheetnames:
                continue
            if token.upper() in EXCEL_FUNCTIONS:
                continue
            if token in {"TRUE", "FALSE"}:
                continue
            refs.append(token)
    return list(set(refs))  # Duplikate entfernen

EXCEL_FUNCTIONS = set("""
ABS ACOS ACOSH ADDRESS AND AREAS ASC ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE 
AVERAGEA AVERAGEIF AVERAGEIFS BASE BAHTTEXT BESSELI BESSELJ BESSELK BESSELY BETADIST BETA.DIST BETAINV BETA.INV 
BIN2DEC BIN2HEX BIN2OCT BINOMDIST BINOM.DIST BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING 
CEILING.MATH CEILING.PRECISE CHAR CHIDIST CHIINV CHITEST CHOOSE CLEAN CODE COLUMN COLUMNS COMBIN COMBINA COMPLEX 
CONCAT CONCATENATE CONFIDENCE CONFIDENCE.NORM CONFIDENCE.T TEST CONVERT CORREL COS COSH COT COTH COUNT COUNTA COUNTBLANK 
COUNTIF COUNTIFS COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR CRITBINOM CSC CSCH CUBEKPIMEMBER 
CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE 
DAVERAGE DAY DAYS DAYS360 DB DBCS DCOUNT DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DECIMAL DEGREES DELTA DEVSQ DGET DISC 
DMAX DMIN DOLLAR DOLLARDE DOLLARFR DPRODUCT DSTDEV DSTDEVP DSUM DURATION DVAR DVARP ECMA.CEILING EDATE EFFECT ENCODEURL 
EOMONTH ERF ERFC ERF.PRECISE ERFC.PRECISE ERROR.TYPE EURO EXACT EXP EXPON.DIST EXPONDIST FACT FACTDOUBLE FALSE FDIST 
FILTERXML FIND FINDB FISHER FISHERINV FIXED FLOOR FLOOR.MATH FLOOR.PRECISE FORECAST FREQUENCY FTEST FV FVSCHEDULE GAMMA 
GAMMADIST GAMMA.DIST GAMMAINV GAMMA.INV GAMMALN GAMMALN.PRECISE GAUSS GCD GEOMEAN GESTEP GETPIVOTDATA GROWTH HARMEAN 
HEX2BIN HEX2DEC HEX2OCT HLOOKUP HOUR HYPERLINK HYPGEOMDIST HYPGEOM.DIST IF IFERROR IFNA IMABS IMAGINARY IMARGUMENT 
IMCONJUGATE IMCOS IMCOSH IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSIN IMSINH IMSQRT 
IMSUB IMSUM INDEX INDIRECT INFO INT INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA 
ISNONTEXT ISNUMBER ISODD ISREF ISTEXT JIS KURT LARGE LCM LEFT LEN LINEST LN LOG LOG10 LOGEST LOGINV LOGNORM.DIST 
LOGNORMINV LOOKUP LOWER MATCH MAX MAXA MAXIFS MDETERM MDURATION MEDIAN MID MIN MINA MINIFS MINUTE MIRR MMULT MOD MODE 
MODE.MULT MODE.SNGL MONTH MROUND MULTINOMIAL N NA NEGBINOMDIST NEGBINOM.DIST NETWORKDAYS NOMINAL NORM.DIST NORM.INV 
NORM.S.DIST NORM.S.INV NOT NOW NPER NPV NUMBERVALUE OCT2BIN OCT2DEC OCT2HEX ODD OFFSET OR PDURATION PEARSON PERCENTILE 
PERCENTILE.EXC PERCENTILE.INC PERCENTRANK PERCENTRANK.EXC PERCENTRANK.INC PERMUT PERMUTATIONA PHI PI PMT POISSON 
POISSON.DIST POWER PPMT PRICE PRICEDISC PRICEMAT PROB PRODUCT PROPER PV QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT 
RADIANS RAND RANDBETWEEN RANK RANK.AVG RANK.EQ RATE RECEIVED REPLACE REPT RIGHT ROMAN ROUND ROUNDDOWN ROUNDUP ROW ROWS 
RRI RSQ RTD SEARCH SEARCHB SECOND SEC SECH SERIESSUM SHEET SHEETS SIGN SIN SINH SKEW SKEW.P SLN SLOPE SMALL SQRT 
SQRTPI STANDARDIZE STDEV STDEV.P STDEV.S STDEVA STDEVP STEYX SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ 
SUMX2MY2 SUMX2PY2 SUMXMY2 SYD T TAN TANH TBILLEQ TBILLPRICE TBILLYIELD T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T 
T.TEST TEXT TIME TIMEVALUE TINV TODAY TRANSPOSE TREND TRIM TRIMMEAN TRUE TRUNC UNICHAR UNICODE UPPER VALUE VAR 
VAR.P VAR.S VARA VARP VDB VLOOKUP WEBSERVICE WEEKDAY WEEKNUM WEIBULL WEIBULL.DIST WORKDAY XIRR XNPV XOR YEAR YEARFRAC 
Z.TEST
""".split())

def build_dep_map(cell_df, range_df):
    # Alle Blattnamen sammeln
    sheetnames = list(cell_df['Blatt'].unique())
    dep_map: Dict[str, List[str]] = {}
    for i, row in cell_df.iterrows():
        key = f"{row['Blatt']}!{row['Adresse']}"
        formula = row['Formel']
        refs = extract_references(formula, sheetnames)
        dep_map[key] = refs
    return dep_map

if __name__ == "__main__":
    cell_df, range_df = load_dataframes()
    dep_map = build_dep_map(cell_df, range_df)
    tafeln = {
        "cell_df": cell_df,
        "range_df": range_df,
        "dep_map": dep_map,
    }
    with open(PKL_OUT, "wb") as f:
        pickle.dump(tafeln, f)
