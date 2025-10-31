import pandas as pd
import csv

def col2num(col):
    num = 0
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

def num2col(n):
    s = ""
    while n > 0:
        n, r = divmod(n-1, 26)
        s = chr(65 + r) + s
    return s

def addr_naked(addr):
    import re
    m = re.match(r"\$?([A-Za-z]+)\$?([0-9]+)", str(addr))
    return f"{m.group(1).upper()}{m.group(2)}" if m else str(addr).upper()

def get_row_col(address):
    import re
    m = re.match(r"\$?([A-Z]+)\$?([0-9]+)", str(address))
    if m:
        return int(m.group(2)), col2num(m.group(1))
    else:
        return None, None

def extract_table(df, sheet, col_start, col_end, header_row, data_row_start, data_row_end=None):
    start_c = col2num(col_start)
    end_c = col2num(col_end)
    # Header aus header_row
    headers = []
    for col in range(start_c, end_c + 1):
        addr = f"{num2col(col)}{header_row}"
        row = df[(df['Blatt'] == sheet) & (df['Adresse'].apply(lambda x: addr_naked(x) == addr))]
        val = row['Wert'].values[0] if not row.empty else ""
        headers.append(val)
    # Daten ab data_row_start bis data_row_end (oder bis max)
    relevant = df[df['Blatt'] == sheet].copy()
    relevant[['row', 'col']] = relevant['Adresse'].map(get_row_col).apply(pd.Series)
    relevant = relevant[(relevant['col'] >= start_c) & (relevant['col'] <= end_c)]
    if data_row_end is None:
        data_row_end = relevant['row'].max()
    data_rows = []
    for r in range(data_row_start, data_row_end + 1):
        vals = []
        for c in range(start_c, end_c + 1):
            addr = f"{num2col(c)}{r}"
            row = df[(df['Blatt'] == sheet) & (df['Adresse'].apply(lambda x: addr_naked(x) == addr))]
            val = row['Wert'].values[0] if not row.empty else ""
            vals.append(val)
        if any(str(x).strip() for x in vals):
            data_rows.append(vals)
    return headers, data_rows

def write_csv(headers, data_rows, filename):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(data_rows)

def main():
    df = pd.read_csv("excelzell.csv")

    # var.csv: Kalkulation!A4:B9 (Header: 4, Daten: 5-9)
    headers, data_rows = extract_table(df, "Kalkulation", "A", "B", 4, 5, 9)
    write_csv(headers, data_rows, "var.csv")

    # tarif.csv: Kalkulation!D4:E11 (Header: 4, Daten: 5-11)
    headers, data_rows = extract_table(df, "Kalkulation", "D", "E", 4, 5, 11)
    write_csv(headers, data_rows, "tarif.csv")

    # grenzen.csv: Kalkulation!G4:H5 (Header: 4, Daten: 5-5)
    headers, data_rows = extract_table(df, "Kalkulation", "G", "H", 4, 5, 5)
    write_csv(headers, data_rows, "grenzen.csv")

    # tafeln.csv: Tafeln!A3:E... (Header: 3, Daten: 4-max)
    headers, data_rows = extract_table(df, "Tafeln", "A", "E", 3, 4)
    write_csv(headers, data_rows, "tafeln.csv")

    # tarif.py – Formel aus Kalkulation!E12 als Python
    e12 = df[(df['Blatt'] == "Kalkulation") & (df['Adresse'].apply(lambda x: addr_naked(x) == "E12"))]
    if e12.empty or pd.isna(e12['Formel'].values[0]):
        raise Exception("Keine Formel in Kalkulation!E12 gefunden.")
    formel = e12['Formel'].values[0].strip("=")
    if formel.upper().startswith("WENN"):
        import re
        m = re.match(r"WENN\((.*?);(.*?);(.*)\)", formel, re.IGNORECASE)
        if not m:
            raise Exception("WENN-Formel in E12 nicht erkannt.")
        cond, iftrue, iffalse = m.groups()
        cond = cond.replace("E11", "zw").replace(",", ".")
        iftrue = iftrue.replace(",", ".")
        iffalse = iffalse.replace(",", ".")
        pyfunc = f"def raten_zuschlag(zw):\n    return {iftrue} if ({cond}) else {iffalse}\n"
    else:
        pyfunc = f"def raten_zuschlag(zw):\n    return {formel}\n"
    with open("tarif.py", "w", encoding="utf-8") as f:
        f.write("# Automatisch generiert\n")
        f.write(pyfunc)

if __name__ == "__main__":
    main()
