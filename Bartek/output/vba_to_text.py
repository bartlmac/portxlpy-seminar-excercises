import re
from pathlib import Path
from oletools.olevba import VBA_Parser

EXCEL_PATH = Path("TARIFRECHNER_KLV.xlsm")

def vba_modules_to_txt(excel_path):
    out_dir = excel_path.parent
    parser = VBA_Parser(str(excel_path))
    for (subfilename, stream_path, vba_filename, vba_code) in parser.extract_macros():
        code_clean = vba_code.strip() if vba_code else ""
        if not code_clean:
            continue  # Skip empty modules
        if not re.search(r'\b(Sub|Function)\b', code_clean, re.IGNORECASE):
            continue  # Skip if no Sub/Function
        # Clean name for file system
        safe_name = re.sub(r"[^A-Za-z0-9_]", "_", vba_filename)
        out_path = out_dir / f"Mod_{safe_name}.txt"
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(code_clean)

if __name__ == "__main__":
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"{EXCEL_PATH} nicht gefunden.")
    vba_modules_to_txt(EXCEL_PATH)
