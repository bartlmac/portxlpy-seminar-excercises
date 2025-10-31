# tests/test_func_parity.py

import re
import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]  # …\davag\08
sys.path.insert(0, str(ROOT))

def extract_vba_functions(mod_dir):
    vba_names = set()
    for file in Path(mod_dir).glob("Mod_*.txt"):
        content = file.read_text(encoding="utf-8")
        for match in re.finditer(
                r'^\s*(?:Public|Private)?\s*(Function|Sub)\s+(\w+)',
                content,
                flags=re.MULTILINE | re.IGNORECASE):
            vba_names.add(match.group(2))
    return vba_names

def extract_python_functions(py_path):
    with open(py_path, "r", encoding="utf-8") as f:
        tree = ast.parse(f.read(), filename=str(py_path))
    return {node.name for node in tree.body if isinstance(node, ast.FunctionDef)}

def test_function_parity():
    all_vba_funcs   = extract_vba_functions(mod_dir=".")
    # Cache-Funktionen sind rein interne Helfer – ausklammern
    vba_funcs       = {n for n in all_vba_funcs if "Cache" not in n}
    ignored_cache   = sorted(all_vba_funcs - vba_funcs)
    py_funcs = extract_python_functions(Path("basfunct.py"))

    # Prints für den XML-Report
    print(f"VBA-Funktionen:   {len(vba_funcs)}")
    print(f"Python-Funktionen:{len(py_funcs)}")

    missing = sorted(vba_funcs - py_funcs)
    extra   = sorted(py_funcs - vba_funcs)

    print(f"\nVBA  ↔  Python-Funktion (Status)\n" + "-" * 38)
    for name in sorted(vba_funcs):
        status = "OK" if name in py_funcs else "MISSING"
        print(f"{name:<25s} → {status}")

    if extra:
        print("\nZusätzliche Python-Funktionen (nicht in VBA):")
        for name in extra:
            print("  •", name)
    if ignored_cache:
        print("\nIgnorierte VBA-Cache-Funktionen:")
        for name in ignored_cache:
            print("  •", name)

    print(f"\nVBA-Gesamt   : {len(vba_funcs)}")
    print(f"Python-Gesamt: {len(py_funcs)}\n")

    assert not missing, f"Fehlende Python-Funktionen: {missing}"
