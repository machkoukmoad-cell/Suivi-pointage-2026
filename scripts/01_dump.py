"""Dump des cellules non vides d'une feuille (valeur/formule, type, format, verrouillage, remplissage)."""
import sys, openpyxl
p, sh = sys.argv[1], sys.argv[2]
r0 = int(sys.argv[3]) if len(sys.argv) > 3 else 1
r1 = int(sys.argv[4]) if len(sys.argv) > 4 else 10**6
cols = sys.argv[5] if len(sys.argv) > 5 else None
wb = openpyxl.load_workbook(p)
ws = wb[sh]
for row in ws.iter_rows(min_row=r0, max_row=min(r1, ws.max_row)):
    for c in row:
        if c.value is None: continue
        if cols and c.column_letter not in cols.split(','): continue
        f = c.fill.fgColor.rgb if c.fill and c.fill.fill_type else None
        print(f"{c.coordinate}\t{type(c.value).__name__}\t{c.value!r}\tfmt={c.number_format}\tlock={c.protection.locked}\tfill={f}")
