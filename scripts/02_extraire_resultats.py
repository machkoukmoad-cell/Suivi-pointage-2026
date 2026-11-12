"""Extrait les résultats recalculés de SUIVI_ES (valeurs) par engin -> CSV + résumé des KPI.
Usage: python3 02_extraire_resultats.py fichier_recalcule.xlsx sortie.csv [premiere_ligne_donnees=8] [nom_feuille=SUIVI_ES]"""
import sys, openpyxl, csv
p, out = sys.argv[1], sys.argv[2]
r0 = int(sys.argv[3]) if len(sys.argv) > 3 else 8
sh = sys.argv[4] if len(sys.argv) > 4 else 'SUIVI_ES'
wb = openpyxl.load_workbook(p, data_only=True)
ws = wb[sh]
hdr = [ws.cell(r0 - 1, c).value for c in range(1, 18)]
rows = []
for r in range(r0, r0 + 70):
    if ws.cell(r, 1).value in (None, ''): continue
    rows.append([r] + [ws.cell(r, c).value for c in range(1, 18)])
with open(out, 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f, delimiter=';'); w.writerow(['ligne'] + hdr); w.writerows(rows)
print('KPI A3:F3 =', [ws.cell(r0 - 5, c).value for c in range(1, 7)], '(ligne', r0 - 5, ')')
from collections import Counter
print('engins:', len(rows), 'statuts:', Counter(r[10] for r in rows))
err = 0
for s in wb.worksheets:
    for row in s.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('#') and c.value.rstrip('!?0/').upper() in ('#REF', '#VALUE', '#NAME', '#DIV', '#N/A', '#NUM', '#NULL') or (isinstance(c.value, str) and c.value in ('#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#N/A', '#NUM!', '#NULL!', 'Err:502', 'Err:504', 'Err:508')):
                err += 1; print('ERREUR', s.title, c.coordinate, c.value)
print('cellules en erreur:', err)
