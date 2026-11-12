"""H6/H7 — cas limites sur une copie jetable (openpyxl + recalcul LibreOffice). Ne touche jamais l'original."""
import openpyxl, subprocess, json, sys
SRC = sys.argv[1]; DST = 'travail/sim_H6_H7.xlsx'
wb = openpyxl.load_workbook(SRC); S = wb['SUIVI_ES']; H = wb['HISTORIQUE_ES']
cas = {
 11: ('B', 62000, 'compteur < dernier ES (62000 < 62356)'),
 12: ('E', 9, 'N° cycle hors 1–8 (9)'),
 13: ('E', 3, 'type ≠ cycle (ES 500 h avec N° 3)'),
 14: ('B', None, 'compteur vide sur engin en service'),
 17: ('B', '2400', 'compteur saisi en texte'),
 18: ('E', 2.5, 'N° cycle décimal (2,5)'),
 19: ('K', None, 'état vide'),
 20: ('K', 'hors service', 'état en minuscules'),
}
for r, (c, v, _) in cas.items(): S[f'{c}{r}'] = v
S['A15'] = 'KOM1'          # doublon exact
S['A16'] = 'KOM2 '         # doublon avec espace final
H['D15'] = '12/10/2026'    # date d'historique en texte (Arroseur HP24)
# H6 : trou dans la colonne A + nouvel engin après des lignes vides
S['A30'] = None            # (D11T4) -> trou au milieu
S['A60'] = 'NOUVEL ENGIN'; S['B60'] = 100; S['C60'] = 0; S['D60'] = 'ES 2000 h'; S['E60'] = 8; S['K60'] = 'En service'
S['AH1'] = '=ROWS(Liste_Engins)'; S['AH2'] = '=COUNTIF(Liste_Engins,"NOUVEL ENGIN")'; S['AH3'] = '=COUNTIF(Liste_Engins,"HELI 15T")'
S['AH4'] = '=MATCH("KOM1",Engins_Plage,0)'; S['AH5'] = '=COUNTIF(Engins_Plage,"KOM2")'
S['AH6'] = '=SUMPRODUCT(--(Liste_Engins=""))'
wb.save(DST)
print(subprocess.run(['bash', 'scripts/recalc_copy.sh', DST, DST], capture_output=True, text=True).stdout)
v = openpyxl.load_workbook(DST, data_only=True)['SUIVI_ES']
res = {}
for r, (c, val, lab) in cas.items():
    res[lab] = {f'{x}{r}': v[f'{x}{r}'].value for x in 'ABCDEFIJKM'}
res['doublon A15=KOM1 / A16="KOM2 "'] = {'J15': v['J15'].value, 'MATCH KOM1 dans Engins_Plage (1re occurrence)': v['AH4'].value, 'COUNTIF KOM2 (sans espace)': v['AH5'].value}
res['H6 Liste_Engins'] = {'ROWS(Liste_Engins) (51 engins non vides)': v['AH1'].value, 'NOUVEL ENGIN (A60) présent': v['AH2'].value, 'HELI 15T (A57) présent': v['AH3'].value, 'cellules vides incluses': v['AH6'].value}
res['KPI A3:F3'] = [v.cell(3, c).value for c in range(1, 7)]
print(json.dumps(res, ensure_ascii=False, indent=1, default=str))
json.dump(res, open('travail/preuve_H6_H7.json', 'w'), ensure_ascii=False, indent=1, default=str)
