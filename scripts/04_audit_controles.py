"""Contrôles généraux d'audit sur V1.1 (lecture seule) :
 - cohérence des formules le long des colonnes
 - nombres stockés en texte, formats de date, espaces superflus, doublons
 - recalcul indépendant (pandas) des résultats par engin et des KPI, comparaison avec le recalcul LibreOffice
 - comparaison SUIVI_ES C:E vs dernière entrée HISTORIQUE_ES (H2)"""
import re, sys, json, openpyxl, pandas as pd, datetime as dt
SRC = sys.argv[1]; RECALC = sys.argv[2]
wb = openpyxl.load_workbook(SRC); wv = openpyxl.load_workbook(RECALC, data_only=True)
S = wb['SUIVI_ES']; H = wb['HISTORIQUE_ES']; G = wb['SUIVI_GRAISSAGE_SOUFFLAGE']; HG = wb['HISTORIQUE_GRAISSAGE_SOUFFLAGE']
out = {}

def norm(f, r):
    return re.sub(r'(?<![A-Za-z_$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?!\d)', lambda m: m.group(1)+m.group(2)+m.group(3)+('{r}' if int(m.group(4))==r else ('{r-1}' if int(m.group(4))==r-1 else m.group(4))), f)

# 1. cohérence des formules
incons = {}
for col in 'FGHIJMNOPQ':
    pats = {}
    for r in range(8, 78):
        v = S[f'{col}{r}'].value
        pats.setdefault(norm(str(v), r), []).append(r)
    incons[f'SUIVI_ES!{col}8:{col}77'] = {k[:90]: (len(v), v[:3]) for k, v in pats.items()}
out['coherence_formules'] = {k: v for k, v in incons.items() if len(v) > 1}
out['coherence_formules_ok'] = [k for k, v in incons.items() if len(v) == 1]

# 2. saisies : nombres en texte, espaces, doublons
txtnum = []
for r in range(8, 78):
    for col in 'BCE':
        v = S[f'{col}{r}'].value
        if isinstance(v, str): txtnum.append(f'SUIVI_ES!{col}{r}={v!r}')
for k in range(70):
    for i in range(4):
        r = 6 + 4 * k + i
        for c in range(3, 43):
            v = H.cell(r, c).value
            if v is None: continue
            if i in (0, 3) and not isinstance(v, (int, float)): txtnum.append(f'HISTORIQUE_ES!{H.cell(r,c).coordinate}={v!r}')
            if i == 1 and not isinstance(v, (dt.datetime, dt.date)): txtnum.append(f'HISTORIQUE_ES!{H.cell(r,c).coordinate} date non-date={v!r}')
for r in range(6, 76):
    for col in 'CDG':
        v = G[f'{col}{r}'].value
        if v is None: continue
        if col == 'C' and not isinstance(v, (dt.datetime, dt.date)): txtnum.append(f'GRAISS!{col}{r}={v!r}')
        if col in 'DG' and not isinstance(v, (int, float)): txtnum.append(f'GRAISS!{col}{r}={v!r}')
out['nombres_texte_ou_dates_invalides'] = txtnum
fmts = {}
for k in range(70):
    for c in range(3, 43):
        fmts.setdefault(H.cell(7 + 4 * k, c).number_format, 0); fmts[H.cell(7 + 4 * k, c).number_format] += 1
out['formats_date_HISTORIQUE_ES'] = fmts
out['formats_date_GRAISS_C'] = list({G[f'C{r}'].number_format for r in range(6, 76)})
out['formats_date_HG'] = list({HG.cell(7 + 3 * k, c).number_format for k in range(70) for c in range(3, 63)})
names = [(r, S[f'A{r}'].value) for r in range(8, 78) if S[f'A{r}'].value]
out['noms_espaces'] = [(r, n) for r, n in names if n != n.strip() or '  ' in n or '\n' in n]
tn = [' '.join(n.split()).upper() for r, n in names]
out['noms_doublons_normalises'] = [n for n in set(tn) if tn.count(n) > 1]
out['nb_engins'] = len(names)

# 3. recalcul indépendant
MARGE, INTERV = S['AF17'].value, S['AF18'].value
CYCLE = [S[f'AF{r}'].value for r in range(8, 16)]
rows = []
for r, n in names:
    B, C, D, E, K = (S[f'{c}{r}'].value for c in 'BCDEK')
    if K == 'Hors service':
        statut, ecart, cible = 'Hors service', None, None
    else:
        ok = all(x is not None for x in (C, D, E)) and isinstance(C, (int, float)) and isinstance(E, int) and 1 <= E <= 8 and D == CYCLE[E - 1]
        if not ok: statut, ecart, cible = 'À confirmer', None, None
        else:
            cible = C + INTERV - MARGE
            if not isinstance(B, (int, float)) or B < C: statut, ecart = 'À confirmer', None
            else:
                ecart = B - cible
                statut = 'Normal' if ecart < -20 else ('Proche' if ecart <= 0 else 'Dépassé')
    lo_s = wv['SUIVI_ES'][f'J{r}'].value; lo_i = wv['SUIVI_ES'][f'I{r}'].value; lo_f = wv['SUIVI_ES'][f'F{r}'].value
    rows.append(dict(ligne=r, engin=n, statut_py=statut, ecart_py=ecart, cible_py=cible, statut_LO=lo_s, ecart_LO=lo_i, cible_LO=lo_f))
df = pd.DataFrame(rows)
diff = df[(df.statut_py != df.statut_LO) | ((df.ecart_py.notna()) & (df.ecart_py != pd.to_numeric(df.ecart_LO, errors='coerce'))) | ((df.cible_py.notna()) & (df.cible_py != pd.to_numeric(df.cible_LO, errors='coerce')))]
out['recalcul_independant_ecarts'] = diff.to_dict('records')
kpi_py = {'en_service': int((df.statut_py != 'Hors service').sum()), **df.statut_py.value_counts().to_dict()}
kpi_lo = dict(zip(['en_service', 'Normal', 'Proche', 'Dépassé', 'À confirmer', 'Hors service'], [wv['SUIVI_ES'].cell(3, c).value for c in range(1, 7)]))
out['kpi_python'] = kpi_py; out['kpi_LO_A3F3'] = kpi_lo
# masqué vs visible
hidden = [r for r in range(8, 78) if S.row_dimensions[r].hidden]
out['lignes_masquees'] = [(r, S[f'A{r}'].value, S[f'K{r}'].value) for r in hidden]
vis = df[~df.ligne.isin(hidden)]
out['statuts_lignes_visibles'] = vis.statut_py.value_counts().to_dict()

# 4. H2 : SUIVI vs dernière entrée HISTORIQUE
h2 = []
for k, (r, n) in enumerate([(r, S[f'A{r}'].value) for r in range(8, 78)]):
    if not n: continue
    base = 6 + 4 * (r - 8)
    entries = [(H.cell(4 + 1, c).value, H.cell(base, c).value, H.cell(base + 1, c).value, H.cell(base + 2, c).value, H.cell(base + 3, c).value) for c in range(3, 43) if any(H.cell(base + i, c).value is not None for i in range(4))]
    suivi = (S[f'C{r}'].value, S[f'D{r}'].value, S[f'E{r}'].value)
    if not entries:
        h2.append((r, n, 'historique vide', suivi)); continue
    last = entries[-1]
    h2.append((r, n, 'OK' if (last[1], last[3], last[4]) == suivi else 'ÉCART', suivi, last))
out['H2_resume'] = {'historique_vide': sum(1 for x in h2 if x[2] == 'historique vide'), 'OK': [x for x in h2 if x[2] == 'OK'], 'ECART': [x for x in h2 if x[2] == 'ÉCART']}
# intervalles d'historique
for x in h2:
    pass
print(json.dumps(out, ensure_ascii=False, indent=1, default=str))
json.dump(out, open('travail/preuve_audit_general.json', 'w'), ensure_ascii=False, indent=1, default=str)
