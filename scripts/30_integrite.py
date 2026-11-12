"""Diff d'intégrité V1.1 -> V2 :
 1. données utilisateur (aucune modification tolérée) : noms/ordre des engins, compteurs, derniers ES, états, catégories,
    historiques ES et graissage, saisies graissage, noms de feuilles ;
 2. structure : validations, MFC, fusions, volets figés, protection, logo, commentaires, noms définis, verrouillage formules/saisies.
Usage : python3 30_integrite.py V1.1.xlsx V2.xlsx [--json sortie.json]"""
import sys, json, zipfile, hashlib, openpyxl, datetime as dt

A, B = sys.argv[1], sys.argv[2]
wa, wb = openpyxl.load_workbook(A), openpyxl.load_workbook(B)
rep = {'donnees_utilisateur': {}, 'structure': {}}

# ---------------------------------------------------------------- 1. données utilisateur
def cells(ws, rows, cols):
    return {f'{c}{r}': ws[f'{c}{r}'].value for r in rows for c in cols}
def colrange(a, b):
    from openpyxl.utils import get_column_letter as L, column_index_from_string as I
    return [L(i) for i in range(I(a), I(b) + 1)]
checks = {
    'SUIVI_ES A:E,K,L (8-77)': ('SUIVI_ES', range(8, 78), list('ABCDEKL')),
    'HISTORIQUE_ES C:AP (6-285)': ('HISTORIQUE_ES', range(6, 286), colrange('C', 'AP')),
    'SUIVI_GRAISSAGE C:E,G (6-75)': ('SUIVI_GRAISSAGE_SOUFFLAGE', range(6, 76), list('CDEG')),
    'HIST_GRAISSAGE C:BJ (6-215)': ('HISTORIQUE_GRAISSAGE_SOUFFLAGE', range(6, 216), colrange('C', 'BJ')),
    'Libellés HISTORIQUE_ES B / formules A': ('HISTORIQUE_ES', range(6, 286), ['A', 'B']),
    'Formules SUIVI_GRAISSAGE A,B,F': ('SUIVI_GRAISSAGE_SOUFFLAGE', range(6, 76), list('ABF')),
}
for lab, (sh, rows, cols) in checks.items():
    x, y = cells(wa[sh], rows, cols), cells(wb[sh], rows, cols)
    d = {k: (x[k], y[k]) for k in x if x[k] != y[k]}
    rep['donnees_utilisateur'][lab] = {'cellules_comparees': len(x), 'non_vides': sum(v is not None for v in x.values()), 'differences': d}
rep['donnees_utilisateur']['feuilles_d_origine_presentes'] = all(n in wb.sheetnames for n in wa.sheetnames)
nonempty_hist = sum(1 for r in range(6, 286) for c in range(3, 43) if wa['HISTORIQUE_ES'].cell(r, c).value is not None)
rep['donnees_utilisateur']['HISTORIQUE_ES_cellules_saisies'] = nonempty_hist

# ---------------------------------------------------------------- 2. structure
def dvs(ws):
    return sorted((str(d.sqref), d.type, d.formula1, d.formula2, d.errorStyle, bool(d.showInputMessage), d.promptTitle) for d in ws.data_validations.dataValidation)
def cfs(ws):
    out = []
    for cf, rules in ws.conditional_formatting._cf_rules.items():
        for r in rules: out.append((str(cf.sqref), tuple(r.formula), r.stopIfTrue))
    return sorted(out, key=str)
def prot(ws):
    p = ws.protection
    return {k: getattr(p, k) for k in ('sheet', 'password', 'sort', 'autoFilter', 'insertRows', 'deleteRows', 'formatColumns', 'formatRows', 'formatCells')}
for ws in wa.worksheets:
    v = wb[ws.title]
    st = {}
    a_dv, b_dv = dvs(ws), dvs(v)
    st['validations'] = {'V1.1': len(a_dv), 'V2': len(b_dv), 'retirees': [x for x in a_dv if x not in b_dv], 'ajoutees': [x for x in b_dv if x not in a_dv]}
    a_cf, b_cf = cfs(ws), cfs(v)
    st['mfc'] = {'V1.1': len(a_cf), 'V2': len(b_cf), 'retirees': [x[0] + ' ' + x[1][0][:90] for x in a_cf if x not in b_cf],
                 'ajoutees': [x[0] + ' ' + x[1][0][:90] for x in b_cf if x not in a_cf]}
    am, bm = set(map(str, ws.merged_cells.ranges)), set(map(str, v.merged_cells.ranges))
    st['fusions'] = {'V1.1': len(am), 'V2': len(bm), 'retirees': sorted(am - bm), 'ajoutees': sorted(bm - am)}
    st['volets'] = (ws.freeze_panes, v.freeze_panes)
    st['protection'] = {'V1.1': prot(ws), 'V2': prot(v)}
    st['images'] = (len(ws._images), len(v._images))
    ac = {c.coordinate: c.comment.text for row in ws.iter_rows() for c in row if c.comment}
    bc = {c.coordinate: c.comment.text for row in v.iter_rows() for c in row if c.comment}
    st['commentaires'] = {'V1.1': sorted(ac), 'V2': sorted(bc), 'texte_modifie': [k for k in ac if k in bc and ac[k] != bc[k]]}
    st['filtre'] = (ws.auto_filter.ref, v.auto_filter.ref)
    st['titres_impression'] = ((ws.print_title_rows, ws.print_title_cols), (v.print_title_rows, v.print_title_cols))
    rep['structure'][ws.title] = st
# commentaires déplacés vers Paramètres
rep['structure']['Paramètres_commentaires'] = {c.coordinate: c.comment.text[:60] for row in wb['Paramètres'].iter_rows() for c in row if c.comment} if 'Paramètres' in wb.sheetnames else None
# logo : image identique octet pour octet
def media(p):
    with zipfile.ZipFile(p) as z:
        return {n: hashlib.sha256(z.read(n)).hexdigest()[:16] for n in z.namelist() if n.startswith('xl/media/')}
rep['structure']['logo_sha256'] = {'V1.1': media(A), 'V2': media(B)}
rep['structure']['logo_ancrage'] = [(type(i.anchor).__name__, i.anchor._from.col, i.anchor._from.row, i.width, i.height) for i in wb['SUIVI_ES']._images]
rep['structure']['noms_definis'] = {'V1.1': {n: d.attr_text for n, d in wa.defined_names.items()}, 'V2': {n: d.attr_text for n, d in wb.defined_names.items()}}

# ---------------------------------------------------------------- verrouillage : formules verrouillées, saisies déverrouillées
bad_formula_unlocked, formula_cells = [], 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                formula_cells += 1
                if c.protection.locked is False: bad_formula_unlocked.append(f'{ws.title}!{c.coordinate}')
inputs = [('SUIVI_ES', r, c) for r in range(8, 78) for c in 'BCDEKL'] + [('SUIVI_ES', r, 'A') for r in range(58, 78)] + \
         [('SUIVI_GRAISSAGE_SOUFFLAGE', r, c) for r in range(6, 76) for c in 'CDEG'] + \
         [('Paramètres', r, 'B') for r in list(range(5, 13)) + [16, 17, 19] + list(range(23, 28)) + list(range(30, 38)) + [41]]
bad_input_locked = [f'{s}!{c}{r}' for s, r, c in inputs if wb[s][f'{c}{r}'].protection.locked]
hist_locked = sum(1 for r in range(6, 286) for c in range(3, 43) if wb['HISTORIQUE_ES'].cell(r, c).protection.locked)
rep['verrouillage'] = {'cellules_formule': formula_cells, 'formules_non_verrouillees': bad_formula_unlocked,
                       'saisies_verrouillees': bad_input_locked, 'HISTORIQUE_ES_saisies_verrouillees': hist_locked,
                       'feuilles_protegees': {ws.title: ws.protection.sheet for ws in wb.worksheets},
                       'mot_de_passe': {ws.title: ws.protection.password for ws in wb.worksheets}}
out = json.dumps(rep, ensure_ascii=False, indent=1, default=str)
if '--json' in sys.argv: open(sys.argv[sys.argv.index('--json') + 1], 'w').write(out)
print(out)
