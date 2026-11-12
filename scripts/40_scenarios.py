"""Étape 4 — tests de scénarios sur copies jetables de V2 (jamais sur le livrable).
Chaque scénario : copie -> modification (openpyxl ou LibreOffice UNO) -> recalcul LibreOffice -> assertions.
Usage : python3 40_scenarios.py travail/SUIVI_ES_V2_openpyxl.xlsx"""
import sys, os, shutil, subprocess, json, datetime as dt
import openpyxl
sys.path.insert(0, '/home/user/Suivi-pointage-2026/scripts')
V2 = sys.argv[1]
OUT = 'travail/scenarios'; shutil.rmtree(OUT, ignore_errors=True); os.makedirs(OUT)
results = []

def recalc(path):
    r = subprocess.run(['bash', 'scripts/recalc_copy.sh', path, path], capture_output=True, text=True)
    j = json.loads(r.stdout[r.stdout.index('{'):])
    return j

def run(sid, titre, mutate, checks, uno_mutate=None, erreurs_attendues=False):
    path = f'{OUT}/{sid}.xlsx'
    shutil.copy(V2, path)
    if mutate:
        wb = openpyxl.load_workbook(path); mutate(wb); wb.save(path)
    if uno_mutate:
        uno_mutate(path)
    rc = recalc(path)
    v = openpyxl.load_workbook(path, data_only=True)
    obs = []
    if erreurs_attendues:   # erreurs #REF! attendues HORS Contrôles (c'est le risque H1 démontré) ; Contrôles doit rester sans erreur
        errs_c = sum(1 for row in v['Contrôles'].iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('#'))
        ok = errs_c == 0
        obs.append(f"erreurs classeur={rc.get('total_errors')} (attendues hors Contrôles) ; erreurs dans Contrôles={errs_c}")
    else:
        ok = rc.get('total_errors', 1) == 0
        obs.append(f"erreurs={rc.get('total_errors')}")
    for lab, fn, exp in checks:
        got = fn(v)
        good = (got == exp) if not callable(exp) else exp(got)
        ok &= bool(good)
        obs.append(f"{lab}={got!r}{'' if good else ' (ATTENDU ' + repr(exp) + ')'}")
    results.append({'id': sid, 'scenario': titre, 'resultat': 'RÉUSSI' if ok else 'ÉCHEC', 'observations': '; '.join(obs)})
    print(sid, 'RÉUSSI' if ok else 'ÉCHEC', '|', titre, '|', '; '.join(obs))

S = lambda ref: (lambda v: v['SUIVI_ES'][ref].value)
Cc = lambda ref: (lambda v: v['Contrôles'][ref].value)
Dd = lambda ref: (lambda v: v['Tableau de bord'][ref].value)
KPI = lambda v: [v['SUIVI_ES'].cell(3, c).value for c in range(1, 8)]

# T1 compteur au-delà de la cible -> Dépassé (Port-Chars HP 21 : cible 6 481 h)
run('T1', 'Compteur au-delà de la cible → Dépassé', lambda wb: wb['SUIVI_ES'].__setitem__('B8', 6496),
    [('J8', S('J8'), 'Dépassé'), ('I8', S('I8'), 15), ('KPI A3:G3', KPI, [45, 36, 4, 5, 0, 5, 1]),
     ('Tableau de bord rang 5', Dd('C26'), 'Port-Chars HP 21')])

# T2 dans le seuil Proche (HP23 cible 15 513 h) + bornes : −20 -> Proche (Arroseur HP24 cible 6 436), −21 -> Normal (UR131 cible 62 556)
def t2(wb):
    s = wb['SUIVI_ES']; s['B9'] = 15500; s['B10'] = 6416; s['B11'] = 62535
run('T2', 'Écart dans le seuil Proche (−13 h) et bornes −20 h / −21 h', t2,
    [('J9', S('J9'), 'Proche'), ('I9', S('I9'), -13), ('J10 (écart −20)', S('J10'), 'Proche'), ('J11 (écart −21)', S('J11'), 'Normal'),
     ('KPI Proche', lambda v: v['SUIVI_ES']['C3'].value, 6)])

# T3 type ≠ N° cycle
run('T3', 'Type ne correspondant pas au N° de cycle → « Type ≠ cycle »', lambda wb: wb['SUIVI_ES'].__setitem__('E11', 3),
    [('M11', S('M11'), 'Type ≠ cycle'), ('J11', S('J11'), 'À confirmer'), ('Contrôles H9', Cc('H9'), 'Type ≠ cycle'),
     ('Total alertes', Cc('D3'), 2), ('Tuile SUIVI_ES G3', S('G3'), 2)])

# T4 compteur vide sur engin en service
run('T4', 'Compteur actuel vide sur engin en service → À confirmer', lambda wb: wb['SUIVI_ES'].__setitem__('B12', None),
    [('J12', S('J12'), 'À confirmer'), ('I12', S('I12'), 'n.d.'), ('Contrôles F10', Cc('F10'), 'Compteur actuel manquant ou non numérique'),
     ('KPI À confirmer', lambda v: v['SUIVI_ES']['E3'].value, 1), ('Tuile Tableau de bord À confirmer', Dd('J4'), 1)])

# T5 nouvel engin dans une ligne libre
def t5(wb):
    s = wb['SUIVI_ES']
    for col, val in zip('ABCDEKL', ['ENGIN TEST V2', 1000, 900, 'ES 250 h', 1, 'En service', 'Camions']): s[f'{col}58'] = val
    p = wb['Paramètres']; p['E1'] = '=ROWS(Liste_Engins)'; p['E2'] = '=COUNTIF(Liste_Engins,"ENGIN TEST V2")'; p['E3'] = '=INDEX(Liste_Engins,ROWS(Liste_Engins))'
run('T5', 'Nouvel engin en ligne libre (58) → listes et tableau de bord', t5,
    [('J58', S('J58'), 'Normal'), ('F58', S('F58'), 1100), ('KPI en service', lambda v: v['SUIVI_ES']['A3'].value, 46),
     ('ROWS(Liste_Engins)', lambda v: v['Paramètres']['E1'].value, 51), ('dans Liste_Engins', lambda v: v['Paramètres']['E2'].value, 1),
     ('dernier de Liste_Engins', lambda v: v['Paramètres']['E3'].value, 'ENGIN TEST V2'),
     ('Tableau de bord Camions/Normal', Dd('D9'), 14), ('Tuile en service', Dd('B4'), 46), ('Tuile sous-titre', Dd('B5'), 'sur 51 engins suivis'),
     ('HISTORIQUE_ES!A206 (bloc 51)', lambda v: v['HISTORIQUE_ES']['A206'].value, 'ENGIN TEST V2'),
     ('SUIVI_GRAISSAGE!A56', lambda v: v['SUIVI_GRAISSAGE_SOUFFLAGE']['A56'].value, 'ENGIN TEST V2'),
     ('Contrôles alertes ligne 56', Cc('M56'), 0), ('Total alertes', Cc('D3'), 1)])

# T5b liste d'engins avec trou (H6) : A30 vidée (copie) + engin ajouté en A60 après deux lignes vides
def t5b(wb):
    s = wb['SUIVI_ES']; s['A30'] = None
    for col, val in zip('ABCDEK', ['ENGIN APRES TROU', 500, 400, 'ES 250 h', 1, 'En service']): s[f'{col}60'] = val
    p = wb['Paramètres']; p['E1'] = '=ROWS(Liste_Engins)'; p['E2'] = '=COUNTIF(Liste_Engins,"ENGIN APRES TROU")'; p['E3'] = '=SUMPRODUCT(--(Liste_Engins=""))'
run('T5b', 'H6 : trou dans la colonne A + engin après des lignes vides → liste complète sans vide', t5b,
    [('ROWS(Liste_Engins)', lambda v: v['Paramètres']['E1'].value, 50), ('engin après trou présent', lambda v: v['Paramètres']['E2'].value, 1),
     ('vides dans la liste', lambda v: v['Paramètres']['E3'].value, 0),
     ('Contrôles K28 (D11T4 vidé)', Cc('K28'), 'Ordre des lignes modifié (attendu : D11T4)')])

# T6 entrée d'historique discordante (HP23 : nouvelle colonne F)
def t6(wb):
    h = wb['HISTORIQUE_ES']; h['F10'] = 15400; h['F11'] = dt.datetime(2023, 11, 1); h['F12'] = 'ES 250 h'; h['F13'] = 3
run('T6', "Entrée d'historique ne correspondant pas à SUIVI_ES → signalée dans Contrôles", t6,
    [('Contrôles I7', Cc('I7'), 'SUIVI ≠ HISTORIQUE'), ('Contrôles O7 (dernier ES historisé)', Cc('O7'), 15400), ('Total alertes', Cc('D3'), 2),
     ('SUIVI_ES C9 inchangé', S('C9'), 15313)])

# T6b vidange exceptionnelle après le dernier ES : non comparée (hors cycle)
def t6b(wb):
    h = wb['HISTORIQUE_ES']; h['F6'] = 6300; h['F7'] = dt.datetime(2026, 11, 1); h['F8'] = 'Vidange exceptionnelle'
run('T6b', 'Vidange exceptionnelle postérieure au dernier ES → pas de faux écart', t6b,
    [('Contrôles I6', Cc('I6'), None), ('Contrôles N6 (entrées)', Cc('N6'), 4), ('Contrôles O6', Cc('O6'), 6281), ('Total alertes', Cc('D3'), 1)])

# T7 Marge 50 -> 30 : cibles = C + 220, comparaison au recalcul Python indépendant
def t7(wb): wb['Paramètres']['B16'] = 30
def expected_t7(v):
    s = v['SUIVI_ES']; bad = []
    for r in range(8, 58):
        if s[f'K{r}'].value == 'Hors service': continue
        B, C = s[f'B{r}'].value, s[f'C{r}'].value
        cible = C + 250 - 30; ecart = B - cible
        st = 'Normal' if ecart < -20 else ('Proche' if ecart <= 0 else 'Dépassé')
        if s[f'F{r}'].value != cible or s[f'I{r}'].value != ecart or s[f'J{r}'].value != st: bad.append(r)
    return bad
run('T7', 'Marge 50 → 30 : toutes les cibles se mettent à jour (+20 h)', t7,
    [('F8 (6 281 + 220)', S('F8'), 6501), ('lignes non conformes au recalcul Python', expected_t7, []),
     ('Paramètres B18 avance effective', lambda v: v['Paramètres']['B18'].value, 220),
     ('légende SUIVI_ES A4 actualisée', lambda v: '− 30 h de marge' in v['SUIVI_ES']['A4'].value, True),
     ('KPI', KPI, lambda k: k[0] == 45 and sum(k[1:5]) == 45)])

# T8 Seuil_Proche 20 -> 30 (D11T5 à −21 h passe Proche)
run('T8', 'Seuil_Proche 20 → 30 : D11T5 (−21 h) passe Proche', lambda wb: wb['Paramètres'].__setitem__('B19', 30),
    [('J31', S('J31'), 'Proche'), ('KPI Proche', lambda v: v['SUIVI_ES']['C3'].value, lambda x: x > 4)])

# T9 date d'historique saisie en texte
run('T9', "Date d'historique saisie en texte → signalée", lambda wb: wb['HISTORIQUE_ES'].__setitem__('C15', '09/08/2025'),
    [('Contrôles L8', Cc('L8'), "Date d'historique invalide")])

# T10 doublon avec espace final
run('T10', 'Nom en double avec espace final (« KOM2 »)', lambda wb: wb['SUIVI_ES'].__setitem__('A16', 'KOM2 '),
    [('Contrôles J14 (« KOM2 »)', Cc('J14'), 'Nom en double'), ('Contrôles J19 (KOM2)', Cc('J19'), 'Nom en double')])

# T11 compteur < dernier ES
run('T11', 'Compteur actuel < dernier ES', lambda wb: wb['SUIVI_ES'].__setitem__('B11', 62000),
    [('J11', S('J11'), 'À confirmer'), ('Contrôles G9', Cc('G9'), 'Compteur < dernier ES')])

# T12-T14 atténuation H1 (LibreOffice UNO : tri / insertion / suppression)
def uno_op(op):
    def f(path):
        from lo_uno import LO, unprotect_all
        import uno
        from com.sun.star.table import TableSortField
        with LO(port=2103) as lo:
            doc = lo.open(path); unprotect_all(doc); s = doc.Sheets.getByName('SUIVI_ES')
            if op == 'tri':
                rng = s.getCellRangeByName('A8:Q57'); sd = rng.createSortDescriptor()
                fld = TableSortField(); fld.Field = 1; fld.IsAscending = False
                for p in sd:
                    if p.Name == 'SortFields': p.Value = uno.Any('[]com.sun.star.table.TableSortField', (fld,))
                rng.sort(sd)
            elif op == 'insertion':
                s.Rows.insertByIndex(19, 1); s.getCellByPosition(0, 19).setString('INSERE')
            elif op == 'suppression':
                s.Rows.removeByIndex(8, 1)
            lo.save_xlsx(doc, path); doc.close(True)
    return f
cnt_ordre = lambda v: sum(1 for r in range(6, 76) if str(v['Contrôles'][f'K{r}'].value or '').startswith('Ordre'))
run('T12', 'H1 : tri de SUIVI_ES → alertes « Ordre des lignes modifié »', None,
    [('positions signalées', cnt_ordre, lambda n: n >= 40), ('Total alertes', Cc('D3'), lambda n: n >= 40)], uno_mutate=uno_op('tri'))
run('T13', 'H1 : insertion de ligne dans SUIVI_ES → alertes', None,
    [('Contrôles K18 (position 13)', Cc('K18'), 'Ordre des lignes modifié (attendu : KOM1)'), ('positions signalées', cnt_ordre, lambda n: n >= 38)],
    uno_mutate=uno_op('insertion'))
run('T14', 'H1 : suppression de ligne dans SUIVI_ES → alertes, Contrôles sans #REF!', None,
    [('Contrôles K7 (position 2)', Cc('K7'), 'Ordre des lignes modifié (attendu : HP23)'), ('positions signalées', cnt_ordre, lambda n: n >= 48),
     ('Total alertes numérique', Cc('D3'), lambda n: isinstance(n, (int, float)))], uno_mutate=uno_op('suppression'), erreurs_attendues=True)

json.dump(results, open('travail/resultats_scenarios.json', 'w'), ensure_ascii=False, indent=1)
print('\nBILAN :', sum(r['resultat'] == 'RÉUSSI' for r in results), '/', len(results), 'réussis')
