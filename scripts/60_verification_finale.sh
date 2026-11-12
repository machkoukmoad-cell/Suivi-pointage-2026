#!/bin/bash
# Vérification finale sur le livrable exact SUIVI_ES_V2.xlsx (toutes les sorties dans travail/verif_finale/)
set -e
cd /home/user/Suivi-pointage-2026
OUT=travail/verif_finale; rm -rf $OUT; mkdir -p $OUT
echo "== 1. Original intact"; sha256sum SUIVI_ES_V1.1.xlsx | tee $OUT/sha_original.txt
grep -q 9f58705cbd04e57813322e2cef41b1f8ee38cb97ce3254561333f7f979a9bbb1 $OUT/sha_original.txt && echo "original inchangé : OK"
echo "== 2. Recalcul LibreOffice d'une copie du livrable"; bash scripts/recalc_copy.sh SUIVI_ES_V2.xlsx $OUT/V2_recalc.xlsx | tee $OUT/recalc.json
echo "== 3. Résultats par engin et KPI"; python3 scripts/02_extraire_resultats.py $OUT/V2_recalc.xlsx $OUT/resultats_V2.csv | tee $OUT/kpi.txt
echo "== 4. Réconciliation avec la baseline V1.1"; python3 scripts/20_reconciliation.py travail/baseline_V1.1.csv $OUT/resultats_V2.csv | tail -1 | tee $OUT/reconciliation.txt
echo "== 5. Intégrité V1.1 -> V2"; python3 scripts/30_integrite.py SUIVI_ES_V1.1.xlsx SUIVI_ES_V2.xlsx --json $OUT/integrite.json > /dev/null
python3 -c "
import json; d=json.load(open('$OUT/integrite.json'))
print('données utilisateur modifiées hors corrections validées :', sum(len(v['differences']) for v in d['donnees_utilisateur'].values() if isinstance(v, dict)))
print('corrections validées par l utilisateur appliquées :', {k: v for x in d['donnees_utilisateur'].values() if isinstance(x, dict) for k, v in x.get('corrections_validees_appliquees', {}).items()})
v=d['verrouillage']; print('formules non verrouillées :', len(v['formules_non_verrouillees']), '| saisies verrouillées :', len(v['saisies_verrouillees']), '| feuilles protégées :', sum(v['feuilles_protegees'].values()), '/', len(v['feuilles_protegees']))
print('logo :', d['structure']['logo_sha256']['V1.1'] == d['structure']['logo_sha256']['V2'], '| commentaires SUIVI_ES :', d['structure']['SUIVI_ES']['commentaires']['V2'], '+ Paramètres', list(d['structure']['Paramètres_commentaires']))
" | tee $OUT/integrite_resume.txt
echo "== 6. Scénarios (copies jetables)"; python3 scripts/40_scenarios.py SUIVI_ES_V2.xlsx | tail -1 | tee $OUT/scenarios.txt
cp travail/resultats_scenarios.json $OUT/
echo "== 7. Fonctions interdites (tableaux dynamiques, LET, XLOOKUP...)"
python3 -c "
import openpyxl,re
wb=openpyxl.load_workbook('SUIVI_ES_V2.xlsx'); bad=[]
pat=re.compile(r'\b(XLOOKUP|XMATCH|LET|LAMBDA|FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|RANDARRAY|TEXTJOIN|IFS|SWITCH|MAXIFS|MINIFS|CONCAT)\s*\(', re.I)
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith('=') and pat.search(c.value): bad.append(ws.title+'!'+c.coordinate)
for n,d in wb.defined_names.items():
    if pat.search(d.attr_text or ''): bad.append('nom '+n)
print('formules avec fonctions non Excel 2016 :', len(bad), bad[:5])
print('macros (vba) :', 'xl/vbaProject.bin' in __import__('zipfile').ZipFile('SUIVI_ES_V2.xlsx').namelist())
" | tee $OUT/fonctions.txt
