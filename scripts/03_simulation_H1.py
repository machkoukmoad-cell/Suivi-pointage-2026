"""H1 — Simulation du couplage par position de ligne (copie jetable, LibreOffice UNO).
Scénarios : (a) insertion d'une ligne dans SUIVI_ES ; (b) suppression d'une ligne ; (c) tri de SUIVI_ES par compteur.
Pour chaque scénario : on regarde quel engin affiche chaque bloc HISTORIQUE_ES / ligne SUIVI_GRAISSAGE et
si les données d'historique saisies restent attachées au bon engin."""
import sys, json
sys.path.insert(0, '/home/user/Suivi-pointage-2026/scripts')
from lo_uno import LO, unprotect_all
from com.sun.star.table import CellRangeAddress, CellAddress
import uno
from com.sun.star.beans import PropertyValue

SRC = sys.argv[1] if len(sys.argv) > 1 else 'SUIVI_ES_V1.1.xlsx'
H_STEP = 4
res = {}

def snapshot(doc, label):
    s = doc.Sheets.getByName('SUIVI_ES'); h = doc.Sheets.getByName('HISTORIQUE_ES'); g = doc.Sheets.getByName('SUIVI_GRAISSAGE_SOUFFLAGE')
    hd = doc.Sheets.getByName('HISTORIQUE_ES')
    out = {}
    # Engins avec historique : Port-Chars HP 21, HP23, Arroseur HP24 + Pelleteuse W1 (G48 graissage)
    for k in range(0, 50):
        hr = 5 + H_STEP * k  # 0-based row of block k
        name_h = h.getCellByPosition(0, hr).getString()
        first_es = h.getCellByPosition(2, hr).getString()
        out[k] = {'bloc_HIST_ligne': hr + 1, 'nom_affiche_HIST': name_h, 'HIST_ES1_compteur': first_es,
                  'HIST_formule': h.getCellByPosition(0, hr).getFormula(),
                  'GRAISS_nom': g.getCellByPosition(0, 5 + k).getString(), 'GRAISS_G': g.getCellByPosition(6, 5 + k).getString(),
                  'GRAISS_B': g.getCellByPosition(1, 5 + k).getString()}
    names = [s.getCellByPosition(0, r).getString() for r in range(7, 78)]
    return out, names

with LO() as lo:
    # (a) insertion d'une ligne avant KOM1 (ligne 20)
    doc = lo.open(SRC); unprotect_all(doc)
    s = doc.Sheets.getByName('SUIVI_ES')
    s.Rows.insertByIndex(19, 1)   # insère avant la ligne 20 (1-based)
    s.getCellByPosition(0, 19).setString('NOUVEL ENGIN TEST')
    doc.calculateAll()
    snap, names = snapshot(doc, 'insert')
    h = doc.Sheets.getByName('HISTORIQUE_ES')
    # navigation "Aller ›" pour KOM1 : MATCH dans Engins_Plage
    h.getCellByPosition(8, 0).setString('KOM1'); doc.calculateAll()
    link_kom1 = h.getCellByPosition(11, 0).getFormula()
    pos = names.index('KOM1') + 1
    res['insertion'] = {
        'SUIVI_ES_A20': names[12], 'SUIVI_ES_A21': names[13],
        'Engins_Plage_apres': doc.NamedRanges.getByName('Engins_Plage').getContent(),
        'bloc_HIST_ligne54_nom': snap[12]['nom_affiche_HIST'], 'bloc_HIST_ligne54_formule': snap[12]['HIST_formule'],
        'nouvel_engin_a_un_bloc_HIST': any(v['nom_affiche_HIST'] == 'NOUVEL ENGIN TEST' for v in snap.values()),
        'nouvel_engin_a_une_ligne_GRAISS': any(v['GRAISS_nom'] == 'NOUVEL ENGIN TEST' for v in snap.values()),
        'Aller_KOM1_cible': f'ligne {6 + 4 * (pos - 1)} (bloc réel de KOM1 = ligne 54, nom affiché ligne {6+4*(pos-1)} = ' + h.getCellByPosition(0, 5 + 4 * (pos - 1)).getString() + ')',
        'derniere_formule_HIST_A282': h.getCellByPosition(0, 281).getFormula(),
    }
    doc.close(True)

    # (b) suppression de la ligne 9 (HP23, qui a 3 ES d'historique)
    doc = lo.open(SRC); unprotect_all(doc)
    s = doc.Sheets.getByName('SUIVI_ES'); s.Rows.removeByIndex(8, 1); doc.calculateAll()
    snap, names = snapshot(doc, 'delete')
    g = doc.Sheets.getByName('SUIVI_GRAISSAGE_SOUFFLAGE')
    res['suppression'] = {
        'bloc_HIST_ligne10_nom': snap[1]['nom_affiche_HIST'], 'bloc_HIST_ligne10_formule': snap[1]['HIST_formule'],
        'bloc_HIST_ligne10_ES1_compteur_orphelin': snap[1]['HIST_ES1_compteur'],
        'GRAISS_A7': g.getCellByPosition(0, 6).getString(), 'GRAISS_A7_formule': g.getCellByPosition(0, 6).getFormula(),
        'GRAISS_B7': g.getCellByPosition(1, 6).getString(), 'GRAISS_F7': g.getCellByPosition(5, 6).getString(),
        'Liste_Engins': doc.NamedRanges.getByName('Liste_Engins').getContent(),
    }
    doc.close(True)

    # (c) tri de SUIVI_ES A8:Q57 par compteur actuel décroissant
    doc = lo.open(SRC); unprotect_all(doc)
    s = doc.Sheets.getByName('SUIVI_ES')
    rng = s.getCellRangeByName('A8:Q57')
    sf = rng.createSortDescriptor()
    from com.sun.star.table import TableSortField
    f = TableSortField(); f.Field = 1; f.IsAscending = False
    for p in sf:
        if p.Name == 'SortFields': p.Value = uno.Any('[]com.sun.star.table.TableSortField', (f,))
        if p.Name == 'ContainsHeader': p.Value = False
    rng.sort(sf); doc.calculateAll()
    snap, names = snapshot(doc, 'sort')
    tri = []
    for k in range(50):
        if snap[k]['HIST_ES1_compteur'] or snap[k]['GRAISS_G']:
            tri.append({'bloc_ligne': snap[k]['bloc_HIST_ligne'], 'nom_affiche_HIST_apres_tri': snap[k]['nom_affiche_HIST'],
                        'HIST_ES1_compteur': snap[k]['HIST_ES1_compteur'], 'GRAISS_ligne': 6 + k,
                        'GRAISS_nom_apres_tri': snap[k]['GRAISS_nom'], 'GRAISS_G_saisi': snap[k]['GRAISS_G'], 'GRAISS_B': snap[k]['GRAISS_B']})
    res['tri'] = {'premiers_noms_SUIVI_apres_tri': names[:3], 'donnees_saisies_et_engin_affiche': tri}
    doc.close(True)

print(json.dumps(res, ensure_ascii=False, indent=1))
json.dump(res, open('/home/user/Suivi-pointage-2026/travail/preuve_H1.json', 'w'), ensure_ascii=False, indent=1)
