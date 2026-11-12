"""Construction de SUIVI_ES_V2 à partir de la copie intacte de V1.1 (openpyxl).
Règles : aucune donnée utilisateur modifiée (compteurs, dates, derniers ES, historique, noms et ordre des engins, noms de feuilles).
Usage : python3 scripts/10_construire_V2.py travail/V1.1_intact.xlsx travail/SUIVI_ES_V2_openpyxl.xlsx"""
import sys, copy
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, '/home/user/Suivi-pointage-2026/scripts')
from registre import MODIFICATIONS

SRC, DST = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
S = wb['SUIVI_ES']; H = wb['HISTORIQUE_ES']; G = wb['SUIVI_GRAISSAGE_SOUFFLAGE']; HG = wb['HISTORIQUE_GRAISSAGE_SOUFFLAGE']

# ---------------------------------------------------------------- styles
ARIAL = 'Arial'
C_SAISIE, C_CALCUL, C_PARAM = 'FFF8DC', 'F2F2F2', 'DDEBF7'
C_VERT_F, C_VERT, C_VERT_C, C_TXT = '1F6B1C', '33992E', '2F7F2B', '1F3A1F'
C_BORD = 'BFC5BF'
fill = lambda c: PatternFill('solid', fgColor=c, bgColor=c)
thin = Side(style='thin', color=C_BORD)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_SAISIE, F_CALCUL, F_PARAM = fill(C_SAISIE), fill(C_CALCUL), fill(C_PARAM)
LOCK, UNLOCK = Protection(locked=True), Protection(locked=False)
FONT_IN = Font(name=ARIAL, size=10, bold=True, color='1F3A8A')       # convention de saisie existante (bleu foncé gras)
FONT_N = Font(name=ARIAL, size=10)
FONT_HDR = Font(name=ARIAL, size=9, bold=True, color='FFFFFF')
FONT_HDR10 = Font(name=ARIAL, size=10, bold=True, color='FFFFFF')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
LEFT_TOP = Alignment(horizontal='left', vertical='top', wrap_text=True)

def put(ws, ref, value=None, font=None, fill_=None, align=None, nf=None, lock=True, border=True):
    c = ws[ref]
    if value is not None: c.value = value
    c.font = font or FONT_N
    if fill_ is not None: c.fill = fill_
    c.alignment = align or CENTER
    if nf: c.number_format = nf
    c.protection = LOCK if lock else UNLOCK
    if border: c.border = BORDER
    return c

def protect(ws):
    """Protection sans mot de passe (comme V1.1) ; tri interdit, filtre autorisé (J12)."""
    p = ws.protection
    p.sheet = True  # aucun mot de passe (comme V1.1)
    p.sort = True                     # True = action interdite
    p.autoFilter = False              # filtre autorisé
    p.formatColumns = False; p.formatRows = False
    p.insertRows = True; p.deleteRows = True; p.insertColumns = True; p.deleteColumns = True
    p.formatCells = True; p.selectLockedCells = False; p.selectUnlockedCells = False

N_POS = 70                    # positions SUIVI_ES 8..77
TODAY_TXT = '12/11/2026'

def text_len(v):
    """Longueur affichée approximative (formule ="..."&X&"..." : on compte les littéraux + 4 car. par référence)."""
    if not isinstance(v, str): return len(str(v or ''))
    if v.startswith('='):
        import re
        lits = re.findall(r'"((?:[^"]|"")*)"', v)
        refs = len(re.findall(r'&', v)) // 2
        return sum(len(x) for x in lits) + 4 * refs
    return len(v)

def autoheight(ws, cols_widths, rows, base=15, line=12.6, size_factor=1.0, minimum=None):
    """Hauteur de ligne = nombre de lignes estimé (texte renvoyé à la ligne) × hauteur de ligne."""
    import math
    for r in rows:
        n = 1
        for col, w in cols_widths.items():
            v = ws[f'{col}{r}'].value
            if v is None: continue
            cpl = max(8, w * 1.02 / size_factor)
            n = max(n, sum(max(1, math.ceil(text_len(seg) / cpl)) for seg in str(v).split('\n')) if not str(v).startswith('=') else max(1, math.ceil(text_len(v) / cpl)))
        ws.row_dimensions[r].height = max(minimum or base, round(n * line * size_factor + 5, 1))

# ============================================================ PARAMÈTRES
P = wb.create_sheet('Paramètres')
P.sheet_properties.tabColor = '5B9BD5'
P.sheet_view.showGridLines = False
for col, w in {'A': 44, 'B': 22, 'C': 92}.items(): P.column_dimensions[col].width = w
P.merge_cells('A1:C1')
put(P, 'A1', 'PARAMÈTRES DU SUIVI DES ES', Font(name=ARIAL, size=14, bold=True, color='FFFFFF'), fill(C_VERT), CENTER, border=False)
P.row_dimensions[1].height = 30
P.merge_cells('A2:C2')
put(P, 'A2', "Cellules bleu clair = paramètres modifiables (feuille protégée sans mot de passe : seules ces cellules sont saisissables). "
    "Toute modification met à jour tous les calculs du classeur. Consigner chaque modification dans « Journal des modifications ».",
    Font(name=ARIAL, size=9, italic=True, color='444444'), None, LEFT, border=False)
P.row_dimensions[2].height = 30

def section(ws, ref_range, text):
    ws.merge_cells(ref_range)
    first = ref_range.split(':')[0]
    put(ws, first, text, FONT_HDR10, fill(C_VERT_F), LEFT)
    for row in ws[ref_range]:
        for c in row: c.border = BORDER; c.fill = fill(C_VERT_F)

section(P, 'A3:C3', 'CYCLE DES ENTRETIENS SYSTÉMATIQUES (ordre des 8 étapes)')
for col, t in zip('ABC', ['N° cycle', 'Type ES (Cycle_Types)', 'Remarque']):
    put(P, f'{col}4', t, FONT_HDR, fill(C_VERT_C))
cycle = [S[f'AF{r}'].value for r in range(8, 16)]
assert cycle == ['ES 250 h', 'ES 500 h', 'ES 250 h', 'ES 1000 h', 'ES 250 h', 'ES 500 h', 'ES 250 h', 'ES 2000 h'], cycle
for i, t in enumerate(cycle, start=1):
    r = 4 + i
    put(P, f'A{r}', i, FONT_N, F_CALCUL, nf='"Cycle "0')
    put(P, f'B{r}', t, FONT_IN, F_PARAM, lock=False)
    put(P, f'C{r}', 'Repris de V1.1 (SUIVI_ES!AF' + str(7 + i) + ').' if i > 1 else
        "Repris de V1.1 (SUIVI_ES!AF8:AF15). Le type de chaque étape doit figurer dans la liste Types_ES.", FONT_N, None, LEFT)
put(P, 'A13', 'Longueur du cycle (calculée)', Font(name=ARIAL, size=10, bold=True), F_CALCUL, LEFT)
put(P, 'B13', '=COUNTA(Cycle_Types)', Font(name=ARIAL, size=10, bold=True), F_CALCUL, nf='0" étapes"')
put(P, 'C13', "Nombre d'étapes du cycle (remplace le « 8 » codé en dur dans SUIVI_ES!M:N, les validations et les MFC d'HISTORIQUE_ES).", FONT_N, None, LEFT)

section(P, 'A15:C15', 'RÈGLES DE CALCUL')
put(P, 'A16', 'Marge de sécurité (h) — Marge', FONT_N, None, LEFT)
put(P, 'B16', S['AF17'].value, FONT_IN, F_PARAM, nf='#,##0" h"', lock=False)
put(P, 'C16', 'Retranchée une seule fois : compteur cible = compteur du dernier ES + Interv_ES − Marge.', FONT_N, None, LEFT)
put(P, 'A17', 'Intervalle entre deux ES (h) — Interv_ES', FONT_N, None, LEFT)
put(P, 'B17', S['AF18'].value, FONT_IN, F_PARAM, nf='#,##0" h"', lock=False)
put(P, 'C17', "Identique quel que soit le type d'ES (8 × 250 h = cycle de 2 000 h).", FONT_N, None, LEFT)
put(P, 'A18', 'Avance effective sur le compteur (h)', FONT_N, F_CALCUL, LEFT)
put(P, 'B18', '=Interv_ES-Marge', FONT_N, F_CALCUL, nf='#,##0" h"')
put(P, 'C18', 'Calcul (repris de SUIVI_ES!AF19).', FONT_N, None, LEFT)
put(P, 'A19', 'Seuil du statut « Proche » (h) — Seuil_Proche', FONT_N, None, LEFT)
put(P, 'B19', 20, FONT_IN, F_PARAM, nf='#,##0" h"', lock=False)
put(P, 'C19', '="Normal : écart < −"&Seuil_Proche&" h · Proche : écart de −"&Seuil_Proche&" h à 0 h · Dépassé : écart > 0 h. '
    'Valeur 20 reprise de V1.1 (−20 codé en dur dans SUIVI_ES!J8:J77 et la MFC de I8:I77)."', FONT_N, None, LEFT)
# commentaires d'origine déplacés (AF17, AF18) -> B16, B17
for src, dst in (('AF17', 'B16'), ('AF18', 'B17')):
    cm = S[src].comment
    P[dst].comment = Comment(cm.text, cm.author or 'SUIVI_ES', width=320, height=110)
    S[src].comment = None

section(P, 'A21:C21', "LISTES DE SAISIE")
put(P, 'A22', "Types d'ES — Types_ES", Font(name=ARIAL, size=10, bold=True), None, LEFT)
put(P, 'B22', 'Type', FONT_HDR, fill(C_VERT_C))
put(P, 'C22', 'Usage', FONT_HDR, fill(C_VERT_C))
for i, t in enumerate(['ES 250 h', 'ES 500 h', 'ES 1000 h', 'ES 2000 h'], start=23):
    put(P, f'A{i}', '', FONT_N, None, LEFT)
    put(P, f'B{i}', t, FONT_IN, F_PARAM, lock=False)
    put(P, f'C{i}', 'Liste « Dernier ES — Type » (SUIVI_ES!D) et « Type ES » (HISTORIQUE_ES).', FONT_N, None, LEFT)
put(P, 'A27', '', FONT_N, None, LEFT)
put(P, 'B27', 'Vidange exceptionnelle', FONT_IN, F_PARAM, lock=False)
put(P, 'C27', "HISTORIQUE_ES uniquement (Types_Historique = Types_ES + cette ligne). Hors cycle : ignorée pour la comparaison SUIVI ≠ HISTORIQUE (question ouverte Q3).",
    FONT_N, None, LEFT)

put(P, 'A29', "Catégories d'engins — Liste_Categories", Font(name=ARIAL, size=10, bold=True), None, LEFT)
put(P, 'B29', 'Catégorie', FONT_HDR, fill(C_VERT_C))
put(P, 'C29', '="Positions utilisées : "&Nb_Categories&" — liste de SUIVI_ES!L (valeur hors liste acceptée avec avertissement) et lignes du Tableau de bord."',
    FONT_N, None, LEFT)
cats = []
for r in range(8, 78):
    v = S[f'L{r}'].value
    if v and v not in cats: cats.append(v)
assert cats == ['Camions', 'Chargeuses', 'Bulldozers', 'Sondeuses', 'Engins de servitude'], cats
for i in range(8):
    r = 30 + i
    put(P, f'A{r}', '', FONT_N, None, LEFT)
    put(P, f'B{r}', cats[i] if i < len(cats) else None, FONT_IN, F_PARAM, LEFT, lock=False)
    put(P, f'C{r}', 'Valeurs reprises de SUIVI_ES!L8:L57 (ordre d\'apparition).' if i < len(cats) else 'Emplacement libre.', FONT_N, None, LEFT)
put(P, 'A38', 'Positions de catégories utilisées (calcul)', FONT_N, F_CALCUL, LEFT)
put(P, 'B38', '=SUMPRODUCT(MAX((B30:B37<>"")*(ROW(B30:B37)-ROW(B30)+1)))', FONT_N, F_CALCUL, nf='0')

section(P, 'A40:C40', 'GRAISSAGE / SOUFFLAGE')
put(P, 'A41', 'Intervalle graissage-soufflage (h) — règle à définir', FONT_N, None, LEFT)
put(P, 'B41', None, FONT_IN, F_PARAM, lock=False)
put(P, 'C41', "Vide volontairement : aucune règle d'échéance n'est fournie. Non utilisé par les calculs (logique graissage inchangée). Question ouverte Q5.",
    Font(name=ARIAL, size=10, italic=True, color='7A4100'), None, LEFT)

section(P, 'A43:C43', 'CONTRÔLE DE COHÉRENCE DES PARAMÈTRES')
put(P, 'A44', 'Anomalies de paramétrage (compté dans Total_Alertes)', Font(name=ARIAL, size=10, bold=True), F_CALCUL, LEFT)
put(P, 'B44', '=SUMPRODUCT(--(COUNTIF(Types_ES,Cycle_Types)=0))+IF(AND(ISNUMBER(Marge),ISNUMBER(Interv_ES),ISNUMBER(Seuil_Proche)),'
    'IF(AND(Interv_ES>Marge,Marge>=0,Seuil_Proche>=0),0,1),1)', Font(name=ARIAL, size=10, bold=True), F_CALCUL, nf='0" anomalie(s)"')
put(P, 'C44', "Vérifie : chaque étape du cycle a un type présent dans Types_ES ; Marge, Interv_ES et Seuil_Proche numériques ; Interv_ES > Marge ≥ 0 ; Seuil_Proche ≥ 0.",
    FONT_N, None, LEFT)
autoheight(P, {'A': 44, 'B': 22, 'C': 92}, range(3, 45), base=18, line=13)
P.freeze_panes = 'A4'
# validations Paramètres
dv = DataValidation(type='list', formula1='=Types_ES', allow_blank=False, showErrorMessage=True, errorTitle='Type ES',
                    error='Choisir un type présent dans la liste Types_ES (B23:B26).')
dv.add('B5:B12'); P.add_data_validation(dv)
for ref, lo_, title in (('B16', 0, 'Marge'), ('B17', 1, 'Intervalle ES'), ('B19', 0, 'Seuil Proche')):
    d = DataValidation(type='decimal', operator='between', formula1=str(lo_), formula2='100000', showErrorMessage=True,
                       errorTitle=title, error='Saisir un nombre d\'heures positif.', showInputMessage=True, promptTitle=title,
                       prompt='Paramètre global : modifie tous les calculs. Noter le changement dans le Journal des modifications.')
    d.add(ref); P.add_data_validation(d)
P.conditional_formatting.add('B44', FormulaRule(formula=['$B$44>0'], fill=fill('F4A6A6'), font=Font(name=ARIAL, bold=True, color='7A0000'), stopIfTrue=True))
P.conditional_formatting.add('B44', FormulaRule(formula=['$B$44=0'], fill=fill('BFE6BF'), font=Font(name=ARIAL, bold=True, color='1F5A1F'), stopIfTrue=True))
P.page_setup.orientation = 'landscape'; P.page_setup.paperSize = 9; P.page_setup.fitToWidth = 1; P.page_setup.fitToHeight = 0
P.sheet_properties.pageSetUpPr.fitToPage = True
P.oddFooter.center.text = '&A — page &P / &N'
protect(P)

# ---------------------------------------------------------------- noms définis
def setname(name, ref):
    if name in wb.defined_names: del wb.defined_names[name]
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

setname('Cycle_Types', "'Paramètres'!$B$5:$B$12")
setname('Longueur_cycle', "'Paramètres'!$B$13")
setname('Marge', "'Paramètres'!$B$16")
setname('Interv_ES', "'Paramètres'!$B$17")
setname('Seuil_Proche', "'Paramètres'!$B$19")
setname('Types_ES', "'Paramètres'!$B$23:$B$26")
setname('Types_Historique', "'Paramètres'!$B$23:$B$27")
setname('Type_Hors_Cycle', "'Paramètres'!$B$27")
setname('Nb_Categories', "'Paramètres'!$B$38")
setname('Liste_Categories', "'Paramètres'!$B$30:INDEX('Paramètres'!$B$30:$B$37,MAX(1,'Paramètres'!$B$38))")
setname('Controle_Parametres', "'Paramètres'!$B$44")
setname('Liste_Engins', "'Contrôles'!$V$6:INDEX('Contrôles'!$V$6:$V$75,MAX(1,COUNT('Contrôles'!$U$6:$U$75)))")
setname('Total_Alertes', "'Contrôles'!$D$3")

# ---------------------------------------------------------------- SUIVI_ES : paramètres masqués -> vidés
for r in range(7, 20):
    for col in ('AE', 'AF'):
        c = S[f'{col}{r}']; c.value = None; c.font = FONT_N; c.comment = None
for col in range(12, 35):   # L..AH : afficher
    S.column_dimensions[L(col)].hidden = False
for col in range(18, 35):
    S.column_dimensions[L(col)].width = 13
S.column_dimensions['L'].width = 20
S.column_dimensions['M'].width = 16
S.column_dimensions.group('N', 'Q', outline_level=1, hidden=True)
S.sheet_properties.outlinePr.summaryRight = True
for r in (26, 30, 37, 38, 50):
    S.row_dimensions[r].hidden = False

# en-têtes
S.unmerge_cells('A1:K1'); S.merge_cells('A1:M1')
S['I6'].value = 'ÉCART À LA CIBLE'
S['I7'].value = 'Écart à la cible (h)'
S['I7'].comment = Comment("Écart à la cible = compteur actuel − compteur cible (en heures).\n"
                          "Négatif : il reste X h avant la cible (ex. −172 = 172 h restantes).\n"
                          "Positif : cible dépassée de X h.\n"
                          "Cible = compteur réel du dernier ES + Interv_ES − Marge (voir Paramètres ; 250 − 50 = +200 h en V2), quel que soit le type d'ES.\n"
                          "Calcul et signe identiques à V1.1 (ancien intitulé « Heures restantes »).",
                          S['I7'].comment.author if S['I7'].comment else 'SUIVI_ES', width=360, height=150)
S.merge_cells('L6:M6'); put(S, 'L6', 'CATÉGORIE ET CONTRÔLE', FONT_HDR10, fill(C_VERT))
S['M6'].fill = fill(C_VERT); S['M6'].border = BORDER
S.merge_cells('N6:Q6'); put(S, 'N6', 'CALCULS INTERMÉDIAIRES', FONT_HDR10, fill(C_VERT_F))
for col in 'OPQ': S[f'{col}6'].fill = fill(C_VERT_F); S[f'{col}6'].border = BORDER
for col in 'LMNOPQ':
    put(S, f'{col}7', S[f'{col}7'].value, FONT_HDR, fill(C_VERT_C))
S['L7'].value = 'Catégorie'
S['P7'].value = 'Intervalle ES (h)'

# corps : formules paramétrées + styles
for r in range(8, 78):
    S[f'J{r}'].value = (f'=IF($A{r}="","",IF($K{r}="Hors service","Hors service",IF(NOT(ISNUMBER($I{r})),"À confirmer",'
                        f'IF($I{r}<-Seuil_Proche,"Normal",IF($I{r}<=0,"Proche","Dépassé")))))')
    old_m = S[f'M{r}'].value
    new_m = old_m.replace(f'$E{r}>8,', f'$E{r}>Longueur_cycle,')
    assert new_m != old_m, (r, old_m)
    S[f'M{r}'].value = new_m
    old_n = S[f'N{r}'].value
    new_n = old_n.replace(f'MOD($E{r},8)', f'MOD($E{r},Longueur_cycle)')
    assert new_n != old_n, (r, old_n)
    S[f'N{r}'].value = new_n
    # Catégorie : saisie
    put(S, f'L{r}', None, FONT_IN, F_SAISIE, CENTER, lock=False)
    # colonnes calculées : fond gris
    for col in 'FGHIJ':
        S[f'{col}{r}'].fill = F_CALCUL
    put(S, f'M{r}', None, Font(name=ARIAL, size=9, bold=False, color='404040'), F_CALCUL, CENTER)
    for col in 'NOPQ':
        put(S, f'{col}{r}', None, Font(name=ARIAL, size=9, color='595959'), F_CALCUL, CENTER)
    S[f'P{r}'].number_format = '#,##0" h"'; S[f'Q{r}'].number_format = '#,##0" h"'
    S[f'N{r}'].number_format = '"Cycle "0'

# tuile alertes G2:G3
put(S, 'G2', 'Alertes contrôles', Font(name=ARIAL, size=9, bold=True, color='7A0000'), fill('FCE4E4'), CENTER)
put(S, 'G3', '=Total_Alertes', Font(name=ARIAL, size=14, bold=True, color='7A0000'), fill('FCE4E4'), CENTER, nf='0')
S.conditional_formatting.add('G3', FormulaRule(formula=['$G$3=0'], fill=fill('BFE6BF'), font=Font(name=ARIAL, bold=True, color='1F5A1F'), stopIfTrue=True))
S['G2'].hyperlink = "#'Contrôles'!A1"
# légende et avertissement (lignes 4-5)
S.merge_cells('A4:M4')
put(S, 'A4', '="Écart à la cible (h) = compteur actuel − compteur cible : NÉGATIF = heures restantes avant la cible ; POSITIF = cible dépassée.   '
    'Statut : Normal si écart < −"&Seuil_Proche&" h · Proche de −"&Seuil_Proche&" h à 0 h · Dépassé si > 0 h · '
    'Cible = dernier ES + "&Interv_ES&" h − "&Marge&" h de marge (voir Paramètres)."',
    Font(name=ARIAL, size=9, color=C_TXT), fill('EEF3EE'), LEFT, border=False)
S.merge_cells('A5:M5')
put(S, 'A5', "⚠ NE PAS TRIER, INSÉRER NI SUPPRIMER DE LIGNES : HISTORIQUE_ES, SUIVI_GRAISSAGE_SOUFFLAGE et HISTORIQUE_GRAISSAGE_SOUFFLAGE sont liés à "
    "cette feuille par position de ligne. Pour masquer ou afficher des engins, utiliser les filtres. Nouvel engin : première ligne libre en bas du tableau.",
    Font(name=ARIAL, size=9, bold=True, color='9C0006'), fill('FDECEC'), LEFT, border=False)
S.row_dimensions[4].height = 26; S.row_dimensions[5].height = 26

# MFC : seuil -20 -> Seuil_Proche ; règle M
n_cf = 0
for cf, rules in S.conditional_formatting._cf_rules.items():
    for rule in rules:
        new = [f.replace('$I8<-20', '$I8<-Seuil_Proche').replace('$I8>=-20', '$I8>=-Seuil_Proche') for f in rule.formula]
        if new != rule.formula: n_cf += 1; rule.formula = new
assert n_cf == 2, n_cf
S.conditional_formatting.add('M8:M77', FormulaRule(formula=['AND($A8<>"",$K8<>"Hors service",$M8<>"OK")'],
                                                   fill=fill('FFD9A8'), font=Font(name=ARIAL, bold=True, color='7A4100')))

# repérage des engins ayant au moins une alerte dans Contrôles (priorité juste après la règle « Hors service »)
rule_alert = FormulaRule(formula=['AND($A8<>"",N(INDEX(\'Contrôles\'!$M$6:$M$75,ROW()-ROW($A$7)))>0)'],
                         fill=fill('FCE4E4'), font=Font(name=ARIAL, bold=True, color='9C0006'), stopIfTrue=True)
S.conditional_formatting.add('A8:A77', rule_alert)
allrules = sorted([r for rules in S.conditional_formatting._cf_rules.values() for r in rules], key=lambda r: r.priority)
allrules.remove(rule_alert); allrules.insert(1, rule_alert)
for i, r_ in enumerate(allrules, start=1): r_.priority = i
S.merge_cells('H2:M3')
put(S, 'H2', "Légende : saisie = jaune clair (texte bleu) · calcul = gris · nom d'engin sur fond rose = alerte à traiter dans Contrôles "
    "(cliquer « Alertes contrôles »). Paramètres (marge, intervalle, seuil Proche, cycle) : feuille Paramètres.",
    Font(name=ARIAL, size=8, italic=True, color='444444'), fill('F7F7F7'), LEFT, border=False)
# validations SUIVI_ES
old = list(S.data_validations.dataValidation)
S.data_validations.dataValidation = []
for d in old:
    sq = str(d.sqref)
    if sq == 'B8:C77':
        for col, title, prompt in (('B', 'Compteur actuel (h)', "Relevé du compteur horaire à la date de mise à jour. Vide = statut « À confirmer »."),
                                   ('C', 'Dernier ES — compteur (h)', "Compteur relevé lors du dernier ES réalisé. Reporter la même entrée dans HISTORIQUE_ES.")):
            n = DataValidation(type='decimal', operator='between', formula1='0', formula2='999999', allow_blank=True,
                               showErrorMessage=True, errorTitle=d.errorTitle, error=d.error,
                               showInputMessage=True, promptTitle=title, prompt=prompt)
            n.add(f'{col}8:{col}77'); S.add_data_validation(n)
    elif sq == 'D8:D77':
        n = DataValidation(type='list', formula1='=Types_ES', allow_blank=True, showErrorMessage=True, errorTitle='Type ES',
                           error="Choisir un type dans la liste (Paramètres › Types_ES).", showInputMessage=True,
                           promptTitle='Dernier ES — Type', prompt="Type du dernier ES réalisé. Doit correspondre au N° de cycle (ordre du cycle dans Paramètres).")
        n.add('D8:D77'); S.add_data_validation(n)
    elif sq == 'E8:E77':
        n = DataValidation(type='whole', operator='between', formula1='1', formula2='Longueur_cycle', allow_blank=True,
                           showErrorMessage=True, errorTitle='N° cycle', error="Le N° de cycle doit être un entier de 1 à la longueur du cycle (8, voir Paramètres).",
                           showInputMessage=True, promptTitle='N° de cycle (1 à 8, voir Paramètres)',
                           prompt="N° d'étape de l'ES réalisé = celui annoncé dans « Prochain ES — N° cycle » (col. H) avant la mise à jour. "
                                  "ES refait deux fois de suite : garder le MÊME N°. Après la dernière étape : revenir à 1.")
        n.add('E8:E77'); S.add_data_validation(n)
    else:
        S.add_data_validation(d)
n = DataValidation(type='list', formula1='=Liste_Categories', allow_blank=True, showErrorMessage=True, errorStyle='warning',
                   errorTitle='Catégorie', error="Catégorie absente de la liste (Paramètres). Continuer quand même ?",
                   showInputMessage=True, promptTitle='Catégorie', prompt="Famille d'engin, utilisée pour le tableau Catégorie × Statut du Tableau de bord.")
n.add('L8:L77'); S.add_data_validation(n)
S.auto_filter.ref = 'A7:M77'
protect(S)

# ---------------------------------------------------------------- HISTORIQUE_ES
n_cf = 0
for cf, rules in H.conditional_formatting._cf_rules.items():
    for rule in rules:
        new = [f.replace(',0))),8)', ',0))),Longueur_cycle)') for f in rule.formula]
        if new != rule.formula: n_cf += 1; rule.formula = new
assert n_cf == 6, n_cf
for d in H.data_validations.dataValidation:
    if d.type == 'list' and d.formula1 and 'ES 250 h' in d.formula1:
        d.formula1 = '=Types_Historique'
        d.error = "Choisir un type dans la liste (Paramètres › Types_ES + Vidange exceptionnelle)."
    elif d.type == 'whole' and d.formula2 == '8':
        d.formula2 = 'Longueur_cycle'
        d.error = "Le N° de cycle doit être un entier de 1 à la longueur du cycle (8, voir Paramètres)."
AVERT = (" ⚠ Ne jamais trier, insérer ni supprimer de lignes (ici comme dans SUIVI_ES) : chaque bloc est lié par position à la ligne "
         "correspondante de SUIVI_ES (contrôle « Ordre / liens » dans Contrôles).")
H['A2'].value = H['A2'].value + AVERT
H.row_dimensions[2].height = 50
protect(H)

# ---------------------------------------------------------------- SUIVI_GRAISSAGE_SOUFFLAGE (logique inchangée)
G['A3'].value = G['A3'].value + AVERT.replace('chaque bloc est lié', 'chaque ligne est liée').replace('correspondante', 'correspondante')
G.row_dimensions[3].height = 46
for r in range(6, 76):
    G[f'B{r}'].fill = F_CALCUL; G[f'F{r}'].fill = F_CALCUL
protect(G)

# ---------------------------------------------------------------- HISTORIQUE_GRAISSAGE_SOUFFLAGE
HG['A2'].value = HG['A2'].value + AVERT
HG.row_dimensions[2].height = 50
protect(HG)

# ============================================================ CONTRÔLES
C = wb.create_sheet('Contrôles')
C.sheet_properties.tabColor = 'C00000'
C.sheet_view.showGridLines = False
widths = {'A': 5, 'B': 8, 'C': 30, 'D': 12, 'E': 12, 'F': 24, 'G': 19, 'H': 12, 'I': 18, 'J': 13, 'K': 30, 'L': 20,
          'M': 9, 'N': 10, 'O': 12, 'P': 12, 'Q': 9, 'R': 30, 'S': 11, 'T': 9, 'U': 9, 'V': 26, 'W': 10}
for k, v in widths.items(): C.column_dimensions[k].width = v
C.merge_cells('A1:R1')
put(C, 'A1', 'CONTRÔLES DE COHÉRENCE — une ligne par position de SUIVI_ES (lignes 8 à 77)', Font(name=ARIAL, size=14, bold=True, color='FFFFFF'),
    fill('8B1E1E'), CENTER, border=False)
C.row_dimensions[1].height = 30
C.merge_cells('A2:R2')
put(C, 'A2', "Lecture seule : aucune donnée n'est modifiée. Chaque alerte se corrige dans la feuille d'origine (SUIVI_ES, HISTORIQUE_ES…) ; "
    "un conflit SUIVI ≠ HISTORIQUE doit être tranché par le responsable (ne pas « forcer » l'une des deux valeurs sans vérification terrain).",
    Font(name=ARIAL, size=9, italic=True, color='444444'), None, LEFT, border=False)
C.row_dimensions[2].height = 28
put(C, 'A3', None, FONT_N, None, border=False)
C.merge_cells('A3:C3'); put(C, 'A3', 'TOTAL DES ALERTES (Total_Alertes)', Font(name=ARIAL, size=11, bold=True, color='7A0000'), fill('FCE4E4'), LEFT)
put(C, 'D3', '=SUM(M6:M75)+Controle_Parametres', Font(name=ARIAL, size=14, bold=True, color='7A0000'), fill('FCE4E4'), nf='0')
C.merge_cells('E3:F3'); put(C, 'E3', 'Engins concernés', Font(name=ARIAL, size=10, bold=True), F_CALCUL, LEFT)
put(C, 'G3', '=COUNTIF(M6:M75,">0")', Font(name=ARIAL, size=12, bold=True), F_CALCUL, nf='0')
C.merge_cells('H3:I3'); put(C, 'H3', 'Anomalies Paramètres', Font(name=ARIAL, size=10, bold=True), F_CALCUL, LEFT)
put(C, 'J3', '=Controle_Parametres', Font(name=ARIAL, size=12, bold=True), F_CALCUL, nf='0')
C.merge_cells('K3:R3')
put(C, 'K3', "Référence d'ordre (col. R) figée le " + TODAY_TXT + " à partir de SUIVI_ES!A8:A57 : un écart signale un tri, une insertion, une suppression ou un renommage.",
    Font(name=ARIAL, size=8, italic=True, color='444444'), None, LEFT, border=False)
C.row_dimensions[3].height = 30
C.merge_cells('A4:R4')
put(C, 'A4', "⚠ Ne pas trier, insérer ni supprimer de lignes dans les feuilles de suivi et d'historique (liaison par position). Colonnes S:W = calculs techniques (groupées).",
    Font(name=ARIAL, size=9, bold=True, color='9C0006'), fill('FDECEC'), LEFT, border=False)
hdrs = ['N°', 'Ligne SUIVI_ES', 'Engin', 'État', 'Statut', 'Saisie incomplète ou invalide', 'Compteur < dernier ES', 'Type ≠ cycle',
        'SUIVI ≠ HISTORIQUE', 'Nom en double', 'Ordre / liens des lignes', "Date d'historique invalide", 'Nb alertes',
        'Nb entrées historique ES', 'Dernier ES historisé — Compteur', 'Dernier ES historisé — Type', 'Dernier ES historisé — N°',
        "Engin attendu (référence d'ordre)", 'Clé priorité (écart)', 'Rang priorité', 'Rang liste', 'Liste engins (compacte)', 'Position dernier ES historisé']
for i, h in enumerate(hdrs, start=1):
    put(C, f'{L(i)}5', h, FONT_HDR, fill('8B1E1E' if 6 <= i <= 13 else (C_VERT_C if i <= 18 else '7F7F7F')))
C.row_dimensions[5].height = 42
HR = "'HISTORIQUE_ES'!$C$6:$AP$285"
SA = lambda col: f"SUIVI_ES!${col}$8:${col}$77"
ref_names = {k: S[f'A{7 + k}'].value for k in range(1, N_POS + 1)}
for k in range(1, N_POS + 1):
    r = 5 + k
    f = {}
    f['A'] = k
    f['B'] = f'=IF($A{r}>ROWS({SA("A")}),"—",ROW(INDEX({SA("A")},$A{r})))'
    f['C'] = f'=IF($A{r}>ROWS({SA("A")}),"",IF(INDEX({SA("A")},$A{r})="","",INDEX({SA("A")},$A{r})))'
    f['D'] = f'=IF($C{r}="","",INDEX({SA("K")},$A{r})&"")'
    f['E'] = f'=IF($C{r}="","",INDEX({SA("J")},$A{r})&"")'
    f['F'] = (f'=IF(OR($C{r}="",$D{r}="Hors service"),"",IF($D{r}="","État non renseigné",IF(NOT(ISNUMBER(INDEX({SA("B")},$A{r}))),'
              f'"Compteur actuel manquant ou non numérique",IF(OR(INDEX({SA("M")},$A{r})="Non renseigné",INDEX({SA("M")},$A{r})="Incomplet"),'
              f'"Dernier ES incomplet",IF(INDEX({SA("M")},$A{r})="Valeur invalide","Dernier ES : valeur invalide","")))))')
    f['G'] = (f'=IF(OR($C{r}="",$D{r}="Hors service"),"",IF(AND(ISNUMBER(INDEX({SA("B")},$A{r})),ISNUMBER(INDEX({SA("C")},$A{r}))),'
              f'IF(INDEX({SA("B")},$A{r})<INDEX({SA("C")},$A{r}),"Compteur < dernier ES",""),""))')
    f['H'] = f'=IF(OR($C{r}="",$D{r}="Hors service"),"",IF(INDEX({SA("M")},$A{r})="Type ≠ cycle","Type ≠ cycle",""))'
    f['I'] = (f'=IF(OR($C{r}="",$W{r}="",$W{r}=0),"",IF(OR($O{r}<>INDEX({SA("C")},$A{r}),$P{r}<>INDEX({SA("D")},$A{r}),'
              f'$Q{r}<>INDEX({SA("E")},$A{r})),"SUIVI ≠ HISTORIQUE",""))')
    f['J'] = f'=IF($C{r}="","",IF(SUMPRODUCT(--(TRIM({SA("A")})=TRIM($C{r})))>1,"Nom en double",""))'
    f['K'] = (f'=IF(AND($R{r}<>"",$C{r}<>$R{r}),"Ordre des lignes modifié (attendu : "&$R{r}&")",IF($C{r}="","",'
              f'IF(IFERROR(INDEX(\'HISTORIQUE_ES\'!$A$6:$A$285,4*$A{r}-3)<>$C{r},TRUE),"Lien HISTORIQUE_ES rompu",'
              f'IF(IFERROR(INDEX(\'SUIVI_GRAISSAGE_SOUFFLAGE\'!$A$6:$A$75,$A{r})<>$C{r},TRUE),"Lien SUIVI_GRAISSAGE rompu",'
              f'IF(IFERROR(INDEX(\'HISTORIQUE_GRAISSAGE_SOUFFLAGE\'!$A$6:$A$215,3*$A{r}-2)<>$C{r},TRUE),"Lien HIST. GRAISSAGE rompu","")))))')
    f['L'] = (f'=IF($C{r}="","",IF(SUMPRODUCT(ISTEXT(INDEX({HR},4*$A{r}-2,0))+ISNUMBER(INDEX({HR},4*$A{r}-2,0))*'
              f'(INDEX({HR},4*$A{r}-2,0)>TODAY()))>0,"Date d\'historique invalide",""))')
    f['M'] = f'=IF(AND($C{r}="",$K{r}=""),"",SUMPRODUCT(--(LEN($F{r}:$L{r})>0)))'
    f['N'] = f'=IF($C{r}="","",COUNT(INDEX({HR},4*$A{r}-3,0)))'
    f['O'] = f'=IF(OR($W{r}="",$W{r}=0),"",INDEX({HR},4*$A{r}-3,$W{r}))'
    f['P'] = f'=IF(OR($W{r}="",$W{r}=0),"",INDEX({HR},4*$A{r}-1,$W{r}))'
    f['Q'] = f'=IF(OR($W{r}="",$W{r}=0),"",INDEX({HR},4*$A{r},$W{r}))'
    f['R'] = ref_names[k]
    f['S'] = f'=IF(OR($E{r}="Dépassé",$E{r}="Proche"),INDEX({SA("I")},$A{r}),"")'
    f['T'] = (f'=IF(ISNUMBER($S{r}),SUMPRODUCT(ISNUMBER($S$6:$S$75)*($S$6:$S$75>$S{r}))+SUMPRODUCT(ISNUMBER($S$6:$S{r})*($S$6:$S{r}=$S{r})),"")')
    f['U'] = f'=IF($C{r}="","",SUMPRODUCT(--($C$6:$C{r}<>"")))'
    f['V'] = f'=IFERROR(INDEX($C$6:$C$75,MATCH($A{r},$U$6:$U$75,0)),"")'
    f['W'] = (f'=IF($C{r}="","",SUMPRODUCT(MAX((INDEX({HR},4*$A{r}-3,0)<>"")*(INDEX({HR},4*$A{r}-1,0)<>Type_Hors_Cycle)*'
              f'(COLUMN(\'HISTORIQUE_ES\'!$C$6:$AP$6)-COLUMN(\'HISTORIQUE_ES\'!$C$6)+1))))')
    for col, val in f.items():
        al = LEFT if col in 'CFGHIJKLRV' else CENTER
        fnt = Font(name=ARIAL, size=9, bold=(col in 'CM'), color=('595959' if col in 'STUVW' else '000000'))
        put(C, f'{col}{r}', val, fnt, F_CALCUL, al)
    C[f'O{r}'].number_format = '#,##0" h"'; C[f'Q{r}'].number_format = '"Cycle "0'; C[f'S{r}'].number_format = '+#,##0" h";-#,##0" h";0" h"'
    C.row_dimensions[r].height = 15
C.column_dimensions.group('S', 'W', outline_level=1, hidden=True)
C.freeze_panes = 'D6'
C.auto_filter.ref = 'A5:R75'
C.conditional_formatting.add('F6:L75', FormulaRule(formula=['LEN(F6)>0'], fill=fill('F8C9C9'), font=Font(name=ARIAL, bold=True, color='9C0006')))
C.conditional_formatting.add('M6:M75', FormulaRule(formula=['AND(ISNUMBER($M6),$M6>0)'], fill=fill('F4A6A6'), font=Font(name=ARIAL, bold=True, color='7A0000')))
for st, fc, tc in (('Normal', 'BFE6BF', '1F5A1F'), ('Proche', 'FFD9A8', '7A4100'), ('Dépassé', 'F4A6A6', '7A0000'), ('À confirmer', 'F0F1F3', '666666'), ('Hors service', 'DADADA', '404040')):
    C.conditional_formatting.add('E6:E75', FormulaRule(formula=[f'$E6="{st}"'], fill=fill(fc), font=Font(name=ARIAL, bold=True, color=tc)))
C.conditional_formatting.add('D3', FormulaRule(formula=['$D$3=0'], fill=fill('BFE6BF'), font=Font(name=ARIAL, bold=True, color='1F5A1F'), stopIfTrue=True))
C.page_setup.orientation = 'landscape'; C.page_setup.paperSize = 9; C.page_setup.fitToWidth = 1; C.page_setup.fitToHeight = 0
C.sheet_properties.pageSetUpPr.fitToPage = True
C.print_title_rows = '5:5'
C.print_area = 'A1:R75'
C.oddFooter.center.text = '&A — page &P / &N'
protect(C)

# ============================================================ TABLEAU DE BORD
D = wb.create_sheet('Tableau de bord')
D.sheet_properties.tabColor = '1F6B1C'
D.sheet_view.showGridLines = False
D.column_dimensions['A'].width = 2
for i in range(2, 16): D.column_dimensions[L(i)].width = 11.5
D.column_dimensions['P'].width = 9
D.merge_cells('B1:O1')
put(D, 'B1', 'TABLEAU DE BORD — SUIVI DES ENTRETIENS SYSTÉMATIQUES (ES) · ENGINS DE BENGUERIR', Font(name=ARIAL, size=15, bold=True, color='FFFFFF'), fill(C_VERT), CENTER, border=False)
for col in range(2, 16): D.cell(1, col).fill = fill(C_VERT)
D.row_dimensions[1].height = 34
D.merge_cells('B2:O2')
put(D, 'B2', '="Cible = dernier ES + "&Interv_ES&" h − "&Marge&" h de marge   ·   Normal : écart < −"&Seuil_Proche&" h   ·   Proche : écart de −"&Seuil_Proche&" à 0 h   ·   Dépassé : écart > 0 h   ·   Paramètres modifiables dans la feuille Paramètres"',
    Font(name=ARIAL, size=9, italic=True, color='444444'), None, CENTER, border=False)
D.row_dimensions[2].height = 20
tiles = [('Engins en service', '=SUIVI_ES!$A$3', '="sur "&COUNTA(SUIVI_ES!$A$8:$A$77)&" engins suivis"', 'E8EEE8', C_TXT),
         ('Normal', '=SUIVI_ES!$B$3', '=IF($B$4=0,"",TEXT(D4/$B$4,"0%")&" des engins en service")', 'BFE6BF', '1F5A1F'),
         ('Proche', '=SUIVI_ES!$C$3', '=IF($B$4=0,"",TEXT(F4/$B$4,"0%")&" des engins en service")', 'FFD9A8', '7A4100'),
         ('Dépassé', '=SUIVI_ES!$D$3', '=IF($B$4=0,"",TEXT(H4/$B$4,"0%")&" des engins en service")', 'F4A6A6', '7A0000'),
         ('À confirmer', '=SUIVI_ES!$E$3', '="saisie à compléter ou corriger"', 'F0F1F3', '666666'),
         ('Hors service', '=SUIVI_ES!$F$3', '="exclus des calculs"', 'D9D9D9', '404040'),
         ('Alertes Contrôles', '=Total_Alertes', '="engins concernés : "&\'Contrôles\'!$G$3', 'FCE4E4', '7A0000')]
for i, (lab, val, sub, bg, fg) in enumerate(tiles):
    c1, c2 = L(2 + 2 * i), L(3 + 2 * i)
    for row, content, fnt, h in ((3, lab, Font(name=ARIAL, size=10, bold=True, color=fg), 22), (4, val, Font(name=ARIAL, size=24, bold=True, color=fg), 40),
                                 (5, sub, Font(name=ARIAL, size=8, italic=True, color=fg), 18)):
        D.merge_cells(f'{c1}{row}:{c2}{row}')
        put(D, f'{c1}{row}', content, fnt, fill(bg), CENTER, nf='0' if row == 4 else None, border=False)
        D[f'{c2}{row}'].fill = fill(bg)
        D.row_dimensions[row].height = h
    if i == 6:
        D[f'{c1}3'].hyperlink = "#'Contrôles'!A1"
D.conditional_formatting.add('N3:O5', FormulaRule(formula=['$N$4=0'], fill=fill('BFE6BF'), font=Font(name=ARIAL, bold=True, color='1F5A1F')))
D.row_dimensions[6].height = 10
# Catégorie x Statut (B7:I18)
D.merge_cells('B7:I7'); put(D, 'B7', 'RÉPARTITION PAR CATÉGORIE ET STATUT', FONT_HDR10, fill(C_VERT_F), LEFT)
D.merge_cells('J7:O7'); put(D, 'J7', 'STATUTS DES ENGINS (nombre)', FONT_HDR10, fill(C_VERT_F), LEFT)
STATUTS = ['Normal', 'Proche', 'Dépassé', 'À confirmer', 'Hors service']
D.merge_cells('B8:C8'); put(D, 'B8', 'Catégorie', FONT_HDR, fill(C_VERT_C)); D['C8'].border = BORDER
for j, st in enumerate(STATUTS):
    put(D, f'{L(4 + j)}8', st, FONT_HDR, fill(C_VERT_C))
put(D, 'I8', 'Total', FONT_HDR, fill(C_VERT_C))
for i in range(8):
    r = 9 + i
    D.merge_cells(f'B{r}:C{r}')
    put(D, f'B{r}', f"=IF('Paramètres'!$B{30 + i}=\"\",\"\",'Paramètres'!$B{30 + i})", Font(name=ARIAL, size=9, bold=True), None, LEFT)
    D[f'C{r}'].border = BORDER
    for j in range(5):
        col = L(4 + j)
        put(D, f'{col}{r}', f'=IF($B{r}="","",COUNTIFS(SUIVI_ES!$L$8:$L$77,$B{r},SUIVI_ES!$J$8:$J$77,{col}$8))', Font(name=ARIAL, size=9), None, CENTER, nf='0;-0;"·"')
    put(D, f'I{r}', f'=IF($B{r}="","",SUM(D{r}:H{r}))', Font(name=ARIAL, size=9, bold=True), F_CALCUL, CENTER, nf='0')
D.merge_cells('B17:C17'); put(D, 'B17', 'Autre / non renseignée', Font(name=ARIAL, size=9, italic=True), None, LEFT); D['C17'].border = BORDER
D.merge_cells('B18:C18'); put(D, 'B18', 'Total', Font(name=ARIAL, size=9, bold=True), F_CALCUL, LEFT); D['C18'].border = BORDER
for j in range(5):
    col = L(4 + j)
    put(D, f'{col}18', f'=COUNTIF(SUIVI_ES!$J$8:$J$77,{col}$8)', Font(name=ARIAL, size=9, bold=True), F_CALCUL, CENTER, nf='0')
    put(D, f'{col}17', f'={col}18-SUM({col}9:{col}16)', Font(name=ARIAL, size=9, italic=True), None, CENTER, nf='0;-0;"·"')
put(D, 'I17', '=SUM(D17:H17)', Font(name=ARIAL, size=9, italic=True), F_CALCUL, CENTER, nf='0;-0;"·"')
put(D, 'I18', '=SUM(D18:H18)', Font(name=ARIAL, size=9, bold=True), F_CALCUL, CENTER, nf='0')
for r in range(8, 19): D.row_dimensions[r].height = 18
for col, (bg, fg) in zip('DEFGH', [('BFE6BF', '1F5A1F'), ('FFD9A8', '7A4100'), ('F4A6A6', '7A0000'), ('F0F1F3', '666666'), ('DADADA', '404040')]):
    D.conditional_formatting.add(f'{col}9:{col}17', FormulaRule(formula=[f'AND(ISNUMBER({col}9),{col}9>0)'], fill=fill(bg), font=Font(name=ARIAL, bold=True, color=fg)))
# graphique des statuts (série unique : pas de légende ; étiquettes de valeur ; couleurs de statut réservées)
ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
ch.title = None; ch.legend = None; ch.style = 10
data = Reference(D, min_col=4, max_col=8, min_row=18, max_row=18)
cats_ref = Reference(D, min_col=4, max_col=8, min_row=8, max_row=8)
ch.add_data(data, from_rows=True, titles_from_data=False)
ch.set_categories(cats_ref)
s = ch.series[0]
s.graphicalProperties = GraphicalProperties(solidFill='0CA30C')
s.graphicalProperties.line = LineProperties(noFill=True)
for idx, col in enumerate(['0CA30C', 'FAB219', 'D03B3B', '9AA0A6', '6B6B6B']):
    pt = DataPoint(idx=idx); pt.graphicalProperties = GraphicalProperties(solidFill=col); pt.graphicalProperties.line = LineProperties(noFill=True)
    s.dPt.append(pt)
s.dLbls = DataLabelList(); s.dLbls.showVal = True; s.dLbls.showSerName = False; s.dLbls.showCatName = False; s.dLbls.showLegendKey = False; s.dLbls.position = 'outEnd'
ch.gapWidth = 60
ch.y_axis.majorGridlines = None
ch.y_axis.delete = True
ch.x_axis.delete = False
ch.x_axis.title = None
ch.height = 6.9; ch.width = 14.2
D.add_chart(ch, 'J8')
# À traiter en priorité
D.row_dimensions[19].height = 10
D.merge_cells('B20:O20')
put(D, 'B20', 'À TRAITER EN PRIORITÉ — Dépassé puis Proche, du plus dépassé au moins dépassé (15 premiers)', FONT_HDR10, fill('8B1E1E'), LEFT)
heads = [('B', 'B', 'Rang'), ('C', 'E', 'Engin'), ('F', 'G', 'Catégorie'), ('H', 'H', 'Statut'), ('I', 'I', 'Écart à la cible (h)'),
         ('J', 'J', 'Compteur actuel'), ('K', 'K', 'Compteur cible'), ('L', 'L', 'Prochain ES'), ('M', 'M', 'N° cycle'), ('N', 'O', 'Ligne SUIVI_ES')]
for a, b, t in heads:
    if a != b: D.merge_cells(f'{a}21:{b}21')
    put(D, f'{a}21', t, FONT_HDR, fill(C_VERT_C))
    for col in range(openpyxl.utils.column_index_from_string(a), openpyxl.utils.column_index_from_string(b) + 1):
        D.cell(21, col).border = BORDER; D.cell(21, col).fill = fill(C_VERT_C)
D.row_dimensions[21].height = 30
put(D, 'P21', 'Position (technique)', Font(name=ARIAL, size=8, color='595959'), F_CALCUL)
for i in range(1, 16):
    r = 21 + i
    pos = f'$P{r}'
    put(D, f'P{r}', f"=IFERROR(MATCH({i},'Contrôles'!$T$6:$T$75,0),\"\")", Font(name=ARIAL, size=8, color='595959'), F_CALCUL)
    cells = {'B': (f'=IF({pos}="","",{i})', '0', CENTER),
             'C': (f"=IF({pos}=\"\",\"\",INDEX('Contrôles'!$C$6:$C$75,{pos}))", None, LEFT),
             'F': (f'=IF({pos}="","",INDEX(SUIVI_ES!$L$8:$L$77,{pos})&"")', None, LEFT),
             'H': (f'=IF({pos}="","",INDEX(SUIVI_ES!$J$8:$J$77,{pos}))', None, CENTER),
             'I': (f'=IF({pos}="","",INDEX(SUIVI_ES!$I$8:$I$77,{pos}))', '+#,##0" h";-#,##0" h";0" h"', CENTER),
             'J': (f'=IF({pos}="","",INDEX(SUIVI_ES!$B$8:$B$77,{pos}))', '#,##0" h"', CENTER),
             'K': (f'=IF({pos}="","",INDEX(SUIVI_ES!$F$8:$F$77,{pos}))', '#,##0" h"', CENTER),
             'L': (f'=IF({pos}="","",INDEX(SUIVI_ES!$G$8:$G$77,{pos}))', None, CENTER),
             'M': (f'=IF({pos}="","",INDEX(SUIVI_ES!$H$8:$H$77,{pos}))', '"Cycle "0', CENTER),
             'N': (f'=IF({pos}="","",ROW(INDEX(SUIVI_ES!$A$8:$A$77,{pos})))', '0', CENTER)}
    for col, (val, nf, al) in cells.items():
        put(D, f'{col}{r}', val, Font(name=ARIAL, size=9, bold=(col in 'CHI')), None, al, nf=nf)
    for a, b in (('C', 'E'), ('F', 'G'), ('N', 'O')):
        D.merge_cells(f'{a}{r}:{b}{r}')
        for col in range(openpyxl.utils.column_index_from_string(a), openpyxl.utils.column_index_from_string(b) + 1):
            D.cell(r, col).border = BORDER
    D.row_dimensions[r].height = 17
for st, bg, fg in (('Proche', 'FFD9A8', '7A4100'), ('Dépassé', 'F4A6A6', '7A0000')):
    D.conditional_formatting.add('H22:I36', FormulaRule(formula=[f'$H22="{st}"'], fill=fill(bg), font=Font(name=ARIAL, bold=True, color=fg)))
D.conditional_formatting.add('B22:O36', FormulaRule(formula=['AND($P22<>"",MOD(ROW(),2)=0)'], fill=fill('F4F8F4')))
D.merge_cells('B37:O37')
put(D, 'B37', "=IF(COUNT('Contrôles'!$S$6:$S$75)=0,\"Aucun engin Dépassé ou Proche.\",\"Engins Dépassé ou Proche : \"&COUNT('Contrôles'!$S$6:$S$75)&"
    "IF(COUNT('Contrôles'!$S$6:$S$75)>15,\" — seuls les 15 premiers sont affichés (liste complète : filtrer SUIVI_ES sur le statut).\",\".\"))",
    Font(name=ARIAL, size=9, italic=True, color='444444'), None, LEFT, border=False)
D.column_dimensions.group('P', 'P', outline_level=1, hidden=True)
D.print_area = 'B1:O37'
D.page_setup.orientation = 'landscape'; D.page_setup.paperSize = 9
D.page_setup.fitToWidth = 1; D.page_setup.fitToHeight = 1
D.sheet_properties.pageSetUpPr.fitToPage = True
D.print_options.horizontalCentered = True
D.page_margins.left = D.page_margins.right = 0.4; D.page_margins.top = D.page_margins.bottom = 0.5
D.oddFooter.center.text = '&A — imprimé le &D'
protect(D)

# ============================================================ MODE D'EMPLOI
M = wb.create_sheet("Mode d'emploi")
M.sheet_properties.tabColor = '7F7F7F'
M.sheet_view.showGridLines = False
M.column_dimensions['A'].width = 2; M.column_dimensions['B'].width = 24; M.column_dimensions['C'].width = 110
M.merge_cells('B1:C1')
put(M, 'B1', "MODE D'EMPLOI — SUIVI DES ENTRETIENS SYSTÉMATIQUES (ES)", Font(name=ARIAL, size=15, bold=True, color='FFFFFF'), fill(C_VERT), CENTER, border=False)
M['C1'].fill = fill(C_VERT)
M.row_dimensions[1].height = 34
rows = []
def sec(t): rows.append(('sec', t, None))
def line(a, b, h=None, fl=None): rows.append(('line', a, b, h, fl))
sec('1. OBJET')
line('Pourquoi ce classeur', "Suivre les entretiens systématiques (ES) des 50 engins de Benguerir (OCP) : à partir du compteur actuel et du dernier ES réalisé, "
     "le classeur calcule le prochain ES (type, N° de cycle, compteur cible), l'écart à la cible et un statut ; il conserve l'historique des ES et des graissages-soufflages.", 32)
line('Ce qui est saisi', "Compteur actuel, dernier ES (compteur, type, N° de cycle), état et catégorie de l'engin (SUIVI_ES) ; historiques ; dernier graissage-soufflage.", 18)
line('Ce qui est calculé', "Prochain ES, écart à la cible, statut, contrôles de saisie, alertes (Contrôles), synthèse (Tableau de bord). Les paramètres sont dans la feuille Paramètres.", 18)
sec('2. PROCÉDURE MENSUELLE EN 5 ÉTAPES')
line('Étape 1 — Compteurs', "SUIVI_ES, colonne B : saisir le compteur horaire actuel de chaque engin en service. Engin arrêté : colonne K = « Hors service » "
     "(son compteur se saisit alors dans SUIVI_GRAISSAGE_SOUFFLAGE, colonne G).", 32)
line('Étape 2 — ES réalisés', "Pour chaque ES réalisé depuis le dernier relevé : SUIVI_ES colonnes C (compteur de réalisation), D (type), E (N° de cycle = celui annoncé "
     "en colonne H avant la mise à jour). Puis reporter la MÊME entrée dans HISTORIQUE_ES (colonne suivante du bloc de l'engin : compteur, date, type, N°).", 32)
line('Étape 3 — Graissage', "SUIVI_GRAISSAGE_SOUFFLAGE : date, compteur et contexte du dernier graissage-soufflage ; puis l'ajouter dans HISTORIQUE_GRAISSAGE_SOUFFLAGE.", 18)
line('Étape 4 — Contrôles', "Ouvrir Contrôles (total en tête de SUIVI_ES et sur le Tableau de bord) : traiter chaque alerte dans la feuille d'origine jusqu'à 0, "
     "ou noter la justification. Un écart « SUIVI ≠ HISTORIQUE » se tranche avec le terrain, jamais en recopiant l'une des valeurs sans vérifier.", 32)
line('Étape 5 — Priorités', "Tableau de bord : lire la liste « À traiter en priorité » (Dépassé puis Proche) et planifier les ES. Impression : A4 paysage, une page.", 18)
sec('3. LÉGENDE DES COULEURS')
line('Saisie', "Cellule jaune clair, texte bleu gras : à remplir par l'utilisateur.", 18, C_SAISIE)
line('Calcul', "Cellule grise : formule, verrouillée — ne pas modifier.", 18, C_CALCUL)
line('Paramètre', "Cellule bleu clair (feuille Paramètres) : réglage global, modifie tous les calculs.", 18, C_PARAM)
line('Hors service', "Ligne grisée en italique : engin exclu des calculs et de la synthèse.", 18, 'DADADA')
sec('4. STATUTS ET SEUILS (valeurs lues dans Paramètres)')
line('Écart à la cible (h)', '="Écart = compteur actuel − compteur cible, avec compteur cible = compteur du dernier ES + "&Interv_ES&" h − "&Marge&" h de marge. '
     'NÉGATIF = heures restantes avant la cible ; POSITIF = cible dépassée."', 32)
line('Normal', '="Écart < −"&Seuil_Proche&" h : plus de "&Seuil_Proche&" h avant la cible. Rien à faire."', 18, 'BFE6BF')
line('Proche', '="Écart entre −"&Seuil_Proche&" h et 0 h : la cible approche, planifier l\'ES."', 18, 'FFD9A8')
line('Dépassé', '="Écart > 0 h : la cible est dépassée, ES à réaliser en priorité (la marge de "&Marge&" h avant l\'intervalle de "&Interv_ES&" h est entamée)."', 18, 'F4A6A6')
line('À confirmer', '="Saisie incomplète ou incohérente : compteur actuel vide, compteur < dernier ES, type ≠ cycle, N° de cycle hors 1–"&Longueur_cycle&". Corriger la saisie (voir Contrôles)."', 18, 'F0F1F3')
line('Hors service', "État « Hors service » en colonne K : engin ignoré dans les calculs.", 18, 'D9D9D9')
sec('5. RÈGLES À RESPECTER')
line('⚠ Ne pas trier', "Ne jamais trier, insérer ni supprimer de lignes dans SUIVI_ES, HISTORIQUE_ES, SUIVI_GRAISSAGE_SOUFFLAGE et HISTORIQUE_GRAISSAGE_SOUFFLAGE : "
     "les historiques et le graissage sont reliés à SUIVI_ES par POSITION DE LIGNE. Un tri attribuerait l'historique d'un engin à un autre.", 32, 'FDECEC')
line('Masquer / afficher', "Utiliser les filtres (flèches des en-têtes), par exemple sur « État de l'engin » pour masquer les engins hors service.", 18)
line('Nouvel engin', "Saisir son nom dans la première ligne libre en bas de SUIVI_ES (lignes 58 à 77) : son bloc d'historique et sa ligne de graissage existent déjà à la même position.", 18)
line('Alerte « Ordre des lignes »', "Signale un tri, une insertion, une suppression ou un renommage : annuler (Ctrl+Z) ou revenir à la dernière sauvegarde. "
     "Un renommage volontaire se reporte aussi dans Contrôles, colonne R (référence d'ordre).", 32)
line('Protection', "Feuilles protégées SANS mot de passe contre les modifications accidentelles. Ne retirer la protection (Révision › Ôter la protection) "
     "que pour une maintenance, puis la remettre. Toute modification de structure ou de paramètre se note dans « Journal des modifications ».", 32)
line('Navigation', "HISTORIQUE_ES et HISTORIQUE_GRAISSAGE_SOUFFLAGE : choisir l'engin en I1 puis cliquer « Aller › ».", 18)
r = 3
for item in rows:
    if item[0] == 'sec':
        M.merge_cells(f'B{r}:C{r}')
        put(M, f'B{r}', item[1], FONT_HDR10, fill(C_VERT_F), LEFT); M[f'C{r}'].fill = fill(C_VERT_F)
        M.row_dimensions[r].height = 20
    else:
        _, a, b, h, fl = item
        put(M, f'B{r}', a, Font(name=ARIAL, size=10, bold=True), fill(fl) if fl else None, LEFT)
        put(M, f'C{r}', b, Font(name=ARIAL, size=10), None, LEFT)
        autoheight(M, {'B': 24, 'C': 110}, [r], base=18, line=13)
    r += 1
M.print_area = f'B1:C{r - 1}'
M.page_setup.orientation = 'landscape'; M.page_setup.paperSize = 9; M.page_setup.fitToWidth = 1; M.page_setup.fitToHeight = 0
M.sheet_properties.pageSetUpPr.fitToPage = True
M.oddFooter.center.text = '&A — page &P / &N'
protect(M)

# ============================================================ JOURNAL DES MODIFICATIONS
J = wb.create_sheet('Journal des modifications')
J.sheet_properties.tabColor = '7F7F7F'
J.sheet_view.showGridLines = False
for col, w in {'A': 7, 'B': 33, 'C': 80, 'D': 46, 'E': 16}.items(): J.column_dimensions[col].width = w
J.merge_cells('A1:E1')
put(J, 'A1', 'JOURNAL DES MODIFICATIONS', Font(name=ARIAL, size=14, bold=True, color='FFFFFF'), fill(C_VERT), CENTER, border=False)
J.row_dimensions[1].height = 30
J.merge_cells('A2:E2')
put(J, 'A2', "Version 2.0 du " + TODAY_TXT + " (audit et mise à niveau de SUIVI_ES_V1.1). « ID audit » renvoie au tableau des constats du rapport d'audit. "
    "Lignes jaunes en bas : à compléter pour toute évolution future (structure, paramètres, règles).", Font(name=ARIAL, size=9, italic=True, color='444444'), None, LEFT, border=False)
J.row_dimensions[2].height = 28
for i, t in enumerate(['ID', 'Feuille', 'Modification', 'Raison', 'ID audit'], start=1):
    put(J, f'{L(i)}4', t, FONT_HDR, fill(C_VERT_C))
r = 5
for m in MODIFICATIONS:
    for i, v in enumerate(m, start=1):
        put(J, f'{L(i)}{r}', v, Font(name=ARIAL, size=9, bold=(i == 1)), F_CALCUL if False else None, LEFT if i in (2, 3, 4) else CENTER)
    autoheight(J, {'B': 33, 'C': 80, 'D': 46}, [r], base=18, line=11.5, size_factor=0.9)
    r += 1
for k in range(15):
    for i in range(1, 6):
        put(J, f'{L(i)}{r}', None, Font(name=ARIAL, size=9, color='1F3A8A'), F_SAISIE, LEFT, lock=False)
    r += 1
J.freeze_panes = 'A5'
J.auto_filter.ref = f'A4:E{r - 1}'
J.print_title_rows = '4:4'
J.page_setup.orientation = 'landscape'; J.page_setup.paperSize = 9; J.page_setup.fitToWidth = 1; J.page_setup.fitToHeight = 0
J.sheet_properties.pageSetUpPr.fitToPage = True
J.oddFooter.center.text = '&A — page &P / &N'
protect(J)

# ---------------------------------------------------------------- ordre des feuilles, feuille active
order = ["Mode d'emploi", 'Tableau de bord', 'SUIVI_ES', 'SUIVI_GRAISSAGE_SOUFFLAGE', 'HISTORIQUE_ES', 'HISTORIQUE_GRAISSAGE_SOUFFLAGE',
         'Contrôles', 'Paramètres', 'Journal des modifications']
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets: ws.sheet_view.tabSelected = False
wb.active = 1
wb['Tableau de bord'].sheet_view.tabSelected = True
wb.save(DST)
print('OK ->', DST)
