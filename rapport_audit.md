# Rapport d'audit — SUIVI_ES_V1.1.xlsx → SUIVI_ES_V2.xlsx

*Suivi des entretiens systématiques (ES) — 50 engins, Benguerir (OCP). Audit réalisé le 12/11/2026.*

## 1. Synthèse

| Sévérité | Constats | Corrigés | Atténués | Signalés (à vérifier) | Ouverts / non corrigés |
|---|---|---|---|---|---|
| **Critique** (résultats faux) | **0** | – | – | – | – |
| **Majeur** (risque de résultats faux) | **7** | 4 (A02*, A07, A10, A11) | 1 (A01) | 1 (A03) | 1 (A19) |
| **Mineur** (hygiène) | **18** | 11 | 1 (A06) | 3 (A15, A17, A18) | 3 (A23 ouvert ; A24, A25 non corrigés) |
| **Total** | **25** | **15** | **2** | **4** | **4** |

\* A02 a été corrigé après la décision de l'utilisateur du 12/11/2026 (Q1 : la valeur SUIVI_ES est la bonne). Voir la mise à jour en fin de §1.

- **Aucun résultat faux dans V1.1.** Le recalcul indépendant en Python des 50 engins (cible, écart, statut) et des 6 KPI donne **0 écart** avec le recalcul LibreOffice. Les risques viennent de la structure : liaison par position de ligne, paramètres en dur et masqués, intitulé trompeur, liste volatile, absence de contrôles.
- **Baseline réconciliée.** V2 donne exactement les mêmes résultats que V1.1 : 50 engins, 45 en service, Normal 37, Proche 4, Dépassé 4, À confirmer 0, Hors service 5, 0 cellule en erreur. Sur 50 engins × 17 colonnes, **0 valeur diffère**. Seul l'intitulé de la colonne I change (A10).
- **Aucune donnée utilisateur modifiée sans validation.** Nous avons comparé cellule par cellule les compteurs, les derniers ES, les états, les catégories, les noms et l'ordre des engins, les deux historiques et les saisies de graissage : 0 différence, en dehors des 3 cellules ajoutées à l'historique de HP24 sur décision de l'utilisateur (voir la mise à jour ci-dessous). L'original `SUIVI_ES_V1.1.xlsx` est intact (SHA-256 `9f58705c…a9bbb1`).
- **18 scénarios sur 18 réussis**, tous exécutés sur des copies jetables.
- **1 alerte active dans V2** : Arroseur HP24, « Date d'historique manquante ou invalide ». La date de l'ES 2 ajouté est à compléter (voir la mise à jour ci-dessous).

**Mise à jour du 12/11/2026 — décision Q1.** L'utilisateur a confirmé que la valeur SUIVI_ES d'Arroseur HP24 est la bonne (6 236 h, ES 250 h, cycle 1). L'historique était donc incomplet.
- L'ES 2 manquant a été ajouté dans HISTORIQUE_ES : D14 = 6 236, D16 = ES 250 h, D17 = 1, recopiés de SUIVI_ES!C10:E10.
- La date D15 reste vide, car elle est inconnue : aucune date n'a été inventée.
- La feuille Contrôles signale désormais aussi une entrée d'historique sans date (J25). L'alerte « SUIVI ≠ HISTORIQUE » a disparu, remplacée par « Date d'historique manquante ou invalide » jusqu'à la saisie de D15 (test T15 : 0 alerte une fois la date saisie).
- Ces 3 cellules sont les **seules** données modifiées par rapport à V1.1. Elles sont tracées dans le Journal (J24) et dans `scripts/registre.py` (`CORRECTIONS_DONNEES`), et le diff d'intégrité les vérifie une à une.
- **Valeurs en cache : non livrées.** L'aller-retour par LibreOffice faisait perdre des éléments (§7). Nous livrons donc la version openpyxl, qui se recalcule à l'ouverture.

## 2. Baseline (étape 0)

Comme le fichier ne contient aucune valeur en cache, nous l'avons recalculé avec LibreOffice 24.2 sur une copie (`travail/V1.1_recalc_LO.xlsx`). Les résultats par engin sont dans `travail/baseline_V1.1.csv`.

| Indicateur | Attendu | Obtenu V1.1 | Obtenu V2 |
|---|---|---|---|
| Engins | 50 | 50 | 50 |
| En service (A3) | 45 | 45 | 45 |
| Normal / Proche / Dépassé | 37 / 4 / 4 | 37 / 4 / 4 | 37 / 4 / 4 |
| À confirmer / Hors service | 0 / 5 | 0 / 5 | 0 / 5 |
| Cellules en erreur | 0 | 0 (1 059 formules) | 0 (2 793 formules) |

Dépassé : HYSTER-1 (+58 h), TERX3 (+54 h), KOM1 (+43 h), Pelle hydraulique CX470B (+26 h). Proche : CH994F2 (0 h), Niv 8 CAT (−1 h), Niv 7 CAT (−5 h), Pelleteuse W2 (−18 h).

## 3. Verdict sur les hypothèses

| Hyp. | Verdict | Preuve principale |
|---|---|---|
| **H1** Couplage par position | **Confirmée** (Majeur, A01) | Simulation LibreOffice UNO (`scripts/03_simulation_H1.py`, `travail/preuve_H1.json`). **Insertion** avant la ligne 20 : le nouvel engin n'a ni bloc d'historique ni ligne de graissage, et « Aller › » vers KOM1 mène à la ligne 58 (bloc de KOM2). **Suppression** de la ligne 9 : #REF! en HISTORIQUE_ES!A10 et en SUIVI_GRAISSAGE!A7/B7/F7, et les 3 ES de HP23 deviennent orphelins. **Tri** par compteur : l'historique de Port-Chars HP 21 s'affiche sous UR132, et le compteur G48 = 3 177 h de Pelleteuse W1 passe sous Magni. Sous protection, l'insertion et la suppression sont bloquées, mais l'option `sort="0"` laissait possible un tri partiel (par exemple B:E seules). La protection n'a pas de mot de passe. |
| **H2** SUIVI C:E = dernier HISTORIQUE | **Confirmée** (A02, A03) | Arroseur HP24 : SUIVI 6 236 h / ES 250 h / cycle 1, mais HISTORIQUE (ES 1) 6 007 h / 09/08/2025 / ES 2000 h / cycle 8. Port-Chars HP 21 et HP23 concordent. 47/50 engins n'ont aucun historique ES. HISTORIQUE_GRAISSAGE est vide (0 cellule sur 12 600) et SUIVI_GRAISSAGE C:E aussi (seule G48 est saisie). |
| **H3** Éléments masqués / verrouillés | **Confirmée** pour le masquage (A04, A05, A06) ; **rejetée** pour le comptage | Colonnes L:AH masquées : L est une saisie déverrouillée sans couleur de saisie, AE:AF contient les paramètres. Lignes 26, 30, 37, 38 et 50 masquées : ce sont les 5 engins hors service. A3:F3 les compte de façon **cohérente** (A3 45 + F3 5 = 50, COUNTIF indépendant du masquage). Protection sans mot de passe avec `formatColumns="0"` et `formatRows="0"` : n'importe qui peut afficher les éléments masqués. |
| **H4** Valeurs en dur | **Confirmée** (A07, A08, A09) | −20 dans J8:J77 (70 formules) et dans 2 règles MFC de I8:I77. 8 dans M8:M77, dans N8:N77, dans les validations N° cycle et dans 6 règles MFC d'HISTORIQUE_ES. Listes de types tapées dans les validations de SUIVI_ES!D et des lignes « Type ES » d'HISTORIQUE_ES. « Vidange exceptionnelle » n'existe que dans la liste de l'historique. |
| **H5** Signe de « Heures restantes » | **Confirmée** (A10) | I8 = −172 pour Port-Chars HP 21, ce qui signifie 172 h *restantes*. L'intitulé suggère l'inverse. Seul le commentaire I7 l'expliquait. |
| **H6** Liste_Engins volatile et fragile | **Confirmée** (A11) | `DECALER(...;NBVAL(A8:A77))`. Test : avec A30 vidée et un engin ajouté en A60, la liste compte 50 lignes, dont 1 vide, et le nouvel engin est absent (`travail/preuve_H6_H7.json`). |
| **H7** Cas limites | **Partiellement confirmée** (A12–A18) | Compteur < dernier ES : statut « À confirmer » correct, mais M affiche « OK ». N° cycle 9 ou 2,5 : « Valeur invalide », correct. Type ≠ cycle : correct. Compteur vide ou en texte : « À confirmer », mais M affiche « OK ». État vide : l'engin est traité comme en service sans aucun signal. Doublon exact : MFC correcte, mais « KOM2 » suivi d'un espace n'est pas détecté et la navigation va vers la 1re occurrence. Noms : double espace (Pay 600  W1/W2/W3) et retour à la ligne (A54). Dates d'historique : toutes valides (7/7) ; seule une MFC signalait une date en texte. |

**Contrôles généraux, tous conformes** :
- formules cohérentes sur toute la hauteur pour les 10 colonnes calculées (F:J, M:Q) ;
- plages complètes (8:77, 70 blocs × 4 lignes, 70 blocs × 3 lignes) ;
- 0 nombre stocké en texte ;
- un seul format de date (jj/mm/aaaa) ;
- 0 cellule en erreur ;
- recalcul indépendant des KPI : 0 écart (`travail/preuve_audit_general.json`).

## 4. Constats

| ID | Sév. | Feuille | Cellule / plage | Constat | Preuve | Correction | Statut |
|---|---|---|---|---|---|---|---|
| A01 | Majeur | SUIVI_ES ↔ 3 feuilles liées | HISTORIQUE_ES!A6:A282, SUIVI_GRAISSAGE!A/B/F6:75, HISTORIQUE_GRAISSAGE!A6:A213, L1 (HYPERLINK) | Les feuilles sont liées par position de ligne. Un tri, une insertion ou une suppression attribue l'historique et le graissage au mauvais engin, ou produit #REF!. | Simulation H1 (§3) | Aucune refonte, puisqu'il n'y a pas de résultat faux en usage normal. Mise en place : avertissement visible (SUIVI_ES ligne 5, aides A2/A3), règle dans le Mode d'emploi, contrôle « Ordre / liens des lignes » dans Contrôles (référence d'ordre figée + INDEX positionnel + test des liens), tri interdit par la protection. Testé par T12 (tri : 50 positions signalées), T13 (insertion : 38) et T14 (suppression : 49, Contrôles sans #REF!). | **Atténué** |
| A02 | Majeur | SUIVI_ES / HISTORIQUE_ES | SUIVI_ES!C10:E10 vs HISTORIQUE_ES!C14:C17 | Conflit sur Arroseur HP24 : SUIVI 6 236 h / ES 250 h / 1 ; HISTORIQUE 6 007 h / 09/08/2025 / ES 2000 h / 8. | Valeurs lues dans les deux feuilles | Signalé dans Contrôles!I8. **Après la décision de l'utilisateur (Q1 : SUIVI juste)**, l'ES 2 manquant a été ajouté en HISTORIQUE_ES!D14, D16 et D17 ; la date D15, inconnue, est à saisir et elle est signalée. | **Corrigé** (décision utilisateur, J24) |
| A03 | Majeur | HISTORIQUE_ES, HISTORIQUE_GRAISSAGE | tous les blocs | 47/50 engins n'ont aucun historique ES et l'historique de graissage est vide. Le contrôle SUIVI ↔ HISTORIQUE ne couvre donc que 3 engins. | 28 cellules saisies sur 11 200 ; 0 sur 12 600 | Contrôles affiche, pour chaque engin, le nombre d'entrées et le dernier ES historisé. | **Signalé** (Q2) |
| A04 | Mineur | SUIVI_ES | L7:L77, AE7:AF19, colonnes L:AH | La colonne de saisie « Catégorie » et les paramètres (cycle, Marge, Interv_ES) sont masqués et en police Calibri. | Inventaire (colonnes masquées, verrouillage) | Feuille Paramètres visible ; L affichée en saisie avec la liste Liste_Categories ; AE:AF vidées et paramètres déplacés avec leurs commentaires. | **Corrigé** |
| A05 | Mineur | SUIVI_ES | lignes 26, 30, 37, 38, 50 | Les 5 engins hors service sont masqués manuellement. Ils restent comptés de façon cohérente (A3+F3 = 50), mais deviennent invisibles. | `row_dimensions.hidden` | Lignes affichées. La MFC existante les grise ; le filtre de la colonne K sert à les masquer. | **Corrigé** |
| A06 | Mineur | 4 feuilles d'origine | sheetProtection | Protection sans mot de passe ; `sort="0"` permettait un tri partiel. | XML `sheetProtection` | Mot de passe toujours absent (décision D4) ; « Trier » désormais interdit. | **Atténué** |
| A07 | Majeur | SUIVI_ES | J8:J77 ; MFC I8:I77 (2 règles) | Seuil « Proche » −20 codé en dur à 3 endroits. Modifier l'un sans les autres désaccorde les statuts et les couleurs. | Formules et MFC | Paramètre Seuil_Proche (20 h) ; formules et MFC le référencent. T8 : seuil passé à 30 → D11T5 passe Proche. | **Corrigé** |
| A08 | Mineur | SUIVI_ES / HISTORIQUE_ES | M8:M77, N8:N77, validations N° cycle, 6 MFC d'HISTORIQUE_ES | Longueur de cycle 8 codée en dur. | Formules | Longueur_cycle = NBVAL(Cycle_Types), référencée partout. | **Corrigé** |
| A09 | Mineur | SUIVI_ES / HISTORIQUE_ES | validations D8:D77 et « Type ES » (70 lignes) | Listes tapées dans les validations ; « Vidange exceptionnelle » présente seulement dans l'historique. | XML `dataValidation` | Types_ES (4 types uniques) et Types_Historique (+ Vidange exceptionnelle) dans Paramètres, avec contrôle de cohérence (B44). | **Corrigé** (+ Q3) |
| A10 | Majeur | SUIVI_ES | I6:I7 | « Heures restantes » est négatif quand il reste des heures, d'où un risque de lecture inversée. | I8 = −172 (172 h restantes) | Intitulé « Écart à la cible (h) » ; calcul et signe inchangés ; légende ligne 4 et Mode d'emploi ; commentaire I7 actualisé. | **Corrigé** |
| A11 | Majeur | Nom Liste_Engins | — | DECALER est volatil, et la liste est tronquée dès qu'il y a un vide dans A8:A77. | Test H6 : nouvel engin absent, 1 vide inclus | Liste compacte non volatile (Contrôles!U:V, INDEX/EQUIV) et nom en `'Contrôles'!$V$6:INDEX(...)`. T5 et T5b : liste complète, 0 vide. | **Corrigé** |
| A12 | Mineur | SUIVI_ES | M8:M77 | Avec un compteur actuel vide, en texte ou inférieur au dernier ES, J affiche « À confirmer » mais M affiche « OK » : la cause n'est pas montrée. | Test H7 (UR131, Camion Volvo, TERX2) | Contrôles : « Compteur actuel manquant ou non numérique » et « Compteur < dernier ES » ; MFC orange sur M ≠ OK. T4 et T11 réussis. | **Corrigé** |
| A13 | Mineur | SUIVI_ES | K8:K77 | Un état vide est compté « en service » sans aucun signal. | Test H7 (TERX4 → Normal) | Contrôles : « État non renseigné ». | **Corrigé** |
| A14 | Mineur | SUIVI_ES / HISTORIQUE | A8:A77, L1 | Doublon avec espace final non détecté par la MFC (COUNTIF) ; la navigation va à la 1re occurrence. | Test H7 : « KOM2 » + espace → COUNTIF = 1 | Contrôles « Nom en double » (comparaison avec SUPPRESPACE). T10 réussi. | **Corrigé** |
| A15 | Mineur | SUIVI_ES | A45:A47, A54 | Noms avec double espace (« Pay 600  W1/W2/W3 ») et retour à la ligne (« MANIPULATEUR(gieen 200)↵702612 (HELI) »). | Valeurs brutes | Aucune modification (données utilisateur). | **Signalé – à vérifier** (Q8) |
| A16 | Mineur | HISTORIQUE_ES | lignes « Date » | Une date saisie en texte ou future n'était signalée que par une MFC, sans alerte consolidée. Aucun cas actuellement. | 7/7 dates valides | Contrôles « Date d'historique manquante ou invalide » : date en texte, date future, ou compteur saisi sans date (J25). T9 réussi. | **Corrigé** |
| A17 | Mineur | HISTORIQUE_ES | C10:D11 (HP23) | 14 710 h le 03/09/2023 puis 15 077 h le 14/09/2023 : +367 h en 11 jours, alors que 264 h est le maximum physique. | Calcul (24 h × 11 j) | Aucune modification. | **Signalé – à vérifier** (Q6) |
| A18 | Mineur | HISTORIQUE_ES | D6:E8 (Port-Chars HP 21) | ES 500 h réalisé à 6 281 h, soit 71 h après l'ES 250 h (6 210 h), pour un intervalle de 250 h. | Valeurs | Aucune modification. | **Signalé – à vérifier** (Q7) |
| A19 | Majeur | Classeur | tout | Aucune valeur en cache (fichier openpyxl, `fullCalcOnLoad`) : les lecteurs sans moteur de calcul affichent des cellules calculées vides. | `data_only=True` : None partout | Aller-retour LibreOffice tenté puis rejeté pour pertes (§7). Le recalcul à l'ouverture d'Excel et de LibreOffice reste actif. | **Non corrigé** (risque résiduel R2) |
| A20 | Mineur | SUIVI_ES | commentaire I7, invite E8:E77 | Textes figés (« 250 h − 50 h (soit +200 h) », « 1 à 8 ») qui deviennent faux si les paramètres changent. | Texte des commentaires | Commentaire et invites renvoient à Paramètres ; légende et Mode d'emploi lisent les valeurs par formule (T7 : légende passée à « − 30 h »). | **Corrigé** |
| A21 | Mineur | SUIVI_ES | filtre A7:K77 | Le filtre automatique n'inclut ni Catégorie ni Contrôle saisie. | `auto_filter.ref` | Filtre étendu à A7:M77. | **Corrigé** |
| A22 | Mineur | SUIVI_ES | L:Q, AE:AF (422 cellules) | Police Calibri au lieu d'Arial. | Inventaire des polices | Passage en Arial. | **Corrigé** |
| A23 | Mineur | SUIVI_GRAISSAGE | — | Aucune règle d'échéance pour le graissage et le soufflage (choix explicite de V1.1). | A3 : « Aucune échéance n'est calculée » | Paramètre vide ajouté (Paramètres!B41), non utilisé. La logique graissage est inchangée. | **Ouvert** (Q5) |
| A24 | Mineur | HISTORIQUE_ES | MFC C6:AP285 (couleurs par type), légende C3:K3 | Les couleurs sont liées aux libellés littéraux (« ES 250 h »…) : renommer un type dans Paramètres casse la couleur. | Formules MFC | Non modifié : la correspondance couleur/type est volontairement explicite. | **Non corrigé** (R4) |
| A25 | Mineur | HISTORIQUE_ES / HIST. GRAISSAGE | mise en page | Impression sans ajustement en largeur : 78 pages et texte d'aide coupé à l'impression. | Rendu PDF | Mise en page d'origine conservée (consigne « garder la disposition »). | **Non corrigé** |

## 5. Décisions prises (journalisées)

| N° | Décision | Justification |
|---|---|---|
| D1 | Original en lecture seule (`chmod 444`) ; travail dans `travail/` ; copie intacte `travail/V1.1_intact.xlsx` ; scripts dans `scripts/`. | Consigne. |
| D2 | Aucun constat classé Critique. | 0 écart au recalcul indépendant ; aucun résultat faux en usage normal (protection active). |
| D3 | Conflit A02 d'abord seulement signalé, puis résolu sur décision de l'utilisateur (Q1, 12/11/2026) : ES 2 ajouté dans l'historique, date laissée vide et signalée. | Un conflit se signale ; seul l'utilisateur le tranche. Aucune date n'est inventée. |
| D4 | H1 : disposition conservée, protection conservée **sans mot de passe** ; seule l'option « Trier » passe à « interdit ». | Atténuation minimale ; un tri partiel restait possible sous protection et ne sert jamais ici (le filtre reste autorisé). |
| D5 | Seuil_Proche = 20 h (valeur de V1.1) ; Longueur_cycle = NBVAL(Cycle_Types). | Aucune règle inventée : seules des constantes sont externalisées. |
| D6 | Types_ES saisis dans Paramètres (4 types) ; Types_Historique = Types_ES + « Vidange exceptionnelle » ; contrôle de cohérence Cycle_Types ⊂ Types_ES. | Liste unique demandée. |
| D7 | Liste_Engins = liste compacte calculée dans Contrôles (colonnes techniques S:W groupées). | Non volatile, sans vide, compatible Excel 2016 sans formule matricielle. |
| D8 | Colonnes L et M et les 5 lignes affichées ; N:Q groupées et repliées ; R:AH affichées (vides) ; AE:AF vidées après déplacement. | Plus aucun contenu masqué ; une seule source pour les paramètres. |
| D9 | Calcul et signe de l'écart inchangés ; seul l'intitulé change. | Consigne ; aucun écart de baseline. |
| D10 | Listes « État » (En service / Hors service) et « Contexte » graissage laissées en dur. | Les formules testent le texte littéral « Hors service » ; logique graissage inchangée. |
| D11 | La comparaison SUIVI ↔ HISTORIQUE ignore les entrées « Vidange exceptionnelle ». | Ce type ne peut pas être saisi dans SUIVI_ES, donc pas de faux écart (T6b). Voir Q3. |
| D12 | Noms d'engins A8:A57 toujours verrouillés (comme en V1.1) ; A58:A77 déverrouillés pour les nouveaux engins. | Habitudes et protection des noms existants. |
| D13 | Couleurs de SUIVI_GRAISSAGE conservées (gris « Aucune donnée » sur les saisies), par exception au code couleur. | « Ne pas changer la logique graissage » ; cohérence avec sa légende. |
| D14 | Catégories : les 5 valeurs existantes + 3 emplacements libres ; validation en mode « Avertissement ». | Ne pas bloquer une catégorie nouvelle. |
| D15 | Référence d'ordre figée le 12/11/2026 dans Contrôles!R (noms A8:A57). | Seul moyen sans macro de détecter un tri. |
| D16 | Total_Alertes = somme des alertes par position + anomalies de Paramètres ; affiché en SUIVI_ES!G3 et sur le Tableau de bord. | Consigne. |
| D17 | Valeurs en cache : aller-retour LibreOffice rejeté, version openpyxl livrée. | Règle « garder seulement si rien n'est perdu » (§7). |
| D18 | Feuille active à l'ouverture : Tableau de bord. | Usage mensuel. |

## 6. Tests de scénarios (étape 4, copies jetables dans `travail/scenarios/`)

| ID | Scénario | Résultat | Observé |
|---|---|---|---|
| T1 | Compteur au-delà de la cible → Dépassé | RÉUSSI | J8 = Dépassé, écart +15 h ; KPI 45/36/4/5/0/5 ; rang 5 au Tableau de bord |
| T2 | Écart −13 h → Proche ; bornes −20 h → Proche et −21 h → Normal | RÉUSSI | J9, J10, J11 conformes |
| T3 | Type ≠ N° cycle | RÉUSSI | M11 = « Type ≠ cycle », J11 = À confirmer, Contrôles!H9, total d'alertes = 2 |
| T4 | Compteur vide (en service) → À confirmer | RÉUSSI | J12 = À confirmer ; Contrôles!F10 ; tuile = 1 |
| T5 | Nouvel engin en ligne libre (58) | RÉUSSI | Liste_Engins 51 (dernier = nouvel engin) ; en service 46 ; Camions/Normal 14 ; HISTORIQUE_ES!A206 et SUIVI_GRAISSAGE!A56 affichent le nom |
| T5b | Trou en A30 + engin en A60 (H6) | RÉUSSI | Liste de 50 noms, 0 vide, engin présent ; alerte « Ordre » sur D11T4 |
| T6 | Entrée d'historique discordante (HP23) | RÉUSSI | Contrôles!I7 = « SUIVI ≠ HISTORIQUE » ; SUIVI_ES!C9 inchangé |
| T6b | Vidange exceptionnelle après le dernier ES | RÉUSSI | Pas de faux écart |
| T7 | Marge 50 → 30 | RÉUSSI | F8 = 6 501 (+220) ; 0 ligne non conforme au recalcul Python ; légende actualisée |
| T8 | Seuil_Proche 20 → 30 | RÉUSSI | D11T5 (−21 h) → Proche |
| T9 | Date d'historique en texte (Port-Chars HP 21) | RÉUSSI | Contrôles!L6 signalé ; total d'alertes = 2 |
| T0 | État livré après la correction Q1 | RÉUSSI | HP24 : SUIVI = HISTORIQUE (6 236 h / ES 250 h / 1), 2 entrées ; date manquante signalée ; total d'alertes = 1 ; KPI inchangés |
| T15 | Saisie de la date de l'ES 2 de HP24 | RÉUSSI | 0 alerte (Contrôles et tuile SUIVI_ES!G3) |
| T10 | Doublon avec espace final | RÉUSSI | Les 2 positions sont signalées |
| T11 | Compteur < dernier ES | RÉUSSI | À confirmer + Contrôles!G9 |
| T12 | Tri de SUIVI_ES (UNO) | RÉUSSI | 50 positions « Ordre des lignes modifié » |
| T13 | Insertion de ligne (UNO) | RÉUSSI | 38 positions signalées, dont KOM1 |
| T14 | Suppression de ligne (UNO) | RÉUSSI | 49 positions signalées ; 0 erreur dans Contrôles (les 5 #REF! hors Contrôles sont la manifestation attendue de H1) |

## 7. Intégrité V1.1 → V2 et valeurs en cache

Les écarts ci-dessous sont tous **voulus** et journalisés (feuille « Journal des modifications », J01–J25). Côté données, la seule différence est la correction validée de HP24 (HISTORIQUE_ES!D14, D16, D17 ; J24) :
- **Validations SUIVI_ES** :
  - B8:C77 scindée en B et C, avec messages ;
  - D8:D77 pointe vers `=Types_ES` ;
  - E8:E77 accepte 1 à `Longueur_cycle` ;
  - L8:L77 ajoutée.
- **Validations HISTORIQUE_ES** : `=Types_Historique` et `Longueur_cycle`. Les validations de graissage sont inchangées.
- **MFC** :
  - SUIVI_ES : 2 règles réécrites (Seuil_Proche) et 3 ajoutées (G3, M, A = alerte) ;
  - HISTORIQUE_ES : 6 règles réécrites (Longueur_cycle) ;
  - aucune règle supprimée.
- **Fusions** : A1:K1 devient A1:M1 ; ajout de A4:M4, A5:M5, L6:M6 et N6:Q6. Les autres feuilles d'origine sont inchangées.
- **Inchangés** :
  - volets figés ;
  - filtres (sauf SUIVI_ES, étendu à M) ;
  - titres d'impression ;
  - noms de feuilles.
- **Protection** : identique, sauf `sort` interdit. Les 9 feuilles sont protégées sans mot de passe ; 0 formule déverrouillée ; 0 saisie verrouillée.
- **Logo** : identique à l'octet près (SHA-256 `d43517cd…`), même ancrage en A1.
- **Commentaires** : I7 (texte actualisé) et K7 sont conservés ; AF17 et AF18 sont déplacés tels quels en Paramètres!B16 et B17.
- **Fonctions et macros** : aucune fonction postérieure à Excel 2016 (pas de XLOOKUP, LET, tableau dynamique, TEXTJOIN…) et aucune macro.

**Aller-retour LibreOffice (valeurs en cache) : rejeté.** Le diff fin (`scripts/31_diff_allerretour.py`) relève 104 écarts. Les pertes réelles sont :
- les couleurs des barres du graphique, revenues à la couleur par défaut ;
- les couleurs d'onglets ;
- les largeurs de colonnes (C de SUIVI_ES 15 → 13, D:BJ des historiques 12,5 → 13…) ;
- l'option « Interrompre si vrai » des MFC et l'ordre de priorité ;
- les auteurs des commentaires ;
- 6 règles MFC d'HISTORIQUE_ES réécrites.

Le livrable est donc la version openpyxl : **il se recalcule à l'ouverture dans Excel ou LibreOffice, mais n'affiche aucune valeur dans un lecteur sans moteur de calcul.**

## 8. Questions ouvertes

1. **Q1 — Arroseur HP24 (A02) : répondue le 12/11/2026.** La valeur SUIVI est la bonne, et l'ES 2 a été ajouté à l'historique. **Reste à fournir : la date de cet ES 2** (HISTORIQUE_ES!D15), réalisé à 6 236 h, donc après le 09/08/2025.
2. **Q2 — Historiques (A03)** : faut-il reprendre l'historique papier des 47 engins sans historique ES et des graissages ? Sans cette reprise, le contrôle SUIVI ↔ HISTORIQUE ne protège que 3 engins.
3. **Q3 — « Vidange exceptionnelle »** : fait-elle avancer ou réinitialiser le cycle ES, ou décale-t-elle la cible ? V2 l'ignore pour le cycle et pour la comparaison, comme V1.1.
4. **Q4 — Seuils** : confirmer Seuil_Proche = 20 h (repris de V1.1, non documenté) et Marge = 50 h (« cahier des charges »).
5. **Q5 — Graissage / soufflage (A23)** : quelle fréquence ou quelle échéance appliquer ? Paramètres!B41 a été laissé vide volontairement.
6. **Q6 — HP23 (A17)** : +367 h en 11 jours entre le 03/09 et le 14/09/2023. Est-ce la date ou le compteur qui est faux ?
7. **Q7 — Port-Chars HP 21 (A18)** : l'ES 500 h réalisé 71 h après l'ES 250 h était-il un ES anticipé volontaire ?
8. **Q8 — Noms (A15)** : faut-il normaliser « Pay 600  W1/W2/W3 » (double espace) et le nom en A54 (retour à la ligne) ? Si oui, il faudra le faire à la main, sous protection levée, puis mettre à jour Contrôles!R.

## 9. Risques résiduels

- **R1 — Liaison par position (H1).** Elle est toujours là. La protection sans mot de passe se retire en deux clics, et le contrôle « Ordre / liens » **détecte** un désalignement sans le **réparer** (il faut annuler ou restaurer une sauvegarde). Une refonte par identifiant d'engin supprimerait ce risque.
- **R2 — Pas de valeurs en cache (A19).** Les aperçus de messagerie ou de téléphone et le Mode protégé d'Excel affichent des cellules calculées vides jusqu'à « Activer la modification ». Excel peut aussi proposer d'enregistrer à la fermeture, à cause du recalcul.
- **R3 — Référence d'ordre figée.** Un renommage volontaire déclenche une alerte tant que Contrôles!R n'est pas mis à jour (cellule protégée).
- **R4 — Couleurs de l'historique (A24)** liées aux libellés des types.
- **R5 — Validation Excel non faite.** Le classeur a été construit avec openpyxl et vérifié uniquement avec LibreOffice 24.2 ; il n'a pas été ouvert dans Microsoft Excel ici. Les formules se limitent aux fonctions d'Excel 2007–2016 (INDEX, EQUIV, SOMMEPROD, NB.SI.ENS, plages INDEX dans les noms). Une ouverture de contrôle dans Excel 2016 est recommandée : tuiles, liste « À traiter en priorité », listes déroulantes Liste_Engins et Liste_Categories.
- **R6 — Fonction volatile.** AUJOURDHUI() est utilisée dans Contrôles!L (dates futures), comme dans la validation et la MFC d'origine. Le recalcul suit chaque modification ; l'impact est négligeable sur 70 lignes.
- **R7 — Couverture des contrôles** : voir A03 (historiques quasi vides).

## 10. Fichiers produits et scripts exécutés

- `SUIVI_ES_V2.xlsx` : livrable, 9 feuilles (Mode d'emploi, Tableau de bord, SUIVI_ES, SUIVI_GRAISSAGE_SOUFFLAGE, HISTORIQUE_ES, HISTORIQUE_GRAISSAGE_SOUFFLAGE, Contrôles, Paramètres, Journal des modifications).
- `rapport_audit.md` : ce rapport.
- `scripts/` :

  | Script | Rôle |
  |---|---|
  | `00_inventaire.py` | Inventaire |
  | `01_dump.py` | Lecture cellule par cellule |
  | `02_extraire_resultats.py` | Résultats par engin et KPI |
  | `03_simulation_H1.py` | Simulation insertion, suppression, tri (UNO) |
  | `04_audit_controles.py` | Contrôles généraux et recalcul pandas |
  | `05_simulation_H6_H7.py` | Cas limites |
  | `10_construire_V2.py` | Construction de V2 |
  | `registre.py` | Journal |
  | `20_reconciliation.py` | Réconciliation |
  | `30_integrite.py` | Diff d'intégrité |
  | `31_diff_allerretour.py` | Diff de l'aller-retour LibreOffice |
  | `40_scenarios.py` | 18 tests |
  | `50_rendu.py` | Rendus PDF et PNG |
  | `60_verification_finale.sh` | Chaîne complète sur le livrable |
  | `lo_uno.py`, `recalc_copy.sh` | Outils |

- `travail/` : copie intacte, recalculs, preuves JSON (`preuve_H1.json`, `preuve_H6_H7.json`, `preuve_audit_general.json`), baseline CSV, scénarios, rendus (`rendu/`, `rendu2/`) et sorties de la vérification finale (`verif_finale/`).
