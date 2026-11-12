"""Réconciliation baseline V1.1 vs V2 : résultats par engin (colonnes A:Q de SUIVI_ES) et KPI."""
import pandas as pd, sys
a = pd.read_csv(sys.argv[1], sep=';', dtype=str).fillna('')
b = pd.read_csv(sys.argv[2], sep=';', dtype=str).fillna('')
print('colonnes V1.1:', list(a.columns)); print('colonnes V2  :', list(b.columns))
a.columns = b.columns = [f'c{i}' for i in range(len(a.columns))]
diff = (a != b)
n = int(diff.values.sum())
print('engins V1.1/V2:', len(a), len(b), '— cellules différentes:', n)
for i, j in zip(*diff.values.nonzero()):
    print('  ligne', a.iloc[i, 0], 'col', j, repr(a.iloc[i, j]), '->', repr(b.iloc[i, j]))
