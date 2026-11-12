"""Diff fin openpyxl V2 vs aller-retour LibreOffice : validations (normalisées), MFC (règle + style dxf), styles de cellules,
formats, largeurs, commentaires, graphique, mise en page. Sert à décider si l'aller-retour (valeurs en cache) est conservable."""
import sys, openpyxl, zipfile, re
A, B = sys.argv[1], sys.argv[2]
wa, wb = openpyxl.load_workbook(A), openpyxl.load_workbook(B)
lost = []
def nf(x): return None if x is None else str(x).lstrip('=').replace("'", '')
for ws in wa.worksheets:
    v = wb[ws.title]
    def dvk(d): return (str(d.sqref), d.type, nf(d.formula1), nf(d.formula2), d.errorStyle or 'stop', bool(d.showInputMessage), bool(d.showErrorMessage),
                        d.promptTitle or None, d.prompt or None, d.errorTitle or None, d.error or None, bool(d.allow_blank), d.operator or None)
    da = {dvk(d) for d in ws.data_validations.dataValidation}; db = {dvk(d) for d in v.data_validations.dataValidation}
    for x in da - db:
        cand = [y for y in db if y[0] == x[0]]
        diffs = [(i, x[i], cand[0][i]) for i in range(len(x)) if cand and x[i] != cand[0][i]] if cand else 'ABSENTE'
        lost.append((ws.title, 'validation', x[0][:40], diffs))
    def cfk(ws_):
        out = {}
        for cf, rules in ws_.conditional_formatting._cf_rules.items():
            for r in rules:
                f = r.dxf.font if r.dxf else None; fl = r.dxf.fill if r.dxf else None
                key = (str(cf.sqref), re.sub(r"'|,TRUE\)|,1\)", '', r.formula[0]))
                out[key] = (r.stopIfTrue, f and f.b, f and f.i, f and f.u, f and f.color and f.color.rgb, fl and fl.fgColor and fl.fgColor.rgb, r.priority)
        return out
    ca, cb = cfk(ws), cfk(v)
    for k in ca:
        if k not in cb: lost.append((ws.title, 'mfc absente', k[0], k[1][:60]))
        elif ca[k][:6] != cb[k][:6]: lost.append((ws.title, 'mfc style', k[0], k[1][:50], ca[k], cb[k]))
    # ordre de priorité relatif
    oa = [k for k, _ in sorted(ca.items(), key=lambda kv: kv[1][6])]; ob = [k for k, _ in sorted(cb.items(), key=lambda kv: kv[1][6])]
    if [k for k in oa if k in cb] != [k for k in ob if k in ca]: lost.append((ws.title, 'mfc ordre de priorité modifié'))
    # styles de cellules (échantillon : toutes les cellules non vides + cellules de saisie)
    n_style = 0; ex = []
    for row in ws.iter_rows():
        for c in row:
            d = v[c.coordinate]
            a_ = (c.number_format, c.font.name, c.font.b, c.font.i, c.font.sz, c.font.color.rgb if c.font.color else None,
                  c.fill.fgColor.rgb if c.fill.fill_type else None, c.protection.locked, c.alignment.horizontal, bool(c.alignment.wrap_text))
            b_ = (d.number_format, d.font.name, d.font.b, d.font.i, d.font.sz, d.font.color.rgb if d.font.color else None,
                  d.fill.fgColor.rgb if d.fill.fill_type else None, d.protection.locked, d.alignment.horizontal, bool(d.alignment.wrap_text))
            if a_ != b_:
                n_style += 1
                if len(ex) < 3: ex.append((c.coordinate, [(i, a_[i], b_[i]) for i in range(len(a_)) if a_[i] != b_[i]]))
    if n_style: lost.append((ws.title, 'styles de cellules différents', n_style, ex))
    wa_ = {k: (d.width, d.hidden, d.outlineLevel) for k, d in ws.column_dimensions.items()}
    wb_ = {k: (v.column_dimensions[k].width, v.column_dimensions[k].hidden, v.column_dimensions[k].outlineLevel) for k in wa_}
    dd = [(k, wa_[k], wb_[k]) for k in wa_ if (round(wa_[k][0] or 0, 1), wa_[k][1], wa_[k][2]) != (round(wb_[k][0] or 0, 1), wb_[k][1], wb_[k][2])]
    if dd: lost.append((ws.title, 'colonnes (largeur/masquage/groupe)', len(dd), dd[:3]))
    pa = (ws.page_setup.orientation, ws.page_setup.paperSize, ws.page_setup.fitToWidth, ws.page_setup.fitToHeight, ws.print_title_rows, ws.print_area, ws.sheet_properties.tabColor.rgb if ws.sheet_properties.tabColor else None)
    pb = (v.page_setup.orientation, v.page_setup.paperSize, v.page_setup.fitToWidth, v.page_setup.fitToHeight, v.print_title_rows, v.print_area, v.sheet_properties.tabColor.rgb if v.sheet_properties.tabColor else None)
    if pa != pb: lost.append((ws.title, 'mise en page/onglet', pa, pb))
    ka = {c.coordinate: (c.comment.text, c.comment.author) for row in ws.iter_rows() for c in row if c.comment}
    kb = {c.coordinate: (c.comment.text, c.comment.author) for row in v.iter_rows() for c in row if c.comment}
    if ka != kb: lost.append((ws.title, 'commentaires', [(k, ka.get(k, ('',''))[1], kb.get(k, ('',''))[1]) for k in set(ka) | set(kb) if ka.get(k) != kb.get(k)]))
# graphique
def chartxml(p):
    with zipfile.ZipFile(p) as z:
        n = [x for x in z.namelist() if x.startswith('xl/charts/chart')]
        return z.read(n[0]).decode() if n else ''
xa, xb = chartxml(A), chartxml(B)
for tag in ('dPt', 'dLbls', 'showVal', 'noFill', '0CA30C', 'D03B3B', 'FAB219'):
    if (tag in xa) != (tag in xb) or xa.count(tag) != xb.count(tag): lost.append(('graphique', tag, xa.count(tag), xb.count(tag)))
for x in lost: print(x)
print('TOTAL écarts :', len(lost))
