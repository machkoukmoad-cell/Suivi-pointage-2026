"""Inventaire structurel d'un classeur : feuilles, dimensions, noms, validations, MFC, fusions, volets, protection, images, commentaires."""
import sys, openpyxl
p = sys.argv[1]
wb = openpyxl.load_workbook(p)
print("SHEETS:", wb.sheetnames)
print("DEFINED NAMES:")
for n, dn in wb.defined_names.items():
    print("  ", n, "=", dn.attr_text, "localSheetId=", dn.localSheetId)
for ws in wb.worksheets:
    print("=" * 80)
    print("SHEET", ws.title, "dims", ws.dimensions, "max_row", ws.max_row, "max_col", ws.max_column, "state", ws.sheet_state)
    print(" freeze", ws.freeze_panes, "autofilter", ws.auto_filter.ref, "tabColor", ws.sheet_properties.tabColor)
    pr = ws.protection
    print(" protection sheet=", pr.sheet, "password=", pr.password, "fmtCols", pr.formatColumns, "fmtRows", pr.formatRows, "sort", pr.sort, "autoFilter", pr.autoFilter, "insertRows", pr.insertRows, "deleteRows", pr.deleteRows)
    print(" merged:", [str(m) for m in ws.merged_cells.ranges])
    print(" images:", len(ws._images), [(type(i.anchor).__name__, getattr(i.anchor, '_from', None) and (i.anchor._from.col, i.anchor._from.row)) for i in ws._images])
    print(" charts:", len(ws._charts))
    hc = [k for k, d in ws.column_dimensions.items() if d.hidden]
    print(" hidden cols:", hc)
    gc = [(k, d.outlineLevel, d.min, d.max) for k, d in ws.column_dimensions.items() if d.outlineLevel]
    print(" grouped cols:", gc)
    hr = [k for k, d in ws.row_dimensions.items() if d.hidden]
    print(" hidden rows:", hr)
    print(" col widths:", {k: d.width for k, d in ws.column_dimensions.items()})
    print(" DV count", len(ws.data_validations.dataValidation))
    for dv in ws.data_validations.dataValidation:
        print("   DV", dv.sqref, dv.type, dv.formula1, dv.formula2, dv.operator, "allowBlank", dv.allow_blank, "prompt", dv.promptTitle, dv.prompt, "err", dv.errorTitle, dv.error, dv.showInputMessage, dv.showErrorMessage)
    print(" CF:")
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        for r in rules:
            print("   CF", rng.sqref, r.type, r.operator, r.formula, "prio", r.priority, "stop", r.stopIfTrue, "dxf fill", r.dxf.fill.fgColor.rgb if r.dxf and r.dxf.fill else None, "font", r.dxf.font.color.rgb if r.dxf and r.dxf.font and r.dxf.font.color else None)
    com = [(c.coordinate, c.comment.text[:80]) for row in ws.iter_rows() for c in row if c.comment]
    print(" comments:", len(com))
    for c in com: print("   ", c)
    print(" print: orient", ws.page_setup.orientation, "paper", ws.page_setup.paperSize, "fitW", ws.page_setup.fitToWidth, "fitH", ws.page_setup.fitToHeight, "titles", ws.print_title_rows, ws.print_title_cols, "area", ws.print_area)
