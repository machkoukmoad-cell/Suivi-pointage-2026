"""Rendu visuel de feuilles (PDF via LibreOffice UNO puis PNG via pdftoppm) pour inspection.
Usage : python3 50_rendu.py classeur.xlsx dossier_sortie "Tableau de bord" "SUIVI_ES" ..."""
import sys, os, subprocess
sys.path.insert(0, '/home/user/Suivi-pointage-2026/scripts')
from lo_uno import LO, unprotect_all, _pv
import uno
src, out = sys.argv[1], sys.argv[2]
targets = sys.argv[3:]
os.makedirs(out, exist_ok=True)
with LO(port=2101) as lo:
    for t in targets:
        doc = lo.open(src)
        doc.calculateAll()
        sheets = doc.Sheets
        names = [sheets.getByIndex(i).Name for i in range(sheets.Count)]
        for n in names:
            if n != t: sheets.removeByName(n) if False else None
        # n'exporter que la feuille cible : masquer les autres via plage d'impression vide n'est pas fiable -> export par sélection
        sh = sheets.getByName(t)
        ctrl = doc.CurrentController; ctrl.setActiveSheet(sh)
        fd = uno.Any('[]com.sun.star.beans.PropertyValue', (_pv('Selection', sh),))
        safe = t.replace(' ', '_').replace("'", '')
        pdf = os.path.abspath(f'{out}/{safe}.pdf')
        doc.storeToURL(uno.systemPathToFileUrl(pdf), (_pv('FilterName', 'calc_pdf_Export'), _pv('FilterData', fd)))
        doc.close(True)
        subprocess.run(['pdftoppm', '-png', '-r', '80', '-f', '1', '-l', '3', pdf, f'{out}/{safe}'], check=True)
        print(t, '->', pdf, subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout.split('Pages:')[1].split('\n')[0].strip(), 'pages')
