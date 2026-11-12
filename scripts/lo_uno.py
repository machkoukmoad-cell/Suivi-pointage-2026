"""Pilotage de LibreOffice headless via UNO (simulations insertion/suppression/tri, recalcul, export).
Utilisé uniquement sur des copies jetables dans travail/."""
import os, subprocess, time, tempfile, uno, sys
from pathlib import Path
from com.sun.star.beans import PropertyValue

SK = '/root/.claude/skills/synced/613869b7-18fd-4e4f-a3ea-f504b0dc1277_6a8a304d-5172-4887-9b7a-248c0ff3e593/xlsx/scripts'
sys.path.insert(0, SK)
from office.soffice import get_soffice_env

def _pv(name, value):
    p = PropertyValue(); p.Name = name; p.Value = value; return p

class LO:
    def __init__(self, port=2099):
        self.port = port
        self.profile = tempfile.mkdtemp(prefix='lo_uno_')
    def __enter__(self):
        self.proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--nologo', '--norestore',
            f'-env:UserInstallation={Path(self.profile).as_uri()}',
            f'--accept=socket,host=127.0.0.1,port={self.port};urp;StarOffice.ComponentContext'],
            env=get_soffice_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(120):
            try:
                self.ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={self.port};urp;StarOffice.ComponentContext'); break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError('LibreOffice UNO non joignable')
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', self.ctx)
        return self
    def open(self, path):
        return self.desktop.loadComponentFromURL(Path(path).resolve().as_uri(), '_blank', 0, (_pv('Hidden', True),))
    def save_xlsx(self, doc, path):
        doc.calculateAll()
        doc.storeToURL(Path(path).resolve().as_uri(), (_pv('FilterName', 'Calc MS Excel 2007 XML'),))
    def export_pdf(self, doc, path, sheet_index=None):
        doc.storeToURL(Path(path).resolve().as_uri(), (_pv('FilterName', 'calc_pdf_Export'),))
    def __exit__(self, *a):
        try: self.desktop.terminate()
        except Exception: pass
        try: self.proc.wait(timeout=20)
        except Exception: self.proc.kill()

def unprotect_all(doc):
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        if sh.isProtected(): sh.unprotect('')
