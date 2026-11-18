# PyInstaller: pyinstaller ElmiMonitorinq.spec
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [("dashboard_template.html", "."), ("config.json", "."), ("muellimler.xlsx", "."), ("logo.ico", "."), ("logo.png", ".")]
binaries, hiddenimports = [], ["scopus_monitor", "openpyxl"]
for pkg in ("webview",):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h

a = Analysis(["desktop_app.py"], pathex=["."], binaries=binaries, datas=datas,
             hiddenimports=hiddenimports, excludes=["tkinter", "matplotlib", "numpy", "PIL"], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="ElmiMonitorinq", icon="logo.ico",
          console=False, upx=False)
coll = COLLECT(exe, a.binaries, a.datas, name="ElmiMonitorinq", upx=False)
