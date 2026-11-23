#!/usr/bin/env python3
"""Builds PO_Hub.gs: the whole app in ONE Apps Script file (paste once, no HTML files to create)."""
import json, os
SRC="/home/claude/gws/src"; OUT="/home/claude/gws/dist"; os.makedirs(OUT, exist_ok=True)
gs=["Code.gs","Setup.gs","Store.gs","Logic.gs","Api.gs","Jobs.gs"]
html=["Hub","Portal","Receive","Client","Styles","Lang","Logo","Message"]
parts=["/**\n * PO Hub · Tootsi Impex · Google Workspace pilot\n * ONE-FILE INSTALL: paste this whole file into Extensions › Apps Script (replace everything in Code.gs), Save, then Run ▶ setup.\n * Source and docs: github.com/jmunevar-afk/tootsi-3pl-demo/tree/main/apps-script\n */\n"]
for f in gs: parts.append(f"\n// ==================== {f} ====================\n"+open(os.path.join(SRC,f)).read())
files={h:open(os.path.join(SRC,h+".html")).read() for h in html}
parts.append("\n// ==================== pages (HTML) ====================\nvar HTML_FILES = "+json.dumps(files, ensure_ascii=False)+";\n")
open(os.path.join(OUT,"PO_Hub.gs"),"w").write("".join(parts))
print(os.path.getsize(os.path.join(OUT,"PO_Hub.gs"))//1024,"KB")
