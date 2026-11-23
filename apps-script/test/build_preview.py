#!/usr/bin/env python3
"""Builds a browser preview of the Apps Script web app (same code + mock Google backend).
Output: /home/claude/gws/preview/{index,hub,tp,dl}.html"""
import re, json, os, sys
SRC = "/home/claude/gws/src"; OUT = "/home/claude/gws/preview"; os.makedirs(OUT, exist_ok=True)
fixed_now = sys.argv[1] if len(sys.argv) > 1 else ""   # e.g. 2026-11-23T11:00 for tests

def read(n): return open(os.path.join(SRC, n), encoding="utf-8").read()

def resolve(name, vars):
    h = read(name + ".html")
    h = re.sub(r"<\?!=\s*include\('(\w+)'\)\s*\?>", lambda m: resolve_inc(m.group(1), vars), h)
    h = re.sub(r"<\?!=\s*JSON\.stringify\((\w+)\)\s*\?>", lambda m: json.dumps(vars[m.group(1)]), h)
    h = re.sub(r"<\?=\s*(\w+)\s*\?>", lambda m: str(vars.get(m.group(1), "")), h)
    return h

def resolve_inc(n, vars):
    return re.sub(r"<\?!=\s*include\('(\w+)'\)\s*\?>", lambda m: resolve_inc(m.group(1), vars), read(n + ".html"))

server = "\n".join(read(f) for f in ["Store.gs", "Logic.gs", "Api.gs", "Setup.gs", "Jobs.gs"])
mock = open("/home/claude/gws/test/mock.js", encoding="utf-8").read()

def bar(active):
    links = [("hub.html", "Tootsi staff: PO Hub (Google sign-in)"), ("3pl.html", "3PL sign-in (TP or DL)")]
    a = "".join(f'<a href="{h}" style="color:{"#fff" if h==active else "#c9cbcd"};{"font-weight:600;text-decoration:underline" if h==active else ""};margin-right:14px">{t}</a>' for h, t in links)
    return ('<div style="background:#16181a;color:#e9e9e6;padding:8px 20px;font:13px IBM Plex Sans,system-ui;display:flex;flex-wrap:wrap;gap:6px;align-items:center">'
            '<b style="color:#fff;margin-right:10px">Google Workspace version · preview</b>' + a +
            '<button onclick="localStorage.removeItem(\'pohub_gws_preview_v4\');localStorage.removeItem(\'tootsi_3pl_session\');location.reload()" style="margin-left:auto;background:none;border:1px solid #555;color:#e9e9e6;border-radius:16px;padding:4px 12px;cursor:pointer">Reset sample data</button></div>')

def inject(html, active, user):
    # The server code runs in its own hidden frame, like Apps Script runs on Google's servers.
    # This keeps server and page names apart. google.script.run forwards calls to that frame.
    boot = "window.MOCK_USER=%s;%s" % (json.dumps(user), ("window.MOCK_NOW=%s;" % json.dumps(fixed_now)) if fixed_now else "")
    srv = json.dumps("<script>" + boot + "</scr"+"ipt><script>" + server + "\n</scr"+"ipt><script>" + mock + "\n</scr"+"ipt>").replace("</", "<\\/")
    shim = """<script>(function(){var fr=document.createElement('iframe');fr.style.display='none';fr.title='server';document.documentElement.appendChild(fr);
var d=fr.contentDocument;d.open();d.write(%s);d.close();var S=fr.contentWindow;window.MOCK=function(){return S};
function runner(ok,fail){return new Proxy({},{get:function(_,p){if(p==='withSuccessHandler')return function(f){return runner(f,fail)};if(p==='withFailureHandler')return function(f){return runner(ok,f)};
return function(){var r=S.google.script.run;if(ok)r=r.withSuccessHandler(ok);if(fail)r=r.withFailureHandler(function(e){fail(new Error(e&&e.message||String(e)))});r[p].apply(null,arguments)}}})}
window.google={script:{run:runner(null,null)}};})();</script>
""" % srv
    html = html.replace("<body>", "<body>\n" + bar(active), 1)
    i = html.index("<script>")
    return html[:i] + shim + html[i:]

open(f"{OUT}/hub.html", "w").write(inject(resolve("Hub", {}), "hub.html", "juan@tootsi.com"))
open(f"{OUT}/3pl.html", "w").write(inject(resolve("Portal", {"lang": "fr", "hubUrl": "hub.html"}), "3pl.html", ""))
for old in ["tp.html", "dl.html"]:   # old preview links now open the sign-in page
    open(f"{OUT}/{old}", "w").write('<meta http-equiv="refresh" content="0;url=3pl.html">')
open(f"{OUT}/index.html", "w").write("""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>PO Hub · Google Workspace version (preview)</title>
<style>body{font:15px IBM Plex Sans,system-ui,sans-serif;background:#f6f6f4;color:#16181a;margin:0}main{max-width:760px;margin:40px auto;padding:0 20px}a.card{display:block;background:#fff;border:1px solid #e3e3df;border-left:5px solid #d71f3c;border-radius:12px;padding:16px 18px;margin:10px 0;color:inherit;text-decoration:none}a.card:hover{border-color:#16181a}a.card b{font-size:17px}.s{color:#62676c;font-size:13.5px}table{border-collapse:collapse;margin-top:8px}td{padding:4px 14px 4px 0;font-size:14px}code{background:#efefec;padding:2px 6px;border-radius:5px}</style></head>
<body><main><h1 style="margin-bottom:4px">PO Hub · Google Workspace version</h1><p class="s">Preview of the free pilot. This is the real Apps Script code running in your browser with sample data. In the pilot it runs on Tootsi's Google account: a Google Sheet is the database and a Drive folder holds the photos.</p>
<a class="card" href="hub.html"><b>Tootsi staff: PO Hub</b><div class="s">Real link: opens only for a signed-in @tootsi.com Google account. No extra password. (The preview signs you in as juan@tootsi.com.)</div></a>
<a class="card" href="3pl.html"><b>3PL partners: sign-in page</b><div class="s">Each 3PL signs in with its own username and password. After sign-in, it only sees its own POs.</div>
<table><tr><td>Trans-Plus</td><td>username <code>transplus</code></td><td>password <code>demo-tp</code></td></tr><tr><td>Deslauriers</td><td>username <code>deslauriers</code></td><td>password <code>demo-dl</code></td></tr></table></a>
<p class="s">Changes you make are saved in this browser only. Use "Reset sample data" at the top of any page to start over.</p></main></body></html>""")
print("built", OUT)
