import asyncio, json, re
from playwright.async_api import async_playwright
P="file:///home/claude/gws/preview/"; R=[]
def ok(name,c,info=""): R.append((name,bool(c))); print(("PASS " if c else "FAIL ")+name+(" · "+str(info) if info and not c else ""))
INTERNAL=["supplier","alloc","so","prodQty","salesQty","instrSap","extra","qc","lbs","cs","issue","notes","docs","files","url","diffNote","claim","Purchasing"]
async def main():
  async with async_playwright() as p:
    b=await p.chromium.launch(); ctx=await b.new_context(viewport={"width":1300,"height":900}); errs=[]
    pg=await ctx.new_page(); pg.on("pageerror",lambda e:errs.append(str(e)))
    await pg.route(re.compile("fonts\\.(googleapis|gstatic)"),lambda r:r.abort())
    # ---------- privacy: what the TP portal receives ----------
    await pg.goto(P+"tp.html"); await pg.wait_for_timeout(1200)
    res=await pg.evaluate("MOCK().MOCK_LAST")
    rows=res["res"]["rows"]; s=json.dumps(res["res"])
    ok("portalLoad returns only TP POs", rows and all(r["po"] in ("90101","90102","90106","90107") for r in rows), [r["po"] for r in rows])
    keys=set(k for r in rows for k in r)
    ok("portalLoad has no internal fields", not (keys & set(INTERNAL)), keys & set(INTERNAL))
    ok("no supplier name / sales order / DL data leaks", not any(x in s for x in ["Vietnam","SO 118","90203","Deslauriers","Cashew exporter","FOR PRODUCTION"]))
    ok("3PL instructions do reach TP", "double-stacked" in s)
    bad=await pg.evaluate("new Promise(r=>google.script.run.withSuccessHandler(x=>r('LEAK')).withFailureHandler(e=>r(e.message)).portalLoad('wrong-token'))")
    ok("wrong token is refused", bad!="LEAK", bad)
    other=await pg.evaluate("new Promise(r=>google.script.run.withSuccessHandler(x=>r('LEAK')).withFailureHandler(e=>r(e.message)).portalSubmit('preview-tp-0123456789abcdef0123','90203|14015',{}))")
    ok("TP cannot submit on a DL PO", other!="LEAK", other)
    hubtry=await pg.evaluate("new Promise(r=>google.script.run.withSuccessHandler(x=>r('LEAK')).withFailureHandler(e=>r(e.message)).hubLoad())")
    ok("portal page cannot call hub (no Google login)", hubtry!="LEAK", hubtry)
    # ---------- TP receiving: blind count, recount, send ----------
    txt=await pg.inner_text("main")
    ok("portal is in French", "Bonjour" in txt)
    ok("late PO shows orange/red", await pg.locator(".lvl-esc, .lvl-nudge").count()>0)
    await pg.locator("button:has-text('90106')").first.click(); await pg.wait_for_timeout(500)
    dtxt=await pg.inner_text("#drawer")
    ok("receive form opens", "90106" in dtxt, dtxt[:120])
    ok("expected qty is hidden (blind count)", "700" not in dtxt)
    # tiny test image
    import base64,os
    png="/tmp/t.png"
    if not os.path.exists(png):
      open(png,"wb").write(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="))
    async def up(f):
      await pg.set_input_files(f"[data-f='{f}']",png); await pg.wait_for_timeout(500)
    await up("loadPics")
    await pg.fill("#f_plt","18"); await pg.fill("#f_cs","690")
    await pg.click("#chk"); await pg.wait_for_timeout(600)
    d=await pg.inner_text("#drawer")
    ok("wrong count asks for a recount", await pg.locator("#f_rcs").count()==1)
    ok("check does not reveal expected qty", "700" not in d)
    await pg.fill("#f_rplt","18"); await pg.fill("#f_rcs","690")
    await pg.click("#chk"); await pg.wait_for_timeout(600)
    ok("second mismatch asks for a note", await pg.locator("#f_note").count()==1)
    await pg.fill("#f_note","Recounted twice, 10 cases short on pallet 4.")
    await up("label"); await pg.fill("#f_lot","L-555"); await pg.fill("#f_exp","2027-10-01"); await pg.check("#f_leg")
    await up("doc")
    await pg.click("[data-c='Good']"); await pg.wait_for_timeout(200)
    await pg.fill("#f_com","Driver late 30 min.")
    await pg.click("#send"); await pg.wait_for_timeout(1500)
    d=await pg.inner_text("#drawer")
    ok("submission accepted", "90106" in d and ("envoy" in d.lower() or "merci" in d.lower()), d[:200])
    db=await pg.evaluate("MOCK().MOCK_DB()")
    sub=[x for x in db["Submissions"] if x["id"]=="90106|10220"]
    ok("submission saved in Sheet tab", len(sub)==1 and str(sub[0]["cs"])=="690", sub)
    iss=[x for x in db["Issues"] if x["id"]=="90106|10220" and x.get("status")=="Open"]
    ok("quantity issue auto-opened, owner Purchasing", iss and iss[0]["owner"]=="Purchasing", iss)
    mail=await pg.evaluate("MOCK().MOCK_MAIL")
    ok("email alert drafted (pilot: only to Juan)", mail and all(m["to"]=="juan@tootsi.com" for m in mail), mail[:2])
    # ---------- hub ----------
    hub=await ctx.new_page(); hub.on("pageerror",lambda e:errs.append("hub "+str(e)))
    await hub.route(re.compile("fonts\\.(googleapis|gstatic)"),lambda r:r.abort())
    await hub.goto(P+"hub.html"); await hub.wait_for_timeout(1500)
    t=await hub.inner_text("main")
    ok("hub lists all warehouses", all(x in t for x in ["90101","90203","90301","90401","90402"]))
    ok("hub shows TP receipt", "Arrived" in t)
    await hub.click("[data-open='90106|10220']"); await hub.wait_for_timeout(900)
    d=await hub.inner_text("#drawer")
    ok("hub PO page shows the 3PL counts and note", "690" in d and "10 cases short" in d, d[:300])
    ok("hub PO page shows open issue", "Open issue" in d)
    await hub.fill("#n_txt","Called supplier about 10 cases."); await hub.click("#n_add"); await hub.wait_for_timeout(900)
    ok("note posted", "Called supplier" in await hub.inner_text("#drawer"))
    await hub.click("#x"); await hub.wait_for_timeout(300)
    # missing cases on imported PO
    await hub.click("[data-open='90401|80541']"); await hub.wait_for_timeout(900)
    await hub.fill("#e_cs","840"); await hub.click("#e_save"); await hub.wait_for_timeout(900)
    db=await hub.evaluate("MOCK().MOCK_DB()")
    ok("expected cases saved", any(x["id"]=="90401|80541" and str(x["cs"])=="840" for x in db["POs"]))
    await hub.click("#x"); await hub.wait_for_timeout(300)
    # add PO
    await hub.click("#addpo"); await hub.wait_for_timeout(300)
    for k,v in [("n_po","45001"),("n_item","10999"),("n_desc","TEST ALMONDS"),("n_sup","Test"),("n_eta","2026-11-27"),("n_plt","2"),("n_cs","80"),("n_lbs","2000")]: await hub.fill("#"+k,v)
    await hub.select_option("#n_wh","BM"); await hub.click("#nf button[type=submit]"); await hub.wait_for_timeout(1200)
    ok("manual PO added", "45001" in await hub.inner_text("main"))
    # BM receiving by Tootsi with shared form
    await hub.click("[data-open='90301|10501']"); await hub.wait_for_timeout(900)
    ok("BM PO offers Tootsi receiving", await hub.locator("#recvown").count()==1)
    await hub.click("#recvown"); await hub.wait_for_timeout(400)
    ok("receiving form opens in English", "Count" in await hub.inner_text("#drawer") or "count" in await hub.inner_text("#drawer"))
    # KPI filter
    await hub.click("#x"); await hub.wait_for_timeout(300)
    await hub.click("[data-k='rcv']"); await hub.wait_for_timeout(300)
    t=await hub.inner_text("tbody")
    ok("late KPI filter shows TP 90107 not received", "90107" in t, t[:200])
    await hub.click("[data-v='week']"); await hub.wait_for_timeout(300)
    ok("week board renders", await hub.locator(".week .wcol").count()==7)
    await hub.screenshot(path="/tmp/gws_hub.png")
    await pg.screenshot(path="/tmp/gws_tp.png")
    await b.close()
    print("ERRORS:",errs)
    print(f"{sum(x for _,x in R)}/{len(R)} passed")
asyncio.run(main())
