// mock.js — runs the real Apps Script code (Store/Logic/Api/Setup/Jobs .gs) in a browser.
// Replaces Google services (Sheets, Drive, Mail, Session, Lock) with an in-memory store saved in localStorage.
// Used for the local tests and the preview. NOT part of the Apps Script project.
(function(){
const KEY="pohub_gws_preview_v1";
const pad=v=>String(v).padStart(2,"0");
function mockNow(){if(window.MOCK_NOW)return new Date(window.MOCK_NOW);return new Date()}
function ymd(d){return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate())}
window.Utilities={
 formatDate:(d,tz,f)=>f==="H"?String(d.getHours()):f==="yyyy-MM-dd HH:mm"?ymd(d)+" "+pad(d.getHours())+":"+pad(d.getMinutes()):ymd(d),
 getUuid:()=>([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g,c=>(c^Math.random()*16>>c/4).toString(16)),
 base64Decode:()=>[],newBlob:()=>({})};
window.Session={getActiveUser:()=>({getEmail:()=>window.MOCK_USER||""})};
window.LockService={getScriptLock:()=>({waitLock(){},releaseLock(){}})};
window.MOCK_MAIL=[];window.MailApp={sendEmail:o=>window.MOCK_MAIL.push(o)};
window.PropertiesService={getScriptProperties:()=>({getProperty:()=>null,setProperty(){}})};
window.Logger={log:()=>{}};

// ---------- in-memory tables ----------
let DB=null;
function save(){try{localStorage.setItem(KEY,JSON.stringify(DB))}catch(e){}}
function seed(){
 const today=ymd(mockNow()),d=n=>{const x=new Date(today+"T12:00");x.setDate(x.getDate()+n);return ymd(x)};
 // business-day offsets so the colors look right whatever day you open the preview
 const bd=n=>{const x=new Date(today+"T12:00");let k=0;while(k<Math.abs(n)){x.setDate(x.getDate()+(n<0?-1:1));if(x.getDay()>0&&x.getDay()<6)k++}return ymd(x)};
 const stamp=ymd(mockNow())+" 08:00";
 DB={POs:[
  {id:"90101|10201",po:"90101",item:"10201",wh:"TP",desc:"RAW WHOLE CASHEWS (W-320)",supplier:"Cashew exporter · Vietnam",eta:bd(1),cntr:"TCLU4826103",plt:28,cs:1100,lbs:55000,load:"FCL 40' · Palletized",alloc:"Production",qc:"Not received",instrSap:"FOR PRODUCTION ONLY. Palletized, stretch-wrapped, no double stacking.",tplInstr:"Do not accept double-stacked pallets. Photograph any torn bag before unloading.",source:"sheet",addedBy:"Import"},
  {id:"90102|20087",po:"90102",item:"20087",wh:"TP",desc:"NATURAL COCOA POWDER 10-12% FAT",supplier:"Cocoa processor · Peru",eta:bd(2),cntr:"MSKU7731902",plt:20,cs:400,lbs:22048,load:"FCL 40' · Palletized",alloc:"Split",prodQty:200,salesQty:200,so:"SO 118377",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90106|10220",po:"90106",item:"10220",wh:"TP",desc:"CASHEW PIECES (SSP2)",supplier:"Cashew processor · Ghana",eta:bd(-2),cntr:"UETU3091547",plt:18,cs:700,lbs:35000,load:"FCL 20' · Palletized",alloc:"Sales",so:"SO 118441",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90107|31009",po:"90107",item:"31009",wh:"TP",desc:"ORGANIC HULLED MILLET",supplier:"Grain exporter · Ukraine",eta:bd(-1),cntr:"CAIU6280415",plt:21,cs:800,lbs:40000,load:"FCL 20' · Palletized",alloc:"Production",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90401|80541",po:"90401",item:"80541",wh:"LY",desc:"GUMMY COLA BOTTLES",supplier:"Candy maker · Spain",eta:bd(3),cntr:"OOLU9841314",plt:21,cs:"",lbs:1670,load:"FCL 40' · Palletized",alloc:"",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90203|14015",po:"90203",item:"14015",wh:"DL",desc:"SULTANA RAISINS (TURKEY)",supplier:"Raisin packer · Türkiye",eta:bd(2),cntr:"MSBU1188204",plt:20,cs:1600,lbs:44096,load:"FCL 20' · Floor loaded",alloc:"Production",qc:"Locked",qcAt:bd(-1),source:"sheet",addedBy:"Import",tplInstr:"Floor-loaded. Photograph the open container before unloading."},
  {id:"90204|10711",po:"90204",item:"10711",wh:"DL",desc:"TURKISH APRICOTS",supplier:"Apricot packer · Türkiye",eta:bd(-1),cntr:"MEDU9047731",plt:17,cs:1299,lbs:36372,load:"FCL 40' · Palletized",alloc:"Sales",so:"SO 118412",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90301|10501",po:"90301",item:"10501",wh:"BM",desc:"FANCY JUMBO PECAN HALVES",supplier:"Pecan sheller · Mexico",eta:today,cntr:"",plt:22,cs:1680,lbs:41580,load:"FTL · Palletized",alloc:"Sales",so:"SO 118455",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90402|10307",po:"90402",item:"10307",wh:"",desc:"DRY ROASTED SALTED CALIFORNIA PISTACHIOS",supplier:"Pistachio processor · California",eta:bd(4),cntr:"",plt:20,cs:"",lbs:10000,load:"FTL · Palletized",alloc:"",qc:"Not received",source:"sheet",addedBy:"Import"}
 ],Submissions:[
  {subId:"s1",id:"90203|14015",seq:1,superseded:false,by:"Deslauriers receiving",at:bd(-1)+" 14:10",recv:bd(-1),plt:20,cs:1580,lbs:43545,firstJson:{plt:20,cs:1578,lbs:43490},diffNote:"Recounted twice. 20 cases missing.",lot:"450072532-14-585",exp:"2027-09-30",cond:"Torn / leaking bags",damage:true,dmgText:"2 bags torn on bottom of pallet 7, product spilled.",comment:"Seal intact. Bottom row crushed.",filesJson:[{cat:"Load photo (before unloading)",name:"container_open.jpg",url:"#"},{cat:"Label photo",name:"label_14015.jpg",url:"#"},{cat:"Receiving document",name:"DL_receiving_90203.pdf",url:"#"},{cat:"Damage photo",name:"damage_pallet7.jpg",url:"#"}]}
 ],Issues:[],Notes:[{id:"90203|14015",who:"Purchasing",text:"Sent photos and counts to supplier.",at:bd(-1)+" 16:02",ext:false}],
 Delays:[{delayId:"d1",id:"90204|10711",reason:"r_customs",note:"Container held at customs. Broker expects release in 2 days.",newDate:bd(2),by:"Deslauriers receiving",at:bd(-1)+" 15:30",date:bd(-1),closedAt:"",filesJson:[]}],
 Recounts:[],Docs:[{id:"90203|14015",cat:"Supplier packing list",name:"packing_list_90203.pdf",url:"#",by:"purchasing@tootsi.com",at:bd(-3)+" 10:00"}],
 SupplierQty:[{id:"90203|14015",kind:"pl",plt:20,cs:1600,lbs:44096,by:"Purchasing",at:bd(-3)+" 10:00"}],
 Users:[{token:"preview-tp-0123456789abcdef0123",wh:"TP",name:"Trans-Plus",email:"",active:true},{token:"preview-dl-0123456789abcdef0123",wh:"DL",name:"Deslauriers",email:"",active:true}],
 Config:Object.entries({DOMAIN:"tootsi.com",SLA_NUDGE_H:9,SLA_ESC_H:18,LB_TOLERANCE:0.01,QC_DAYS:2,HOLIDAYS:"",EMAIL_ON:"yes",PILOT_ONLY_ME:"yes",FALLBACK_EMAIL:"juan@tootsi.com",EMAIL_3PL:"no",LANG_TP:"fr",LANG_DL:"fr",DIGEST_TO:"juan@tootsi.com"}).map(([key,value])=>({key,value})),
 Log:[]};
 save();
}
try{DB=JSON.parse(localStorage.getItem(KEY))}catch(e){}
if(!DB||window.MOCK_RESET)seed();
window.MOCK_DB=()=>DB;window.MOCK_RESEED=()=>{seed()};

// ---------- replace Store.gs primitives ----------
const clone=o=>JSON.parse(JSON.stringify(o));
window.readAll=name=>{const cols=TABLES[name];return (DB[name]||[]).map(r=>{const o={};cols.forEach(c=>{o[c]=r[c]===undefined?(/Json$/.test(c)?null:""):r[c]});return clone(o)})};
window.append=(name,o)=>{const r={};TABLES[name].forEach(c=>{if(o[c]!==undefined)r[c]=clone(o[c]===null?"":o[c])});(DB[name]=DB[name]||[]).push(r);save()};
window.upsert=(name,key,o)=>{const arr=DB[name]=DB[name]||[];const i=arr.findIndex(r=>String(r[key])===String(o[key]));const r=i>=0?arr[i]:{};TABLES[name].forEach(c=>{if(c in o&&c[0]!=="_")r[c]=clone(o[c]===undefined?"":o[c])});if(i<0)arr.push(r);save()};
window.upsertWhere=(name,match,o)=>{const arr=DB[name]=DB[name]||[];const i=arr.findIndex(r=>match(r));if(i>=0)arr[i]=clone(o);else arr.push(clone(o));save()};
window.clearCache_=()=>{};
window.saveFile_=(id,f)=>({name:f.name,url:"#preview-file",fileId:"x"});
window.today_=()=>ymd(mockNow());
window.nowStamp_=()=>{const d=mockNow();return ymd(d)+" "+pad(d.getHours())+":"+pad(d.getMinutes())};
window.ctx_=()=>({cfg:config_(),today:today_(),nowH:mockNow().getHours(),stamp:nowStamp_()});

// ---------- google.script.run (async, JSON round-trip like the real thing) ----------
function runner(ok,fail){return new Proxy({},{get(_,p){
 if(p==="withSuccessHandler")return f=>runner(f,fail);
 if(p==="withFailureHandler")return f=>runner(ok,f);
 return(...args)=>setTimeout(()=>{try{const fn=window[p];if(typeof fn!=="function")throw new Error("No server function "+p);
  const res=fn(...JSON.parse(JSON.stringify(args)));window.MOCK_LAST={fn:p,res:res===undefined?null:JSON.parse(JSON.stringify(res))};ok&&ok(res===undefined?null:JSON.parse(JSON.stringify(res)))}catch(e){fail&&fail(e)}},60)}})}
window.google={script:{run:runner(null,null)}};
})();
