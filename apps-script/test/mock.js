// mock.js — runs the real Apps Script code (Store/Logic/Api/Setup/Jobs .gs) in a browser.
// Replaces Google services (Sheets, Drive, Mail, Session, Lock) with an in-memory store saved in localStorage.
// Used for the local tests and the preview. NOT part of the Apps Script project.
(function(){
const KEY="pohub_gws_preview_v4";
const pad=v=>String(v).padStart(2,"0");
function mockNow(){if(window.MOCK_NOW)return new Date(window.MOCK_NOW);return new Date()}
function ymd(d){return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate())}
window.Utilities={
 formatDate:(d,tz,f)=>f==="H"?String(d.getHours()):f==="yyyy-MM-dd HH:mm"?ymd(d)+" "+pad(d.getHours())+":"+pad(d.getMinutes()):ymd(d),
 getUuid:()=>([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g,c=>(c^Math.random()*16>>c/4).toString(16)),
 base64Decode:s=>{s=String(s).replace(/[^A-Za-z0-9+/]/g,'');while(s.length%4)s+='=';return Array.from(atob(s)).map(c=>{const x=c.charCodeAt(0);return x>127?x-256:x})},base64Encode:a=>btoa(String.fromCharCode.apply(null,a.map(x=>x&255))),newBlob:()=>({}),
 DigestAlgorithm:{SHA_256:"SHA_256"},Charset:{UTF_8:"UTF_8"},computeDigest:(alg,str)=>sha256(String(str))};
// small synchronous SHA-256 (Apps Script's computeDigest is synchronous too)
function sha256(str){const K=[],H=[];let n=2,c=0;const ip=x=>{for(let d=2;d*d<=x;d++)if(x%d===0)return false;return true},fr=x=>((x-Math.floor(x))*4294967296)|0;
 while(c<64){if(ip(n)){if(c<8)H[c]=fr(Math.pow(n,1/2));K[c]=fr(Math.pow(n,1/3));c++}n++}
 const b=Array.from(new TextEncoder().encode(str)),l=b.length*8;b.push(128);while(b.length%64!==56)b.push(0);for(let i=7;i>=0;i--)b.push(i>3?0:(l>>>(i*8))&255);
 const r=(x,y)=>(x>>>y)|(x<<(32-y)),h=H.slice();
 for(let o=0;o<b.length;o+=64){const w=[];for(let i=0;i<16;i++)w[i]=(b[o+i*4]<<24)|(b[o+i*4+1]<<16)|(b[o+i*4+2]<<8)|b[o+i*4+3];
  for(let i=16;i<64;i++){const s0=r(w[i-15],7)^r(w[i-15],18)^(w[i-15]>>>3),s1=r(w[i-2],17)^r(w[i-2],19)^(w[i-2]>>>10);w[i]=(w[i-16]+s0+w[i-7]+s1)|0}
  let [A,B,C,D,E,F,G,I]=h;for(let i=0;i<64;i++){const t1=(I+(r(E,6)^r(E,11)^r(E,25))+((E&F)^(~E&G))+K[i]+w[i])|0,t2=((r(A,2)^r(A,13)^r(A,22))+((A&B)^(A&C)^(B&C)))|0;I=G;G=F;F=E;E=(D+t1)|0;D=C;C=B;B=A;A=(t1+t2)|0}
  [A,B,C,D,E,F,G,I].forEach((v,i)=>h[i]=(h[i]+v)|0)}
 const out=[];h.forEach(v=>{for(let i=3;i>=0;i--){const x=(v>>>(i*8))&255;out.push(x>127?x-256:x)}});return out}
window.MOCK_SHA=s=>sha256(s).map(x=>("0"+(x&255).toString(16)).slice(-2)).join("");
const CACHE={};window.CacheService={getScriptCache:()=>({get:k=>CACHE[k]===undefined?null:CACHE[k],put:(k,v)=>{CACHE[k]=v},remove:k=>{delete CACHE[k]}})};
window.Session={getActiveUser:()=>({getEmail:()=>window.MOCK_USER||""})};
window.LockService={getScriptLock:()=>({waitLock(){},tryLock(){return true},releaseLock(){}})};
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
  {id:"90410|80450",po:"90410",item:"80450",wh:"TP",desc:"STRAWBERRY CANDY BRICKS",supplier:"Candy maker · Spain",eta:bd(-1),cntr:"OOLU5550001",plt:21,cs:"",lbs:1670,load:"FCL 40' · Palletized",alloc:"",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90410|80451",po:"90410",item:"80451",wh:"TP",desc:"RASPBERRY CANDY BRICKS",supplier:"Candy maker · Spain",eta:bd(-1),cntr:"OOLU5550001",plt:21,cs:"",lbs:1670,load:"FCL 40' · Palletized",alloc:"",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90410|80490",po:"90410",item:"80490",wh:"TP",desc:"SUGARED JELLY MANDARINS",supplier:"Candy maker · Spain",eta:bd(-1),cntr:"OOLU5550001",plt:21,cs:"",lbs:2200,load:"FCL 40' · Palletized",alloc:"",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90410|80492",po:"90410",item:"80492",wh:"TP",desc:"GUMMY MULTICOLOR RINGS",supplier:"Candy maker · Spain",eta:bd(-1),cntr:"OOLU5550001",plt:21,cs:"",lbs:1900,load:"FCL 40' · Palletized",alloc:"",qc:"Not received",source:"sheet",addedBy:"Import"},
  {id:"90402|10307",po:"90402",item:"10307",wh:"",desc:"DRY ROASTED SALTED CALIFORNIA PISTACHIOS",supplier:"Pistachio processor · California",eta:bd(4),cntr:"",plt:20,cs:"",lbs:10000,load:"FTL · Palletized",alloc:"",qc:"Not received",source:"sheet",addedBy:"Import"}
 ],Submissions:[
  {subId:"s1",id:"90203|14015",seq:1,superseded:false,by:"Deslauriers receiving",at:bd(-1)+" 14:10",recv:bd(-1),plt:20,cs:1580,lbs:43545,firstJson:{plt:20,cs:1578,lbs:43490},diffNote:"Recounted twice. 20 cases missing.",lot:"450072532-14-585",exp:"2027-09-30",cond:"Torn / leaking bags",damage:true,dmgText:"2 bags torn on bottom of pallet 7, product spilled.",comment:"Seal intact. Bottom row crushed.",filesJson:[{cat:"Load photo (before unloading)",name:"container_open.jpg",url:"#"},{cat:"Label photo",name:"label_14015.jpg",url:"#"},{cat:"Receiving document",name:"DL_receiving_90203.pdf",url:"#"},{cat:"Damage photo",name:"damage_pallet7.jpg",url:"#"}]}
 ],Issues:[],Notes:[{id:"90203|14015",who:"Purchasing",text:"Sent photos and counts to supplier.",at:bd(-1)+" 16:02",ext:false}],
 Delays:[{delayId:"d1",id:"90204|10711",reason:"r_customs",note:"Container held at customs. Broker expects release in 2 days.",newDate:bd(2),by:"Deslauriers receiving",at:bd(-1)+" 15:30",date:bd(-1),closedAt:"",filesJson:[]}],
 Recounts:[],Docs:[{id:"90203|14015",cat:"Supplier packing list",name:"packing_list_90203.pdf",url:"#",by:"purchasing@tootsi.com",at:bd(-3)+" 10:00"}],
 SupplierQty:[{id:"90203|14015",kind:"pl",plt:20,cs:1600,lbs:44096,by:"Purchasing",at:bd(-3)+" 10:00"}],
 Users:[{token:"u-tp-0000000000000000000000000001",wh:"TP",name:"Trans-Plus",email:"",active:true,user:"transplus",salt:"s1",passHash:window.MOCK_SHA("s1|demo-tp")},{token:"u-dl-0000000000000000000000000002",wh:"DL",name:"Deslauriers",email:"",active:true,user:"deslauriers",salt:"s2",passHash:window.MOCK_SHA("s2|demo-dl")}],Sessions:[],Checks:[],MasterSync:[],MasterSeen:[],Slips:[{po:"90410",status:"FOUND",url:"#",fileId:"slipfile1",source:"From PACKING LIST OOLU5550001.pdf (pages 1 of 1)",action:"",checked:today},{po:"90101",status:"MISSING",url:"",fileId:"",source:"Checked 1 PDF(s) in the folder, none is a packing slip",action:"Add the packing slip to the PO folder.",checked:today}],
 Config:Object.entries({DOMAIN:"tootsi.com",SESSION_HOURS:12,MASTER_SHEET_ID:"master-preview",MASTER_MODE:"preview",SLIP_TO_3PL:"after",GO_LIVE:bd(-6),PORTAL_URL:"3pl.html",SLA_NUDGE_H:9,SLA_ESC_H:18,LB_TOLERANCE:0.01,QC_DAYS:2,HOLIDAYS:"",EMAIL_ON:"yes",PILOT_ONLY_ME:"yes",FALLBACK_EMAIL:"juan@tootsi.com",EMAIL_3PL:"no",LANG_TP:"fr",LANG_DL:"fr",DIGEST_TO:"juan@tootsi.com"}).map(([key,value])=>({key,value})),
 Log:[]};
 save();
}
try{DB=JSON.parse(localStorage.getItem(KEY))}catch(e){}
if(!DB||window.MOCK_RESET)seed();
window.MOCK_DB=()=>DB;window.MOCK_RESEED=()=>{seed()};

// ---------- replace Store.gs primitives ----------
const clone=o=>JSON.parse(JSON.stringify(o));
window.readAll_=name=>{const cols=TABLES[name];return (DB[name]||[]).map(r=>{const o={};cols.forEach(c=>{o[c]=r[c]===undefined?(/Json$/.test(c)?null:""):r[c]});return clone(o)})};
window.append_=(name,o)=>{const r={};TABLES[name].forEach(c=>{if(o[c]!==undefined)r[c]=clone(o[c]===null?"":o[c])});(DB[name]=DB[name]||[]).push(r);save()};
window.upsert_=(name,key,o)=>{const arr=DB[name]=DB[name]||[];const i=arr.findIndex(r=>String(r[key])===String(o[key]));const r=i>=0?arr[i]:{};TABLES[name].forEach(c=>{if(c in o&&c[0]!=="_")r[c]=clone(o[c]===undefined?"":o[c])});if(i<0)arr.push(r);save()};
window.upsertWhere_=(name,match,o)=>{const arr=DB[name]=DB[name]||[];const i=arr.findIndex(r=>match(r));if(i>=0)arr[i]=clone(o);else arr.push(clone(o));save()};
window.deleteWhere_=(name,match)=>{DB[name]=(DB[name]||[]).filter(r=>!match(r));save()};
window.clearCache_=()=>{};
window.saveFile_=(id,f)=>({name:f.name,url:"#preview-file",fileId:"x"});
const INC={};let incN=0;
window.saveIncoming_=(who,f)=>{const fid="inc"+(++incN);INC[fid]={who,name:f.name};return{fid,name:f.name}};
window.verifyFile_=(fid,who)=>{const x=INC[fid];if(!x||x.who!==who)throw new Error("A photo was not uploaded. Please add it again.");return{name:x.name,url:"#preview-file"}};
window.finishMoves_=moves=>{(moves||[]).forEach(mv=>{delete INC[mv.fid]})};
window.slipFile_=fid=>({name:"PO packing slip.pdf",type:"application/pdf",data:"JVBERi0xLjQKJcOkw7zDtsOfCjEgMCBvYmo8PC9UeXBlL0NhdGFsb2c+PgplbmRvYmoKdHJhaWxlcjw8L1Jvb3QgMSAwIFI+PgolJUVPRgo="});
window.poFolder_=()=>({getUrl:()=>"#preview-folder"});
window.syncSources_=()=>({upcoming:0,slips:0,master:{seen:0,written:0,preview:0}});
window.today_=()=>ymd(mockNow());
window.nowStamp_=()=>{const d=mockNow();return ymd(d)+" "+pad(d.getHours())+":"+pad(d.getMinutes())};
window.ctx_=()=>({cfg:config_(),today:today_(),nowH:mockNow().getHours(),stamp:nowStamp_()});

// ---------- google.script.run (async, JSON round-trip like the real thing) ----------
function runner(ok,fail){return new Proxy({},{get(_,p){
 if(p==="withSuccessHandler")return f=>runner(f,fail);
 if(p==="withFailureHandler")return f=>runner(ok,f);
 return(...args)=>setTimeout(()=>{try{const fn=/_$/.test(String(p))?null:window[p];if(typeof fn!=="function")throw new Error("No server function "+p);
  const res=fn(...JSON.parse(JSON.stringify(args)));window.MOCK_LAST={fn:p,res:res===undefined?null:JSON.parse(JSON.stringify(res))};ok&&ok(res===undefined?null:JSON.parse(JSON.stringify(res)))}catch(e){fail&&fail(e)}},60)}})}
window.google={script:{run:runner(null,null)}};
})();
