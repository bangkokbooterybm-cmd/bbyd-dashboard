// Build the view-only, password-protected snapshot of the Goal Board.
// usage: node snapshot.js <dataDir> <password> <out.html>
// dataDir holds db docs as saved by ArtifactData out_dir: <dataDir>/<collection>/<id>.json ({data:{...}} or plain)
const fs=require('fs'),path=require('path'),crypto=require('crypto');
const [dataDir,password,out]=process.argv.slice(2);
if(!dataDir||!password||!out){console.error('usage: node snapshot.js <dataDir> <password> <out.html>');process.exit(1);}
const rd=f=>{const j=JSON.parse(fs.readFileSync(f,'utf8'));return j&&j.data&&typeof j.data==='object'&&!Array.isArray(j.data)?j.data:j;};
const col=c=>{const d=path.join(dataDir,c);if(!fs.existsSync(d))return {};const o={};for(const f of fs.readdirSync(d))if(f.endsWith('.json'))o[f.slice(0,-5)]=rd(path.join(d,f));return o;};
const docs={};for(const [c,o] of Object.entries({config:col('config'),sync:col('sync')}))for(const [id,v] of Object.entries(o))docs[`${c}/${id}`]=v;
const cols={monthly:col('monthly'),team:col('team'),stock:col('stock')};
const src=process.env.PAGE||path.join(__dirname,'goal-setting.html');
const page=fs.readFileSync(src,'utf8');
const stub=`<script>(function(){const D=${JSON.stringify({docs,cols}).replace(/</g,'\\u003c')};
const snap=d=>({docs:d.map(([id,x])=>({id,exists:true,data:()=>x})),size:d.length,empty:!d.length});
const db={doc:p=>({onSnapshot:n=>{setTimeout(()=>{const v=D.docs[p];n({exists:!!v,data:()=>v});},0);return()=>{}},set:async()=>{throw {code:'read_only'}}}),
 collection:p=>({onSnapshot:n=>{setTimeout(()=>n(snap(Object.entries(D.cols[p]||{}))),0);return()=>{}}})};
const user={canEdit:async()=>false,can:async()=>false,id:async()=>'viewer'};
window.claude={use:async n=>n==='db'?db:n==='user'?user:null};})();</script>`;
const inner='<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow"></head><body>'+stub+page+'</body></html>';
// encrypt: PBKDF2-SHA256 (250k) -> AES-256-GCM
const salt=crypto.randomBytes(16),iv=crypto.randomBytes(12),key=crypto.pbkdf2Sync(password,salt,250000,32,'sha256');
const c=crypto.createCipheriv('aes-256-gcm',key,iv);const enc=Buffer.concat([c.update(inner,'utf8'),c.final(),c.getAuthTag()]);
const b64=x=>x.toString('base64');
const stamp=new Date().toLocaleString('th-TH',{timeZone:'Asia/Bangkok',dateStyle:'medium',timeStyle:'short'});
const wrap=`<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow"><title>Supervisor Goal Board</title>
<style>:root{--bg:#F1F3EF;--card:#fff;--ink:#17211C;--mute:#6F7B73;--acc:#2F6B4F;--line:#D5DBD1;--bad:#B3412F}
@media (prefers-color-scheme:dark){:root{--bg:#111613;--card:#18201B;--ink:#E4EAE5;--mute:#8A968E;--acc:#6FBF95;--line:#2E3A33;--bad:#EE8B79}}
body{margin:0;min-height:100vh;display:grid;place-items:center;background:var(--bg);color:var(--ink);font-family:"IBM Plex Sans Thai",system-ui,sans-serif;padding:16px;box-sizing:border-box}
form{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:24px;width:100%;max-width:340px;display:flex;flex-direction:column;gap:12px}
h1{font-size:20px;margin:0}p{margin:0;color:var(--mute);font-size:13px}input{font:inherit;padding:10px 12px;border:1px solid var(--line);border-radius:8px;background:transparent;color:var(--ink)}
button{font:600 14px inherit;padding:10px;border:0;border-radius:8px;background:var(--acc);color:#fff;cursor:pointer}label{font-size:13px;color:var(--mute);display:flex;gap:6px;align-items:center}.err{color:var(--bad);font-size:13px;min-height:18px}</style></head>
<body><form id="f"><h1>Supervisor Goal Board</h1><p>ข้อมูล ณ ${stamp}</p><input id="p" type="password" placeholder="รหัสผ่าน" autocomplete="current-password" required autofocus>
<label><input id="r" type="checkbox"> จำรหัสในเครื่องนี้</label><button>เปิดดู</button><div class="err" id="e"></div></form>
<script>const S="${b64(salt)}",I="${b64(iv)}",C="${b64(enc)}";const u=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
async function open_(pw,quiet){try{const k0=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']);
const k=await crypto.subtle.deriveKey({name:'PBKDF2',salt:u(S),iterations:250000,hash:'SHA-256'},k0,{name:'AES-GCM',length:256},false,['decrypt']);
const html=new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:u(I)},k,u(C)));return html;}catch(e){if(!quiet)document.getElementById('e').textContent='รหัสผ่านไม่ถูกต้อง';return null;}}
function show(h){document.open();document.write(h);document.close();}
let saved=null;try{saved=localStorage.getItem('gb.pw')}catch(e){}
if(saved)open_(saved,true).then(h=>{if(h)show(h);else{try{localStorage.removeItem('gb.pw')}catch(e){}}});
document.getElementById('f').addEventListener('submit',async ev=>{ev.preventDefault();const pw=document.getElementById('p').value;const h=await open_(pw);
if(h){if(document.getElementById('r').checked){try{localStorage.setItem('gb.pw',pw)}catch(e){}}show(h);}});</script></body></html>`;
fs.mkdirSync(path.dirname(out),{recursive:true});fs.writeFileSync(out,wrap);
console.log('wrote',out,(wrap.length/1024).toFixed(0)+'KB','docs',Object.keys(docs).length,'monthly',Object.keys(cols.monthly).length,'stock',Object.keys(cols.stock).length);
