// Refresh pipeline for the Supervisor Goal Board snapshot.
// node run.js <xlsxDir> <dbDir> <password>
//   xlsxDir: staff.xlsx, sales.xlsx, audit.xlsx, demand.xlsx (Google Sheets exported as xlsx)
//   dbDir:   current artifact db docs (config/, monthly/, stock/, team/, sync/) as saved by ArtifactData out_dir
// Writes: <dbDir>/sync/{staff,sales,audit,demand}.json refreshed, _out/sync_*.json (to write back to the db),
//         ../index.html (encrypted page)
const fs=require('fs'),path=require('path');global.XLSX=require('xlsx');
const P=require('./proc.js');const [xd,dd,pw]=process.argv.slice(2);
if(!xd||!dd||!pw){console.error('usage: node run.js <xlsxDir> <dbDir> <password>');process.exit(1);}
const rd=f=>XLSX.read(fs.readFileSync(path.join(xd,f)));const now=Date.now();
const out={};const log=[];
const step=(name,file,fn)=>{if(!fs.existsSync(path.join(xd,file))){log.push(`${name}: ไม่มีไฟล์ ใช้ข้อมูลเดิม`);return;}try{out[name]=Object.assign({updatedAt:now},fn(rd(file)));log.push(`${name}: OK`);}catch(e){log.push(`${name}: ERROR ${e.message}`);}};
step('staff','staff.xlsx',wb=>P.parseStaff(wb));
step('sales','sales.xlsx',wb=>P.parseSales(wb));
step('audit','audit.xlsx',wb=>P.parseAudit(wb));
step('demand','demand.xlsx',wb=>P.parseDemand(wb));
fs.mkdirSync(path.join(dd,'sync'),{recursive:true});fs.mkdirSync(path.join(__dirname,'_out'),{recursive:true});
for(const [k,v] of Object.entries(out)){const j=JSON.parse(JSON.stringify(v));fs.writeFileSync(path.join(dd,'sync',k+'.json'),JSON.stringify(j));fs.writeFileSync(path.join(__dirname,'_out','sync_'+k+'.json'),JSON.stringify(j));}
// inline proc into the page, then encrypt
const page=fs.readFileSync(path.join(__dirname,'page.src.html'),'utf8').replace('/*PROC*/',()=>fs.readFileSync(path.join(__dirname,'proc.js'),'utf8'));
const tmp=path.join(__dirname,'_out','goal-setting.html');fs.writeFileSync(tmp,page);
process.env.PAGE=tmp;process.argv=[process.argv[0],'snapshot.js',dd,pw,path.join(__dirname,'..','index.html')];require('./snapshot.js');
const s=out.sales||JSON.parse(fs.readFileSync(path.join(dd,'sync','sales.json'),'utf8'));const m=s.months||(s.data&&s.data.months)||{};const last=Object.keys(m).sort().pop();
console.log(log.join('\n'));if(last)console.log('sales last date',m[last].lastDate,'BKK',m[last].BKK.sales,'UPC',m[last].UPC.sales);
