// Focus Product data from the central Drive "BBY Data Update" (raw Odoo exports), replacing the manual Settings upload.
// node drive_odoo.js <stock.xlsx> <pos.xlsx> <dbDir> <stockDateISO> [stockFileName] [posFileName]
//   stock.xlsx : latest raw Odoo stock.quant export (Stock\ / Form Uploads\Stock...)  -- chosen by newest upload time (rule 7, 14 Nov 2026)
//   pos.xlsx   : latest raw Odoo "Point of Sale Orders" export (Sales Cut\ / Form Uploads\Sales Cut...)
//   dbDir      : artifact db folder (needs config/main.json for the Focus list)
//   stockDateISO: file date = YYYY-MM-DD in the file name, else upload time in Thai time (the raw export carries no date)
// Writes <dbDir>/stock/<date>.json, <dbDir>/stock/wh-<date>.json, <dbDir>/pos/<cutDate>.json and copies them to _out/db/<collection>/<id>.json
// (write those back with ArtifactData set). Exit code 2 = a check failed: nothing is written for that file, keep the old data.
const fs=require('fs'),path=require('path');global.XLSX=require('xlsx');
const P=require('./proc.js');
const [stockF,posF,dd,stockDate,stockName,posName]=process.argv.slice(2);
if(!dd||!/^\d{4}-\d{2}-\d{2}$/.test(stockDate||'')){console.error('usage: node drive_odoo.js <stock.xlsx|-> <pos.xlsx|-> <dbDir> <stockDate YYYY-MM-DD> [stockName] [posName]');process.exit(1);}
const cj=JSON.parse(fs.readFileSync(path.join(dd,'config','main.json'),'utf8'));const cfg=cj.data&&typeof cj.data==='object'?cj.data:cj;
if(!cfg.focus||!cfg.focus.length){console.error('ERROR no focus list in config/main.json');process.exit(2);}
const outDir=path.join(__dirname,'_out','db');
const save=(col,id,obj)=>{for(const base of [dd,outDir]){const d=path.join(base,col);fs.mkdirSync(d,{recursive:true});fs.writeFileSync(path.join(d,id+'.json'),JSON.stringify(JSON.parse(JSON.stringify(obj))));}};
let failed=0;
// ---- stock: raw Odoo columns are read BY NAME (the template changes: 9 / 11 / 12 / 21 columns) ----
const STOCK_NEED=['Lot/Serial Number','Model','Warehouse/Warehouse','Inventoried Quantity','Product Category','Skin','Variant Values'];
const STOCK_MIN=16500,STOCK_MAX=19500;
if(stockF&&stockF!=='-')try{
  const wb=XLSX.read(fs.readFileSync(stockF));const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});const h=(rows[0]||[]).map(x=>String(x==null?'':x).trim());
  const miss=STOCK_NEED.filter(n=>h.indexOf(n)<0);if(miss.length)throw new Error('stock file is missing columns: '+miss.join(', ')+' (header: '+h.join(' | ')+')');
  const ix=n=>h.indexOf(n),B=ix('Lot/Serial Number'),Q=ix('Inventoried Quantity');
  // Odoo also exports lot-less rows (empty Lot/Serial); only rows with a lot are pieces
  const lots=rows.slice(1).filter(r=>r&&r[B]);const pcs=new Set(lots.map(r=>String(r[B]))).size;
  const qty=lots.reduce((a,r)=>a+(+r[Q]||0),0);
  if(pcs<STOCK_MIN||pcs>STOCK_MAX)throw new Error(`stock pieces ${pcs} outside ${STOCK_MIN}-${STOCK_MAX}`);
  if(Math.abs(qty-pcs)>pcs*0.01)throw new Error(`stock quantity ${qty} differs from unique lots ${pcs}`);
  const [y,m,d]=stockDate.split('-');const sheetName=`Stock ${d}.${m}.${String(+y+543).slice(2)}`;
  const aoa=[['Barcode','Model','Branch','Qty.','Prod.Group','Skin','Color']].concat(lots.map(r=>[r[B],r[ix('Model')],r[ix('Warehouse/Warehouse')],1,r[ix('Product Category')],r[ix('Skin')],r[ix('Variant Values')]]));
  const nwb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(nwb,XLSX.utils.aoa_to_sheet(aoa),sheetName);
  const r=P.parseStock(nwb,cfg.focus,cfg.focusExclude);
  if(r.stockDate!==stockDate)throw new Error('stock date mismatch '+r.stockDate+' vs '+stockDate);
  if(r.rowsRead!==pcs)throw new Error(`parsed ${r.rowsRead} pieces but file has ${pcs} unique lots`);
  const whRows=r.wh;delete r.wh;
  save('stock',stockDate,Object.assign(r,{uploadedAt:Date.now(),fileName:stockName||path.basename(stockF),source:'drive'}));
  if(whRows)save('stock','wh-'+stockDate,{whFor:stockDate,wh:whRows,uploadedAt:Date.now()});
  const pc=Object.values(r.branches).reduce((a,b)=>a+b.pcs,0);
  console.log(`stock OK ${stockDate} pieces ${r.rowsRead} (qty ${qty}) 17-branch focus pcs ${pc} dupBarcodes ${r.dupBarcodes}`);
}catch(e){failed=2;console.log('stock ERROR '+e.message);}
// ---- Odoo sales cut ----
if(posF&&posF!=='-')try{
  const wb=XLSX.read(fs.readFileSync(posF));const r=P.parsePOS(wb,cfg.focus);
  if(!r.cutDate)throw new Error('cannot read the Date column');
  if(!r.diag||!r.diag.counted)throw new Error('no pieces counted in the 17 main branches');
  save('pos',r.cutDate,Object.assign(r,{uploadedAt:Date.now(),fileName:posName||path.basename(posF),source:'drive'}));
  console.log(`pos OK cut ${r.cutDate} from ${r.fromDate} counted ${r.diag.counted} pcs focus ${r.diag.focusPcs} pcs`);
}catch(e){failed=2;console.log('pos ERROR '+e.message);}
process.exit(failed);
