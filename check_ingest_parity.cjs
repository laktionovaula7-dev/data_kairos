// Проверка: браузерный пересчёт (kairos_ingest.js) даёт ту же модель, что и rebuild.py.
// Запуск:  node check_ingest_parity.cjs   (нужны Excel-выгрузки в папке и собранный Kairos_dashboard_proto.html)
const fs=require('fs'), path=require('path'), KI=require('./kairos_ingest.js');
const R=__dirname+path.sep;
(async()=>{
  const html=fs.readFileSync(R+'Kairos_dashboard_proto.html','utf8'), k='window.KRS_DATA = ', i=html.indexOf(k)+k.length;
  let depth=0,j=i,inStr=false,esc=false; // границы JSON-объекта
  for(;j<html.length;j++){const c=html[j]; if(inStr){ if(esc)esc=false; else if(c==='\\')esc=true; else if(c==='"')inStr=false; continue;} if(c==='"')inStr=true; else if(c==='{')depth++; else if(c==='}'&&--depth===0){j++;break;}}
  const D=JSON.parse(html.slice(i,j));
  const files={sales:await fs.openAsBlob(R+'Продажи по бизнес регионам 24г-26г.xlsx'),stock:await fs.openAsBlob(R+'Остатки и доступность по сериям.xlsx'),expiry:await fs.openAsBlob(R+'Отчет по товарам на складах с окончанием срока годности.xlsx')};
  files.sales.name='sales.xlsx';
  const t0=Date.now(); const K=await KI.run(files,D,()=>{});
  console.log('пересчёт за',((Date.now()-t0)/1000).toFixed(1),'с; данные на',K.dataAsOfHuman,'| нетто',Math.round(K.overview.sales));
  const diffs=[];
  const cmp=(a,b,p)=>{ if(a===b)return;
    if(typeof a==='number'&&typeof b==='number'){ if(Math.abs(a-b)>Math.max(1,Math.abs(b)*1e-6))diffs.push([p,a,b]); return;}
    if(a&&b&&typeof a==='object'&&typeof b==='object'){
      if(Array.isArray(a)){ if(a.length!==b.length)diffs.push([p,'длина '+a.length+' ≠ '+b.length]); const n=Math.min(a.length,b.length); let c=0; for(let q=0;q<n&&c<3;q++){const w=diffs.length;cmp(a[q],b[q],p+'['+q+']');if(diffs.length>w)c++;} return;}
      const ka=Object.keys(a),kb=Object.keys(b); if(ka.length!==kb.length)diffs.push([p,'ключей '+ka.length+' ≠ '+kb.length]);
      let c=0; for(const x of kb){ if(x in a){const w=diffs.length;cmp(a[x],b[x],p+'.'+x);if(diffs.length>w&&++c>=4)break;} else diffs.push([p+'.'+x,'нет в JS']); } return;}
    diffs.push([p,String(a).slice(0,50),String(b).slice(0,50)]); };
  for(const key of Object.keys(D)){ if(key==='refs'||key==='meta')continue; cmp(K[key],D[key],key); }
  if(diffs.length){ console.log('РАСХОЖДЕНИЯ с rebuild.py:',diffs.length); diffs.slice(0,30).forEach(d=>console.log(' ',JSON.stringify(d))); process.exit(1); }
  console.log('OK: браузерный расчёт совпал с rebuild.py по всем разделам модели');
})().catch(e=>{console.error(e);process.exit(1)});
