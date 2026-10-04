const fs=require('fs'),vm=require('vm'),assert=require('assert');
function load(path){const html=fs.readFileSync(path,'utf8');const scripts=[...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]);const ctx={console,kpi(){return ""},window:null,document:{getElementById(){return null}},Intl};ctx.window=ctx;vm.createContext(ctx);vm.runInContext(scripts[0],ctx);const s=scripts[1];const helpers=s.slice(s.indexOf('  const fmtMoney'),s.indexOf('  const ico'));
vm.runInContext('var D=window.KRS_DATA,state={sf:{},sfBy:"clients"};'+helpers+';var SF_LAYOUT_STYLE="";function ov2Assets(){};function sfChart(d,m){return "<chart>"+JSON.stringify(d)+"</chart>"};function pageHead(a,b){return a+b};function dataTag(){return ""};function periodRange(){return window.range||[20260101,20261231]};function render(){};',ctx);
vm.runInContext(s.slice(s.indexOf('  function salesFilterPage(){'),s.indexOf('  window.setSF=function')),ctx);
if(s.includes('  window.selectSFRow='))vm.runInContext(s.slice(s.indexOf('  window.selectSFRow='),s.indexOf('  function salesFilterPage(){')),ctx);
return {ctx,html,render(){return vm.runInContext('salesFilterPage()',ctx)}};}
const before=load('output/Kairos_dashboard_proto-before-sales-layout.html'),after=load('Kairos_dashboard_proto.html');
assert.deepEqual(after.ctx.KRS_DATA,before.ctx.KRS_DATA);
for(const filters of [{},{filial:['Москва']},{filial:['Владивосток'],brand:['HeadRock']},{filial:['Сети']},{brand:['KRONbuild','ENKI']}]){
for(const app of [before,after])app.ctx.state.sf=JSON.parse(JSON.stringify(filters));
const a=before.render(),b=after.render();assert.deepEqual(JSON.parse(JSON.stringify(before.ctx.__sfExport)),JSON.parse(JSON.stringify(after.ctx.__sfExport)),'assort calculations preserved');assert(b.indexOf('Нажмите название строки')<b.indexOf('Динамика продаж<span'));assert(b.includes('Территория и ответственность'));assert(b.includes('Объект анализа'));assert(b.includes('setSFBy(\'brands\')'));}
after.ctx.state.sf={};after.ctx.state.sfBy='clients';after.render();const name=after.ctx.__sfRowNames[0];after.ctx.selectSFRow(0);assert.equal(after.ctx.state.sfFocus.name,name);const detail=after.render();assert(detail.includes('id="sfFocus"'));assert(detail.includes('Открыть карточку клиента'));assert.deepEqual(JSON.parse(JSON.stringify(after.ctx.state.sf)),{filial:[],okrug:[],city:[],manager:[],brand:[],client:[]});after.ctx.closeSFFocus();assert.equal(after.ctx.state.sfFocus,null);
for(const by of ['products','cities','okrugs','brands','managers']){after.ctx.state.sfBy=by;after.ctx.state.sfFocus=null;after.render();after.ctx.selectSFRow(0);assert(after.render().includes('id="sfFocus"'),by)}
after.ctx.state.sfFocus=null;after.ctx.state.sfBy='clients';after.ctx.state.sfLimit=30;after.ctx.showMoreSF();after.render();assert.equal(after.ctx.__sfRowNames.length,60);after.ctx.setSFDisclosure('assort',true);assert(after.render().includes('class="sf-disclosure" open'));
after.ctx.state.sf={client:['missing-client']};assert(after.render().includes('Нет данных. Проверьте'));
const template=fs.readFileSync('Kairos_proto_template.html','utf8');assert.equal(template.slice(template.indexOf('  function salesFilterPage(){'),template.indexOf('  window.setSF=function')),after.html.slice(after.html.indexOf('  function salesFilterPage(){'),after.html.indexOf('  window.setSF=function')));
console.log('PASS: data unchanged; five filter scenarios preserve assortment calculations; stable groups; six row-detail modes; global filters preserved; pagination; disclosure state; empty state; template synchronized.');

