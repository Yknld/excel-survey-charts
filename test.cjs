const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const ts = require('typescript');
const path = require('node:path');
const sourcePath = path.join(__dirname, 'create_survey_charts.ts');
const source = fs.readFileSync(sourcePath, 'utf8');

// Compile against Microsoft's actual Office Scripts release declarations.
const declarations = fs.readFileSync(path.join(__dirname, 'excel.d.ts'), 'utf8').replace('export declare namespace ExcelScript', 'declare namespace ExcelScript');
const virtual = path.join(__dirname, 'global-excel.d.ts');
const options = {noEmit:true, strict:true, target:ts.ScriptTarget.ES2020, skipLibCheck:true, types:[]};
const host = ts.createCompilerHost(options);
const read = host.readFile.bind(host);
host.readFile = name => name === virtual ? declarations : read(name);
const program = ts.createProgram([sourcePath, virtual], options, host);
const diagnostics = ts.getPreEmitDiagnostics(program);
assert.equal(diagnostics.length, 0, ts.formatDiagnosticsWithColorAndContext(diagnostics, {getCanonicalFileName:n=>n,getCurrentDirectory:()=>__dirname,getNewLine:()=> '\n'}));

const ExcelScript = {
  ChartType:{barClustered:'bar',pie:'pie',barStacked100:'stacked',columnClustered:'column'},
  ChartDataLabelPosition:{bestFit:'bestFit',outsideEnd:'outsideEnd'},
  ChartLegendPosition:{right:'right',bottom:'bottom'}
};
const context = {ExcelScript, console:{log:()=>{}}};
vm.createContext(context);
vm.runInContext(ts.transpileModule(source,{compilerOptions:{target:ts.ScriptTarget.ES2020}}).outputText,context);
let passed=0;
function test(name, fn) {fn();passed++;console.log('PASS '+name);}
function grid(french=false,text=false) {
  const labels=french?['Oui','Non']:['Yes','No'];
  return [['Q4 — Transfer payments','','',''],[french?'Catégorie':'Category','Institutions','Base (n)','2025–2026'],
    [labels[0],text?'40':40,text?'92':92,text?'43,5 %':40/92],
    [labels[1],text?'52':52,text?'92':92,text?'56,5\u00a0%':52/92],['','','',''],['Skipped',11,'','']];
}
function columnName(n){let v='';for(n++;n>0;n=Math.floor((n-1)/26))v=String.fromCharCode(65+(n-1)%26)+v;return v;}
class Range {
  constructor(sheet,r,c,h=1,w=1){Object.assign(this,{sheet,r,c,h,w});}
  getValues(){return Array.from({length:this.h},(_,i)=>Array.from({length:this.w},(_,j)=>this.sheet.cells[this.r+i]?.[this.c+j]??''));}
  getRowIndex(){return this.r;} getColumnIndex(){return this.c;}
  getAddress(){return "'"+this.sheet.name.replaceAll("'","''")+"'!"+columnName(this.c)+(this.r+1);}
  setValue(v){this.setFormulas([[v]]);}
  setFormulas(rows){rows.forEach((row,i)=>row.forEach((v,j)=>{this.sheet.cells[this.r+i]??=[];this.sheet.cells[this.r+i][this.c+j]=v;}));}
  setNumberFormat(v){this.numberFormat=v;}
}
function format(){return {getFill:()=>({setSolidColor:()=>{}}),getFont:()=>({setSize:()=>{}})};}
class Series {
  constructor(chart,name){this.chart=chart;this.name=name;}
  delete(){this.chart.series=this.chart.series.filter(s=>s!==this);}
  setValues(r){this.values=r;} setXAxisValues(r){this.categories=r;}
  getFormat(){return format();} setVaryByCategories(){}
  getDataLabels(){return {setShowValue(){},setShowPercentage(){},setShowCategoryName(){},setShowSeriesName(){},setLinkNumberFormat(){},setNumberFormat(){},setPosition(){}};}
}
class Chart {
  constructor(type){this.type=type;this.series=[new Series(this,'default')];this.axis={setMinimum:v=>this.min=v,setMaximum:v=>this.max=v,setNumberFormat(){},setReversePlotOrder(){}};}
  getSeries(){return [...this.series];} addChartSeries(name){const s=new Series(this,name);this.series.push(s);return s;}
  setName(v){this.name=v;} getTitle(){return {setText:v=>this.title=v,setVisible(){},getFormat:format};}
  getFormat(){return format();} getLegend(){return {setVisible(){},setPosition(){}};}
  getAxes(){return {getValueAxis:()=>this.axis,getCategoryAxis:()=>this.axis};}
  setLeft(v){this.left=v;} setTop(v){this.top=v;} setWidth(v){this.width=v;} setHeight(v){this.height=v;}
}
class Sheet {
  constructor(name,cells,book){Object.assign(this,{name,cells,book});this.charts=[];}
  getName(){return this.name;} getUsedRange(){if(!this.cells.length)return undefined;return new Range(this,0,0,this.cells.length,Math.max(...this.cells.map(r=>r.length)));}
  getRangeByIndexes(r,c,h,w){return new Range(this,r,c,h,w);}
  getRange(address){assert.equal(address,'AA1');return new Range(this,0,26);}
  setShowGridlines(){} addChart(type){const chart=new Chart(type);this.charts.push(chart);return chart;}
  activate(){this.book.active=this;}
}
function workbook(cells,name='Analysis'){
  const book={sheets:[],getActiveWorksheet(){return this.active;},getWorksheet(name){return this.sheets.find(s=>s.name===name);},addWorksheet(name){const s=new Sheet(name,[],this);this.sheets.push(s);return s;}};
  const s=book.addWorksheet(name);s.cells=cells;book.active=s;return book;
}
test('numeric, English percentage text, decimal-comma text, NBSP and narrow NBSP',()=>{
 for(const [v,w] of [[0.343,0.343],['34.3%',0.343],['34,3 %',0.343],['34,3\u00a0%',0.343],['34,3\u202f%',0.343],['0,343',0.343]])assert.ok(Math.abs(context.readNumber(v,true)-w)<1e-12);
 for(const v of ['',true,'abc','1,234.56'])assert.ok(Number.isNaN(context.readNumber(v,true)));
});
test('numeric and French text tables give identical exact ratios',()=>{
 const a=context.findQuestions(grid(),0,0)[0],b=context.findQuestions(grid(true,true),5,3)[0];
 assert.equal(a.kind,'pie');assert.equal(b.kind,'pie');assert.equal(b.row,7);assert.equal(b.column,3);assert.equal(b.percentages[0],40/92);
});
test('full main() path produces linked charts and preserves source',()=>{
 const data=grid(true,true),before=JSON.stringify(data),book=workbook(data,'Analysis FR');context.main(book);
 const out=book.active;assert.equal(out.name,'Charts FR Auto');assert.equal(out.charts.length,1);assert.equal(JSON.stringify(data),before);
 assert.match(out.cells[3][27],/NUMBERVALUE/);assert.ok(out.cells[3][27].includes("'Analysis FR'!B3"));assert.ok(out.cells[3][27].includes("'Analysis FR'!C3"));
 assert.equal(out.charts[0].series[0].values.sheet,out);assert.equal(out.charts[0].series[0].categories.sheet,out);
 book.active=book.sheets[0];context.main(book);assert.equal(book.active.name,'Charts FR Auto 2');
});
test('invalid data fails before creating a sheet and identifies the row',()=>{
 const data=grid();data[2][3]='broken';const book=workbook(data);assert.throws(()=>context.main(book),/source row 3/);assert.equal(book.sheets.length,1);
 data[2][3]=0.9;assert.throws(()=>context.main(book),/does not match/);assert.equal(book.sheets.length,1);
});
test('all four chart types execute; no overlap; valid 0–100% axes',()=>{
 const data=grid();data.push(['Q1 — Extent','','',''],['Category','Institutions','Base (n)','Year'],...['Great','Moderate','Small','None','N/A'].map(x=>[x,20,100,0.2]),['','','',''],['Q2 — Processes','','',''],['Category','Institutions','Base (n)','Year'],['Recruitment',80,100,.8],['Retention',70,100,.7],['','','',''],['Q5 — Areas','','',''],['Category','Institutions','Base (n)','Year'],['Employment',50,100,.5],['Education',60,100,.6]);
 const book=workbook(data);context.main(book);assert.deepEqual(book.active.charts.map(c=>c.type),['pie','stacked','column','bar']);
 book.active.charts.forEach((c,i)=>{if(i)assert.ok(c.top>=book.active.charts[i-1].top+book.active.charts[i-1].height);if(c.type!=='pie'){assert.equal(c.min,0);assert.equal(c.max,1);}});
 assert.equal(book.active.charts[1].series.length,5);
});
// Optional local real-data validation. The supplied file is never published.
if(process.argv[2])test('all 33 real tables and chart branches, with French text variants',()=>{
 const data=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
 for(const variant of [data.grid,data.grid.map(row=>row.map(v=>typeof v==='number'?String(v).replace('.',','):v==='Category'?'Catégorie':v==='Yes'?'Oui':v==='No'?'Non':v))]){
  const blocks=context.findQuestions(variant,0,0);assert.equal(blocks.length,33);
  blocks.forEach((b,i)=>{const expected=data.tables[i].entries.filter(e=>!e[0].startsWith('Any '));assert.equal(b.labels.length,expected.length);b.percentages.forEach((v,j)=>assert.ok(Math.abs(v-expected[j][1]/data.tables[i].base)<1e-12));});
  const book=workbook(variant,'Analysis FR');context.main(book);assert.equal(book.active.charts.length,33);
 }
});
console.log(`PASS: Microsoft API type check and ${passed} test groups. Simulated workbook execution, not a live Excel runtime.`);
