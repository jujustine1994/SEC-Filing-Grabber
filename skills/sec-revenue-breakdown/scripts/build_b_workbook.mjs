/** Fixed B presentation engine. Classifications/values come from a validated
 * Python plan; templates own layout, never company facts. No dependency installs.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';

const args=process.argv.slice(2),opt=k=>args[args.indexOf(k)+1];
const names=['Quarterly History','Annual History','Coverage','Sources'];
const rawFields=['Record ID','Fiscal period','Period start','Period end','Duration','Metric','Original label','Parent','Dimension','Classification version','Presentation','Original value','Original unit','Currency','Normalized value','Normalized unit','Basis / geography definition','Filed / accession','Source ID','Source locator','Official URL','Method','Input record IDs','Verification','Original footnote / definition','Display cell'];
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const number='#,##0.000;[Red](#,##0.000);0.000';
const percent='0.0%;[Red](0.0%);0.0%';
const header={fill:'#17324D',font:{name:'Microsoft JhengHei',size:12,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:30};
const safe=v=>typeof v==='string'&&/^[=+@]/.test(v)?"'"+v:v;
function layout(s,lastRow,lastCol){
 s.showGridLines=true;
 s.getRange(`A1:${col(lastCol)}${lastRow}`).format.font={name:'Microsoft JhengHei',size:12};
 s.getRange(`A1:A${lastRow}`).format.columnWidth=4;
 s.getRange(`B1:B${lastRow}`).format.columnWidth=s.name==='Sources'?20:58;
 if(lastCol>=3)s.getRange(`C1:${col(lastCol)}${lastRow}`).format.columnWidth=s.name==='Sources'?24:18;
 s.getRange(`A1:${col(lastCol)}${lastRow}`).format.rowHeight=23;
 if(names.slice(0,2).includes(s.name)){
  s.getRange('B2').format.font={name:'Microsoft JhengHei',size:16,bold:true,color:'#17324D'};
  s.getRange(`B3:${col(Math.max(lastCol,8))}3`).format={wrapText:true,rowHeight:48};
  s.getRange(`B5:${col(Math.max(lastCol,8))}5`).format=header;
  s.getRange(`C6:${col(Math.max(lastCol,8))}${Math.max(lastRow,7)}`).format.numberFormat=number;
  s.getRange(`B6:B${Math.max(lastRow,7)}`).format.wrapText=true;
  s.freezePanes.freezeRows(5);s.freezePanes.freezeColumns(2);
 }else if(s.name==='Coverage'){
  s.getRange('B2:H2').format=header;
  s.getRange(`C3:C${Math.max(lastRow,3)}`).format.columnWidth=90;
  s.getRange(`B3:H${Math.max(lastRow,3)}`).format={wrapText:true,rowHeight:48};
 }else{
  s.getRange('A1:F1').format=header;
  s.getRange(`E1:E${lastRow}`).format.columnWidth=55;
  s.freezePanes.freezeRows(1);
 }
}
async function save(wb,out,previewSheet='Quarterly History'){
 await fs.mkdir(path.dirname(out),{recursive:true});
 try{await fs.access(out);throw Error('Refuse overwrite: '+out);}catch(e){if(e.code!=='ENOENT')throw e;}
 wb.recalculate();
 const exported=await SpreadsheetFile.exportXlsx(wb);await exported.save(out);
 if(args.includes('--preview')){
  const preview=await wb.render({sheetName:previewSheet,range:'B2:H32',scale:1.25,format:'png'});
  await fs.writeFile(out.replace(/\.xlsx$/i,'_preview.png'),new Uint8Array(await preview.arrayBuffer()));
 }
}
async function templates(dir){
 const catalog={schema_version:'b-template-catalog-1',version:1,templates:[]};
 for(const [id,title] of [['flat','平面拆分'],['hierarchy','父子階層'],['segment-op','報告部門與利益']]){
  const wb=Workbook.create();for(const n of names)wb.worksheets.add(n);
  for(const n of names){const s=wb.worksheets.getItem(n);layout(s,50,n==='Sources'?26:8);}
  for(const [i,n] of names.slice(0,2).entries()){
   const s=wb.worksheets.getItem(n);
   s.getRange('B2').values=[['<Ticker / Company> — '+title]];
   s.getRange('B3').values=[['空白版型；只填官方來源。比重分母為公司總營收；QoQ／YoY需同版本同口徑。']];
   s.getRange('B5:H5').values=[['分類／指標','<Period 1>','<Period 2>','<Period 3>','<Period 4>','<Period 5>','<Period 6>']];
   s.getRange('B6').values=[['公司總營收（官方揭露）']];
   s.getRange('B8').values=[['最新分類 — <Dimension / version>']];s.getRange('B8:H8').format={fill:'#DDEAF2',font:{bold:true}};
   const labels=['<Category>','    占公司總營收',...(i?[]:['    QoQ']),'    YoY',...(id==='segment-op'?['    營業利益（有披露才使用）','    營業利益率']:[]),...(id==='hierarchy'?['    <Child category>','        占公司總營收']:[])];
   for(let j=0;j<labels.length;j++)s.getRange(`B${9+j}`).values=[[labels[j]]];
   for(let c=3;c<=8;c++){
    const l=col(c),totalSource=c+4,categorySource=c+10;s.getRange(`${l}6`).formulas=[[`=IF(ISNUMBER(Sources!$O$${totalSource}),Sources!$O$${totalSource},"")`]];
    s.getRange(`${l}9`).formulas=[[`=IF(ISNUMBER(Sources!$O$${categorySource}),Sources!$O$${categorySource},"")`]];
    s.getRange(`${l}10`).formulas=[[`=IF(COUNT(${l}9,${l}$6)=2,IF(${l}$6>0,${l}9/${l}$6,""),"")`]];
   }
   s.getRange('C10:H10').format.numberFormat=percent;
   s.getRange('B25').values=[['Inactive Segments — <Dimension / old version>']];s.getRange('B25:H25').format={fill:'#ECE7DF',font:{bold:true}};
  }
  const cov=wb.worksheets.getItem('Coverage');cov.getRange('B2:C2').values=[['B template',`${id} / formula-v1`]];
  cov.getRange('B3:C6').values=[['用途',title],['數字來源','空白模板沒有公司數字；使用 builder 填入已核對 records。'],['公式示意','公式位置只示範來源與比重；真實 period／record／growth 由 planner 產生，不直接填此骨架當完整成果。'],['維護','同類披露結構可共用；不能把示範分類當公司資料。']];
  const src=wb.worksheets.getItem('Sources');src.getRange('A1:F1').values=[['Source ID','Form','Document type','Filed','SEC official URL','SHA-256']];src.getRange('A6:Z6').values=[rawFields];src.getRange('A6:Z6').format=header;
  const file=`Revenue_Breakdown_${id}_formula_v1.xlsx`;await save(wb,path.join(dir,file));
  catalog.templates.push({id,file,version:1,description:title,sha256:crypto.createHash('sha256').update(await fs.readFile(path.join(dir,file))).digest('hex')});
 }
 await fs.writeFile(path.join(dir,'catalog.json'),JSON.stringify(catalog,null,2));
}
async function build(planFile,template,out){
 const p=JSON.parse(await fs.readFile(planFile,'utf8'));
 if(p.schema_version!=='b-formula-1')throw Error('Unknown plan schema');
 const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(template));
 const recorded=wb.worksheets.getItem('Coverage').getRange('C2').values[0][0];
 if(recorded!==`${p.template_id} / formula-v1`)throw Error('Template ID/version mismatch');
 for(const n of names){
  const s=wb.worksheets.getItem(n),sp=p.sheets[n];
  // Clear only content of the small template scaffold; preserve styles.
  s.getRange('A1:Z50').clear({applyTo:'contents'});
  // Example bands are scaffolding, not facts. Remove their body formatting
  // so a template's inactive example cannot paint an unrelated data row.
  s.getRange('B6:Z50').clear({applyTo:'formats'});
  const matrix=Array.from({length:sp.last_row},()=>Array(sp.last_col).fill(null));
  const formulas=[];
  for(const c of sp.cells){const m=c.address.match(/^([A-Z]+)(\d+)$/);let x=0;for(const a of m[1])x=x*26+a.charCodeAt(0)-64;
   if(c.formula){formulas.push(c);if(args.includes('--defer-calculation'))matrix[Number(m[2])-1][x-1]="'"+c.formula;}else matrix[Number(m[2])-1][x-1]=safe(c.value);
  }
  s.getRange(`A1:${col(sp.last_col)}${sp.last_row}`).values=matrix;
  layout(s,sp.last_row,sp.last_col);
  const rows=new Map();
  for(const c of formulas){const m=c.address.match(/^([A-Z]+)(\d+)$/);let x=0;for(const a of m[1])x=x*26+a.charCodeAt(0)-64;
   const row=rows.get(m[2])??[];row.push({col:x,formula:c.formula});rows.set(m[2],row);
  }
  if(!args.includes('--defer-calculation'))for(const [r,cells] of rows){const first=Math.min(...cells.map(c=>c.col)),last=Math.max(...cells.map(c=>c.col));const matrix=Array(last-first+1).fill('');for(const c of cells)matrix[c.col-first]=c.formula;s.getRange(`${col(first)}${r}:${col(last)}${r}`).formulas=[matrix];}
  const percentageRows=new Set(sp.cells.filter(c=>['share','qoq','yoy','op_margin','net_margin'].includes(c.kind)).map(c=>c.address.match(/\d+/)[0]));
  for(const r of percentageRows)s.getRange(`C${r}:${col(sp.last_col)}${r}`).format.numberFormat=percent;
  for(const b of sp.bands)s.getRange(`B${b.row}:${col(sp.last_col)}${b.row}`).format={fill:b.status==='inactive'?'#ECE7DF':'#DDEAF2',font:{bold:true,color:'#17324D'},wrapText:true,rowHeight:34};
  if(n==='Sources'){
   s.getRange(`A${p.raw_header}:Z${p.raw_header}`).format=header;
   s.getRange(`O${p.raw_header+1}:O${sp.last_row}`).format.numberFormat=number;
  }
 }
 await save(wb,out);console.log('BUILT',out);
}
if(args.includes('--templates'))await templates(path.resolve(opt('--templates')));
else if(args.includes('--plan')&&args.includes('--template')&&args.includes('--out'))await build(opt('--plan'),opt('--template'),opt('--out'));
else throw Error('Use --templates DIR or --plan JSON --template XLSX --out XLSX');
