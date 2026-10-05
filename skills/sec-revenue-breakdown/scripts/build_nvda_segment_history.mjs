import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';

const repoArg=process.argv.indexOf('--repo');
const repo=repoArg>=0?path.resolve(process.argv[repoArg+1]):path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const out=path.join(repo,'output','NVDA_segment_history');
const v=JSON.parse(await fs.readFile(path.join(out,'views.json'),'utf8'));
const checks=JSON.parse(await fs.readFile(path.join(out,'checks.json'),'utf8'));
const pack=JSON.parse(await fs.readFile(path.join(repo,'output','NVDA_history_sources','manifest.json'),'utf8'));
const raw=JSON.parse(await fs.readFile(path.join(out,'breakdown.json'),'utf8'));
const wb=Workbook.create();
const history=wb.worksheets.add('Quarterly History');
const annual=wb.worksheets.add('Annual History');
const coverage=wb.worksheets.add('Coverage');
const sources=wb.worksheets.add('Sources');
const positions=[];
const col=n=>{let s='';for(;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const fmt='#,##0.000;[Red](#,##0.000);"–"';

function band(sheet,row,last,text,inactive=false){
 sheet.getRange(`B${row}:${last}${row}`).format={fill:inactive?'#ECE7DF':'#DDEAF2',font:{bold:true,color:'#16324F'},rowHeight:text.length>70?48:30};
 sheet.getRange(`B${row}`).values=[[text]];
 sheet.getRange(`B${row}`).format.wrapText=true;
}

function build(sheet,periods,isAnnual=false){
 const last=col(periods.length+2),rows=[];
 sheet.getRange('B2').values=[['NVDA 營收拆分'+(isAnnual?' — 年度':' — 季度歷史')]];
 sheet.getRange('B2').format.font={name:'Arial',size:14,bold:true,color:'#16324F'};
 sheet.mergeCells(`B3:${last}3`);
 sheet.getRange('B3').values=[['USD million。最新分類在上，Inactive Segments 在下。空白表示未取得同口徑數值；新舊分類不可相加。']];
 sheet.getRange(`B3:${last}3`).format={font:{name:'Arial',size:10,italic:true,color:'#526174'},rowHeight:30,wrapText:true};
 sheet.getRange(`B5:${last}5`).values=[['分類',...periods.map(p=>p==='FY1998Transition'?'FY1998 (1 month)':p)]];
 sheet.getRange(`B5:${last}5`).format={fill:'#16324F',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:28};
 let row=7;
 const blocks=v.blocks;
 for(const b of blocks){
  if(isAnnual&&!b.rows.some(r=>periods.some(p=>r.annual_values[p]!==undefined)))continue;
  band(sheet,row,last,b.title,b.status==='inactive');row++;
  const starts={};
  for(const r of b.rows){
   if(!periods.some(p=>(isAnnual?r.annual_values:r.values)[p]!==undefined)&&b.status==='inactive')continue;
   const values=isAnnual?r.annual_values:r.values;
   sheet.getRange(`B${row}:${last}${row}`).values=[[r.parent?'    '+r.node:r.node,...periods.map(p=>values[p]??null)]];
   sheet.getRange(`C${row}:${last}${row}`).format.numberFormat=fmt;
   starts[r.node]=row;
   positions.push({sheet:sheet.name,row,node:r.node,version:b.version_id,status:b.status,periods,records:isAnnual?r.annual_records:r.records});
   row++;
   const opValues=isAnnual?r.annual_op_values:r.op_values;
   if(b.dimension==='reportable_segment'&&Object.keys(opValues).length){
    const revenueRow=row-1;
    sheet.getRange(`B${row}:${last}${row}`).values=[['    Operating income',...periods.map(p=>opValues[p]??null)]];
    positions.push({sheet:sheet.name,row,node:r.node+' / Operating income',version:b.version_id,status:b.status,periods,records:isAnnual?r.annual_op_records:r.op_records});
    sheet.getRange(`C${row}:${last}${row}`).format.numberFormat=fmt;row++;
    sheet.getRange(`B${row}`).values=[['    OP%']];
    for(let i=0;i<periods.length;i++){
     const c=col(i+3);
     sheet.getRange(`${c}${row}`).formulas=[[`=IF(COUNT(${c}${row-1},${c}${revenueRow})=2,IF(${c}${revenueRow}=0,"",${c}${row-1}/${c}${revenueRow}),"")`]];
    }
    sheet.getRange(`B${row}:${last}${row}`).format.font={italic:true};
    sheet.getRange(`C${row}:${last}${row}`).format.numberFormat='0.00%;[Red](0.00%);"–"';row++;
   }
  }
  if(b.dimension==='market_platform'&&b.status==='active'){
   sheet.getRange(`B${row}:${last}${row}`).values=[['公司總營收（官方揭露）',...periods.map(p=>v.consolidated[p]??null)]];
   sheet.getRange(`B${row}:${last}${row}`).format={font:{bold:true},numberFormat:fmt};
   const totalRow=row++;let testRow=row;
   sheet.getRange(`B${row}`).values=[['Data Center + Edge − 公司總營收']];
   for(let i=0;i<periods.length;i++){
    const c=col(i+3),a=starts['Data Center'],e=starts['Edge Computing'];
    if(a&&e)sheet.getRange(`${c}${row}`).formulas=[[`=IF(COUNT(${c}${a},${c}${e},${c}${totalRow})=3,SUM(${c}${a},${c}${e})-${c}${totalRow},"")`]];
   }
   sheet.getRange(`C${row}:${last}${row}`).format.numberFormat='0.00';
   sheet.getRange(`C${row}:${last}${row}`).conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:'#FCE3E3',font:{color:'#B42318'}}});row+=2;
  }
  row+=2;
 }
 sheet.getRange(`B6:${last}${row}`).format.font.name='Arial';
 sheet.getRange(`B6:${last}${row}`).format.font.size=10;
 sheet.getRange(`B6:${last}${row}`).format.rowHeight=23;
 // Restore section heights after common body formatting.
 for(let r=7;r<=row;r++)if(sheet.getRange(`B${r}`).values[0]?.[0]?.startsWith('Inactive Segments'))sheet.getRange(`B${r}`).format.rowHeight=44;
 sheet.getRange(`B1:B${row}`).format.columnWidth=55;
 sheet.getRange(`C1:${last}${row}`).format.columnWidth=isAnnual?17:15;
 sheet.freezePanes.freezeRows(5);sheet.freezePanes.freezeColumns(2);
 sheet.showGridLines=false;
}

build(history,v.periods);
build(annual,v.annual_periods,true);

const summaries=[
 ['項目','結果'],
 ['官方來源',`${pack.sources.length} 檔案／${new Set(pack.sources.map(s=>s.accession)).size} 申報；原始申報日期 1999–2026。`],
 ['抽取範圍','季度檢查 FY1999 Q1–FY2027 Q2；早期 Product / Royalty 原文另列，1998 是一個月過渡期間。'],
 ['季度完整性','下表逐期列出。未取得可用拆分不等於公司未披露，不能宣稱全部歷史已補完。'],
 ['最新市場分類','Hyperscale / ACIE 使用 FY2027 Q2 的原報及重編比較值。其他舊期不推估。'],
 ['Q1 原報保留','Hyperscale 37,869、ACIE 37,377 放 Inactive Segments；最新分類的 Q1 是 43,050、32,196。'],
 ['分類停用','只代表舊披露方式停用，不代表業務消失。不同維度、新舊分類不可相加。'],
 ['報告部門 OP%','依該部門 OP／營收；不是市場 OP%。分部 OP 合計可能有未分攤費用，不能當公司 GAAP OP。'],
 ['Q4 推導','僅使用同年度、同分類、相容來源的全年減九個月累計。缺資料留空。'],
 ['歷史原文','來源表格與資料集保留所有比較期及分類版本，未通過核對的本地資料不填入主表。'],
 ['年度轉換','1998 是一個月過渡期間，不與一般全年直接比較；1996/1997 原財年結束於 12 月。'],
 ['數值核對',`${checks.reconciliations.filter(c=>c.status==='pass').length} 組通過；${checks.reconciliations.filter(c=>c.status==='incomplete').length} 組本地分類不完整；${checks.reconciliations.filter(c=>c.status==='fail').length} 組失敗。`],
 ['來源定位','Sources 包含官方來源清單與完整揭露紀錄；數值已標準化為 USD millions。原始數字字串與註腳尚未逐筆擷取，請查原文件；Display cell 可追溯主表。'],
 ['尚待補查','早期季度的純文字與敘述型產品拆分，以及未解析表格，保留待補查，不以零替代。']
];
coverage.getRange(`B2:C${summaries.length+1}`).values=summaries;
coverage.getRange('B2:C2').format={fill:'#16324F',font:{bold:true,color:'#FFFFFF'}};
coverage.getRange(`B3:C${summaries.length+1}`).format={wrapText:true,rowHeight:42};
coverage.getRange('B1:B130').format.columnWidth=24;coverage.getRange('C1:C130').format.columnWidth=105;
coverage.getRange('B18:G18').values=[['季度','公司營收','市場分類','報告部門','地域分類','早期營收類型']];
coverage.getRange('B18:G18').format={fill:'#16324F',font:{bold:true,color:'#FFFFFF'}};
coverage.getRange(`B19:G${18+v.coverage.length}`).values=v.coverage.map(r=>[r.period,...['consolidated','market','reportable','geography','revenue_type'].map(k=>r[k]?'已取得':'未取得可用拆分')]);
coverage.getRange('C18:G134').format.columnWidth=23;
for(let r=2;r<=summaries.length+1;r++)coverage.mergeCells(`C${r}:G${r}`);
coverage.getRange(`B19:G${18+v.coverage.length}`).format.rowHeight=23;
coverage.freezePanes.freezeRows(18);
coverage.getRange('B16:C16').values=[['Template version','Revenue breakdown workbook v1']];
coverage.mergeCells('C16:G16');
const registryRow=22+v.coverage.length;
coverage.getRange(`B${registryRow}:H${registryRow}`).values=[['Version ID','Dimension','Status','Original definition','First observed','Last observed','Recast / mapping evidence']];
coverage.getRange(`B${registryRow}:H${registryRow}`).format={fill:'#16324F',font:{bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:42};
coverage.getRange(`B${registryRow+1}:H${registryRow+v.blocks.length}`).values=v.blocks.map(b=>[b.version_id,b.dimension,b.status,b.title,b.first_observed,b.last_observed,[...new Set(b.rows.flatMap(r=>r.source_definitions))].join('; ')]);
coverage.getRange(`B${registryRow+1}:H${registryRow+v.blocks.length}`).format={wrapText:true,rowHeight:72};
sources.getRange('A1:F1').values=[['Source ID','Form','Document type','Filed','SEC official URL','SHA-256']];
sources.getRange(`A2:F${pack.sources.length+1}`).values=pack.sources.map(s=>[s.source_id,s.form,s.document_type,s.filing_date,s.url,s.sha256]);
sources.getRange('A1:F1').format={fill:'#16324F',font:{bold:true,color:'#FFFFFF'}};
sources.getRange(`A1:A${pack.sources.length+1}`).format.columnWidth=60;
sources.getRange(`B1:D${pack.sources.length+1}`).format.columnWidth=16;
sources.getRange(`E1:F${pack.sources.length+1}`).format.columnWidth=95;
sources.freezePanes.freezeRows(1);
const rawHeader=pack.sources.length+5;
const rawFields=['Record ID','Fiscal period','Period start','Period end','Duration','Metric','Original label','Parent','Dimension','Classification version','Presentation','Original value','Original unit','Currency','Normalized value','Normalized unit','Basis / geography definition','Filed / accession','Source ID','Source locator','Official URL','Method','Input record IDs','Verification','Original footnote / definition','Display cell'];
const displayCells=new Map();
for(const pos of positions)for(let i=0;i<pos.periods.length;i++){
 const rid=pos.records?.[pos.periods[i]];if(!rid)continue;
 const refs=displayCells.get(rid)??[];refs.push(`${pos.sheet}!${col(i+3)}${pos.row}`);displayCells.set(rid,refs);
}
sources.getRange(`A${rawHeader}:Z${rawHeader}`).values=[rawFields];
sources.getRange(`A${rawHeader}:Z${rawHeader}`).format={fill:'#16324F',font:{bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:42};
sources.getRange(`A${rawHeader+1}:Z${rawHeader+raw.records.length}`).values=raw.records.map(r=>[r.id,r.period,r.period_start,r.period_end,r.duration,r.metric,r.original_label,r.parent_id,r.dimension,r.classification_version,r.presentation,null,null,r.currency,r.value,r.unit,r.basis,`${r.filing_date} / ${r.published_in_accession}`,r.source.source_id,r.source.locator,r.source.url,r.derivation.method,JSON.stringify(r.derivation.inputs??[]),r.status,null,(displayCells.get(r.id)??[]).join('; ')]);
sources.getRange(`G1:Z${rawHeader+raw.records.length}`).format.columnWidth=28;
sources.getRange(`O${rawHeader+1}:O${rawHeader+raw.records.length}`).format.numberFormat=fmt;
wb.recalculate();
await fs.writeFile(path.join(out,'cell_sources.json'),JSON.stringify(positions,null,2));
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(path.join(out,'NVDA_Revenue_Segment_History.xlsx'));
const image=await wb.render({sheetName:'Quarterly History',range:'B2:J32',scale:1,format:'png'});
await fs.writeFile(path.join(out,'quarterly_preview.png'),new Uint8Array(await image.arrayBuffer()));
const coverageImage=await wb.render({sheetName:'Coverage',range:'B2:G15',scale:1,format:'png'});
await fs.writeFile(path.join(out,'coverage_preview.png'),new Uint8Array(await coverageImage.arrayBuffer()));
for(const name of ['Annual History','Sources']){
 const preview=await wb.render({sheetName:name,range:name==='Sources'?'A1:F9':'B2:J20',scale:1,format:'png'});
 await fs.writeFile(path.join(out,name.replaceAll(' ','_')+'_preview.png'),new Uint8Array(await preview.arrayBuffer()));
}
console.log('Exported NVDA_Revenue_Segment_History.xlsx');
