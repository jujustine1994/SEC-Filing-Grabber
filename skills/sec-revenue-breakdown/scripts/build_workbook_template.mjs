import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';

const out=path.resolve(process.argv[2]??'output/revenue_breakdown_template');
await fs.mkdir(out,{recursive:true});
const wb=Workbook.create();
const names=['Quarterly History','Annual History','Coverage','Sources'];
const sheets=names.map(n=>wb.worksheets.add(n));
const header={fill:'#16324F',font:{name:'Arial',bold:true,color:'#FFFFFF'},rowHeight:30,wrapText:true};
for(const sheet of sheets){
 sheet.showGridLines=false;
 sheet.getRange('A1:Z60').format.font={name:'Arial',size:10};
 sheet.getRange('B1:B60').format.columnWidth=48;
 sheet.getRange('C1:H60').format.columnWidth=20;
}
for(const [index,sheet] of sheets.slice(0,2).entries()){
 sheet.getRange('B2').values=[['<Company / ticker> Revenue breakdown — '+(index?'Annual':'Quarterly')]];
 sheet.getRange('B2').format.font={bold:true,size:14,color:'#16324F'};
 sheet.mergeCells('B3:H3');
 sheet.getRange('B3').values=[['填入來源數值；指定幣別／單位。空白＝未取得，0＝官方零值。原報與重編分開保留。']];
 sheet.getRange('B3:H3').format={wrapText:true,rowHeight:42,font:{color:'#526174'}};
 sheet.getRange('B5:H5').values=[['Classification / metric','<Period 1>','<Period 2>','<Period 3>','<Period 4>','<Period 5>','<Period 6>']];
 sheet.getRange('B5:H5').format=header;
 const labels={7:'Latest classification — <Dimension / version>',8:'<Category> — Revenue',9:'    Reported operating income (if disclosed)',10:'    OP% (simple calculation)',11:'    External revenue (if disclosed)',12:'    Intersegment revenue (if disclosed)',14:'<Child category> — Revenue',16:'Company total revenue (official)',18:'Other disclosed metrics — <exact metric / unit>',19:'<Reported gross profit / cost / KPI, if relevant>',22:'Inactive Segments — <Dimension / old version>',23:'<Old category> — Revenue',24:'    Reported operating income (if disclosed)',25:'    External revenue (if disclosed)',26:'    Intersegment revenue (if disclosed)'};
 for(const [row,label] of Object.entries(labels))sheet.getRange(`B${row}`).values=[[label]];
 for(const r of [7,18,22])sheet.getRange(`B${r}:H${r}`).format={fill:r===22?'#ECE7DF':'#DDEAF2',font:{bold:true,color:'#16324F'},rowHeight:30};
 sheet.getRange('B8:H26').format.rowHeight=26;
 sheet.getRange('B8:B26').format.wrapText=true;
 sheet.getRange('C8:H26').format.numberFormat='#,##0.000;[Red](#,##0.000);"–"';
 sheet.getRange('C10:H10').formulas=[['C','D','E','F','G','H'].map(c=>`=IF(COUNT(${c}8,${c}9)=2,IF(${c}8=0,"",${c}9/${c}8),"")`)];
 sheet.getRange('C10:H10').format.numberFormat='0.00%;[Red](0.00%);"–"';
 sheet.freezePanes.freezeRows(5);sheet.freezePanes.freezeColumns(2);
}
const coverage=sheets[2];
coverage.getRange('B2:C2').values=[['Template version','Revenue breakdown workbook v1']];
coverage.getRange('B2:H2').format=header;
const rows=[['Company / ticker / CIK',null],['Currency / normalized unit',null],['Financial periods / filing cutoff',null],['Classification and geography definitions',null],['Original / official recast policy','使用公司公布的重編比較值；保留原報，不按比例回推。'],['Calculated values policy','預設僅 OP%；Q4 全年減九個月若使用，必須另標 derived 及來源。'],['Missing-data policy','未取得留空；來源未披露、尚未解析、核對失敗須分別說明。'],['Raw-value policy','保留原標籤、原單位與原始值；標準化值另列。'],['Coverage / limitations',null],['Source integrity / reconciliation',null],['Extracted at / tool version',null]];
coverage.getRange('B3:C13').values=rows;
for(let r=3;r<=13;r++)coverage.mergeCells(`C${r}:H${r}`);
coverage.getRange('B3:H13').format={wrapText:true,rowHeight:42};
coverage.getRange('B16:H16').values=[['Period','Consolidated','Product / market','Reportable segment','Geography','Other dimensions','Gap / reason']];
coverage.getRange('B16:H16').format=header;
coverage.getRange('B21:H21').values=[['Version ID','Dimension','Status','Original definition','First observed','Last observed','Recast / mapping evidence']];
coverage.getRange('B21:H21').format=header;
coverage.freezePanes.freezeRows(2);
const sources=sheets[3];
sources.getRange('A1:F1').values=[['Source ID','Form','Document type','Filed','SEC official URL','SHA-256']];
sources.getRange('A1:F1').format=header;
sources.getRange('A6:Z6').values=[['Record ID','Fiscal period','Period start','Period end','Duration','Metric','Original label','Parent','Dimension','Classification version','Presentation','Original value','Original unit','Currency','Normalized value','Normalized unit','Basis / geography definition','Filed / accession','Source ID','Source locator','Official URL','Method','Input record IDs','Verification','Original footnote / definition','Display cell']];
sources.getRange('A6:Z6').format=header;
sources.getRange('A1:Z30').format.columnWidth=25;
sources.getRange('E1:E30').format.columnWidth=65;
sources.freezePanes.freezeRows(1);
wb.recalculate();
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(path.join(out,'Revenue_Breakdown_Template.xlsx'));
for(const name of names){
 const image=await wb.render({sheetName:name,range:name==='Sources'?'A1:F8':name==='Coverage'?'B2:H13':'B2:H26',scale:1,format:'png'});
 await fs.writeFile(path.join(out,name.replaceAll(' ','_')+'_preview.png'),new Uint8Array(await image.arrayBuffer()));
}
console.log('Exported Revenue_Breakdown_Template.xlsx');
