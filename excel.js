(function(root){
'use strict';
const xml=s=>String(s??'').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
const parse=s=>{const doc=new DOMParser().parseFromString(s,'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('文件中的XML格式损坏');return doc;};
const els=(node,tag)=>Array.from(node.getElementsByTagNameNS('*',tag));
const first=(node,tag)=>els(node,tag)[0];
const col=i=>{let s='';for(i++;i;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s;};
const colIndex=s=>[...s.replace(/\d/g,'')].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;
const resolve=p=>{const out=[];for(const s of p.split('/')){if(s==='..')out.pop();else if(s&&s!=='.')out.push(s);}return out.join('/');};
async function readXlsx(buffer,source){
 const zip=await JSZip.loadAsync(buffer);
 if(!zip.file('xl/workbook.xml'))throw Error(`${source} 不是有效的xlsx工作簿，请将xls文件另存为xlsx`);
 const load=async p=>{const entry=zip.file(p);return entry?parse(await entry.async('string')):null;};
 const book=await load('xl/workbook.xml'),rels=await load('xl/_rels/workbook.xml.rels');
 const relationships=Object.fromEntries(els(rels,'Relationship').map(n=>[n.getAttribute('Id'),n.getAttribute('Target')]));
 const sharedDoc=await load('xl/sharedStrings.xml');
 const strings=sharedDoc?els(sharedDoc,'si').map(n=>els(n,'t').map(t=>t.textContent).join('')):[];
 const styles=await load('xl/styles.xml');
 const formats=styles?Object.fromEntries(els(styles,'numFmt').map(n=>[n.getAttribute('numFmtId'),n.getAttribute('formatCode')])):{};
 const xfs=styles?Array.from(first(styles,'cellXfs')?.children||[]):[];
 const date1904=first(book,'workbookPr')?.getAttribute('date1904')==='1';
 const sheets=[];
 for(const sh of els(book,'sheet')){
  const rid=sh.getAttribute('r:id')||sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
  const target=relationships[rid];if(!target)continue;
  const path=resolve(target.startsWith('/')?target.slice(1):`xl/${target}`),doc=await load(path);if(!doc)continue;
  const grid=[],formulas=new Set();
  for(const row of els(doc,'row')){
   const number=Number(row.getAttribute('r'))||grid.length+1;const values=[];values._row=number;
   for(const c of Array.from(row.children).filter(n=>n.localName==='c')){
    const ci=colIndex(c.getAttribute('r')||`A${number}`);const type=c.getAttribute('t');
    const v=first(c,'v')?.textContent??'';const f=first(c,'f');if(f)formulas.add(`${number}:${ci}`);
    let value=type==='s'?strings[Number(v)]??'':type==='inlineStr'?els(c,'t').map(n=>n.textContent).join(''):type==='e'?'':v;
    if(type!=='s'&&type!=='inlineStr'&&type!=='e'&&v!==''){
     const fmtId=Number(xfs[Number(c.getAttribute('s')||0)]?.getAttribute('numFmtId')||0),fmt=formats[fmtId]||'';
     if(/^0+$/.test(fmt)&&/^[0-9]+$/.test(v))value=v.padStart(fmt.length,'0');
     else if(([14,15,16,17,18,19,20,21,22,45,46,47].includes(fmtId)||(/[ymd]/i.test(fmt.replace(/"[^"]*"/g,''))&&!/[Ee][+-]/.test(fmt)))&&Number.isFinite(Number(v))){
      const millis=(Number(v)+(date1904?1462:0)-25569)*86400000;const date=new Date(Math.round(millis));
      if(Number.isFinite(date.getTime()))value=date.toISOString().replace('T',' ').slice(0,/h|s/i.test(fmt)||[18,19,20,21,22,45,46,47].includes(fmtId)?19:10);
     }
    }
    values[ci]=value;
   }
   if(values.some(v=>v!==''&&v!=null))grid.push(values);
  }
  const hi=VCore.findHeader(grid);if(hi<0)continue;
  const headers=Array.from({length:grid[hi].length},(_,i)=>String(grid[hi][i]??''));
  const mapping=VCore.mapping(headers);
  // Ignore blank helper sheets, retain identity sheets for manual mapping.
  if(mapping.name<0&&mapping.id<0)continue;
  sheets.push({source:`${source} / ${sh.getAttribute('name')}`,headers,rows:grid.slice(hi+1),formulas,mapping,enabled:true});
 }
 return sheets;
}
async function readFiles(files,progress){
 const all=[];let completed=0;
 for(const file of files){
  if(file.size>60*1024*1024)throw Error(`${file.name} 超过60MB，请分批导入`);
  const bytes=await file.arrayBuffer();
  if(/\.zip$/i.test(file.name)){
   const zip=await JSZip.loadAsync(bytes);const entries=Object.values(zip.files).filter(e=>!e.dir&&/\.xlsx$/i.test(e.name)&&!e.name.includes('__MACOSX')&&!e.name.split('/').at(-1).startsWith('~$'));
   if(!entries.length)throw Error(`${file.name} 中未找到xlsx评分表`);
   for(const entry of entries){all.push(...await readXlsx(await entry.async('arraybuffer'),`${file.name} / ${entry.name}`));completed++;progress?.(`已读取 ${completed} 份工作簿`);}
  }else if(/\.xlsx$/i.test(file.name)){all.push(...await readXlsx(bytes,file.name));completed++;progress?.(`已读取 ${completed} 份工作簿`);}
  else throw Error(`${file.name} 格式不支持，请使用xlsx或含xlsx的ZIP`);
 }
 return all;
}
const charWidth=s=>[...String(s??'')].reduce((n,c)=>n+(c.charCodeAt(0)>255?2:1),0);
function exportColumns(result,base,full=true){
 const reserved=['普通面试归一化成绩','面试原始总分（15分）','闪电得分（10分）','摄影得分（10分）','面试评分来源','核对备注','志愿者类别'];
 const core=['姓名','学号'];
 const fields=full?[...core,...base.fields.filter(x=>!core.includes(x)&&!reserved.includes(x)&&x!=='面试总分（15分）')]:['姓名','学号','性别','学院/单位','班级','手机号','邮箱'];
 const headers=[...fields,...reserved];
 const rows=result.selected.map(r=>[...fields.map(k=>k==='学号'?r.id:k==='姓名'?r.name:r.data[k]??''),r.normalized,r.general,r.lightning,r.photo,r.scoreSource,[...r.notes].join('；'),r.role]);
 return {headers,rows};
}
async function makeXlsx(headers,rows,{fontSize=11,rowHeight=28,maxWidth=44,sheetName='志愿者总名单'}={}){
 if(!rows.length)throw Error('没有可导出的人员');
 const widths=headers.map((h,i)=>{
  if(h==='姓名')return 14;if(/学号/.test(h))return 19;if(/证件号/.test(h))return 25;if(/手机号|QQ号/.test(h))return 20;if(/得分|总分/.test(h))return 18;if(h==='志愿者类别')return 20;
  const dataMax=rows.reduce((n,r)=>Math.max(n,...String(r[i]??'').split('\n').map(charWidth)),0);
  return Math.max(16,Math.min(maxWidth,Math.max(Math.min(charWidth(h)+2,28),dataMax+2)));
 });
 const lineHeight=fontSize*1.4;
 const heightFor=row=>Math.min(409,Math.max(rowHeight,...row.map((v,i)=>String(v??'').split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil(charWidth(line)/(widths[i]-2))),0)*lineHeight+8)));
 const heights=[Math.max(40,heightFor(headers)),...rows.map(heightFor)];
 const last=`${col(headers.length-1)}${rows.length+1}`;
 const header='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 const sheet=header+`<worksheet xmlns="${ns}"><dimension ref="A1:${last}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C2" sqref="C2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="${rowHeight}"/><cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>`+
 [headers,...rows].map((row,ri)=>`<row r="${ri+1}" ht="${heights[ri]}" customHeight="1">`+row.map((v,ci)=>{
  const ref=`${col(ci)}${ri+1}`;const numeric=ri>0&&typeof v==='number'&&Number.isFinite(v);const style=ri===0?1:numeric?3:2;
  return numeric?`<c r="${ref}" s="${style}"><v>${v}</v></c>`:`<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xml(v)}</t></is></c>`;
 }).join('')+'</row>').join('')+`</sheetData><autoFilter ref="A1:${last}"/><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup orientation="landscape" paperSize="9"/></worksheet>`;
 const styles=header+`<styleSheet xmlns="${ns}"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.######"/></numFmts><fonts count="2"><font><sz val="${fontSize}"/><color rgb="FF111111"/><name val="Arial"/></font><font><b/><sz val="${fontSize}"/><color rgb="FF111111"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="0" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
 const zip=new JSZip();
 zip.file('[Content_Types].xml',header+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
 zip.file('_rels/.rels',header+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
 zip.file('xl/workbook.xml',header+`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="${xml(sheetName)}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`);
 zip.file('xl/_rels/workbook.xml.rels',header+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
 zip.file('xl/styles.xml',styles);zip.file('xl/worksheets/sheet1.xml',sheet);
 return zip.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
root.VExcel={readXlsx,readFiles,exportColumns,makeXlsx,charWidth};
})(typeof window==='undefined'?globalThis:window);
