(function(root){
'use strict';
const text=v=>v==null?'':String(v).trim();
const num=v=>{const s=text(v);return s!==''&&/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(s)&&Number.isFinite(Number(s))?Number(s):null;};
const clean=v=>text(v).replace(/^Q\s*\d+\s*[.．、:：]?\s*/i,'').replace(/\s+/g,'');
function field(h){
 const s=clean(h);
 if(/^(您的)?学号$/.test(s))return '学号';
 if(/^(您的)?姓名$/.test(s))return '姓名';
 if(/姓名拼音/.test(s))return '姓名拼音';
 if(/手机号/.test(s))return '手机号';
 if(/证件号/.test(s))return '证件号';
 if(/出生日期/.test(s))return '出生日期';
 if(/学院\/单位/.test(s))return '学院/单位';
 if(/您的邮箱/.test(s))return '邮箱';
 if(/您的性别/.test(s))return '性别';
 if(/您的班级/.test(s))return '班级';
 if(/您的年级/.test(s))return '年级';
 if(/所在的校区/.test(s))return '校区';
 if(/身高/.test(s))return '身高（cm）';
 if(/衣服尺码/.test(s))return '衣服尺码';
 if(/是否(选择|申请)线上面试/.test(s))return '是否线上面试';
 if(/方便参加面试的时间段/.test(s))return '面试时间段';
 if(/您的学生组织/.test(s))return '学生组织';
 if(/是否按时能参加/.test(s))return '能否按时参加';
 return s.replace(/^您的/,'')||'';
}
function findHeader(rows){
 let best=-1,rank=0;
 rows.slice(0,20).forEach((row,i)=>{let v=0;row.forEach(h=>{if(['姓名','学号'].includes(field(h)))v+=5;if(/总分|得分/.test(text(h)))v++;});if(v>rank){rank=v;best=i;}});
 return best;
}
function mapping(headers){
 const h=headers.map(clean);const idx=f=>h.findIndex(f);
 return {name:headers.findIndex(x=>field(x)==='姓名'),id:headers.findIndex(x=>field(x)==='学号'),
 general:idx(x=>/总分|普通.*得分|面试.*得分/.test(x)&&!/(归一|标准|闪电|摄影)/.test(x)),
 lightning:idx(x=>/闪电.*(得分|总分)|小闪电.*分数/.test(x)&&!/(归一|标准)/.test(x)),
 photo:(()=>{const score=idx(x=>/摄影.*(得分|总分)/.test(x)&&!/(归一|标准)/.test(x));return score>=0?score:idx(x=>/^摄影评价/.test(x));})()};
}
function numericScore(sheet,row,map,kind){
 const cell=map[kind]>=0?row[map[kind]]:null;
 let direct=num(cell);
 const patterns=kind==='general'?[/^活动感想/,/^活动思考/,/^应变能力/,/^情景对话.*[（(]3[）)]/,/^综合印象/]:kind==='lightning'?[/^工作理解/,/^情景对话.*[（(]6[）)]/] : [/^摄影评价/];
 const indexes=patterns.map(p=>sheet.headers.findIndex(h=>p.test(clean(h))));
 const parts=indexes.map(i=>i<0?null:num(row[i]));
 const any=parts.some(n=>n!==null),complete=parts.every(n=>n!==null);
 if(complete && parts.length)return {value:parts.reduce((a,b)=>a+b,0),mismatch:direct!==null&&Math.abs(direct-parts.reduce((a,b)=>a+b,0))>1e-7};
 if(direct!==null&&(direct!==0||any||!sheet.formulas?.has(`${row._row}:${map[kind]}`)))return {value:direct};
 return {value:null};
}
function prepare(infoSheets,scoreSheets){
 const people=new Map(),byName=new Map(),scores=new Map(),errors=[],fields=[],fileInfo=[];
 let ordinal=0;
 for(const s of infoSheets.filter(s=>s.enabled!==false)){
  const map=s.mapping||mapping(s.headers);if(map.name<0||map.id<0){errors.push(`${s.source}：请指定姓名、学号列`);continue;}
  const hs=s.headers.map(field);hs.forEach(h=>{if(h&&!fields.includes(h))fields.push(h);});
  for(const row of s.rows){
   const name=text(row[map.name]),id=text(row[map.id]);if(!name&&!id)continue;
   if(!name||!id){errors.push(`${s.source} 第${row._row}行缺少姓名或学号`);continue;}
   let p=people.get(id);const data=Object.fromEntries(hs.map((h,i)=>[h,row[i]??'']).filter(([h])=>h));
   data['姓名']=name;data['学号']=id;
   if(p){
    for(const [h,v] of Object.entries(data))if(text(v)&&text(p.data[h])&&text(v)!==text(p.data[h]))p.notes.add(`重复问卷“${h}”不同，保留首次值`);
    for(const [h,v] of Object.entries(data))if(!text(p.data[h]))p.data[h]=v;
    p.sources.push(s.source);
   }else{p={id,name,data,order:ordinal++,sources:[s.source],notes:new Set(),hasInfo:true};people.set(id,p);}
  }
 }
 for(const s of scoreSheets.filter(s=>s.enabled!==false)){
  const map=s.mapping||mapping(s.headers);if(map.name<0||map.id<0){errors.push(`${s.source}：请指定姓名、学号列`);continue;}
  if([map.general,map.lightning,map.photo].every(x=>x<0)){errors.push(`${s.source}：未指定任何得分列`);continue;}
  let count=0;
  for(const row of s.rows){
   const id=text(row[map.id]),name=text(row[map.name]);if(!name&&!id)continue;if(!name||!id)continue;
   const record={id,name,source:s.source,row:row._row,order:ordinal++,general:null,lightning:null,photo:null,notes:[]};
   for(const kind of ['general','lightning','photo']){
    if(map[kind]<0)continue;
    const svalue=numericScore(s,row,map,kind);record[kind]=svalue.value;
    if(svalue.mismatch)record.notes.push(`${kind==='general'?'普通面试':kind==='photo'?'摄影':'闪电'}总分与分项不符，已使用分项之和`);
    if(record[kind]!==null&&(record[kind]<0||record[kind]>(kind==='general'?15:10))){record.notes.push('分数超出常用范围，请核对');}
   }
   if([record.general,record.lightning,record.photo].every(x=>x===null))continue;
   if(!scores.has(id))scores.set(id,[]);scores.get(id).push(record);count++;
  }
  fileInfo.push({source:s.source,count});
 }
 for(const [id,recs] of scores){for(const name of new Set(recs.map(r=>r.name))){if(!byName.has(name))byName.set(name,[]);byName.get(name).push(id);}}
 return {people,scores,byName,fields,errors,fileInfo};
}
function evaluate(base,opts={},edits={}){
 const options={lightning:28,photo:14,ties:'all',priority:'lightning',duplicate:'average',includeMissing:false,scope:'all',...opts};
 const rows=[],claimed=new Set(),pending=[];
 const aggregate=recs=>{
  const result={};
  for(const k of ['general','lightning','photo']){
   const values=recs.filter(r=>r[k]!==null).map(r=>r[k]);
   result[k]=!values.length?null:options.duplicate==='highest'?Math.max(...values):options.duplicate==='last'?values.at(-1):values.reduce((a,b)=>a+b,0)/values.length;
  }return result;
 };
 const add=(p,scoreId,matched,edit)=>{
  const recs=base.scores.get(scoreId)||[],note=[...p.notes];
  if(recs.length){claimed.add(scoreId);if(recs.some(r=>r.name!==p.name))note.push(`学号一致，评分表姓名：${[...new Set(recs.map(r=>r.name))].join('、')}`);}
  if(scoreId!==p.id&&recs.length)note.push(`已确认按姓名匹配，问卷学号${p.id}，评分表学号${scoreId}`);
  if(recs.length>1)note.push(`${recs.length}条面试记录，按${{average:'平均分',highest:'最高分',last:'导入顺序最后一条'}[options.duplicate]}汇总`);
  recs.forEach(r=>note.push(...r.notes));
  if(!p.hasInfo)note.push('未匹配到人员问卷');
  const r={...p,...aggregate(recs),scoreId,matched,notes:[...new Set(note)],scoreSource:recs.map(r=>`${r.source} 第${r.row}行`).join('；'),override:edit?.role||'auto',role:'未入选',reason:''};
  if(edit?.note)r.notes.push(edit.note);rows.push(r);return r;
 };
 for(const p of base.people.values()){
  const edit=edits[p.id]||{};
  if(base.scores.has(p.id)){add(p,p.id,true,edit);continue;}
  const same=base.byName.get(p.name)||[];
  if(same.length===1&&edit.acceptName){add(p,same[0],true,edit);}
  else {const r=add(p,p.id,false,edit);if(same.length){r.nameCandidates=same;r.notes.push(`同名评分学号：${same.join('、')}，请核实后匹配`);pending.push(r);}else r.notes.push('未找到有效面试评分');}
 }
 for(const [id,recs] of base.scores){if(claimed.has(id)||base.people.has(id))continue;
  const person={id,name:recs[0].name,data:{'姓名':recs[0].name,'学号':id},order:recs[0].order,sources:[],notes:new Set(),hasInfo:false};
  const r=add(person,id,true,edits[id]);r.missingOnly=true;
 }
 const eligible=r=>(r.hasInfo||options.includeMissing)&&r.matched&&[r.general,r.lightning,r.photo].some(x=>x!==null);
 const compare=k=>(a,b)=>(b[k]??-Infinity)-(a[k]??-Infinity)||(b.general??-Infinity)-(a.general??-Infinity)||a.order-b.order;
 const cuts={};
 const roles={lightning:'小闪电',photo:'摄影志愿者'};
 for(const r of rows){
  if(r.override==='exclude'){r.reason='手动排除';continue;}
  if(eligible(r)){r.role='普通志愿者';}
  else r.reason=!r.hasInfo?'无人员问卷':r.nameCandidates?'学号不一致，待确认':'无有效面试评分';
  if(r.override!=='auto'&&r.override!=='exclude'&&eligible(r))r.role=r.override;
 }
 const sequence=options.priority==='photo'?['photo','lightning']:['lightning','photo'];
 for(const kind of sequence){
  const quota=Math.max(0,Math.floor(Number(options[kind])||0));const role=roles[kind];
  const locked=rows.filter(r=>r.role===role&&r.override!=='auto');
  const candidates=rows.filter(r=>eligible(r)&&r.override==='auto'&&r.role==='普通志愿者'&&r[kind]!==null&&
   (options.scope==='all'||Object.entries(r.data).some(([k,v])=>/报名.*闪电|报名.*摄影/.test(k)&&text(v).includes(kind==='photo'?'摄影':'小闪电')))).sort(compare(kind));
  const n=Math.max(0,quota-locked.length);let picked=candidates.slice(0,n);let cutoff=picked.length?picked.at(-1)[kind]:null;
  if(options.ties==='all'&&cutoff!==null&&n>0)picked=candidates.filter(r=>r[kind]>=cutoff);
  picked.forEach(r=>{r.role=role;});
  cuts[kind]={quota,cutoff,selected:locked.length+picked.length,tiedExtra:Math.max(0,locked.length+picked.length-quota),available:locked.length+candidates.length};
 }
 const priority={'摄影志愿者':0,'小闪电':1,'普通志愿者':2,'未入选':3};
 rows.sort((a,b)=>priority[a.role]-priority[b.role]||(a.role==='摄影志愿者'?compare('photo')(a,b):a.role==='小闪电'?compare('lightning')(a,b):compare('general')(a,b)));
 const selected=rows.filter(r=>r.role!=='未入选');
 return {rows,selected,pending,cuts,options,counts:{photo:selected.filter(r=>r.role==='摄影志愿者').length,lightning:selected.filter(r=>r.role==='小闪电').length,ordinary:selected.filter(r=>r.role==='普通志愿者').length,total:selected.length,excluded:rows.length-selected.length},errors:base.errors};
}
root.VCore={text,num,clean,field,findHeader,mapping,prepare,evaluate};
})(typeof window==='undefined'?globalThis:window);
