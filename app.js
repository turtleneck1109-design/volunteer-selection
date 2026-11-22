(function(){
'use strict';
const $=id=>document.getElementById(id),esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const state={info:[],score:[],base:null,result:null,edits:{},tab:'selected',page:0,stale:false,demo:false};
const number=v=>v===null||v===undefined?'—':Number.isInteger(v)?String(v):Number(v.toFixed(2)).toString();
function message(text,error=false){$('notice').hidden=!text;$('notice').textContent=text;$('notice').classList.toggle('error',error);}
function busy(text){$('busy').hidden=!text;if(text)$('busy-text').textContent=text;}
function opts(){return {lightning:Number($('lightning-quota').value),photo:Number($('photo-quota').value),ties:$('ties').value,priority:$('priority').value,scope:$('scope').value,duplicate:$('duplicate').value,includeMissing:$('include-missing').checked};}
function validate(){for(const id of ['lightning-quota','photo-quota']){const el=$(id);if(!el.validity.valid||el.value==='')throw Error('名额请填写0至10000之间的整数');}}
function status(){
 for(const type of ['info','score']){const sheets=state[type];const el=$(`${type}-status`);el.textContent=sheets.length?`${sheets.length} 个工作表 · ${sheets.reduce((n,s)=>n+s.rows.length,0)} 行数据${state.demo?'（虚构示例）':''}`:'尚未导入';el.classList.toggle('loaded',!!sheets.length);}
 $('generate').disabled=!state.info.length||!state.score.length;
}
function dirty(){state.stale=true;$('export').disabled=true;if(state.result){$('result-caption').textContent='设置已更改，请重新生成名单';}}
function renderMapping(){
 const parts=[];
 for(const type of ['info','score'])state[type].forEach((s,i)=>{
  const keys=type==='info'?['name','id']:['name','id','general','lightning','photo'];
  const labels={name:'姓名',id:'学号',general:'普通面试总分（15分）',lightning:'闪电原始得分（10分）',photo:'摄影原始得分（10分）'};
  parts.push(`<div class="sheet-mapping"><label class="sheet-label"><input type="checkbox" data-enable="${type}:${i}" ${s.enabled!==false?'checked':''}><span>${esc(s.source)}<br>${s.rows.length} 行</span></label>${keys.map(k=>`<label>${labels[k]}<select data-map="${type}:${i}:${k}"><option value="-1">未使用 / 未识别</option>${s.headers.map((h,c)=>h?`<option value="${c}" ${s.mapping[k]===c?'selected':''}>${esc(h)}</option>`:'').join('')}</select></label>`).join('')}</div>`);
 });
 $('mapping-panel').innerHTML=parts.join('')||'<p class="small muted">导入文件后显示字段。</p>';
}
async function importFiles(type,files){
 if(!files.length)return;busy('正在读取文件');message('');
 try{
  if(state.demo){state.info=[];state.score=[];state.edits={};state.result=null;state.demo=false;$('empty').hidden=false;$('table-wrap').hidden=true;$('pagination').hidden=true;$('selection-notes').hidden=true;for(const key of ['total','photo','lightning','ordinary'])$(`count-${key}`).textContent='—';$('result-caption').textContent='完成导入后，名单将在这里生成';}
  const sheets=await VExcel.readFiles(files,t=>$('busy-text').textContent=t);
  if(!sheets.length)throw Error('未找到包含姓名或学号的工作表。请确认表头在前20行，文件为xlsx格式。');
  const existing=new Set(state[type].map(s=>s.source));
  const added=sheets.filter(s=>!existing.has(s.source));state[type].push(...added);
  status();renderMapping();dirty();message(added.length?`已导入${added.length}个工作表。确认名额后点击“生成名单”。`:'同名文件已导入，无需重复添加。');
 }catch(e){message(`导入失败：${e.message}`,true);status();}finally{busy('');$(`${type}-files`).value='';}
}
function recalculate(rebuild=true){
 validate();if(rebuild)state.base=VCore.prepare(state.info,state.score);
 if(state.base.errors.length)throw Error(state.base.errors.slice(0,8).join('\n')+(state.base.errors.length>8?`\n另有${state.base.errors.length-8}条，请检查字段识别` :''));
 if(!state.base.people.size)throw Error('没有识别到有效人员，请检查姓名和学号列。');
 if(!state.base.scores.size)throw Error('没有识别到有效分数，请检查得分列或在Excel中保存计算后的评分表。');
 state.result=VCore.evaluate(state.base,opts(),state.edits);state.stale=false;state.page=0;render();
}
function render(){
 if(!state.result)return;
 const r=state.result,c=r.counts;
 for(const key of ['total','photo','lightning','ordinary'])$(`count-${key}`).textContent=c[key];
 for(const key of ['photo','lightning'])$(`cut-${key}`).textContent=r.cuts[key].cutoff===null?`设置名额 ${r.cuts[key].quota}`:`入选线 ${number(r.cuts[key].cutoff)} 分`;
 $('result-caption').textContent=`${state.base.people.size} 人有问卷 · ${c.total} 人入选 · ${c.excluded} 条记录未入选${state.demo?' · 虚构示例':''}`;
 const notes=[];
 for(const key of ['photo','lightning']){const cut=r.cuts[key],label=key==='photo'?'摄影':'小闪电';if(cut.tiedExtra)notes.push(`${label}设置${cut.quota}人，按同分/手动选择实际保留${cut.selected}人。`);if(cut.selected<cut.quota)notes.push(`${label}符合条件的候选不足，已选${cut.selected}人（名额${cut.quota}）。`);}
 if(r.pending.length)notes.push(`${r.pending.length} 人有同名评分但学号不一致，尚未入选。请在“核对事项”中确认。`);
 const nMissing=r.rows.filter(p=>!p.hasInfo&&p.role==='未入选').length;if(nMissing)notes.push(`${nMissing} 人有评分但无问卷，当前不纳入。可在更多筛选规则中调整。`);
 $('selection-notes').hidden=!notes.length;$('selection-notes').textContent=notes.join(' ');
 $('review-count').textContent=r.rows.filter(p=>p.notes.length).length;$('excluded-count').textContent=c.excluded;
 $('export').disabled=!r.selected.length||state.stale;
 $('empty').hidden=true;$('table-wrap').hidden=false;$('pagination').hidden=false;renderRows();
}
function renderRows(){
 if(!state.result)return;const term=$('search').value.toLowerCase().trim();
 let rows=state.result.rows.filter(r=>state.tab==='selected'?r.role!=='未入选':state.tab==='review'?r.notes.length:r.role==='未入选');
 if(term)rows=rows.filter(r=>[r.name,r.id,r.data['学院/单位'],r.data['学生组织']].some(v=>String(v??'').toLowerCase().includes(term)));
 const pageSize=40,pages=Math.max(1,Math.ceil(rows.length/pageSize));state.page=Math.max(0,Math.min(state.page,pages-1));
 const shown=rows.slice(state.page*pageSize,(state.page+1)*pageSize);
 $('roster-body').innerHTML=shown.map(r=>{
  const cls=r.role==='摄影志愿者'?'photo':r.role==='普通志愿者'?'ordinary':r.role==='未入选'?'excluded':'';
  const allowed=r.matched&&(r.hasInfo||$('include-missing').checked);
  return `<tr><td><span class="person-name">${esc(r.name)}</span><span class="person-id">${esc(r.id)}</span></td><td>${esc(r.data['学院/单位']||r.data['单位']||r.data['学生组织']||'—')}</td><td class="num">${number(r.general)}</td><td class="num">${number(r.lightning)}</td><td class="num">${number(r.photo)}</td><td><span class="role-chip ${cls}">${r.role}</span><select aria-label="调整${esc(r.name)}的岗位" class="edit-role" data-role="${esc(r.id)}"><option value="auto" ${r.override==='auto'?'selected':''}>自动分配</option>${['摄影志愿者','小闪电','普通志愿者'].map(role=>`<option value="${role}" ${r.override===role?'selected':''} ${!allowed?'disabled':''}>指定${role}</option>`).join('')}<option value="exclude" ${r.override==='exclude'?'selected':''}>排除该人员</option></select></td><td>${r.notes.length?`<div class="row-note">${r.notes.map(esc).join('<br>')}</div>`:`<div class="row-note neutral">${r.reason?esc(r.reason):'学号匹配'}</div>`}${r.nameCandidates?.length===1?`<button class="secondary accept-match" data-match="${esc(r.id)}">确认同一人，按问卷学号保留</button>`:''}${state.edits[r.id]?.acceptName?`<button class="text-button" data-unmatch="${esc(r.id)}">撤销姓名匹配</button>`:''}</td></tr>`;
 }).join('')||'<tr><td colspan="7" style="padding:45px;text-align:center;color:#60716f">没有符合当前筛选的人员</td></tr>';
 $('page-info').textContent=`共 ${rows.length} 人 · 第 ${state.page+1} / ${pages} 页`;$('prev-page').disabled=state.page===0;$('next-page').disabled=state.page>=pages-1;
 document.querySelectorAll('[data-tab]').forEach(b=>b.classList.toggle('active',b.dataset.tab===state.tab));
}
function download(blob,name){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
async function exportFile(){
 if(state.stale||!state.result?.selected.length)return;busy('正在生成 Excel');message('');
 try{
  for(const id of ['font-size','row-height','max-width'])if(!$(id).validity.valid||!$(id).value)throw Error('请检查导出排版数值');
  const {headers,rows}=VExcel.exportColumns(state.result,state.base,$('export-fields').value==='full');
  const blob=await VExcel.makeXlsx(headers,rows,{fontSize:Number($('font-size').value),rowHeight:Number($('row-height').value),maxWidth:Number($('max-width').value)});
  download(blob,`志愿者总名单_${rows.length}人.xlsx`);message(`已导出${rows.length}人。摄影与小闪电置顶，所有单元格为白底、黑色边框。`);
 }catch(e){message(`导出失败：${e.message}`,true);}finally{busy('');}
}
function demo(){
 const infoHeaders=['姓名','学号','性别','学院/单位','手机号','邮箱','备注'];
 const scoreHeaders=['姓名','学号','总分(15)','闪电得分(10)','摄影得分(10)'];
 const names=['林知夏','陈星野','周予安','许清禾','沈亦舟','苏沐宁','陆景行','程晚晴','叶思远','顾念初','方以南','何雨桐','白云舒','江望川','宋若溪','梁清越'];
 const infoRows=names.map((n,i)=>{const r=[n,`DEMO${String(i+1).padStart(3,'0')}`,i%2?'男':'女',['计算机学院','人文学院','设计学院','理学院'][i%4],`示例手机号 ${i+1}`,`demo${i+1}@example.com`,'虚构示例，仅用于体验'];r._row=i+2;return r;});
 const scoreRows=names.slice(0,15).map((n,i)=>{const r=[n,i===13?'DEMO999':infoRows[i][1],15-(i%9)*.5,i%3===0?null:10-(i%6)*.5,i%4===0?9-(i%3)*.5:null];r._row=i+2;return r;});
 state.info=[{source:'虚构示例 / 人员问卷',headers:infoHeaders,rows:infoRows,mapping:VCore.mapping(infoHeaders),enabled:true}];
 state.score=[{source:'虚构示例 / 面试评分',headers:scoreHeaders,rows:scoreRows,mapping:VCore.mapping(scoreHeaders),enabled:true}];
 state.edits={};state.demo=true;$('lightning-quota').value=4;$('photo-quota').value=2;$('ties').value='all';$('priority').value='lightning';$('scope').value='all';$('include-missing').checked=false;
 status();renderMapping();state.tab='selected';$('search').value='';try{recalculate();message('已载入虚构示例。你可以修改名额、调整岗位，或导出体验。导入真实文件时会清除示例。');}catch(e){message(e.message,true);}
}
for(const type of ['info','score']){
 $(`${type}-files`).addEventListener('change',e=>importFiles(type,[...e.target.files]));
 const zone=$(`${type}-drop`);zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('dragover');});zone.addEventListener('dragleave',()=>zone.classList.remove('dragover'));zone.addEventListener('drop',e=>{e.preventDefault();zone.classList.remove('dragover');importFiles(type,[...e.dataTransfer.files]);});
}
$('mapping-panel').addEventListener('change',e=>{
 if(e.target.dataset.map){const [type,index,key]=e.target.dataset.map.split(':');state[type][index].mapping[key]=Number(e.target.value);}
 if(e.target.dataset.enable){const [type,index]=e.target.dataset.enable.split(':');state[type][index].enabled=e.target.checked;}dirty();
});
for(const id of ['lightning-quota','photo-quota','ties','priority','scope','duplicate','include-missing'])$(id).addEventListener('change',dirty);
$('generate').addEventListener('click',()=>{try{message('');recalculate();}catch(e){state.stale=true;$('export').disabled=true;message(e.message,true);}});
$('export').addEventListener('click',exportFile);$('demo').addEventListener('click',demo);$('empty-demo').addEventListener('click',demo);
$('reset').addEventListener('click',()=>{if(!state.info.length&&!state.score.length)return;if(confirm('清空导入文件和手动调整？已下载的Excel不会受影响。'))location.reload();});
document.querySelectorAll('[data-tab]').forEach(b=>b.addEventListener('click',()=>{state.tab=b.dataset.tab;state.page=0;renderRows();}));
$('search').addEventListener('input',()=>{state.page=0;renderRows();});$('prev-page').onclick=()=>{state.page--;renderRows();};$('next-page').onclick=()=>{state.page++;renderRows();};
$('roster-body').addEventListener('change',e=>{if(!e.target.dataset.role)return;if(state.stale){message('设置已更改，请先重新生成名单再调整人员。',true);renderRows();return;}const id=e.target.dataset.role;state.edits[id]={...state.edits[id],role:e.target.value};try{recalculate(false);}catch(err){message(err.message,true);}});
$('roster-body').addEventListener('click',e=>{
 const id=e.target.dataset.match||e.target.dataset.unmatch;if(!id)return;if(state.stale){message('设置已更改，请先重新生成名单再确认匹配。',true);return;}
 if(e.target.dataset.match){const p=state.result.rows.find(r=>r.id===id),target=p.nameCandidates[0];if(state.result.rows.some(r=>r.id!==id&&r.hasInfo&&r.scoreId===target&&r.matched)){message('该评分已匹配到另一位有问卷的人员，不能重复使用。请先核实源文件。',true);return;}}
 state.edits[id]={...state.edits[id],acceptName:!!e.target.dataset.match};try{message('');recalculate(false);}catch(err){message(err.message,true);}
});
// Public local APIs also make the parsing and selection logic independently testable.
window.VolunteerApp={state,recalculate,importFiles,exportFile,demo};
renderMapping();status();
})();
