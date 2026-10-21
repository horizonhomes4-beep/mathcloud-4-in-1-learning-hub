(()=>{'use strict';
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let db=null;if(window.FIREBASE_READY&&window.firebase){try{firebase.initializeApp(window.FIREBASE_CONFIG);db=firebase.database();}catch(e){console.error(e);}}
const CP=p=>db.ref('comprehension/'+p),MQ=p=>db.ref('mathQuiz/'+p);
const st={classes:{},students:{},passages:{},sessions:{}};
function toast(t){const x=$('toast');x.textContent=t;x.classList.remove('hidden');clearTimeout(toast.t);toast.t=setTimeout(()=>x.classList.add('hidden'),5000);}
const onErr=e=>{console.error(e);toast('Database error: '+(e&&e.message||e));};
if(!db)toast('Firebase is not configured. Check firebase-config.js.');else{
 MQ('classes').on('value',s=>{st.classes=s.val()||{};drawTarget();},onErr);MQ('students').on('value',s=>{st.students=s.val()||{};drawTarget();},onErr);
 CP('passages').on('value',s=>{st.passages=s.val()||{};draw();},onErr);CP('sessions').orderByChild('completedAt').limitToLast(300).on('value',s=>{st.sessions=s.val()||{};draw();},onErr);}
function drawTarget(){const cur=$('target').value,cs=Object.entries(st.classes).sort((a,b)=>(a[1].order||0)-(b[1].order||0)),ss=Object.entries(st.students).sort((a,b)=>String(a[1].name).localeCompare(String(b[1].name)));
 $('target').innerHTML='<option value="">Select a class or a student…</option>'+(cs.length?'<optgroup label="Classes">'+cs.map(([id,c])=>`<option value="c:${esc(id)}">Class: ${esc(c.name)}</option>`).join('')+'</optgroup>':'')+(ss.length?'<optgroup label="Students">'+ss.map(([id,s])=>`<option value="s:${esc(id)}">Student: ${esc(s.name)} (${esc(s.className||s.passcode||'')})</option>`).join('')+'</optgroup>':'');
 if([...$('target').options].some(o=>o.value===cur))$('target').value=cur;draw();}
const label=v=>{if(!v)return '';const id=v.slice(2);return v[0]==='c'?(st.classes[id]||{}).name||id:(st.students[id]||{}).name||id;};
const mine=([,p])=>{const v=$('target').value;return v&&((v[0]==='c'&&p.classId===v.slice(2))||(v[0]==='s'&&p.studentId===v.slice(2)));};
function draw(){const all=Object.entries(st.passages),ss=Object.values(st.sessions);
 $('aPass').textContent=all.length;$('aQs').textContent=all.reduce((a,[,p])=>a+(p.questions||[]).length,0);$('aTests').textContent=ss.length;$('aAvg').textContent=ss.length?Math.round(ss.reduce((a,s)=>a+(s.score||0),0)/ss.length)+'%':'—';
 const q=$('wSearch').value.trim().toLowerCase(),rows=all.filter(mine).filter(([,p])=>!q||((p.title||'')+' '+(p.topic||'')).toLowerCase().includes(q)).sort((a,b)=>String(a[1].topic).localeCompare(b[1].topic)||String(a[1].title).localeCompare(b[1].title));
 $('wTable').innerHTML=rows.length?rows.map(([id,p])=>`<tr><td><b>${esc(p.title)}</b></td><td>${esc(p.topic)}</td><td>${(p.vocab||[]).length}</td><td>${(p.questions||[]).length}</td><td><button class="btn" data-tg="${esc(id)}" style="padding:6px 10px">${esc(p.status||'published')}</button></td><td><button class="btn danger" data-del="${esc(id)}" style="padding:6px 10px">Delete</button></td></tr>`).join(''):`<tr><td colspan="6" class="mq-muted">${$('target').value?'No passages for this target yet. Download the template, add passages, and upload it.':'Select a class or student above.'}</td></tr>`;
 const a={};ss.forEach(s=>Object.entries(s.skills||{}).forEach(([k,[c,t]])=>{a[k]=a[k]||[0,0];a[k][0]+=c;a[k][1]+=t;}));
 $('aSkills').innerHTML=Object.entries(a).map(([k,[c,t]])=>[k,Math.round(100*c/t)]).sort((x,y)=>x[1]-y[1]).map(([k,p])=>`<div class="cp-sk"><span>${esc(k)}</span><div><i style="width:${p}%;background:${p<50?'#b42318':p<75?'#f79009':'#12b76a'}"></i></div><b>${p}%</b></div>`).join('')||'<p class="mq-muted">No results yet.</p>';
 $('rTable').innerHTML=ss.sort((a,b)=>(b.completedAt||0)-(a.completedAt||0)).slice(0,100).map(s=>`<tr><td>${new Date(s.completedAt||0).toLocaleString([],{dateStyle:'medium',timeStyle:'short'})}</td><td>${esc(s.studentName)}</td><td>${esc(s.title)}</td><td>${Math.round(s.score||0)}%</td><td>${s.correct}/${s.total}</td><td>${Math.floor((s.seconds||0)/60)}m ${(s.seconds||0)%60}s</td></tr>`).join('')||'<tr><td colspan="6" class="mq-muted">No results yet.</td></tr>';}
$('target').onchange=draw;$('wSearch').oninput=draw;
document.addEventListener('click',async e=>{const d=e.target.closest&&e.target.closest('[data-del]'),t=e.target.closest&&e.target.closest('[data-tg]');try{
 if(d&&confirm('Delete this passage?')){await CP('passages/'+d.dataset.del).remove();toast('Passage deleted.');}
 if(t){const p=st.passages[t.dataset.tg];await CP('passages/'+t.dataset.tg+'/status').set(p.status==='draft'?'published':'draft');}}catch(x){onErr(x);}});
$('delAll').onclick=async()=>{const v=$('target').value;if(!v)return toast('Select a class or student first.');const ids=Object.entries(st.passages).filter(mine).map(([id])=>id);if(!ids.length)return toast('Nothing to delete.');if(!confirm('Delete '+ids.length+' passages for '+label(v)+'? This cannot be undone.'))return;const up={};ids.forEach(id=>up[id]=null);try{await CP('passages').update(up);toast('Deleted '+ids.length+' passages.');}catch(x){onErr(x);}};
/* ---------- Excel template ---------- */
$('tplBtn').onclick=()=>{const v=$('target').value;if(!v)return toast('Select a class or a student first.');if(!window.XLSX)return toast('Excel library did not load. Check your internet.');
 const D=window.CP_DEMO,P=[['passage_id','title','topic','text','status']],V=[['passage_id','word','meaning']],Q=[['passage_id','question','A','B','C','D','answer','skill','explanation']];
 D.forEach(d=>{P.push([d.id,d.title,d.topic,d.text,'published']);d.vocab.forEach(x=>V.push([d.id,x[0],x[1]]));d.questions.forEach(q=>Q.push([d.id,q[0],...q[1],'ABCD'[q[2]],q[3],q[4]]));});
 const mk=(a,w)=>{const s=XLSX.utils.aoa_to_sheet(a);s['!cols']=w.map(x=>({wch:x}));return s;};
 const info=[['TARGET',v+'|'+label(v)],['FOR',(v[0]==='c'?'Class: ':'Student: ')+label(v)],[''],['HOW TO USE',''],['Passages','One row per passage. passage_id is any short code you invent (p1, p2…). Separate paragraphs with a blank line inside the text cell. status: published or draft.'],['Vocabulary','One row per hard word. Use the same passage_id. The word is highlighted in the passage and Coach Reed reads its meaning.'],['Questions','One row per MCQ. Use the same passage_id. Fill A, B, C, D (at least A and B). answer is the letter A, B, C or D.'],['skill','Literal, Inference, Vocabulary, Main idea, Sequence… Used for the skill analysis. Keep spelling consistent.'],['Important','Replace the 2 sample passages. Do not edit the TARGET row. Passages whose title already exists for this target are skipped.']];
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,mk(P,[10,26,14,80,11]),'Passages');XLSX.utils.book_append_sheet(wb,mk(V,[10,16,60]),'Vocabulary');XLSX.utils.book_append_sheet(wb,mk(Q,[10,50,24,24,24,24,8,14,50]),'Questions');XLSX.utils.book_append_sheet(wb,mk(info,[14,120]),'Instructions');
 XLSX.writeFile(wb,'comprehension-'+label(v).replace(/[^a-z0-9]+/gi,'-').toLowerCase()+'.xlsx');};
const rows=(wb,n)=>wb.Sheets[n]?XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:''}).map(r=>{const L={};Object.entries(r).forEach(([k,x])=>L[String(k).trim().toLowerCase()]=String(x).trim());return L;}).filter(L=>Object.values(L).some(Boolean)):[];
$('xlsFile').onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;const v=$('target').value;if(!v)return toast('Select the class or student this template is for, then upload again.');if(!db||!window.XLSX)return toast('Firebase or the Excel library is not ready.');
 try{const wb=XLSX.read(await f.arrayBuffer(),{type:'array'}),tg=wb.Sheets.Instructions?XLSX.utils.sheet_to_json(wb.Sheets.Instructions,{header:1}).find(r=>r[0]==='TARGET'):null;
  if(!tg||String(tg[1]).split('|')[0]!==v)return toast('This template was made for '+(tg?String(tg[1]).split('|')[1]:'another target')+', not '+label(v)+'. Download the correct template.');
  const have=new Set(Object.values(st.passages).filter(p=>(v[0]==='c'?p.classId:p.studentId)===v.slice(2)).map(p=>String(p.title).toLowerCase())),map={},errs=[];
  rows(wb,'Passages').forEach((r,i)=>{if(!r.passage_id||!r.title||!r.text){errs.push('Passages row '+(i+2)+': needs id, title and text');return;}map[r.passage_id]={title:r.title,topic:r.topic||'General',text:r.text.replace(/\r/g,''),status:/^draft$/i.test(r.status)?'draft':'published',vocab:[],questions:[],createdAt:Date.now()};map[r.passage_id][v[0]==='c'?'classId':'studentId']=v.slice(2);});
  rows(wb,'Vocabulary').forEach((r,i)=>{const p=map[r.passage_id];if(!p||!r.word||!r.meaning){errs.push('Vocabulary row '+(i+2));return;}p.vocab.push([r.word,r.meaning]);});
  rows(wb,'Questions').forEach((r,i)=>{const p=map[r.passage_id],o=[r.a,r.b,r.c,r.d].filter(Boolean),a='ABCD'.indexOf((r.answer||'').toUpperCase()[0]||'?');if(!p||!r.question||o.length<2||a<0||a>=o.length||!r[['a','b','c','d'][a]]){errs.push('Questions row '+(i+2));return;}p.questions.push([r.question,o,a,r.skill||'General',r.explanation||'']);});
  const up={};let n=0,dup=0;Object.values(map).forEach(p=>{if(!p.questions.length){errs.push('"'+p.title+'" has no valid questions');return;}if(have.has(p.title.toLowerCase())){dup++;return;}have.add(p.title.toLowerCase());
   up['p'+Date.now().toString(36)+'_'+n++]={title:p.title,topic:p.topic,text:p.text,status:p.status,createdAt:p.createdAt,classId:p.classId||null,studentId:p.studentId||null,vocab:p.vocab.map(([word,meaning])=>({word,meaning})),questions:p.questions.map(([q,o,a,skill,why])=>({q,o,a,skill,why}))};});
  Object.values(up).forEach(o=>{if(o.classId===null)delete o.classId;if(o.studentId===null)delete o.studentId;});
  if(!n)throw Error('No new valid passages found.'+(errs.length?' '+errs.slice(0,3).join('; ')+'.':''));
  await CP('passages').update(up);toast('Added '+n+' passages to '+label(v)+(dup?'; skipped '+dup+' duplicates':'')+(errs.length?'; '+errs.length+' problem rows skipped ('+errs.slice(0,2).join('; ')+')':'')+'.');
 }catch(x){console.error(x);toast(x.message||'Upload failed.');}};
})();
