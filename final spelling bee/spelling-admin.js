(()=>{'use strict';
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let db=null,root=null;if(window.FIREBASE_READY&&window.firebase){try{firebase.initializeApp(window.FIREBASE_CONFIG);db=firebase.database();root=db.ref();}catch(e){console.error(e);}}
const MQ=p=>root.child('mathQuiz/'+p),SP=p=>root.child('spelling/'+p);
const st={classes:{},students:{},words:{},sessions:{}};
function toast(t){const x=$('toast');x.textContent=t;x.classList.remove('hidden');clearTimeout(toast.t);toast.t=setTimeout(()=>x.classList.add('hidden'),4200);}
if(!db)toast('Firebase is not configured. Check firebase-config.js.');
const onErr=e=>{console.error(e);toast('Database error: '+(e&&e.message||e));};
if(db){MQ('classes').on('value',s=>{st.classes=s.val()||{};drawTarget();},onErr);MQ('students').on('value',s=>{st.students=s.val()||{};drawTarget();},onErr);SP('words').on('value',s=>{st.words=s.val()||{};draw();},onErr);SP('sessions').orderByChild('completedAt').limitToLast(300).on('value',s=>{st.sessions=s.val()||{};draw();},onErr);}
function drawTarget(){const cur=$('target').value,cs=Object.entries(st.classes).sort((a,b)=>(a[1].order||0)-(b[1].order||0)),ss=Object.entries(st.students).sort((a,b)=>String(a[1].name).localeCompare(String(b[1].name)));
 $('target').innerHTML='<option value="">Select a class or a student…</option>'+(cs.length?'<optgroup label="Classes">'+cs.map(([id,c])=>`<option value="c:${esc(id)}">Class: ${esc(c.name)}</option>`).join('')+'</optgroup>':'')+(ss.length?'<optgroup label="Students">'+ss.map(([id,s])=>`<option value="s:${esc(id)}">Student: ${esc(s.name)} (${esc(s.className||s.passcode||'')})</option>`).join('')+'</optgroup>':'');
 if([...$('target').options].some(o=>o.value===cur))$('target').value=cur;draw();}
const label=v=>{if(!v)return '';const [k,id]=[v[0],v.slice(2)];return k==='c'?(st.classes[id]||{}).name||id:(st.students[id]||{}).name||id;};
const mine=([,w])=>{const v=$('target').value;return v&&((v[0]==='c'&&w.classId===v.slice(2))||(v[0]==='s'&&w.studentId===v.slice(2)));};
function draw(){const all=Object.entries(st.words),regs=new Set(all.map(([,w])=>w.register)),ss=Object.values(st.sessions);
 $('aWords').textContent=all.length;$('aRegs').textContent=regs.size;$('aTests').textContent=ss.length;$('aAvg').textContent=ss.length?Math.round(ss.reduce((a,s)=>a+(s.score||0),0)/ss.length)+'%':'—';
 const q=$('wSearch').value.trim().toLowerCase(),rows=all.filter(mine).filter(([,w])=>!q||(w.word+' '+w.register).toLowerCase().includes(q)).sort((a,b)=>String(a[1].register).localeCompare(b[1].register)||String(a[1].word).localeCompare(b[1].word));
 $('wTable').innerHTML=rows.length?rows.slice(0,400).map(([id,w])=>`<tr><td><b>${esc(w.word)}</b></td><td>${esc(w.register)}</td><td>${esc(w.definition)}</td><td>${esc(w.status||'published')}</td><td><button class="btn danger" data-del="${esc(id)}" style="padding:6px 10px">Delete</button></td></tr>`).join(''):`<tr><td colspan="5" class="mq-muted">${$('target').value?'No words for this target yet. Download the template, add words, and upload it.':'Select a class or student above.'}</td></tr>`;
 $('rTable').innerHTML=ss.sort((a,b)=>(b.completedAt||0)-(a.completedAt||0)).slice(0,100).map(s=>`<tr><td>${new Date(s.completedAt||0).toLocaleString([],{dateStyle:'medium',timeStyle:'short'})}</td><td>${esc(s.studentName)}</td><td>${esc(s.title)}</td><td>${Math.round(s.score||0)}%</td><td>${s.correct}/${s.total}</td></tr>`).join('')||'<tr><td colspan="5" class="mq-muted">No results yet.</td></tr>';}
$('target').onchange=draw;$('wSearch').oninput=draw;
document.addEventListener('click',async e=>{const b=e.target.closest&&e.target.closest('[data-del]');if(!b)return;try{await SP('words/'+b.dataset.del).remove();toast('Word deleted.');}catch(x){onErr(x);}});
$('delAll').onclick=async()=>{const v=$('target').value;if(!v)return toast('Select a class or student first.');const ids=Object.entries(st.words).filter(mine).map(([id])=>id);if(!ids.length)return toast('Nothing to delete.');if(!confirm('Delete '+ids.length+' words for '+label(v)+'? This cannot be undone.'))return;const up={};ids.forEach(id=>up[id]=null);try{await SP('words').update(up);toast('Deleted '+ids.length+' words.');}catch(x){onErr(x);}};
$('demoBtn').onclick=async()=>{if(!db)return toast('Configure Firebase first.');const up={};(window.SP_WORDS||[]).forEach((r,i)=>{const k='demo_'+String(i+1).padStart(3,'0');if(!st.words[k])up[k]={word:r[0],register:r[1],definition:r[2],sentence:r[3],origin:r[4],demo:true,status:'published',createdAt:Date.now()};});if(!Object.keys(up).length)return toast('The demo words are already loaded.');try{await SP('words').update(up);toast('Added '+Object.keys(up).length+' demo words. Passcode 123123 will use them.');}catch(x){onErr(x);}};
/* ---------- Excel template: one per class or student ---------- */
$('tplBtn').onclick=()=>{const v=$('target').value;if(!v)return toast('Select a class or a student first.');if(!window.XLSX)return toast('Excel library did not load. Check your internet.');
 const rows=[['word','register','definition','sentence','origin','status'],...(window.SP_WORDS||[]).map(r=>[...r,'published'])];
 const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[16,14,48,48,44,11].map(w=>({wch:w}));
 const info=[['TARGET',v+'|'+label(v)],['FOR',(v[0]==='c'?'Class: ':'Student: ')+label(v)],[''],['HOW TO USE',''],['word','Required. Letters only (hyphen and apostrophe allowed).'],['register','The word group, e.g. School, Farm, Geology, Kitchen, Health. Students pick a register to practise. Use 20+ words per register for a full 20-word test.'],['definition','Shown as a clue. The word itself is hidden automatically.'],['sentence','The word used in a sentence. Shown as a clue with the word hidden.'],['origin','Where the word comes from. Shown as a clue.'],['status','published or draft. Blank means published.'],[''],['Important','Replace the 16 sample rows with this target’s words. Do not edit the TARGET row on this sheet. Words are added, not overwritten: existing word and register pairs are skipped.']];
 const w2=XLSX.utils.aoa_to_sheet(info);w2['!cols']=[{wch:14},{wch:110}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Words');XLSX.utils.book_append_sheet(wb,w2,'Instructions');
 XLSX.writeFile(wb,'spelling-words-'+label(v).replace(/[^a-z0-9]+/gi,'-').toLowerCase()+'.xlsx');};
$('xlsFile').onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;const v=$('target').value;if(!v)return toast('Select the class or student this template is for, then upload again.');if(!db||!window.XLSX)return toast('Firebase or the Excel library is not ready.');
 try{const wb=XLSX.read(await f.arrayBuffer(),{type:'array'}),ins=wb.Sheets.Instructions,tg=ins?XLSX.utils.sheet_to_json(ins,{header:1}).find(r=>r[0]==='TARGET'):null;
  if(!tg||String(tg[1]).split('|')[0]!==v)return toast('This template was made for '+(tg?String(tg[1]).split('|')[1]:'another target')+', not '+label(v)+'. Download the correct template.');
  const sh=wb.Sheets.Words;if(!sh)throw Error('The Words sheet is missing.');
  const have=new Set(Object.values(st.words).filter(w=>(v[0]==='c'?w.classId:w.studentId)===v.slice(2)).map(w=>(w.word+'|'+w.register).toLowerCase())),up={},bad=[];let n=0,dup=0;
  XLSX.utils.sheet_to_json(sh,{defval:''}).forEach((r,i)=>{const L={};Object.entries(r).forEach(([k,x])=>L[String(k).trim().toLowerCase()]=String(x).trim());if(!Object.values(L).some(Boolean))return;
   const word=(L.word||'').toLowerCase();if(!/^[a-z'\- ]{2,40}$/.test(word)){bad.push(i+2);return;}const reg=L.register||'General',key=(word+'|'+reg).toLowerCase();if(have.has(key)){dup++;return;}have.add(key);
   const o={word,register:reg,definition:L.definition||'',sentence:L.sentence||'',origin:L.origin||'',demo:false,status:/^draft$/i.test(L.status)?'draft':'published',createdAt:Date.now()};o[v[0]==='c'?'classId':'studentId']=v.slice(2);
   up['w'+Date.now().toString(36)+'_'+n++]=o;});
  if(!n)throw Error('No new valid words found.'+(bad.length?' Check rows '+bad.slice(0,5).join(', ')+'.':''));
  await SP('words').update(up);toast('Added '+n+' words to '+label(v)+(dup?'; skipped '+dup+' duplicates':'')+(bad.length?'; skipped '+bad.length+' invalid rows':'')+'.');
 }catch(x){console.error(x);toast(x.message||'Upload failed.');}};
})();
