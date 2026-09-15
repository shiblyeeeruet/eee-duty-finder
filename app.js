'use strict';
const data=window.DUTY_DATA;
const $=id=>document.getElementById(id);
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fullDate=d=>new Intl.DateTimeFormat('en-GB',{day:'numeric',month:'long',year:'numeric',timeZone:'Asia/Dhaka'}).format(new Date(d+'T12:00:00+06:00'));
const dayName=d=>new Intl.DateTimeFormat('en-GB',{weekday:'long',timeZone:'Asia/Dhaka'}).format(new Date(d+'T12:00:00+06:00'));
const names=new Map(data.faculty.map(f=>[f.code,f.name]));
const departments=new Map(data.faculty.map(f=>[f.code,f.department]));
for(const d of [...new Set(data.faculty.map(f=>f.department))].sort())$('department').add(new Option(d,d));
function populateFaculty(){
 const previous=$('faculty').value;
 $('faculty').innerHTML='';$('faculty').add(new Option('All faculty / staff',''));
 for(const f of [...data.faculty].filter(f=>!$('department').value||f.department===$('department').value).sort((a,b)=>a.name.localeCompare(b.name)))$('faculty').add(new Option(f.name+' ('+f.department+')',f.code));
 $('faculty').value=[...$('faculty').options].some(o=>o.value===previous)?previous:'';
}
populateFaculty();
const clashes=new Map();
for(const d of data.duties){const key=[d.faculty,d.date,d.slot].join('|');clashes.set(key,(clashes.get(key)||0)+1)}
for(const d of data.dates)$('date').add(new Option(fullDate(d),d));
const params=new URLSearchParams(location.search);
for(const k of ['department','faculty','date','slot'])if([...$(k).options].some(o=>o.value===params.get(k)))$(k).value=params.get(k);
populateFaculty();
function filterDuties(faculty,date,slot){return data.duties.filter(d=>(!$('department').value||d.department===$('department').value)&&(!faculty||d.faculty===faculty)&&(!date||d.date===date)&&(!slot||d.slot===slot))}
function render(){
 const faculty=$('faculty').value,date=$('date').value,slot=$('slot').value;
 const rows=filterDuties(faculty,date,slot);
 const q=new URLSearchParams();for(const k of ['department','faculty','date','slot'])if($(k).value)q.set(k,$(k).value);
 try{history.replaceState(null,'',location.pathname+(q.size?'?'+q:'')+location.hash)}catch{}
 $('count').textContent=`${rows.length} assignment${rows.length===1?'':'s'} · ${new Set(rows.map(d=>d.date)).size} exam dates`;
 $('selection').textContent=[$('department').value||'All departments',names.get(faculty)||'All faculty / staff',date?fullDate(date):'All dates',slot||'All sessions'].join(' · ');
 const warnings=data.discrepancies.filter(e=>(!$('department').value||departments.get(e[0])===$('department').value)&&(!faculty||e[0]===faculty)&&(!date||e[1]===date));
 $('notice').hidden=!warnings.length;
 $('notice').textContent=warnings.length?'Schedule check: on 28 September, detailed assignments and master totals differ for '+warnings.map(e=>names.get(e[0])).join(', ')+'. Detailed room assignments are shown here; confirm with the exam office.':'';
 if(!rows.length){$('results').innerHTML='<section class="empty"><h2>No duties listed</h2><p>No assignments match these filters in the uploaded report. Try another date or clear the filters.</p></section>';return}
 const groups=Map.groupBy?Map.groupBy(rows,d=>d.date):rows.reduce((m,d)=>(m.has(d.date)?m.get(d.date).push(d):m.set(d.date,[d]),m),new Map());
 $('results').innerHTML=[...groups].map(([date,items])=>`<section class="day"><div class="day-header"><h2>${esc(fullDate(date))}</h2><span>${esc(dayName(date))}</span></div><div class="table-wrap"><table><thead><tr><th scope="col">Faculty / staff</th><th scope="col">Time</th><th scope="col">Room</th><th scope="col">Co-invigilators</th></tr></thead><tbody>${items.map(d=>`<tr><td data-label="Faculty / staff"><span class="staff">${esc(names.get(d.faculty)||d.faculty)}</span><small>${esc(d.faculty)}</small>${clashes.get([d.faculty,d.date,d.slot].join('|'))>1?'<small class="conflict">Multiple rooms in this session — confirm assignment</small>':''}</td><td data-label="Time" class="session">${esc(d.time)}<small>${d.slot==='S-1'?'Morning':'Afternoon'} · ${esc(d.slot)}</small></td><td data-label="Room"><span class="room">${esc(d.room)}</span><small>Building ${esc(d.building)}</small></td><td data-label="Co-invigilators" class="team">${d.coInvigilators.length?d.coInvigilators.map(esc).join('<br>'):'No co-invigilator listed'}<small>Source: ${esc(d.source)}</small></td></tr>`).join('')}</tbody></table></div></section>`).join('');
}
for(const k of ['department','faculty','date','slot'])$(k).addEventListener('change',()=>{if(k==='department')populateFaculty();render()});
$('reset').addEventListener('click',()=>{for(const k of ['department','faculty','date','slot'])$(k).value='';populateFaculty();render()});
$('print').addEventListener('click',()=>window.print());
$('checks').innerHTML=data.discrepancies.map(([f,d,master,detail])=>`<p>${esc(names.get(f))} · ${esc(fullDate(d))}: master summary ${master}; detailed schedule ${detail}.</p>`).join('');
render();
