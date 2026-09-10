'use strict';
const data=window.DUTY_DATA;
// Friday morning exams: 10:00 AM–12:00 PM
data.duties.forEach(d => {
  const isFriday = new Date(d.date + 'T12:00:00Z').getUTCDay() === 5;
  if (isFriday && d.slot === 'S-1') {
    d.time = '10:00 AM – 12:00 PM';
  }
});
const $=id=>document.getElementById(id);
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fullDate=d=>new Intl.DateTimeFormat('en-GB',{day:'numeric',month:'long',year:'numeric',timeZone:'Asia/Dhaka'}).format(new Date(d+'T12:00:00+06:00'));
const dayName=d=>new Intl.DateTimeFormat('en-GB',{weekday:'long',timeZone:'Asia/Dhaka'}).format(new Date(d+'T12:00:00+06:00'));
const names=new Map(data.faculty.map(f=>[f.code,f.name]));
for(const f of [...data.faculty].sort((a,b)=>a.name.localeCompare(b.name))){const o=new Option(f.name,f.code);$('faculty').add(o)}
for(const d of data.dates)$('date').add(new Option(fullDate(d),d));
const params=new URLSearchParams(location.search);
for(const k of ['faculty','date','slot'])if([...$(k).options].some(o=>o.value===params.get(k)))$(k).value=params.get(k);
function filterDuties(faculty,date,slot){return data.duties.filter(d=>(!faculty||d.faculty===faculty)&&(!date||d.date===date)&&(!slot||d.slot===slot))}
function render(){
 const faculty=$('faculty').value,date=$('date').value,slot=$('slot').value;
 const rows=filterDuties(faculty,date,slot);
 const q=new URLSearchParams();for(const k of ['faculty','date','slot'])if($(k).value)q.set(k,$(k).value);
 try{history.replaceState(null,'',location.pathname+(q.size?'?'+q:'')+location.hash)}catch{}
 $('count').textContent=`${rows.length} assignment${rows.length===1?'':'s'} · ${new Set(rows.map(d=>d.date)).size} exam dates`;
 $('selection').textContent=[names.get(faculty)||'All EEE staff',date?fullDate(date):'All dates',slot||'All sessions'].join(' · ');
 const warnings=data.discrepancies.filter(e=>(!faculty||e[0]===faculty)&&(!date||e[1]===date));
 $('notice').hidden=!warnings.length;
 $('notice').textContent=warnings.length?'Schedule check: on 28 September, detailed assignments and master totals differ for '+warnings.map(e=>names.get(e[0])).join(', ')+'. Detailed room assignments are shown here; confirm with the exam office.':'';
 if(!rows.length){$('results').innerHTML='<section class="empty"><h2>No duties listed</h2><p>No assignments match these filters in the uploaded report. Try another date or clear the filters.</p></section>';return}
 const groups=Map.groupBy?Map.groupBy(rows,d=>d.date):rows.reduce((m,d)=>(m.has(d.date)?m.get(d.date).push(d):m.set(d.date,[d]),m),new Map());
 $('results').innerHTML=[...groups].map(([date,items])=>`<section class="day"><div class="day-header"><h2>${esc(fullDate(date))}</h2><span>${esc(dayName(date))}</span></div><div class="table-wrap"><table><thead><tr><th scope="col">EEE faculty / staff</th><th scope="col">Time</th><th scope="col">Room</th><th scope="col">Co-invigilators</th></tr></thead><tbody>${items.map(d=>`<tr><td data-label="EEE faculty / staff"><span class="staff">${esc(names.get(d.faculty)||d.faculty)}</span><small>${esc(d.faculty)}</small></td><td data-label="Time" class="session">${esc(d.time)}<small>${d.slot==='S-1'?'Morning':'Afternoon'} · ${esc(d.slot)}</small></td><td data-label="Room"><span class="room">${esc(d.room)}</span><small>Building ${esc(d.building)}</small></td><td data-label="Co-invigilators" class="team">${d.coInvigilators.length?d.coInvigilators.map(esc).join('<br>'):'No co-invigilator listed'}<small>Source: ${esc(d.source)}</small></td></tr>`).join('')}</tbody></table></div></section>`).join('');
}
for(const k of ['faculty','date','slot'])$(k).addEventListener('change',render);
$('reset').addEventListener('click',()=>{for(const k of ['faculty','date','slot'])$(k).value='';render()});
$('print').addEventListener('click',()=>window.print());
$('checks').innerHTML=data.discrepancies.map(([f,d,master,detail])=>`<p>${esc(names.get(f))} · ${esc(fullDate(d))}: master summary ${master}; detailed schedule ${detail}.</p>`).join('');
render();
