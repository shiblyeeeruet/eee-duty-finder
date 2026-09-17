'use strict';
(() => {
const config=window.DUTY_CONFIG;
const $=id=>document.getElementById(id);
let data=null,names=new Map(),departments=new Map(),clashes=new Map(),busy=false,serial=0,firstLoad=true;
const params=new URLSearchParams(location.search);
function options(id,items,placeholder,wanted){const e=$(id);e.innerHTML='';e.add(new Option(placeholder,''));for(const [label,value] of items)e.add(new Option(label,value));e.value=[...e.options].some(o=>o.value===wanted)?wanted:''}
function populateFaculty(wanted=$('faculty').value){options('faculty',[...data.faculty].filter(f=>!$('department').value||f.department===$('department').value).sort((a,b)=>a.name.localeCompare(b.name)).map(f=>[f.name+' ('+f.department+')',f.code]),'All faculty / staff',wanted)}
function status(message,error=false){$('sync-status').textContent=message;$('sync-status').classList.toggle('sync-error',error)}
function enabled(ready){for(const id of ['department','faculty','date','slot','reset','print'])$(id).disabled=!ready}
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fullDate=d=>new Intl.DateTimeFormat('en-GB',{day:'numeric',month:'long',year:'numeric',timeZone:'Asia/Dhaka'}).format(new Date(d+'T12:00:00+06:00'));
const dayName=d=>new Intl.DateTimeFormat('en-GB',{weekday:'long',timeZone:'Asia/Dhaka'}).format(new Date(d+'T12:00:00+06:00'));

function filterDuties(faculty,date,slot){return data.duties.filter(d=>(!$('department').value||d.department===$('department').value)&&(!faculty||d.faculty===faculty)&&(!date||d.date===date)&&(!slot||d.slot===slot))}
function render(){
 if(!data)return;
 const faculty=$('faculty').value,date=$('date').value,slot=$('slot').value;
 const rows=filterDuties(faculty,date,slot);
 const q=new URLSearchParams();for(const k of ['department','faculty','date','slot'])if($(k).value)q.set(k,$(k).value);
 try{history.replaceState(null,'',location.pathname+(q.size?'?'+q:'')+location.hash)}catch{}
 $('count').textContent=`${rows.length} assignment${rows.length===1?'':'s'} · ${new Set(rows.map(d=>d.date)).size} exam dates`;
 $('selection').textContent=[$('department').value||'All departments',names.get(faculty)||'All faculty / staff',date?fullDate(date):'All dates',slot||'All sessions'].join(' · ');
 const warnings=data.discrepancies.filter(e=>(!$('department').value||departments.get(e[0])===$('department').value)&&(!faculty||e[0]===faculty)&&(!date||e[1]===date));
 $('notice').hidden=!warnings.length;
 $('notice').textContent=warnings.length?`${warnings.length} master-summary mismatch(es) affect this selection. Detailed room assignments are shown; see Data checks below and confirm with the exam office.`:'';
 if(!rows.length){$('results').innerHTML='<section class="empty"><h2>No duties listed</h2><p>No assignments match these filters in the source sheet. Try another date or clear the filters.</p></section>';return}
 const groups=Map.groupBy?Map.groupBy(rows,d=>d.date):rows.reduce((m,d)=>(m.has(d.date)?m.get(d.date).push(d):m.set(d.date,[d]),m),new Map());
 $('results').innerHTML=[...groups].map(([date,items])=>`<section class="day"><div class="day-header"><h2>${esc(fullDate(date))}</h2><span>${esc(dayName(date))}</span></div><div class="table-wrap"><table><thead><tr><th scope="col">Faculty / staff</th><th scope="col">Time</th><th scope="col">Room</th><th scope="col">Co-invigilators</th></tr></thead><tbody>${items.map(d=>`<tr><td data-label="Faculty / staff"><span class="staff">${esc(names.get(d.faculty)||d.faculty)}</span><small>${esc(d.faculty)}</small>${clashes.get([d.faculty,d.date,d.slot].join('|'))>1?'<small class="conflict">Multiple rooms in this session — confirm assignment</small>':''}</td><td data-label="Time" class="session">${esc(d.time)}<small>${d.slot==='S-1'?'Morning':'Afternoon'} · ${esc(d.slot)}</small></td><td data-label="Room"><span class="room">${esc(d.room)}</span><small>Building ${esc(d.building)}</small></td><td data-label="Co-invigilators" class="team">${d.coInvigilators.length?d.coInvigilators.map(esc).join('<br>'):'No co-invigilator listed'}<small>Source: ${esc(d.source)}</small></td></tr>`).join('')}</tbody></table></div></section>`).join('');
}

function loadFeed(){return new Promise((resolve,reject)=>{
 let url;try{url=new URL(config.feedUrl)}catch{reject(Error('Setup needed: paste the Google Apps Script /exec URL into config.js.'));return}
 if(url.origin!=='https://script.google.com'||!/^\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(url.pathname)){reject(Error('Use the Google Apps Script deployment URL ending in /exec.'));return}
 const callback='dutyCallback_'+Date.now()+'_'+(++serial),script=document.createElement('script');let done=false;
 const cleanup=()=>{clearTimeout(timer);script.remove();delete window[callback]};
 const fail=message=>{if(done)return;done=true;cleanup();reject(Error(message))};
 const timer=setTimeout(()=>fail('The live feed did not respond. Check the connection or ask the administrator to verify public web-app access.'),30000);
 window[callback]=payload=>{if(done)return;done=true;cleanup();resolve(payload)};
 script.onerror=()=>fail('Unable to reach the schedule feed. Please try again.');
 url.search='';url.searchParams.set('callback',callback);url.searchParams.set('_',Date.now());script.src=url.href;document.head.appendChild(script);
})}
function validPayload(p){
 if(!p?.ok)throw Error(p?.message||'The schedule feed reported an error.');const d=p.data;
 if(p.schemaVersion!==1||!Number.isFinite(Date.parse(p.fetchedAt))||!d||!['faculty','duties','dates','discrepancies','notes'].every(k=>Array.isArray(d[k])))throw Error('The schedule feed returned an unsupported format.');
 if(!d.faculty.every(f=>typeof f.code==='string'&&typeof f.name==='string'&&typeof f.department==='string')||!d.duties.every(r=>['faculty','department','date','slot','time','room','building','source'].every(k=>typeof r[k]==='string')&&Array.isArray(r.coInvigilators)))throw Error('The schedule is incomplete. Ask the administrator to check the importer.');return d;
}
async function sync(){
 if(busy)return;busy=true;$('refresh').disabled=true;status('Checking the latest schedule…');
 try{const payload=await loadFeed(),next=validPayload(payload);
 const selected=Object.fromEntries(['department','faculty','date','slot'].map(k=>[k,firstLoad?params.get(k)||'':$(k).value]));
 data=next;names=new Map(data.faculty.map(f=>[f.code,f.name]));departments=new Map(data.faculty.map(f=>[f.code,f.department]));clashes=new Map();
 for(const d of data.duties){const k=[d.faculty,d.date,d.slot].join('|');clashes.set(k,(clashes.get(k)||0)+1)}
 options('department',[...new Set(data.faculty.map(f=>f.department))].sort().map(d=>[d,d]),'All departments',selected.department);populateFaculty(selected.faculty);
 options('date',data.dates.map(d=>[fullDate(d),d]),'All dates',selected.date);$('slot').value=['','S-1','S-2'].includes(selected.slot)?selected.slot:'';firstLoad=false;
 $('checks').innerHTML=data.discrepancies.map(([f,d,m,n])=>`<p>${esc(names.get(f))} · ${esc(fullDate(d))}: master summary ${esc(m)}; detailed schedule ${esc(n)}.</p>`).join('')+data.notes.map(n=>`<p>${esc(n)}</p>`).join('');
 if(!data.discrepancies.length&&!data.notes.length)$('checks').textContent='No master-summary discrepancies detected.';
 const stamp=new Intl.DateTimeFormat('en-GB',{dateStyle:'medium',timeStyle:'medium',timeZone:'Asia/Dhaka'}).format(new Date(payload.fetchedAt));
 $('last-sync').textContent='Last successfully synced from source: '+stamp+' (Bangladesh time)';
 const old=Date.now()-Date.parse(payload.fetchedAt)>5*60000;
 status(old?'The returned schedule is more than 5 minutes old. Confirm updates with the exam office.':'Schedule loaded. Checks every 2 minutes; the feed may cache results for up to 1 minute.',old);enabled(true);render();
 }catch(e){status(e.message+(data?' Displaying the last successful sync; it may be outdated.':''),true);if(!data)$('results').innerHTML='<section class="empty"><h2>Schedule unavailable</h2><p>Please check the update message above and try Refresh duties.</p></section>'}
 finally{busy=false;$('refresh').disabled=false}
}
for(const k of ['department','faculty','date','slot'])$(k).addEventListener('change',()=>{if(!data)return;if(k==='department')populateFaculty();render()});
$('reset').addEventListener('click',()=>{if(!data)return;for(const k of ['department','faculty','date','slot'])$(k).value='';populateFaculty();render()});
$('print').addEventListener('click',()=>window.print());$('refresh').addEventListener('click',sync);
document.addEventListener('visibilitychange',()=>{if(!document.hidden)sync()});
setInterval(()=>{if(!document.hidden)sync()},Math.max(60000,Number(config?.refreshMs)||120000));enabled(false);sync();
})();
