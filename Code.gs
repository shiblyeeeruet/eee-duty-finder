'use strict';
const DutyParser = (() => {
 const clean=v=>String(v??'').replace(/\s+/g,' ').trim();
 const months={jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
 function dateOf(value){
  const m=clean(value).match(/^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+),?\s+(\d{4})$/);
  if(!m||!months[m[2].slice(0,3).toLowerCase()])return null;
  const mo=months[m[2].slice(0,3).toLowerCase()],day=+m[1],year=+m[3],d=new Date(Date.UTC(year,mo-1,day));
  return d.getUTCDate()===day&&d.getUTCMonth()===mo-1?`${year}-${String(mo).padStart(2,'0')}-${String(day).padStart(2,'0')}`:null;
 }
 function column(i){let s='';for(i++;i;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s}
 function parse(sheets){
  const master=sheets.find(s=>s.title==='Master Sheet');if(!master)throw Error('The Master Sheet tab is missing. Please check the source workbook.');
  const roster=new Map();let department='';
  for(const row of master.values){const name=clean(row[1]),code=clean(row[2]);if(name.startsWith('DEPARTMENT OF '))department=name.slice(14).trim();
   if(code&&code!=='Code'){if(roster.has(code)||!department)throw Error('The master roster layout has changed. Please review the importer.');roster.set(code,{code,name,department,category:clean(row[3])})}
  }
  if(!roster.size)throw Error('No staff roster was found. Please review the source layout.');
  const duties=[],dates=[],notes=[];
  for(const sheet of sheets){const date=dateOf(sheet.title);if(!date)continue;dates.push(date);let building='';
   const rows=sheet.values;
   for(let r=0;r<rows.length;r++){
    const label=clean(rows[r]?.[1]),bm=label.match(/Building\s*-\s*(\d+)/i);if(bm)building=bm[1];
    const sm=label.match(/^(S-\d+)\s+(.+)$/);if(!sm)continue;
    const teachers=[];for(let k=r+1;k<rows.length&&/^Teacher\s+\d+$/i.test(clean(rows[k]?.[1]));k++)teachers.push(k);
    for(let c=2;c<Math.max(0,...teachers.map(k=>rows[k].length));c++){
     const people=teachers.map(k=>({code:clean(rows[k][c]),cell:column(c)+(k+1)})).filter(p=>p.code);if(!people.length)continue;
     const room=clean(rows[r-1]?.[c]);if(!room||!building)throw Error('An assigned room or building could not be identified. Please review the sheet layout.');
     let time=sm[2].replace(/\s*[-–]\s*/g,' – ');
     if(new Date(date+'T12:00:00Z').getUTCDay()===5&&sm[1]==='S-1')time='10:00 AM – 12:00 PM';
     for(const p of people){if(!roster.has(p.code)){const dep=p.code.match(/\(([^)]+)\)$/)?.[1]?.toUpperCase()||'Unlisted';roster.set(p.code,{code:p.code,name:p.code,department:dep,category:'Unlisted'});notes.push(`${p.code}: assigned in the dated schedule but missing from the master roster.`)}
      duties.push({faculty:p.code,department:roster.get(p.code).department,date,slot:sm[1],time,room,building,coInvigilators:people.filter(other=>other.cell!==p.cell).map(other=>other.code),source:sheet.title+'!'+p.cell});
     }
    }
   }
  }
  if(!dates.length)throw Error('No dated schedule tabs were found. The workbook format may have changed.');
  const counts=new Map();for(const d of duties){const key=d.faculty+'|'+d.date;counts.set(key,(counts.get(key)||0)+1)}
  const headers=master.values.find(row=>clean(row[2])==='Code')||[],discrepancies=[];
  for(const row of master.values){const code=clean(row[2]);if(!roster.has(code))continue;for(let c=5;c<headers.length;c++){const date=dateOf(headers[c]);if(!date||!dates.includes(date))continue;const raw=clean(row[c]);if(raw&&!/^\d+(?:\.0+)?$/.test(raw)){notes.push(`${code}: summary count for ${date} could not be checked.`);continue}const expected=Number(raw||0),actual=counts.get(code+'|'+date)||0;if(expected!==actual)discrepancies.push([code,date,expected,actual])}}
  return {faculty:[...roster.values()],dates:dates.sort(),duties:duties.sort((a,b)=>a.date.localeCompare(b.date)||a.slot.localeCompare(b.slot)||a.faculty.localeCompare(b.faculty)),discrepancies,notes};
 }
 return {dateOf,parse};
})();
const SOURCE_ID='1a-v3gtk5KOUcyfFlz6hk2qGPTapiwkCfilEoQUXYijY';
const CACHE_KEY='duty-feed-v1';
function loadSchedule_(){
 const metadata=Sheets.Spreadsheets.get(SOURCE_ID,{fields:'sheets.properties(title)'});
 const titles=(metadata.sheets||[]).map(s=>s.properties.title).filter(t=>t==='Master Sheet'||DutyParser.dateOf(t));
 if(!titles.includes('Master Sheet'))throw Error('Master Sheet missing.');
 const ranges=titles.map(t=>"'"+t.replace(/'/g,"''")+"'!A:Z");
 const result=Sheets.Spreadsheets.Values.batchGet(SOURCE_ID,{ranges:ranges,valueRenderOption:'FORMATTED_VALUE'});
 if(!result.valueRanges||result.valueRanges.length!==titles.length)throw Error('Incomplete source response.');
 const parsed=DutyParser.parse(titles.map((title,i)=>({title:title,values:result.valueRanges[i].values||[]})));
 parsed.faculty=parsed.faculty.map(f=>({code:f.code,name:f.name,department:f.department}));
 return {ok:true,schemaVersion:1,fetchedAt:new Date().toISOString(),data:parsed};
}
function cachedSchedule_(){
 const cache=CacheService.getScriptCache();
 function read(){try{const text=cache.get(CACHE_KEY);if(!text)return null;return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(text))).getDataAsString())}catch(e){return null}}
 let saved=read();if(saved)return saved;
 const lock=LockService.getScriptLock();if(!lock.tryLock(10000))throw Error('Schedule refresh is busy.');
 try{saved=read();if(saved)return saved;const fresh=loadSchedule_();
 const compressed=Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(fresh),'application/json')).getBytes());
 if(compressed.length<95000){try{cache.put(CACHE_KEY,compressed,60)}catch(e){}}
 return fresh;
 }finally{lock.releaseLock()}
}
function doGet(e){
 const callback=String(e&&e.parameter&&e.parameter.callback||'');
 if(callback&&!/^dutyCallback_\d+_\d+$/.test(callback))return ContentService.createTextOutput('Invalid callback.').setMimeType(ContentService.MimeType.TEXT);
 let payload;try{payload=cachedSchedule_()}catch(error){console.error(String(error));payload={ok:false,message:'The schedule could not be refreshed. Please try again later or contact the schedule administrator.'}}
 const json=JSON.stringify(payload).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
 return ContentService.createTextOutput(callback?callback+'('+json+');':json).setMimeType(callback?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);
}
function testConnection(){const r=loadSchedule_();console.log(JSON.stringify({success:r.ok,staff:r.data.faculty.length,assignments:r.data.duties.length,fetchedAt:r.fetchedAt}))}
function clearFeedCache(){CacheService.getScriptCache().remove(CACHE_KEY)}
