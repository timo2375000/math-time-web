// Google Apps Script: run as the spreadsheet owner. Secrets stay in Script Properties.
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('수학 월별 레슨플랜');
}
function authLogin(password) {
  const p = PropertiesService.getScriptProperties();
  const expected = p.getProperty('APP_PASSWORD');
  if (!expected || !/^(?=.*[A-Za-z])(?=.*[0-9])[A-Za-z0-9]{8,}$/.test(expected)) throw new Error('관리자가 APP_PASSWORD를 영문과 숫자를 섞어 8글자 이상으로 설정해야 합니다.');
  const cache = CacheService.getScriptCache(), lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const attempts = Number(cache.get('login-attempts') || 0);
    if (attempts >= 30) throw new Error('잠시 후 다시 시도해주세요.');
    cache.put('login-attempts', String(attempts + 1), 60);
    const digest = value => Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value));
    const a = digest(password), b = digest(expected);
    let different = 0;for(let i=0;i<a.length;i++) different |= a[i]^b[i];
    if (different) throw new Error('접속 비밀번호를 확인해주세요.');
    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put('session:' + token, 'ok', 21600);
    return token;
  } finally { lock.releaseLock(); }
}
function apiRequest(token, path, method, input) {
  if (!token || CacheService.getScriptCache().get('session:' + token) !== 'ok') return {status:401,data:{error:'접속 시간이 만료되었습니다. 새로고침 후 다시 로그인해주세요.'}};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const book = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const registry = loadRegistry_(book);
    return route_(book, registry, path, method || 'GET', input || {});
  } catch(e) {
    console.error(String(e.message));
    return {status:e.status || 500,data:{error:e.status ? e.message : '저장소 처리에 실패했습니다. 잠시 후 다시 시도해주세요.'}};
  } finally { if(lock.hasLock())lock.releaseLock(); }
}
function fail_(status,message){const e=new Error(message);e.status=status;throw e;}
function text_(value,label,max){if(typeof value!=='string'||value.length>(max||5000))fail_(400,label+' 입력을 확인해주세요.');return value.trim();}
function recordSheet_(book,name){let s=book.getSheetByName(name);if(!s){s=book.insertSheet(name);s.appendRow(['ID','JSON']);}return s;}
function readRecords_(sheet){return sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues().map(r=>JSON.parse(r[1])):[];}
function putRecord_(sheet,value){const rows=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues():[];const i=rows.findIndex(r=>String(r[0])===value.id);sheet.getRange(i<0?sheet.getLastRow()+1:i+2,1,1,2).setValues([[value.id,JSON.stringify(value)]]);}
function monthSheets_(book){return book.getSheets().filter(s=>/^\d{4}-(0[1-9]|1[0-2])$/.test(s.getName()));}
function loadRegistry_(book){
  const ts=recordSheet_(book,'_앱_선생님'),cs=recordSheet_(book,'_앱_반');
  const teachers=readRecords_(ts),classes=readRecords_(cs);
  // Preserve IDs already present in the connected monthly source sheets.
  for(const sheet of monthSheets_(book))for(const l of readLessons_(sheet)){
    if(!teachers.some(t=>t.id===l.teacher_id)){const t={id:l.teacher_id,name:l.teacher_name||'선생님',active:true};teachers.push(t);putRecord_(ts,t);}
    let cl=classes.find(c=>c.id===l.class_id);
    if(!cl){cl={id:l.class_id,name:l.class_name||'반',level:'',teacher_ids:[l.teacher_id],active:true};classes.push(cl);putRecord_(cs,cl);}
    else if(cl.active!==false&&!cl.teacher_ids.includes(l.teacher_id)){cl.teacher_ids.push(l.teacher_id);putRecord_(cs,cl);}
  }
  return {teachers,classes,ts,cs};
}
function readLessons_(sheet){
  if(!sheet||sheet.getLastRow()<2)return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,18).getValues().filter(r=>r[0]).map(r=>({id:String(r[0]),teacher_id:String(r[1]),teacher_name:String(r[2]),class_id:String(r[3]),class_name:String(r[4]),date:String(r[5]),start_time:String(r[6]),end_time:String(r[7]),title:String(r[8]),activities:String(r[9]),objectives:String(r[10]),materials:String(r[11]),homework:String(r[12]),assessment:String(r[13]),notes:String(r[14]),status:String(r[15]||'draft'),revision:Number(r[16]),updated_at:String(r[17]),created_at:String(r[17])}));
}
function activeClasses_(r){const ids=new Set(r.teachers.filter(t=>t.active!==false).map(t=>t.id));return r.classes.filter(c=>c.active!==false).map(c=>({...c,teacher_ids:c.teacher_ids.filter(id=>ids.has(id))})).filter(c=>c.teacher_ids.length);}
function teacher_(r,id){const t=r.teachers.find(t=>t.id===id&&t.active!==false);if(!t)fail_(404,'선생님을 찾을 수 없습니다.');return t;}
function visible_(r,sheet,id){const ids=new Set(activeClasses_(r).map(c=>c.id));return readLessons_(sheet).filter(l=>l.teacher_id===id&&ids.has(l.class_id));}
function route_(book,r,path,method,input){
  const pathname=path.split('?')[0];
  const ok=data=>({status:200,data});
  if(pathname==='/api/bootstrap'&&method==='GET')return ok({teachers:r.teachers.filter(t=>t.active!==false),classes:activeClasses_(r)});
  if(pathname==='/api/lesson-plan-archive'&&method==='GET'){
    const groups=[];
    for(const s of monthSheets_(book)){
      const list=[];for(const t of r.teachers.filter(t=>t.active!==false)){
        const lessons=visible_(r,s,t.id);if(lessons.length)list.push({teacher_id:t.id,teacher_name:t.name,lesson_count:lessons.length,class_count:new Set(lessons.map(l=>l.class_id)).size,updated_at:lessons.reduce((a,l)=>l.updated_at>a?l.updated_at:a,'')});
      }
      if(list.length)groups.push({month:s.getName(),teachers:list});
    }
    return ok(groups.sort((a,b)=>b.month.localeCompare(a.month)));
  }
  if(pathname==='/api/teachers'&&method==='POST'){
    const name=text_(input.name,'선생님 이름',80);if(!name)fail_(400,'이름을 입력해주세요.');const t={id:Utilities.getUuid(),name,active:true};putRecord_(r.ts,t);return {status:201,data:t};
  }
  if(pathname==='/api/classes'&&method==='POST'){
    const name=text_(input.name,'반 이름',100),level=text_(input.level||'','학년',100);if(!name)fail_(400,'반 이름을 입력해주세요.');
    if(!Array.isArray(input.teacher_ids)||!input.teacher_ids.length||input.teacher_ids.length>100)fail_(400,'담당 선생님을 선택해주세요.');input.teacher_ids.forEach(id=>teacher_(r,id));
    const c={id:Utilities.getUuid(),name,level,teacher_ids:[...new Set(input.teacher_ids)],active:true};putRecord_(r.cs,c);return {status:201,data:c};
  }
  const deletion=pathname.match(/^\/api\/(teachers|classes)\/([a-zA-Z0-9-]+)$/);
  if(deletion&&method==='DELETE'){
    const list=deletion[1]==='teachers'?r.teachers:r.classes;const target=list.find(x=>x.id===deletion[2]&&x.active!==false);if(!target)fail_(404,'대상을 찾을 수 없습니다.');target.active=false;target.deleted_at=new Date().toISOString();putRecord_(deletion[1]==='teachers'?r.ts:r.cs,target);
    for(const source of monthSheets_(book))for(const t of r.teachers)if(readLessons_(source).some(l=>l.teacher_id===t.id))updateCalendar_(book,source,{month:source.getName(),teacher_name:t.name,lesson:{teacher_id:t.id}});
    return ok({deleted:true,id:target.id});
  }
  const m=pathname.match(/^\/api\/teachers\/([a-zA-Z0-9-]+)\/months\/(\d{4}-(?:0[1-9]|1[0-2]))\/(lessons|export\.docx)(?:\/([a-zA-Z0-9-]+))?$/);
  if(!m)fail_(404,'요청한 기능을 찾을 수 없습니다.');
  const [,tid,month,type,lid]=m,teacher=teacher_(r,tid),source=book.getSheetByName(month);
  if(type==='export.docx'&&method==='GET'){
    const match=path.match(/[?&]class_id=([^&]+)/),cid=match?decodeURIComponent(match[1]):'';
    const classes=activeClasses_(r),selectedClass=cid?classes.find(c=>c.id===cid&&c.teacher_ids.includes(tid)):null;if(cid&&!selectedClass)fail_(400,'담당 반을 선택해주세요.');
    const lessons=visible_(r,source,tid).filter(l=>!cid||l.class_id===cid);if(!lessons.length)fail_(404,'내려받을 계획이 없습니다.');
    const bytes=createLessonDocument_({teacher,classes,lessons,month,selectedClass});return ok({base64:Utilities.base64Encode(Array.from(bytes,x=>x>127?x-256:x))});
  }
  if(type!=='lessons')fail_(404,'요청한 기능을 찾을 수 없습니다.');
  if(method==='GET'&&!lid)return ok(visible_(r,source,tid));
  if(!((method==='POST'&&!lid)||(method==='PUT'&&lid)))fail_(404,'요청한 기능을 찾을 수 없습니다.');
  const existing=lid?readLessons_(source).find(l=>l.id===lid&&l.teacher_id===tid):null;
  if(lid&&!existing)fail_(404,'수업을 찾을 수 없습니다.');if(existing&&Number(input.revision)!==existing.revision)fail_(409,'다른 화면에서 수정했습니다. 새로고침 후 다시 수정해주세요.');
  const data=input.content_only?{...existing,...input}:input;
  const cl=activeClasses_(r).find(c=>c.id===data.class_id&&c.teacher_ids.includes(tid));if(!cl)fail_(400,'담당 반을 선택해주세요.');
  const date=text_(data.date,'날짜',10),parsed=new Date(date+'T12:00:00Z');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||date.slice(0,7)!==month||!Number.isFinite(parsed.getTime())||parsed.toISOString().slice(0,10)!==date)fail_(400,'선택한 월의 날짜를 입력해주세요.');
  const lesson={id:existing?.id||Utilities.getUuid(),teacher_id:tid,class_id:cl.id,date,title:text_(data.title||'','주제',200),status:data.status||'draft',start_time:text_(data.start_time||'','시작 시간',5),end_time:text_(data.end_time||'','종료 시간',5)};
  for(const k of ['activities','objectives','materials','homework','assessment','notes'])lesson[k]=text_(data[k]||'',k);
  if(!lesson.title&&!data.content_only)fail_(400,'주제를 입력해주세요.');if(data.content_only&&!lesson.activities)fail_(400,'수업 내용을 입력해주세요.');
  if(!['draft','confirmed','completed'].includes(lesson.status)||(lesson.status!=='draft'&&!lesson.objectives))fail_(400,'상태와 목표를 확인해주세요.');
  const a=lesson.start_time,b=lesson.end_time;if([a,b].some(t=>t&&!/^([01]\d|2[0-3]):[0-5]\d$/.test(t))||!!a!==!!b||(a&&b<=a))fail_(400,'시작·종료 시간을 확인해주세요.');
  lesson.revision=(existing?.revision||0)+1;lesson.created_at=existing?.created_at||new Date().toISOString();lesson.updated_at=new Date().toISOString();
  let sheet=source;if(!sheet){sheet=book.insertSheet(month);sheet.appendRow(['수업 ID','선생님 ID','선생님','반 ID','반','날짜','시작 시간','종료 시간','주제','수업 내용','학습 목표','교재','과제','평가','비고','상태','수정 버전','수정 시각']);sheet.setFrozenRows(1);}
  const ids=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues():[],i=ids.findIndex(v=>String(v[0])===lesson.id);
  const safe=value=>{const s=String(value??'');return /^[=+@-]/.test(s)?"'"+s:s;};
  const values=[lesson.id,tid,teacher.name,cl.id,cl.name,date,a,b,lesson.title,lesson.activities,lesson.objectives,lesson.materials,lesson.homework,lesson.assessment,lesson.notes,lesson.status,lesson.revision,lesson.updated_at].map(safe);
  sheet.getRange(i<0?sheet.getLastRow()+1:i+2,1,1,18).setNumberFormat('@').setValues([values]);
  let sheets_sync={status:'synced'};try{updateCalendar_(book,sheet,{month,teacher_name:teacher.name,lesson});}catch(e){console.error(String(e.message));sheets_sync={status:'calendar_failed'};}
  return {status:existing?200:201,data:{...lesson,sheets_sync}};
}


function calendarGrid_(month, rows) {
  const [year, m] = month.split('-').map(Number);
  const first = new Date(Date.UTC(year, m - 1, 1)).getUTCDay();
  const days = new Date(Date.UTC(year, m, 0)).getUTCDate();
  const weeks = Math.ceil((first + days) / 7);
  const values = Array.from({length:weeks}, () => Array(7).fill(''));
  rows.sort((a,b) => String(a[5]).localeCompare(String(b[5])) || String(a[6]).localeCompare(String(b[6])) || String(a[0]).localeCompare(String(b[0])));
  for (let day = 1; day <= days; day++) {
    const date = month + '-' + String(day).padStart(2, '0');
    const lessons = rows.filter(r => String(r[5]) === date);
    const parts = [String(day)];
    const holiday = holidayName_(date);
    if (holiday) parts.push(holiday);
    for (const r of lessons) {
      const content = [];
      if (r[6] || r[7]) content.push(r[6] + '–' + r[7]);
      if (r[8]) content.push(String(r[8]));
      for (const [i,label] of [[9,'활동'],[10,'목표'],[11,'교재'],[12,'과제'],[13,'평가'],[14,'비고']]) {
        if (r[i]) content.push((i === 9 ? '활동' : label) + ': ' + r[i]);
      }
      if (content.length) parts.push(content.join('\n'));
    }
    const index = first + day - 1;
    values[Math.floor(index / 7)][index % 7] = parts.join('\n\n');
  }
  return {values, first, days, weeks};
}
function updateCalendar_(book, source, input) {
  const registry = loadRegistry_(book);
  const active = new Set(activeClasses_(registry).map(c => c.id));
  const enabled = registry.teachers.some(t => t.id === input.lesson.teacher_id && t.active !== false);
  const rows = source.getLastRow() > 1
    ? source.getRange(2,1,source.getLastRow()-1,18).getValues().filter(r => enabled && active.has(String(r[3])) && String(r[1]) === String(input.lesson.teacher_id)) : [];
  const props = PropertiesService.getScriptProperties();
  const key = 'CALENDAR_' + input.month + '_' + input.lesson.teacher_id;
  const stored = props.getProperty(key);
  let sheet = stored ? book.getSheets().find(s => String(s.getSheetId()) === stored) : null;
  if (!sheet) {
    const base = (input.month + '_' + input.teacher_name).replace(/[\[\]:*?\/\\]/g, '_').slice(0,80);
    let name = base, count = 2;
    while (book.getSheetByName(name)) name = base + '_' + count++;
    sheet = book.insertSheet(name);
    props.setProperty(key, String(sheet.getSheetId()));
  }
  const grid = calendarGrid_(input.month, rows);
  const area = sheet.getRange(1,1,9,7);
  area.breakApart(); area.clear();
  sheet.getRange(1,1,1,7).merge().setValue(input.month + ' 레슨플랜 캘린더').setFontSize(18).setFontWeight('bold');
  sheet.getRange(2,1,1,7).merge().setValue(input.teacher_name + ' 선생님 · ' + rows.length + '개 수업');
  sheet.getRange(3,1,1,7).setValues([['일','월','화','수','목','금','토']]).setBackground('#e9eef5').setFontWeight('bold');
  const body = sheet.getRange(4,1,grid.weeks,7);
  body.setNumberFormat('@').setValues(grid.values).setWrap(true).setVerticalAlignment('top').setFontSize(10);
  const backgrounds = grid.values.map(row => row.map(v => v ? '#ffffff' : '#f3f4f6'));
  const colors = grid.values.map((row,w) => row.map((v,c) => {
    const day = w*7+c-grid.first+1;
    const holiday = day > 0 && day <= grid.days ? holidayName_(input.month+'-'+String(day).padStart(2,'0')) : '';
    return c === 0 || holiday ? '#b22c40' : c === 6 ? '#2563a6' : '#202938';
  }));
  body.setBackgrounds(backgrounds).setFontColors(colors);
  sheet.getRange(3,1,grid.weeks+1,7).setBorder(true,true,true,true,true,true);
  sheet.setColumnWidths(1,7,150);
  sheet.setRowHeights(4,grid.weeks,135);
  sheet.autoResizeRows(4,grid.weeks);
  for(let row=4;row<4+grid.weeks;row++) if(sheet.getRowHeight(row)<135)sheet.setRowHeight(row,135);
  sheet.setFrozenRows(3);
}
function holidayName_(date) {
  const holidays = {"2026":{"01-01":"신정","02-16":"설날 연휴","02-17":"설날","02-18":"설날 연휴","03-01":"삼일절","03-02":"대체공휴일 (삼일절)","05-01":"노동절","05-05":"어린이날","05-24":"부처님오신날","05-25":"대체공휴일 (부처님오신날)","06-03":"전국동시지방선거","06-06":"현충일","07-17":"제헌절","08-15":"광복절","08-17":"대체공휴일 (광복절)","09-24":"추석 연휴","09-25":"추석","09-26":"추석 연휴","10-03":"개천절","10-05":"대체공휴일 (개천절)","10-09":"한글날","12-25":"성탄절"},"2027":{"01-01":"신정","02-06":"설날 연휴","02-07":"설날","02-08":"설날 연휴","02-09":"대체공휴일 (설날)","03-01":"삼일절","05-01":"노동절","05-03":"대체공휴일 (노동절)","05-05":"어린이날","05-13":"부처님오신날","06-06":"현충일","07-17":"제헌절","07-19":"대체공휴일 (제헌절)","08-15":"광복절","08-16":"대체공휴일 (광복절)","09-14":"추석 연휴","09-15":"추석","09-16":"추석 연휴","10-03":"개천절","10-04":"대체공휴일 (개천절)","10-09":"한글날","10-11":"대체공휴일 (한글날)","12-25":"성탄절","12-27":"대체공휴일 (성탄절)"}};
  return holidays[date.slice(0,4)]?.[date.slice(5)] || '';
}


// Minimal byte helpers for the existing dependency-free DOCX builder.
class WordBytes extends Array {
  static from(value) {return WordBytes.fromArray(Utilities.newBlob(String(value)).getBytes().map(x=>x<0?x+256:x));}
  static fromArray(values) {const out=new WordBytes();for(const x of values)out.push(x);return out;}
  static alloc(size) {return WordBytes.fromArray(Array(size).fill(0));}
  static concat(arrays) {const out=new WordBytes();for(const a of arrays)for(const x of a)out.push(x);return out;}
  writeUInt16LE(value,offset) {this[offset]=value&255;this[offset+1]=(value>>>8)&255;}
  writeUInt32LE(value,offset) {for(let i=0;i<4;i++)this[offset+i]=(value>>>(i*8))&255;}
}
const Buffer=WordBytes;
'use strict';
// Word Open XML package, stored ZIP entries; no external dependencies.
const holidays={supports:year=>[2026,2027].includes(Number(year)),name:date=>holidayName_(date)};
const xml=value=>String(value??'').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g,'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const declaration='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const crcTable=Array.from({length:256},(_,n)=>{for(let i=0;i<8;i++)n=(n&1)?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
function crc32_(buffer){let crc=0xffffffff;for(const byte of buffer)crc=crcTable[(crc^byte)&255]^(crc>>>8);return (crc^0xffffffff)>>>0;}
function zip_(entries){let offset=0;const local=[],central=[];for(const [name,text] of Object.entries(entries)){
  const filename=Buffer.from(name),data=Buffer.from(text,'utf8'),crc=crc32_(data),header=Buffer.alloc(30);
  header.writeUInt32LE(0x04034b50,0);header.writeUInt16LE(20,4);header.writeUInt16LE(0x0800,6);header.writeUInt16LE(33,12);header.writeUInt32LE(crc,14);header.writeUInt32LE(data.length,18);header.writeUInt32LE(data.length,22);header.writeUInt16LE(filename.length,26);
  local.push(header,filename,data);const record=Buffer.alloc(46);record.writeUInt32LE(0x02014b50,0);record.writeUInt16LE(20,4);record.writeUInt16LE(20,6);record.writeUInt16LE(0x0800,8);record.writeUInt16LE(33,14);record.writeUInt32LE(crc,16);record.writeUInt32LE(data.length,20);record.writeUInt32LE(data.length,24);record.writeUInt16LE(filename.length,28);record.writeUInt32LE(offset,42);central.push(record,filename);offset+=header.length+filename.length+data.length;
}const directory=Buffer.concat(central),end=Buffer.alloc(22);end.writeUInt32LE(0x06054b50,0);end.writeUInt16LE(Object.keys(entries).length,8);end.writeUInt16LE(Object.keys(entries).length,10);end.writeUInt32LE(directory.length,12);end.writeUInt32LE(offset,16);return Buffer.concat([...local,directory,end]);}
function paragraph_(value,style='Normal',color='') {
  const chunks=String(value??'').split(/\r\n|\n|\r/);
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r>${color?`<w:rPr><w:color w:val="${color}"/></w:rPr>`:''}${chunks.map((line,i)=>(i?'<w:br/>':'')+`<w:t xml:space="preserve">${xml(line)}</w:t>`).join('')}</w:r></w:p>`;
}
function createLessonDocument_({teacher,classes,lessons,month,selectedClass}) {
  const palette=[['8370C4','F0ECFB'],['328B85','E9F6F3'],['5588B9','EDF3FB'],['B98A40','FCF4E7']];
  const owned=classes.filter(c=>c.teacher_ids.includes(teacher.id));
  const colorFor=id=>palette[Math.max(0,owned.findIndex(c=>c.id===id))%palette.length];
  // Match the calendar's date order, then start time; stable order for equal times.
  const visible=lessons.filter(l=>!selectedClass||l.class_id===selectedClass.id);
  const byDate=new Map();
  for(const lesson of visible){if(!byDate.has(lesson.date))byDate.set(lesson.date,[]);byDate.get(lesson.date).push(lesson);}
  for(const entries of byDate.values())entries.sort((a,b)=>(a.start_time||'').localeCompare(b.start_time||''));
  const [year,m]=month.split('-').map(Number),first=new Date(year,m-1,1),last=new Date(year,m,0),weeks=Math.ceil((first.getDay()+last.getDate())/7);
  // Compact minimum heights let typical months fit one page while long content expands.
  const totalWidth=15878,cellWidth=2268,rowHeight=weeks===6?1000:1100;
  const borders=['top','left','bottom','right','insideH','insideV'].map(side=>`<w:${side} w:val="single" w:sz="4" w:color="E1E2EA"/>`).join('');
  let body=paragraph_(`${year}년 ${m}월 레슨플랜 캘린더`,'Title')+paragraph_(`${teacher.name} 선생님 · ${visible.length}개 수업`,'Subtitle');
  if(!holidays.supports(year))body+=paragraph_(`${year}년 공휴일 정보는 아직 준비되지 않았습니다.`,'Subtitle');
  body+=`<w:tbl><w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>${borders}</w:tblBorders><w:tblCellMar><w:top w:w="55" w:type="dxa"/><w:left w:w="65" w:type="dxa"/><w:bottom w:w="55" w:type="dxa"/><w:right w:w="65" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${Array.from({length:7},(_,i)=>`<w:gridCol w:w="${cellWidth+(i===6?2:0)}"/>`).join('')}</w:tblGrid>`;
  body+='<w:tr><w:trPr><w:tblHeader/></w:trPr>'+['일','월','화','수','목','금','토'].map((day,i)=>`<w:tc><w:tcPr><w:tcW w:w="${cellWidth+(i===6?2:0)}" w:type="dxa"/><w:shd w:fill="F5F4FB"/></w:tcPr>${paragraph_(day,'Weekday',i===0?'B22C40':i===6?'235AA5':'34415A')}</w:tc>`).join('')+'</w:tr>';
  for(let week=0;week<weeks;week++) {
    body+=`<w:tr><w:trPr><w:trHeight w:val="${rowHeight}" w:hRule="atLeast"/></w:trPr>`;
    for(let weekday=0;weekday<7;weekday++) {
      const date=new Date(year,m-1,1-first.getDay()+week*7+weekday);
      const outside=date.getMonth()!==m-1;
      const key=`${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
      const entries=outside?[]:(byDate.get(key)||[]);
      const holiday=outside?'':holidays.name(key);
      body+=`<w:tc><w:tcPr><w:tcW w:w="${cellWidth+(weekday===6?2:0)}" w:type="dxa"/><w:shd w:fill="${outside?'F7F8FC':holiday?'FFF8F8':'FFFFFF'}"/><w:vAlign w:val="top"/></w:tcPr>`+paragraph_(date.getDate(),'CalendarDate',outside?'C7CAD6':holiday||weekday===0?'B22C40':weekday===6?'235AA5':'34415A');
      if(holiday)body+=paragraph_(holiday,'EventHeading','B22C40');
      for(const l of entries) {
        const [color,fill]=colorFor(l.class_id);
        let content=(l.start_time?paragraph_(l.start_time+'–'+l.end_time,'EventHeading',color):'')+(l.title?paragraph_(l.title,'EventTitle',color):'');
        for(const [key,label] of [['objectives','목표'],['activities','활동'],['materials','교재·준비물'],['homework','과제'],['assessment','평가'],['notes','비고']])if(l[key])content+=paragraph_(`${label}: ${l[key]}`,'EventText');
        body+=`<w:tbl><w:tblPr><w:tblW w:w="${cellWidth-180}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:left w:val="single" w:sz="12" w:color="${color}"/></w:tblBorders><w:tblCellMar><w:top w:w="55" w:type="dxa"/><w:left w:w="70" w:type="dxa"/><w:bottom w:w="55" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="${cellWidth-180}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${cellWidth-180}" w:type="dxa"/><w:shd w:fill="${fill}"/></w:tcPr>${content}</w:tc></w:tr></w:tbl>`+paragraph_('','Spacer');
      }
      body+='</w:tc>';
    }
    body+='</w:tr>';
  }
  body+='</w:tbl>'+paragraph_('','Spacer');
  const document=declaration+`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="480" w:right="480" w:bottom="480" w:left="480" w:header="240" w:footer="240"/></w:sectPr></w:body></w:document>`;
  const style=(id,size,extra='',props='')=>`<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/><w:basedOn w:val="Normal"/><w:pPr>${props}</w:pPr><w:rPr><w:sz w:val="${size}"/>${extra}</w:rPr></w:style>`;
  const styles=declaration+'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="맑은 고딕"/><w:sz w:val="17"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'+style('Title',28,'<w:b/><w:color w:val="40328B"/>','<w:keepNext/><w:spacing w:after="60"/>')+style('Subtitle',18,'','<w:keepNext/><w:spacing w:after="40"/>')+style('Legend',16,'','<w:keepNext/><w:spacing w:after="60"/>')+style('Weekday',18,'<w:b/>','<w:jc w:val="center"/>')+style('CalendarDate',18,'<w:b/>','<w:keepNext w:val="0"/><w:spacing w:after="40"/>')+style('EventHeading',17,'<w:b/>')+style('EventTitle',17,'<w:b/>')+style('EventText',17)+style('Spacer',2,'','<w:keepNext w:val="0"/><w:spacing w:after="0" w:line="20" w:lineRule="exact"/>')+'</w:styles>';
  return zip_({
    '[Content_Types].xml':declaration+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    '_rels/.rels':declaration+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml':document,'word/styles.xml':styles,
    'word/_rels/document.xml.rels':declaration+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
  });
}

