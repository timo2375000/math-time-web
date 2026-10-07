'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('node:assert/strict'),crypto=require('crypto'),path=require('path');
class Sheet {
 constructor(name,id){this.name=name;this.id=id;this.rows=[];this.heights={};}
 getName(){return this.name;}getSheetId(){return this.id;}getLastRow(){return this.rows.length;}
 appendRow(row){this.rows.push([...row]);return this;}
 getRange(r,c,n=1,m=1){const sheet=this;const range={
  getValues(){return Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>sheet.rows[r-1+i]?.[c-1+j]??''));},
  getValue(){return this.getValues()[0][0];},
  setValues(values){for(let i=0;i<n;i++){sheet.rows[r-1+i]??=[];for(let j=0;j<m;j++){let v=values[i][j];if(typeof v==='string'&&v.startsWith("'"))v=v.slice(1);sheet.rows[r-1+i][c-1+j]=v;}}return this;},
  setValue(v){return this.setValues([[v]]);},
  clear(){for(let i=0;i<n;i++)for(let j=0;j<m;j++)if(sheet.rows[r-1+i])sheet.rows[r-1+i][c-1+j]='';return this;}
 };for(const name of ['setNumberFormat','setFrozenRows','setFontSize','setFontWeight','merge','breakApart','setBackground','setBackgrounds','setWrap','setVerticalAlignment','setFontColors','setBorder'])range[name]=()=>range;return range;}
 setFrozenRows(){}setColumnWidths(){}setRowHeights(r,n,h){for(let i=0;i<n;i++)this.heights[r+i]=h;}autoResizeRows(){}getRowHeight(r){return this.heights[r]||21;}setRowHeight(r,h){this.heights[r]=h;}
}
const book={sheets:[],getSheets(){return this.sheets;},getSheetByName(n){return this.sheets.find(s=>s.name===n)||null;},insertSheet(n){if(this.getSheetByName(n))throw Error('duplicate');const s=new Sheet(n,this.sheets.length+1);this.sheets.push(s);return s;}};
const props=new Map([['SPREADSHEET_ID','mock'],['APP_PASSWORD','Test1234']]),cache=new Map();
const lock={held:false,waitLock(){this.held=true;},hasLock(){return this.held;},releaseLock(){this.held=false;}};
const context={console:{error(){}},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v)})},CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v)})},LockService:{getScriptLock:()=>lock},SpreadsheetApp:{openById:id=>{assert.equal(id,'mock');return book;}},Utilities:{getUuid:()=>crypto.randomUUID(),DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(algo,value)=>Array.from(crypto.createHash(algo).update(value).digest()),newBlob:s=>({getBytes:()=>Array.from(Buffer.from(s))}),base64Encode:bytes=>Buffer.from(bytes.map(x=>x<0?x+256:x)).toString('base64')}};
vm.createContext(context);for(const name of ['Code.gs'])vm.runInContext(fs.readFileSync(path.join(__dirname,'../apps-script',name),'utf8'),context,{filename:name});
for(const invalid of ['Abc1234','abcdefgh','12345678','Abcd12!@']){props.set('APP_PASSWORD',invalid);assert.throws(()=>context.authLogin(invalid));}
props.set('APP_PASSWORD','Test1234');
assert.throws(()=>context.authLogin('wrong'));const session=context.authLogin('Test1234');
const api=(p,m='GET',b={})=>context.apiRequest(session,p,m,b);
assert.equal(context.apiRequest('bad','/api/bootstrap','GET',{}).status,401);
// Import legacy connected sheet without changing IDs.
const source=book.insertSheet('2026-10');source.appendRow(Array(18).fill('header'));source.appendRow(['old-lesson','old-teacher','기존 선생님','old-class','기존 반','2026-10-06','','','','기존 수업','','','','','','draft','1','2026-10-06T00:00:00Z']);
let bootstrap=api('/api/bootstrap');assert.equal(bootstrap.data.teachers[0].id,'old-teacher');assert.equal(bootstrap.data.classes[0].id,'old-class');
const teacher=api('/api/teachers','POST',{name:'교사 A'}).data;
const other=api('/api/teachers','POST',{name:'교사 B'}).data;
const cl=api('/api/classes','POST',{name:'초등 반',level:'초등',teacher_ids:[teacher.id]}).data;
const url=`/api/teachers/${teacher.id}/months/2026-10/lessons`;
const input={date:'2026-10-07',class_id:cl.id,activities:'덧셈\n문제풀이',content_only:true};
let response=api(url,'POST',input);assert.equal(response.status,201);assert.equal(response.data.sheets_sync.status,'synced');const lesson=response.data;
assert.equal(api(url).data.length,1);assert.equal(api(`/api/teachers/${other.id}/months/2026-10/lessons`).data.length,0);
let calendar=book.getSheetByName('2026-10_교사 A');assert(calendar);assert(calendar.rows.flat().some(v=>String(v).includes('덧셈\n문제풀이')));
response=api(url+'/'+lesson.id,'PUT',{...input,activities:'뺄셈',revision:1});assert.equal(response.status,200);assert.equal(response.data.revision,2);assert.equal(source.getLastRow(),3);assert(calendar.rows.flat().some(v=>String(v).includes('뺄셈')));assert(!calendar.rows.flat().some(v=>String(v).includes('덧셈')));
assert.equal(api(url+'/'+lesson.id,'PUT',{...input,revision:1}).status,409);
assert.equal(api(url,'POST',{...input,date:'2026-10-32'}).status,400);
assert.equal(api(url,'POST',{...input,status:'completed'}).status,400);
assert.equal(api('/api/lesson-plan-archive').data[0].teachers.length,2);
const doc=api(`/api/teachers/${teacher.id}/months/2026-10/export.docx`);assert.equal(doc.status,200);const bytes=Buffer.from(doc.data.base64,'base64');assert.equal(bytes.readUInt32LE(0),0x04034b50);assert(bytes.includes(Buffer.from('뺄셈')));assert(!bytes.includes(Buffer.from('w:type="paragraph_"')));assert(bytes.includes(Buffer.from('w:type="paragraph"')));fs.writeFileSync('/tmp/math-time-free-test.docx',bytes);
assert.equal(api('/api/classes/'+cl.id,'DELETE').status,200);assert.equal(api(url).data.length,0);assert(!calendar.rows.flat().some(v=>String(v).includes('뺄셈')));assert.equal(source.getLastRow(),3);
cache.delete('session:'+session);assert.equal(api('/api/bootstrap').status,401);
console.log('PASS authentication, legacy import, registration, teacher separation, save/edit/revision, validation, archive, calendar refresh/deletion, DOCX, expired session');
