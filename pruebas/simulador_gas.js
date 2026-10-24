const fs=require('fs');
function Sheet(name, data){ this.name=name; this.d=data||[]; this.hidden=false; }
const chain=()=>new Proxy({}, {get:(t,p)=> p==='build'?()=>({}):()=>chain()});
Sheet.prototype={
  getName(){return this.name},
  getDataRange(){ return this.getRange(1,1,Math.max(this.d.length,1),Math.max(...this.d.map(r=>r.length),1)); },
  getLastRow(){ let n=0; this.d.forEach((r,i)=>{ if(r.some(v=>v!==''&&v!=null)) n=i+1;}); return n; },
  getLastColumn(){ return Math.max(0,...this.d.map(r=>{let n=0;r.forEach((v,i)=>{if(v!==''&&v!=null)n=i+1});return n;})); },
  getMaxRows(){ return this._max || 1000; }, getFrozenRows(){ return 1; }, insertRowsAfter(a,n){ this._max=(this._max||1000)+n; }, getMaxColumns(){ return 40; }, insertColumnsAfter(){},
  setFrozenRows(){}, setTabColor(){}, hideSheet(){this.hidden=true}, showSheet(){this.hidden=false}, isSheetHidden(){return this.hidden},
  getProtections(){ return []; }, deleteRow(i){ this.d.splice(i-1,1); }, deleteRows(i,n){ if(this._estricto && this.getMaxRows()-n-this.getFrozenRows()<1) throw new Error('No se pueden eliminar todas las filas que no estén inmovilizadas.'); this.d.splice(i-1,n); if(this._max) this._max-=n; }, hideColumns(){return this}, setColumnWidths(){return this}, setColumnWidth(){return this}, activate(){}, clear(){ this.d=[]; }, clearContents(){ this.d=[]; }, setFrozenColumns(){}, getRangeList(){ return chain(); },
  getRange(r,c,nr=1,nc=1){ const sh=this; if(typeof r==='string'){c=1;r=1;} return {
    getValues(){ const out=[]; for(let i=0;i<nr;i++){const row=[];for(let j=0;j<nc;j++){row.push((sh.d[r-1+i]||[])[c-1+j]??'');}out.push(row);} return out; },
    getFormulas(){ return this.getValues().map(r=>r.map(()=>'')); },
    getRichTextValues(){ return this.getValues().map(r=>r.map(v=>({getLinkUrl:()=>null,getRuns:()=>[]}))); },
    getValue(){ return this.getValues()[0][0]; },
    setValues(v){ v.forEach((row,i)=>row.forEach((x,j)=>{ while(sh.d.length<r+i) sh.d.push([]); const rr=sh.d[r-1+i]; while(rr.length<c+j) rr.push(''); rr[c-1+j]=x; })); return this; },
    setValue(x){ return this.setValues([[x]]); }, setWrap(){return this}, setFontSize(){return this}, setVerticalAlignment(){return this}, setFontStyle(){return this}, setNumberFormat(){ return this; }, setFontWeight(){return this}, setBackground(){return this},
    setFontColor(){return this}, setFontSize(){return this}, setDataValidation(){return this}, clearDataValidations(){return this}, setRichTextValues(){return this}, setBorder(){return this}, protect(){ return chain(); } }; },
};
const store={};
global.CacheService={getScriptCache:()=>({getAll:ks=>{const o={};ks.forEach(k=>{if(store[k]!=null)o[k]=store[k]});return o;},putAll:o=>Object.assign(store,o),get:k=>store[k]??null,put:(k,v)=>store[k]=v,remove:k=>delete store[k]})};
const props={}, uprops={}; global.PropertiesService={getScriptProperties:()=>({getProperty:k=>props[k]??null,setProperty:(k,v)=>props[k]=v}),getUserProperties:()=>({getProperty:k=>uprops[k]??null,setProperty:(k,v)=>uprops[k]=v,deleteProperty:k=>delete uprops[k]})};
global.LockService={getScriptLock:()=>({waitLock(){},tryLock(){return true},releaseLock(){}})};
global.Session={ getActiveUser:()=>({getEmail:()=>'Ana@ej.mx'}), getEffectiveUser:()=>({getEmail:()=>'ana@ej.mx'}), getScriptTimeZone:()=>'UTC' };
let n=0; global.Utilities={ getUuid:()=>'u'+(++n)+'-xxxxxxxx', formatDate:()=> '2026-10-19', computeDigest:(a,s)=>[...require('crypto').createHash('md5').update(s).digest()].map(b=>b>127?b-256:b),
  DigestAlgorithm:{MD5:1}, Charset:{UTF_8:1}, newBlob:(c,t,n)=>({c,t,n,setName(x){this.n=x;return this},getName(){return this.n}}), zip:(arr,n)=>({arr,n,setName(x){this.n=x;return this},getName(){return this.n}}), base64Encode:x=>'B64' };
const files={};
global.DriveApp={ getFileById:id=>{ if(!files[id]) throw new Error('no'); return files[id]; }, getFilesByName:nm=>{const l=Object.values(files).filter(f=>f.name===nm);let i=0;return{hasNext:()=>i<l.length,next:()=>l[i++]}},
  Access:{ANYONE_WITH_LINK:'L',ANYONE:'A',PRIVATE:'P'}, Permission:{VIEW:'V'},
  getRootFolder:()=>folder('root') };
function folder(nm){ return { getId:()=>'f_'+nm, getFoldersByName:()=>({hasNext:()=>false}), createFolder:(x)=>folder(x), createFile:(a,b,c)=>{ const name=typeof a==='string'?a:a.getName(); const f={name,getUrl:()=>'https://drive/'+name,getName:()=>name,getId:()=>'id_'+name,setContent(){},isTrashed:()=>false}; files['exp_'+name]=f; return f; }, getFiles:()=>({hasNext:()=>false}), getFolders:()=>({hasNext:()=>false}) }; }
function archivo(id,name){ files[id]={name, getId:()=>id, isTrashed:()=>false, getSharingAccess(){return this.acc||'P'}, setSharing(a){this.acc=a}, getBlob:()=>({getBytes:()=>[1,2],getContentType:()=>'audio/mpeg'}) }; }
global.ContentService={createTextOutput:t=>({setMimeType(){return t;}}),MimeType:{JSON:1}};
global.ScriptApp={ getService:()=>({getUrl:()=>'https://script/exec'}), getProjectTriggers:()=>[], newTrigger:()=>chain(), deleteTrigger(){}, getOAuthToken:()=>'tok' };
global.UrlFetchApp={ fetch:()=>({getHeaders:()=>({'Content-Type':'image/jpeg'}),getResponseCode:()=>200,getContent:()=>[1],getBlob:()=>({setName(n){this.n=n;return this},getName(){return this.n}})}) };
let sheets;
const SS={ getSheetByName:n=>sheets.find(s=>s.name===n)||null, insertSheet:n=>{const s=new Sheet(n);sheets.push(s);return s;}, getSheets:()=>sheets.slice(), deleteSheet:s=>{sheets=sheets.filter(x=>x!==s)}, getName:()=>'Prueba', getId:()=>'ssid', editores:new Set(), addEditor(c){this.editores.add(c)}, removeEditor(c){this.editores.delete(c)}, getOwner:()=>({getEmail:()=>'ana@ej.mx'}) };
global.SpreadsheetApp={ flush(){}, getActiveSpreadsheet:()=>SS, newDataValidation:()=>chain(), newRichTextValue:()=>chain(), ProtectionType:{RANGE:'r'}, getUi:()=>({ alert:(...a)=>{ global.ULTIMO_ALERT=a; return 'YES'; }, Button:{YES:'YES',NO:'NO',CANCEL:'C',CLOSE:'X'}, ButtonSet:{YES_NO:1,YES_NO_CANCEL:2} }) };
files['ssid']={getParents:()=>({hasNext:()=>false})};
let src=fs.readFileSync(process.argv[2],'utf8').replace(/^const /gm,'var ').replace(/^let /gm,'var '); eval(src);
module.exports={uprops, SS, setSheets:x=>sheets=x, getSheets:()=>sheets, archivo, files, store, props, Sheet};
global.T={uprops, SS, setSheets:x=>sheets=x, sheets:()=>sheets, archivo, files, Sheet};
