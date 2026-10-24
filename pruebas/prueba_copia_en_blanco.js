const fs=require('fs');
eval(fs.readFileSync(process.argv[3],'utf8'));
const ss=T.SS; let fallos=0; const ok=(c,m)=>{console.log((c?'✔ ':'✘ ')+m); if(!c) fallos++;};
const libros={}; let k=0;
SpreadsheetApp.create=n=>{ const hojas=[new T.Sheet('Hoja')]; hojas[0].setName=function(x){this.name=x;return this}; const id='lib'+(++k);
  const l={getId:()=>id,getUrl:()=>'https://doc/'+id,getSheets:()=>hojas,insertSheet:x=>{const s=new T.Sheet(x);hojas.push(s);return s;},getSheetByName:x=>hojas.find(h=>h.name===x)}; libros[id]=l; T.files[id]={moveTo(){}}; return l; };
T.setSheets([new T.Sheet('Hoja 1')]);
instalarCatalogo({proyecto:'Acervo',lenguas:[{nombre_propio:'Wixárika',codigo_iso:'wix',atajos:'+=ɨ'},{nombre_completo:'Español',codigo_iso:'spa'}],categorias:'casa',carpetas:false});
_activarAutomatizacion=()=>{};
equipoGuardar({correo:'luis@ej.mx',rol:'Capturista',activo:'Sí'});
_agregar(_tabla(ss,'PALABRAS'),[{id_palabra:'P1',estado:'Publicado'}]); _agregar(_tabla(ss,'TRADUCCIONES'),[{id_traduccion:'T1',id_palabra:'P1',id_idioma:_s(readSheet(ss,'IDIOMAS')[0].id_idioma),texto:'ha',estado:'Publicado'}]);
_escribirConfig(ss,{ID_CARPETA_RAIZ:'carpetaVieja',APPSHEET_APP_ID:'app-vieja'}); _cfgMemo=null;
const idViejo=getAppConfig(ss).idCatalogo, tokViejo=getAppConfig(ss).apiToken;
ok(_esCopia(ss)===false,'la original se reconoce como original');
let o={tipo:'enBlanco',conservarLenguas:true,conservarAtajos:true,conservarEquipo:false};
let v=limpiezaVista(o); let e=''; try{ limpiezaAplicar(o,v.total,'BORRAR'); }catch(x){ e=x.message; } ok(e.includes('ORIGINAL'),'en la original no se deja en blanco');
// Simular copia: el ID guardado es de otra hoja
_escribirConfig(ss,{ID_HOJA:'otra'}); _cfgMemo=null;   // el ID guardado en la hoja es de otra (así llega una copia)
ok(_esCopia(ss)===true,'copia detectada');
v=limpiezaVista(o); console.log('   ',JSON.stringify(v.hojas));
const r=limpiezaAplicar(Object.assign({nombreNuevo:'Catálogo náayeri'},o),v.total,'BORRAR');
const cfg=getAppConfig(ss);
ok(readSheet(ss,'PALABRAS').length===0 && readSheet(ss,'TRADUCCIONES').length===0 && readSheet(ss,'CATEGORIAS').length===0,'vacío');
ok(readSheet(ss,'IDIOMAS').length===2 && readSheet(ss,'REGLAS_ESCRITURA').length>0,'conserva lenguas y atajos');
ok(readSheet(ss,'EQUIPO').length===1 && readSheet(ss,'FUENTES').length===1,'queda solo quien administra y la fuente del equipo');
ok(cfg.idCatalogo!==idViejo && cfg.apiToken!==tokViejo && !cfg.raiz && !cfg.appsheetId && cfg.titulo==='Catálogo náayeri','identidad nueva');
ok(_esCopia(ss)===false && !PropertiesService.getScriptProperties().getProperty('DIRECTORIO_VOCES_ID'),'ya es un catálogo propio');
ok(readSheet(ss,'HISTORIAL').length===1,'historial nuevo con el registro de la copia');
console.log(fallos?'FALLOS '+fallos:'TODO OK');
