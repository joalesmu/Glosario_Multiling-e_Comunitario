const fs=require('fs');
eval(fs.readFileSync(process.argv[3],'utf8'));
const ss=T.SS; let fallos=0; const ok=(c,m)=>{console.log((c?'✔ ':'✘ ')+m); if(!c) fallos++;};
const libros={}; let k=0;
SpreadsheetApp.create=n=>{ const hojas=[new T.Sheet('Hoja')]; hojas[0].setName=function(x){this.name=x;return this}; const id='lib'+(++k);
  const l={getId:()=>id,getUrl:()=>'https://doc/'+id,getSheets:()=>hojas,insertSheet:x=>{const s=new T.Sheet(x);hojas.push(s);return s;},getSheetByName:x=>hojas.find(h=>h.name===x)}; libros[id]=l; T.files[id]={moveTo(){}}; return l; };
SpreadsheetApp.openById=id=>libros[id];
global.MimeType={GOOGLE_SHEETS:'s'};
_listaRespaldos=()=>Object.keys(libros).map(id=>({id,nombre:'R '+id,fecha:'hoy'}));
T.setSheets([new T.Sheet('Hoja 1')]);
instalarCatalogo({proyecto:'Acervo',lenguas:[{nombre_propio:'Wixárika',codigo_iso:'wix',atajos:'+=ɨ'},{nombre_completo:'Español',codigo_iso:'spa'}],categorias:'casa',carpetas:false});
_activarAutomatizacion=()=>{};
const W=_s(readSheet(ss,'IDIOMAS')[0].id_idioma);
_agregar(_tabla(ss,'PALABRAS'),[{id_palabra:'P1',estado:'Publicado',usar_en_juegos:'TRUE'},{id_palabra:'P2',estado:'Publicado'}]);
_agregar(_tabla(ss,'TRADUCCIONES'),[{id_traduccion:'T1',id_palabra:'P1',id_idioma:W,texto:'ha',estado:'Publicado'},{id_traduccion:'T2',id_palabra:'P2',id_idioma:W,texto:'ki',estado:'Publicado'}]);
_escribirConfig(ss,{ID_CARPETA_RAIZ:'carpetaVieja',APPSHEET_APP_ID:'app-vieja'}); _cfgMemo=null;
const c0=getAppConfig(ss);
// 1. Vaciar el contenido: no cambia identidad
const o={tipo:'contenido'}; let v=limpiezaVista(o);
ok(v.hojas.some(h=>h.hoja==='PALABRAS') && !v.hojas.some(h=>['IDIOMAS','EQUIPO','FUENTES','REGLAS_ESCRITURA'].includes(h.hoja)),'contenido: '+JSON.stringify(v.hojas));
const r=limpiezaAplicar(o,v.total,'BORRAR'); _cfgMemo=null; const c1=getAppConfig(ss);
ok(readSheet(ss,'PALABRAS').length===0 && readSheet(ss,'TRADUCCIONES').length===0 && readSheet(ss,'CATEGORIAS').length===0 && readSheet(ss,'IDIOMAS').length===2,'vacía y conserva lenguas');
ok(c1.idCatalogo===c0.idCatalogo && c1.apiToken===c0.apiToken && c1.raiz===c0.raiz && c1.titulo===c0.titulo && !r.enBlanco,'identidad intacta');
// 2. Recuperar desde el respaldo
const idR=Object.keys(libros).pop();
let rv=limpiezaRestaurarVista(idR); ok(rv.total===v.total && !rv.config,'vista de recuperación: '+rv.total);
limpiezaRestaurar(idR,false,rv.total);
ok(readSheet(ss,'PALABRAS').length===2 && readSheet(ss,'TRADUCCIONES').length===2 && readSheet(ss,'CATEGORIAS').length===1,'filas de vuelta');
ok(readSheet(ss,'PALABRAS')[0].usar_en_juegos===true,'«TRUE» del respaldo vuelve como casilla');
rv=limpiezaRestaurarVista(idR); ok(rv.total===0 && rv.yaEstan===v.total,'recuperar dos veces no duplica');
// 3. Copia en blanco en la ORIGINAL: se rechaza
ok(_esCopia(ss)===false,'es la original');
let e=''; try{ const b={tipo:'enBlanco',conservarLenguas:true,conservarAtajos:true,nombreOriginal:c0.titulo}; limpiezaAplicar(b,limpiezaVista(b).total,'BORRAR'); }catch(x){ e=x.message; } ok(e.includes('ORIGINAL'),'original: no se deja en blanco aunque escriba el nombre');
// 4. En una copia: respaldo con configuración y carpetas; recuperar la configuración
_escribirConfig(ss,{ID_HOJA:'otra'}); _cfgMemo=null;   // el ID guardado en la hoja es de otra (así llega una copia)
ok(!!_asegurarIdentidad(ss),'la copia recibe identidad propia antes de cualquier cosa');
_cfgMemo=null; const cC=getAppConfig(ss); ok(cC.idCatalogo!==c0.idCatalogo && !cC.raiz && !cC.appsheetId,'copia: código nuevo, sin carpetas ni app del original');
_escribirConfig(ss,{ID_CARPETA_RAIZ:'carpetaCopia'}); _cfgMemo=null;
const tI=_tabla(ss,'IDIOMAS'); _actualizar(tI,0,{carpeta_audios:'audiosWix'});
const b={tipo:'enBlanco',conservarLenguas:true,conservarAtajos:true};
const rb=limpiezaAplicar(b,limpiezaVista(b).total,'BORRAR'); _cfgMemo=null; const c2=getAppConfig(ss);
ok(c2.idCatalogo!==cC.idCatalogo && c2.titulo===c0.titulo && rb.titulo===c0.titulo,'copia: identidad nueva y el título NO cambia solo');
ok(!_s(readSheet(ss,'IDIOMAS')[0].carpeta_audios),'olvida la carpeta de audios');
const idB=Object.keys(libros).pop(); rv=limpiezaRestaurarVista(idB);
ok(rv.config && rv.config.titulo===c0.titulo,'el respaldo trae la configuración');
limpiezaRestaurar(idB,true,rv.total); _cfgMemo=null; const c3=getAppConfig(ss);
ok(c3.idCatalogo===cC.idCatalogo && c3.apiToken===cC.apiToken && c3.raiz==='carpetaCopia' && !c3.appsheetId,'configuración recuperada (la de la copia, nunca la del original)');
ok(readSheet(ss,'IDIOMAS')[0].carpeta_audios==='audiosWix' && readSheet(ss,'PALABRAS').length===2,'carpetas de lenguas y palabras de vuelta');
console.log(fallos?'FALLOS '+fallos:'TODO OK');
