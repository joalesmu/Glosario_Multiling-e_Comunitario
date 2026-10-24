const fs=require('fs'); eval(fs.readFileSync(process.argv[3],'utf8'));
const ss=T.SS; let fallos=0; const ok=(c,m)=>{console.log((c?'✔ ':'✘ ')+m); if(!c) fallos++;};
global.DLG={}; global.HtmlService={createHtmlOutput:h=>({h,setWidth(){return this},setHeight(){return this},setTitle(){return this}})};
const ui0=SpreadsheetApp.getUi; SpreadsheetApp.getUi=()=>Object.assign(ui0(),{showModalDialog:(o,t)=>{DLG[t]=o.h;},showSidebar:o=>{DLG.sidebar=o.h;}});
T.setSheets([new T.Sheet('Hoja 1')]);
instalarCatalogo({proyecto:'Acervo',lenguas:[{nombre_propio:'Wixárika',codigo_iso:'wix',atajos:'+=ɨ'},{nombre_completo:'Español',codigo_iso:'spa'}],categorias:'casa',carpetas:false});
_activarAutomatizacion=()=>{};
const W=_s(readSheet(ss,'IDIOMAS')[0].id_idioma), E=_s(readSheet(ss,'IDIOMAS')[1].id_idioma);
// catálogo anterior: sin columna descripcion
const shT=ss.getSheetByName('TRADUCCIONES'); ok(_encabezados(shT).includes('descripcion'),'la estructura nueva trae la columna descripcion');
_agregar(_tabla(ss,'PALABRAS'),[{id_palabra:'P1',estado:'Publicado'},{id_palabra:'P2',estado:'Publicado'}]);
_agregar(_tabla(ss,'TRADUCCIONES'),[{id_traduccion:'T1',id_palabra:'P1',id_idioma:W,texto:'kiitɨ',estado:'Publicado'},{id_traduccion:'T2',id_palabra:'P1',id_idioma:E,texto:'casa',estado:'Publicado'},
  {id_traduccion:'T3',id_palabra:'P2',id_idioma:W,texto:'tsɨkɨ',estado:'Publicado'}]);
let d=editarDatos(); ok(d.palabras.every(p=>!Object.keys(p.d).length),'sin descripciones al inicio');
let r=editarGuardar({id_palabra:'P1',trads:{[W]:{descripcion:'Ki k+ teteuyeika'}}});
ok(readSheet(ss,'TRADUCCIONES').find(t=>t.id_traduccion==='T1').descripcion==='Ki kɨ teteuyeika','guarda la descripción con atajos de escritura');
ok(r.ficha.trads[W].descripcion==='Ki kɨ teteuyeika' && editarDatos().palabras.find(p=>p.id==='P1').d[W]===1,'ficha y filtro saben que tiene descripción');
r=editarGuardar({id_palabra:'P2',trads:{[W]:{descripcion:'x'.repeat(900)}}}); ok(readSheet(ss,'TRADUCCIONES').find(t=>t.id_traduccion==='T3').descripcion.length===600,'tope de largo');
// capturista: corrige descripción publicada → revisión
const adm=_soyAdmin; _soyAdmin=()=>false;
r=editarGuardar({id_palabra:'P1',trads:{[W]:{descripcion:'Ki kɨ teteuyeika, tatei'}}});
ok(readSheet(ss,'TRADUCCIONES').find(t=>t.id_traduccion==='T1').estado==='En revisión','quien captura cambia una descripción: vuelve a revisión');
_soyAdmin=adm; editarGuardar({id_palabra:'P1',trads:{[W]:{estado:'publicado'}}});
// publicación
const pub=construirDatosPublicos(ss); const t1=pub.traducciones.find(t=>t.id_traduccion==='T1'), t2=pub.traducciones.find(t=>t.id_traduccion==='T2');
ok(t1.descripcion==='Ki kɨ teteuyeika, tatei' && !('descripcion' in t2),'publica la descripción solo donde existe');
// plantilla y traer
const hdr=[]; const fake={create:()=>({getActiveSheet:()=>({setName(){return this},getRange:()=>({setValues(v){hdr.push(v);return this},setBackground(){return this},setFontColor(){return this},setFontWeight(){return this}}),setFrozenRows(){},setColumnWidths(){}}),insertSheet:()=>({getRange:()=>({setValues(){return this},setFontWeight(){return this},setFontSize(){return this}}),setColumnWidth(){return this}}),getId:()=>'id',getUrl:()=>'u'})};
try { const o=SpreadsheetApp.create; SpreadsheetApp.create=fake.create; _aColaboraciones=(s,x)=>x; servicioGenerarPlantillaPro([W+'|WIX'],'NUEVAS','','',''); SpreadsheetApp.create=o; } catch(e){ console.log('(plantilla)',e.message); }
ok(hdr.some(v=>v[0].includes('DESCRIPCION_WIX')),'la hoja de colaboración trae DESCRIPCION_<lengua>');
_actualizar(_tabla(ss,'IDIOMAS'),0,{codigo_internacional:'hch'});
const ext=new T.Sheet('DATOS_A_IMPORTAR',[['ID_PALABRA','CATEGORIA','TRADUCCION_WIX','DESCRIPCION_WIX'],['P2','','tsɨkɨ','Teteuyeika me yuawi'],['','casa','maxa','Kiekari mieme']]);
SpreadsheetApp.openByUrl=()=>({getSheetByName:n=>n==='DATOS_A_IMPORTAR'?ext:null,getSheets:()=>[ext],getName:()=>'x'});
servicioProcesarImportacionMasiva('x','revisión',[]);
const TR=readSheet(ss,'TRADUCCIONES');
ok(TR.find(t=>t.id_traduccion==='T3').descripcion==='Teteuyeika me yuawi' && TR.some(t=>t.texto==='maxa' && t.descripcion==='Kiekari mieme'),'traer: actualiza y agrega descripciones');
// intercambio
let paq=null; try { const g=Object.getOwnPropertyNames(globalThis).filter(n=>/^servicioExportar|^exportarPaquete|^_paqueteIntercambio/.test(n)); } catch(e){}
// grafo
const g=_construirGrafo(ss,true); const nt=_grafoNTriples(g);
ok(/skos\/core#definition> "Ki kɨ teteuyeika, tatei"/.test(nt),'N-Triples: skos:definition con la descripción');
ok(_grafoGraphML(g).includes('<data key="descripcion">'),'GraphML con descripción');
// leyenda de impresión
visorAjustesGuardar({impresion:{leyenda:'Material comunitario,   prohibida su venta'}});
ok(_estadoParaVisor(_estadoVisor(),'publico').leyenda==='Material comunitario, prohibida su venta','leyenda de impresión llega al visor');
// panel y ventanas
uiPanel(); ok(DLG.sidebar.includes('Sistema diseñado y desarrollado por Alejandro Estrada'),'crédito en el Panel');
uiEditar(); uiVerEnlaceApi();
for (const [t,h] of Object.entries(DLG)) [...h.matchAll(/<script>([\s\S]*?)<\/script>/g)].forEach((x,i)=>{ try{ new Function(x[1]); }catch(err){ ok(false,'sintaxis '+t+' #'+i+': '+err.message); } });
ok(Object.values(DLG).some(h=>h.includes('vLeyenda')),'campo de leyenda en Ajustes del visor');
console.log(fallos?'FALLOS '+fallos:'TODO OK');
