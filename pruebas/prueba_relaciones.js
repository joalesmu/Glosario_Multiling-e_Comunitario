const fs=require('fs'); eval(fs.readFileSync(process.argv[3],'utf8'));
const ss=T.SS; let fallos=0; const ok=(c,m)=>{console.log((c?'✔ ':'✘ ')+m); if(!c) fallos++;};
global.DLG={}; global.HtmlService={createHtmlOutput:h=>({h,setWidth(){return this},setHeight(){return this},setTitle(){return this}})};
const ui0=SpreadsheetApp.getUi; SpreadsheetApp.getUi=()=>Object.assign(ui0(),{showModalDialog:(o,t)=>{DLG[t]=o.h;},showSidebar:o=>{DLG.sidebar=o.h;}});
T.setSheets([new T.Sheet('Hoja 1')]);
instalarCatalogo({proyecto:'Acervo',lenguas:[{nombre_propio:'Wixárika',codigo_iso:'wix'},{nombre_completo:'Español',codigo_iso:'spa'}],categorias:'casa',carpetas:false});
_activarAutomatizacion=()=>{};
const W=_s(readSheet(ss,'IDIOMAS')[0].id_idioma), E=_s(readSheet(ss,'IDIOMAS')[1].id_idioma);
ok(!!ss.getSheetByName('TIPOS_RELACION') && readSheet(ss,'TIPOS_RELACION').length===TIPOS_REL_BASE.length,'hoja de tipos con los básicos');
_agregar(_tabla(ss,'PALABRAS'),['P1','P2','P3','P4'].map(id=>({id_palabra:id,estado:'Publicado'})));
_agregar(_tabla(ss,'TRADUCCIONES'),[['P1','casa','kii'],['P2','techo','kitsi'],['P3','hogar','x'],['P4','perro','tsɨkɨ']].flatMap(([p,e,w],i)=>[
 {id_traduccion:'TE'+i,id_palabra:p,id_idioma:E,texto:e,estado:'Publicado'},{id_traduccion:'TW'+i,id_palabra:p,id_idioma:W,texto:w,estado:'Publicado'}]));
_escribirConfig(ss,{API_TOKEN:''}); _cfgMemo=null;
// relaciones escritas a mano con nombres sueltos
_agregar(_tabla(ss,'RELACIONES'),[{id_relacion:'R0',id_origen:'P1',tipo:'Tiene como parte',id_destino:'P2',estado:'Publicado'},
 {id_relacion:'R00',id_origen:'P1',tipo:'Se usa en',id_destino:'P4',estado:'Publicado'}]);
let td=tiposRelacionDatos();
const r0=readSheet(ss,'RELACIONES').find(r=>r.id_relacion==='R0');
ok(td.normalizadas===1 && r0.tipo==='parte_de' && r0.id_origen==='P2' && r0.id_destino==='P1','«Tiene como parte» escrito a mano → parte_de al revés');
ok(td.tipos.some(t=>t.id==='Se usa en' && t.nombre==='Se usa en') && readSheet(ss,'TIPOS_RELACION').some(t=>t.id_tipo==='Se usa en'),'tipo desconocido se agrega al vocabulario');
// opciones
let d=editarDatos(); ok(d.relOpciones.some(o=>o.v==='parte_de|<' && o.t==='Tiene como parte') && !d.relOpciones.some(o=>o.v.startsWith('unida a')),'opciones con dirección; «unida a» no se elige');
// agregar
let r=relacionAgregar({desde:'P1',hacia:'P3',opcion:'sinonimo|>',nota:'en la sierra'}); ok(r.relaciones.some(x=>x.tipo==='sinonimo' && x.texto==='hogar' && x.estado==='publicado'),'agrega sinónimo publicado (admin)');
let e=''; try{relacionAgregar({desde:'P3',hacia:'P1',opcion:'sinonimo|>'});}catch(x){e=x.message;} ok(e.includes('ya existe'),'simétrica: no duplica al revés');
e=''; try{relacionAgregar({desde:'P1',hacia:'P1',opcion:'sinonimo|>'});}catch(x){e=x.message;} ok(e.includes('consigo'),'no consigo misma');
e=''; try{relacionAgregar({desde:'P1',hacia:'P2',opcion:'unida a|>'});}catch(x){e=x.message;} ok(e.includes('no existe'),'no permite «unida a»');
let f=editarFicha('P1'); const rp=f.relaciones.find(x=>x.tipo==='parte_de'); ok(rp && rp.entra && rp.etiqueta==='Tiene como parte' && rp.texto==='techo','ficha lee la dirección inversa');
ok(editarFicha('P2').relaciones.find(x=>x.tipo==='parte_de').etiqueta==='Parte de','y la directa desde la otra palabra');
// tipo propio
td=tipoRelacionGuardar({id:'',nombre:'Se cocina con',nombre_inverso:'Se usa para cocinar',simetrica:false,publica:true});
const nuevo=td.tipos.find(t=>t.nombre==='Se cocina con'); ok(nuevo && nuevo.id==='se_cocina_con','tipo propio con código estable');
e=''; try{tipoRelacionGuardar({id:'',nombre:'sinónimo'});}catch(x){e=x.message;} ok(e.includes('Ya hay'),'no repite nombres (aunque sea de fábrica)');
tipoRelacionGuardar({id:'sinonimo',nombre:'Mismo significado',publica:true}); ok(_etiquetaTipoRel(_tiposRelacion(ss).find(t=>t.id==='sinonimo'))==='Mismo significado','renombrar uno de fábrica');
// privado y retirado
relacionAgregar({desde:'P4',hacia:'P1',opcion:'se_cocina_con|>'});
tipoRelacionGuardar({id:'se_cocina_con',nombre:'Se cocina con',nombre_inverso:'Se usa para cocinar',publica:false});
let pub=construirDatosPublicos(ss);
ok(pub.relaciones.some(x=>x.tipo==='sinonimo') && pub.relaciones.some(x=>x.tipo==='parte_de') && !pub.relaciones.some(x=>x.tipo==='se_cocina_con'),'publica solo tipos visibles');
ok(pub.tiposRelacion.find(t=>t.id==='sinonimo').nombre==='Mismo significado' && pub.tiposRelacion.find(t=>t.id==='parte_de').clave==='parte_de' && !pub.relaciones[0].estado && !pub.relaciones[0].creado_por,'tipos publicados y sin datos internos');
// retirar relación
const idS=editarFicha('P1').relaciones.find(x=>x.tipo==='sinonimo').id;
r=relacionRetirar(idS,'P1'); ok(!r.relaciones.some(x=>x.tipo==='sinonimo') && readSheet(ss,'RELACIONES').find(x=>x.id_relacion===idS).estado==='Retirado','retirar no borra');
r=relacionAgregar({desde:'P1',hacia:'P3',opcion:'sinonimo|>'}); ok(readSheet(ss,'RELACIONES').filter(x=>x.tipo==='sinonimo').length===1 && r.relaciones.some(x=>x.tipo==='sinonimo'),'reactivar en vez de duplicar');
// quien captura
const adm=_soyAdmin; _soyAdmin=()=>false;
r=relacionAgregar({desde:'P2',hacia:'P4',opcion:'relacionada|>'}); ok(r.aviso && r.relaciones.find(x=>x.tipo==='relacionada').estado==='revisión','captura: queda en revisión');
e=''; try{relacionRetirar(idS,'P1');}catch(x){e=x.message;} ok(e.includes('Solo quien administra'),'captura no quita publicadas');
_soyAdmin=adm;
// tipo retirado deja de verse
tipoRelacionGuardar({id:'Se usa en',nombre:'Se usa en',activo:false}); pub=construirDatosPublicos(ss); ok(!pub.relaciones.some(x=>x.tipo==='Se usa en') && !editarDatos().relOpciones.some(o=>o.v.startsWith('Se usa en')),'tipo retirado: ni se ve ni se elige');
// grafo
const g=_construirGrafo(ss,true); const ar=g.aristas.filter(a=>a.tipo==='RELACION'); ok(ar.some(a=>a.props.tipo==='parte_de' && a.props.nombre==='Parte de' && a.props.inverso==='Tiene como parte') && !ar.some(a=>a.props.tipo==='se_cocina_con'),'grafo publicado con nombres y sin privados');
const nt=_grafoNTriples(g); ok(nt.includes('relacion:parte_de> <http://www.w3.org/2000/01/rdf-schema#label> "Parte de"'),'N-Triples declara el nombre');
ok(_construirGrafo(ss,false).aristas.some(a=>a.props.tipo==='se_cocina_con'),'grafo completo incluye internos');
const csv=_grafoCSV(g); ok(true,'csv '+(Array.isArray(csv)?csv.length:typeof csv));
const gm=_grafoGraphML(g); ok(gm.includes('<data key="nombre">Parte de</data>'),'GraphML con nombre');
// importación: tipo desconocido
const n=_tiposDesdePaquete(ss,[{id:'ritual',nombre:'Se ofrece en',inverso:'Recibe',simetrica:false}],['ritual','sinonimo']);
ok(n===1 && readSheet(ss,'TIPOS_RELACION').find(t=>t.id_tipo==='ritual').nombre==='Se ofrece en','importa tipos que no existen');
// panel
ok(panelEstado().sinDireccion===true,'panel avisa falta de dirección');
// diálogos: sintaxis de sus scripts
uiTiposRelacion(); uiEditar(); uiPanel();
for (const [t,h] of Object.entries(DLG)) { const sc=[...h.matchAll(/<script>([\s\S]*?)<\/script>/g)]; sc.forEach((x,i)=>{ try{ new Function(x[1]); }catch(err){ ok(false,'sintaxis '+t+' #'+i+': '+err.message); } }); }
ok(Object.keys(DLG).length===3,'tres ventanas generadas');
console.log(fallos?'FALLOS '+fallos:'TODO OK');
