// Replicar un catálogo: copia de la hoja (Archivo → Hacer una copia o «Copiar app» de AppSheet), plantilla y casos límite.
const fs=require('fs'); eval(fs.readFileSync(process.argv[3],'utf8'));
const ss=T.SS; let fallos=0; const ok=(c,m)=>{console.log((c?'✔ ':'✘ ')+m); if(!c) fallos++;};
const libros={}; let k=0;
SpreadsheetApp.create=n=>{ const hojas=[new T.Sheet('Hoja')]; hojas[0].setName=function(x){this.name=x;return this}; const id='lib'+(++k);
  const l={getId:()=>id,getUrl:()=>'https://doc/'+id,getSheets:()=>hojas,insertSheet:x=>{const s=new T.Sheet(x);hojas.push(s);return s;},getSheetByName:x=>hojas.find(h=>h.name===x)}; libros[id]=l; T.files[id]={moveTo(){}}; return l; };
const P=PropertiesService.getScriptProperties(); P.deleteProperty=k=>{ delete P._d; };
const props={}; PropertiesService.getScriptProperties=()=>({getProperty:k=>props[k]??null,setProperty:(k,v)=>{props[k]=v},deleteProperty:k=>{delete props[k]}});
const dprops={}; PropertiesService.getDocumentProperties=()=>({getProperty:k=>dprops[k]??null,setProperty:(k,v)=>{dprops[k]=v},deleteProperty:k=>{delete dprops[k]}});
let duenio='ana@ej.mx'; ss.getOwner=()=>({getEmail:()=>duenio});
let idHoja='orig'; ss.getId=()=>idHoja;
T.files['orig']={getParents:()=>({hasNext:()=>false})}; T.files['copia1']=T.files['orig']; T.files['plant']=T.files['orig']; T.files['copia2']=T.files['orig']; T.files['legado']=T.files['orig'];
_activarAutomatizacion=()=>{};
const cfg=()=>{ _cfgMemo=null; return getAppConfig(ss); }, cr=()=>_configCruda(ss), eq=()=>readSheet(ss,'EQUIPO');

// 1. Catálogo original
T.setSheets([new T.Sheet('Hoja 1')]);
instalarCatalogo({proyecto:'Acervo Wixárika',lenguas:[{nombre_propio:'Wixárika',codigo_iso:'wix',atajos:'+=ɨ'},{nombre_completo:'Español',codigo_iso:'spa'}],categorias:'casa',carpetas:false,clave:true});
equipoGuardar({correo:'luis@ej.mx',rol:'Capturista',activo:'Sí'});
_escribirConfig(ss,{ID_CARPETA_RAIZ:'carpetaOrig',APPSHEET_APP_ID:'app-orig',URL_API:'https://script/orig/exec',CONTACTO_CORREO:'equipo@ej.mx',VISOR_AJUSTES:JSON.stringify({propuestas:{activo:true,clave:['k1']}})});
props[PROP_DIRECTORIO]='dirOrig'; props[CONFIG.PROPS.SNAPSHOT_FILE]='snapOrig'; dprops.CONSTANCIA_FIRMAS='Autoridad de San Andrés';
_agregar(_tabla(ss,'PALABRAS'),[{id_palabra:'P1',estado:'Publicado'}]);
const c0=cfg();
ok(cr().ID_HOJA==='orig' && _esCopia(ss)===false,'la original guarda su ID en la propia hoja');
ok(_asegurarIdentidad(ss)===null,'la original nunca se reinicia');
ok(publicarCatalogo(ss,true).cambio!==undefined && cfg().idCatalogo===c0.idCatalogo,'la original publica normal');

// 2. Copia (lo que llega con Archivo → Hacer una copia o «Copiar app» con datos): mismo contenido, otro ID de hoja
idHoja='copia1'; duenio='rosa@ej.mx'; Session.getActiveUser=()=>({getEmail:()=>'rosa@ej.mx'}); Session.getEffectiveUser=()=>({getEmail:()=>'rosa@ej.mx'});
props[CONFIG.PROPS.SNAPSHOT_FILE]='snapOrig';   // el peor caso: las propiedades del script viajaron con la copia
_cfgMemo=null;
ok(_copiaSinIdentidad(ss) && _esCopia(ss)===true,'la copia se reconoce por el ID guardado en la hoja');
let menu=''; const ui0=SpreadsheetApp.getUi; let items=[]; SpreadsheetApp.getUi=()=>Object.assign(ui0(),{createMenu:t=>{const m={addItem:(a,f)=>{items.push(f);return m},addSeparator:()=>m,addSubMenu:()=>m,addToUi:()=>m};return m;}});
_construirMenu(); ok(items[0]==='uiAsistenteInicial','en la copia, «Primeros pasos» aparece primero en el menú');
const st=panelEstado();   // cualquier entrada del programa le da identidad
const c1=cfg();
ok(c1.idCatalogo && c1.idCatalogo!==c0.idCatalogo,'código de catálogo propio');
ok(c1.apiToken && c1.apiToken!==c0.apiToken,'clave propia (la original tenía clave)');
ok(!c1.urlApi && !c1.appsheetId && !c1.raiz && !c1.visorAjustes,'olvida dirección, app, carpetas y ajustes del visor del original');
ok(c1.titulo==='Acervo Wixárika' && readSheet(ss,'PALABRAS').length===1,'no borra el contenido ni el nombre');
ok(!props[PROP_DIRECTORIO] && !props[CONFIG.PROPS.SNAPSHOT_FILE] && !dprops.CONSTANCIA_FIRMAS,'olvida directorio de voces, publicación y firmantes del original');
ok(props.HOJA_ID==='copia1' && cr().ID_HOJA==='copia1' && _esCopia(ss)===false && _vieneDeCopia(ss),'queda registrada como hoja propia que viene de una copia');
const e1=eq();
ok(e1.some(x=>x.correo==='rosa@ej.mx' && x.rol==='Administrador' && x.activo==='Sí'),'quien hizo la copia queda como administradora');
ok(e1.filter(x=>x.correo!=='rosa@ej.mx').every(x=>x.activo==='No') && e1.length===3,'el equipo del original queda inactivo, sin borrarse');
ok(st.copiaDe && st.copiaDe.startsWith('Acervo Wixárika'),'el panel avisa de dónde viene');
ok(_asegurarIdentidad(ss)===null && cfg().idCatalogo===c1.idCatalogo,'solo ocurre una vez');
const pub=publicarCatalogo(ss,true); ok(props[CONFIG.PROPS.SNAPSHOT_FILE] && props[CONFIG.PROPS.SNAPSHOT_FILE]!=='snapOrig','publica en un archivo propio, nunca en el del original');

// 3. De esa copia, la plantilla para compartir
let v=limpiezaVista({tipo:'enBlanco',conservarLenguas:true,conservarAtajos:true,plantilla:true});
limpiezaAplicar({tipo:'enBlanco',conservarLenguas:true,conservarAtajos:true,plantilla:true},v.total,'BORRAR');
const cp=cfg();
ok(_esPlantilla(ss) && !cp.titulo && !cp.idCatalogo && !cp.apiToken && !cp.contacto.correo,'plantilla: sin nombre, código, clave ni contacto');
ok(eq().length===0 && readSheet(ss,'PALABRAS').length===0 && readSheet(ss,'IDIOMAS').length===2,'plantilla: sin equipo ni palabras; conserva las lenguas elegidas');
asegurarEstructura(ss);
ok(eq().length===0 && !cfg().idCatalogo,'«Revisar y reparar» en la plantilla no le pone código ni equipo');
ok(publicarCatalogo(ss,true).plantilla===true,'la plantilla no publica');
let err=''; try{ instalarCatalogo({proyecto:'X',lenguas:[],categorias:''}); }catch(x){ err=x.message; } ok(err.includes('plantilla'),'en la plantilla, Primeros pasos pide hacer una copia');
items=[]; _construirMenu(); ok(items[0]==='uiAsistenteInicial','en la plantilla, «Primeros pasos» aparece primero');
const idPlant=cr().ID_HOJA;

// 4. Una comunidad copia la plantilla (por AppSheet o por Drive)
idHoja='copia2'; duenio='maria@ej.mx'; Session.getActiveUser=()=>({getEmail:()=>'maria@ej.mx'}); Session.getEffectiveUser=()=>({getEmail:()=>'maria@ej.mx'});
_cfgMemo=null;
ok(idPlant!=='copia2' && _copiaSinIdentidad(ss),'la copia de la plantilla se reconoce');
const r=instalarCatalogo({proyecto:'Catálogo Náayeri',lenguas:[],categorias:'animales',carpetas:false});   // sin lenguas nuevas: ya venían en la plantilla
const c2=cfg();
ok(r.proyecto==='Catálogo Náayeri' && c2.idCatalogo && !_esPlantilla(ss),'Primeros pasos la convierte en catálogo propio, con su código');
ok(eq().length===1 && eq()[0].correo==='maria@ej.mx' && eq()[0].rol==='Administrador','queda como única administradora');
ok(publicarCatalogo(ss,true).plantilla!==true && !!props[CONFIG.PROPS.SNAPSHOT_FILE],'ya publica');

// 5. La dirección del visor de una copia también la reinicia (si alguien implementa antes de abrir el menú)
idHoja='copia3'; T.files['copia3']=T.files['orig']; _cfgMemo=null; CacheService.getScriptCache().remove('IDENTIDAD_OK');
const antes=cfg().idCatalogo; doGet({parameter:{action:'manifiesto'}}); ok(cfg().idCatalogo!==antes && cr().ID_HOJA==='copia3','la dirección del visor de una copia le da identidad propia');

// 6. Copia de trabajo (rama): sigue ligada a su principal
idHoja='rama1'; T.files['rama1']=T.files['orig']; _escribirConfig(ss,{RAMA_DE:'copia3'}); _cfgMemo=null;
const antesR=cfg().idCatalogo; ok(_asegurarIdentidad(ss)===null && cfg().idCatalogo===antesR,'una copia de trabajo no se reinicia');
_escribirConfig(ss,{RAMA_DE:''});

// 7. Catálogo anterior (sin ID en la hoja) con una propiedad que no coincide: se toma como original, nunca se reinicia solo
idHoja='legado'; _escribirConfig(ss,{ID_HOJA:''}); props.HOJA_ID='otraCosa'; _cfgMemo=null;
const antesL=cfg().idCatalogo; asegurarEstructura(ss);
ok(cfg().idCatalogo===antesL && cr().ID_HOJA==='legado' && props.HOJA_ID==='legado','catálogo anterior: se registra sin reiniciarse');

// 8. Quien es dueño de la hoja siempre administra
Session.getActiveUser=()=>({getEmail:()=>'maria@ej.mx'}); duenio='maria@ej.mx';
_agregar(_tabla(ss,'EQUIPO'),[{correo:'otra@ej.mx',rol:'Administrador',activo:'Sí'}]); _actualizar(_tabla(ss,'EQUIPO'),0,{activo:'No'});
ok(_soyAdmin(ss)===true,'el dueño de la hoja siempre puede administrar');
console.log(fallos?'FALLOS '+fallos:'TODO OK');
