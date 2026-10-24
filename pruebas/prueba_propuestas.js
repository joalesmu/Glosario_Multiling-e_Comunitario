const fs=require('fs'); eval(fs.readFileSync(process.argv[3],'utf8'));
const ss=T.SS; let fallos=0; const ok=(c,m)=>{console.log((c?'✔ ':'✘ ')+m); if(!c) fallos++;};
global.DLG={}; global.HtmlService={createHtmlOutput:h=>({h,setWidth(){return this},setHeight(){return this},setTitle(){return this}})};
const ui0=SpreadsheetApp.getUi; SpreadsheetApp.getUi=()=>Object.assign(ui0(),{showModalDialog:(o,t)=>{DLG[t]=o.h;},showSidebar:o=>{DLG.sidebar=o.h;}});
// Drive simulado para imágenes
const MIS={}; let nf=0;
const carpeta=nm=>({getId:()=>'f_'+nm,createFile:b=>{const id='file'+String(++nf).padStart(28,'0');MIS[id]={id,b,acc:'P',trash:false,carpeta:nm,getId:()=>id,moveTo(c){this.carpeta=c.getId()},setSharing(a){this.acc=a},setTrashed(v){this.trash=v}};return MIS[id];},getFoldersByName:()=>({hasNext:()=>false}),createFolder:x=>carpeta(x)});
_carpetaDelLibro=()=>carpeta('libro'); _subcarpeta=(p,n)=>carpeta(n); _carpetaImagenes=()=>carpeta('IMAGENES');
const g0=DriveApp.getFileById; DriveApp.getFileById=id=>MIS[id]||g0(id); DriveApp.getStorageUsed=()=>8e9; DriveApp.getStorageLimit=()=>10e9;
Utilities.base64Decode=s=>Buffer.from(s,'base64');
T.setSheets([new T.Sheet('Hoja 1')]);
instalarCatalogo({proyecto:'Acervo',lenguas:[{nombre_propio:'Wixárika',codigo_iso:'wix',atajos:'+=ɨ'},{nombre_completo:'Español',codigo_iso:'spa'}],categorias:'casa',carpetas:false});
_activarAutomatizacion=()=>{};
const W=_s(readSheet(ss,'IDIOMAS')[0].id_idioma), E=_s(readSheet(ss,'IDIOMAS')[1].id_idioma);
_agregar(_tabla(ss,'PALABRAS'),[{id_palabra:'P1',estado:'Publicado',imagen_referencia:'https://drive.google.com/file/d/AAAAAAAAAAAAAAAAAAAAAAAAA1/view',imagen_id:'AAAAAAAAAAAAAAAAAAAAAAAAA1'},{id_palabra:'P2',estado:'Publicado'}]);
_agregar(_tabla(ss,'TRADUCCIONES'),[{id_traduccion:'T1',id_palabra:'P1',id_idioma:W,texto:'kiitɨ',estado:'Publicado'},{id_traduccion:'T2',id_palabra:'P1',id_idioma:E,texto:'casa',estado:'Publicado'},{id_traduccion:'T3',id_palabra:'P2',id_idioma:W,texto:'tsɨkɨ',estado:'Publicado'}]);
_escribirConfig(ss,{API_TOKEN:'',CONTACTO_CORREO:''}); _cfgMemo=null;
ok(!!ss.getSheetByName('PROPUESTAS') && !!ss.getSheetByName('IMAGENES'),'hojas PROPUESTAS e IMAGENES');
const post=d=>JSON.parse(doPost({postData:{contents:JSON.stringify(d)}}));
const PNG='data:image/png;base64,'+Buffer.from('fakepng').toString('base64');
const quien={nombre:'Mtra. Rosa',publico:true,rol:'docente',contacto:'rosa@ej.mx',acepto:true,compartir:true};
ok(post({accion:'proponer',quien,items:[{tipo:'palabra',texto:'maxa',id_idioma:W}]}).status===403,'cerradas por omisión (se activan en Participación)');
propuestasAjustesGuardar({activo:true,maxImagenes:3,maxPendDesc:1});
ok(post({accion:'proponer',quien:Object.assign({},quien,{acepto:false}),items:[{tipo:'palabra',texto:'maxa',id_idioma:W}]}).status===400,'sin consentimiento no se recibe');
let r=post({accion:'proponer',web:'spam',quien,items:[{tipo:'palabra',texto:'x',id_idioma:W}]}); ok(r.status===200 && readSheet(ss,'PROPUESTAS').length===0,'trampa: finge éxito y no guarda');
r=post({accion:'proponer',dispositivo:'d1',quien,items:[
  {tipo:'palabra',texto:'maxa',id_idioma:W,significado:'venado',id_idioma_significado:E,descripcion:'Teuta mieme',variante:'San Andrés',categoria:'casa',actividad:'Pregunta en casa · 7K3P'},
  {tipo:'descripcion',id_palabra:'P1',id_idioma:W,descripcion:'Ki ɨ́ teteuyeika'},
  {tipo:'descripcion',id_palabra:'P1',id_idioma:W,descripcion:'otra más'},
  {tipo:'correccion',id_palabra:'P2',id_idioma:W,texto:'tsɨkɨɨ'},
  {tipo:'imagen',id_palabra:'P1',imagen:PNG,tipo_imagen:'dibujo',licencia:'CC BY 4.0'},
  {tipo:'imagen',id_palabra:'P1',imagen:PNG,tipo_imagen:'foto'},
  {tipo:'imagen',id_palabra:'P1',imagen:PNG,tipo_imagen:'foto'},
  {tipo:'imagen',id_palabra:'P2',imagen:'data:text/html;base64,PHNjcmlwdD4=',tipo_imagen:'foto'},
  {tipo:'fotos',fotos:[PNG,PNG],actividad:'7K3P'},
  {tipo:'variante',id_palabra:'P2',id_idioma:W,texto:'chɨkɨ',variante:'Tuapurie'},
  {tipo:'palabra',texto:'',id_idioma:W},
  {tipo:'palabra',texto:'bad',id_idioma:'NOEXISTE'}]});
const R=r.resultados; ok(r.status===200 && /^P-[A-Z0-9]{5}$/.test(r.codigo),'recibe y da un código: '+r.codigo);
ok(R[2].error==='completa','límite de descripciones pendientes por palabra');
ok(R[4].ok && R[5].ok && R[6].error==='completa','cupo de imágenes: cuenta la principal y las pendientes (máx 3)');
ok(R[7].error==='imagen' && R[10].error==='vacia' && R[11].error==='lengua','rechaza imagen no válida, vacía y lengua inexistente');
const PR=readSheet(ss,'PROPUESTAS'); ok(PR.length===7 && PR.every(p=>p.estado==='Pendiente' && p.codigo===r.codigo),'todo entra como pendiente');
console.log(Object.values(MIS).map(f=>f.b.n).join(','));ok(Object.values(MIS).filter(f=>/^P-/.test(f.b.n)).length===4 && Object.values(MIS).every(f=>f.acc==='P'),'imágenes guardadas en privado: '+Object.values(MIS).length);
// estado por código
let est=JSON.parse(doGet({parameter:{action:'propuestas',codigos:r.codigo+',P-ZZZZZ'}}));
ok(est.propuestas.length===1 && est.propuestas[0].items.length===7 && !JSON.stringify(est).includes('rosa@'),'«Mis propuestas»: estado por código, sin contacto');
// visor ve ajustes públicos sin clave
let m=JSON.parse(doGet({parameter:{action:'manifiesto'}})); ok(m.estado.propuestas.activo===true && !('clave' in m.estado.propuestas),'el visor sabe que hay propuestas (sin la clave)');
// límite diario
propuestasAjustesGuardar({porDia:7}); ok(post({accion:'proponer',dispositivo:'d1',quien,items:[{tipo:'palabra',texto:'otra',id_idioma:W}]}).status===429,'límite por dispositivo al día');
propuestasAjustesGuardar({porDia:30});
// clave de colaboración
let cd=propuestasAjustesGuardar({quien:'clave'}); const clave=cd.propuestas.clave; ok(/^col/.test(clave),'clave de colaboración creada');
ok(post({accion:'proponer',dispositivo:'d2',quien,items:[{tipo:'palabra',texto:'nawa',id_idioma:W}]}).status===403,'sin clave no entra cuando se pide');
r2=post({accion:'proponer',dispositivo:'d2',colab:clave,quien:{nombre:'Juan',publico:false,acepto:true},items:[{tipo:'palabra',texto:'nawa',id_idioma:W}]}); ok(r2.status===200,'con clave sí');
ok(readSheet(ss,'PROPUESTAS').find(p=>p.texto==='nawa').prioridad==='alta','marcada para revisar primero');
propuestasAjustesGuardar({quien:'todos'});
// revisión
let D=propuestasDatos(false); ok(D.propuestas.length===8 && D.propuestas[0].texto==='nawa','datos de revisión (clave primero)');
const id=t=>D.propuestas.find(t).id_propuesta;
D=propuestaResolver(id(p=>p.texto==='maxa'),'nueva',{id_categoria:readSheet(ss,'CATEGORIAS')[0].id_categoria,respuesta:'¡Gracias!'});
const TR=readSheet(ss,'TRADUCCIONES'); const tm=TR.find(t=>t.texto==='maxa');
ok(tm && TR.some(t=>t.id_palabra===tm.id_palabra && t.texto==='venado') && tm.descripcion==='Teuta mieme' && /San Andrés/.test(tm.nota_variante),'acepta palabra nueva con significado, descripción y variante');
const fu=readSheet(ss,'FUENTES').find(f=>f.id_fuente===tm.id_fuente); ok(fu && fu.base_uso==='Aportación con consentimiento' && fu.publicar_credito==='Sí' && fu.uso_compartir==='Sí' && /Mtra. Rosa/.test(fu.credito),'crédito como fuente con su consentimiento');
ok(readSheet(ss,'ACTORES').some(a=>a.nombre==='Mtra. Rosa'),'quien aporta queda como actor (sin duplicar)');
const pub=construirDatosPublicos(ss); ok(pub.traducciones.some(t=>t.texto==='maxa'),'lo aceptado se publica (la fuente permite publicar)');
ok(pub.fuentes.some(f=>/Mtra. Rosa/.test(f.credito)),'su crédito aparece en las fuentes publicadas');
D=propuestaResolver(id(p=>p.tipo==='descripcion'),'descripcion',{modo:'reemplazar'}); ok(readSheet(ss,'TRADUCCIONES').find(t=>t.id_traduccion==='T1').descripcion==='Ki ɨ́ teteuyeika','acepta descripción');
D=propuestaResolver(id(p=>p.tipo==='correccion'),'correccion',{texto:'tsɨkɨ',modo:'sumar'}); ok(readSheet(ss,'TRADUCCIONES').find(t=>t.id_traduccion==='T3').texto==='tsɨkɨ','corrección igual a lo que ya había: no duplica');
D=propuestaResolver(id(p=>p.tipo==='variante'),'variante',{texto:'chɨkɨ'}); const t3=readSheet(ss,'TRADUCCIONES').find(t=>t.id_traduccion==='T3'); ok(t3.texto==='tsɨkɨ, chɨkɨ' && /Tuapurie/.test(t3.nota_variante),'variante sumada como forma de esa comunidad');
const imgs=D.propuestas.filter(p=>p.tipo==='imagen');
D=propuestaResolver(imgs[0].id_propuesta,'imagen',{alt:'Casa dibujada con crayón'});
const IM=readSheet(ss,'IMAGENES'); ok(IM.length===1 && IM[0].alt==='Casa dibujada con crayón' && MIS[IM[0].archivo].acc==='L' && MIS[IM[0].archivo].carpeta==='f_IMAGENES','imagen a la galería: pública y en la carpeta de imágenes');
D=propuestaResolver(D.propuestas.find(p=>p.tipo==='imagen').id_propuesta,'imagen',{principal:true,alt:'Foto de una casa'});
const P1=readSheet(ss,'PALABRAS').find(p=>p.id_palabra==='P1'); ok(/drive\.google\.com\/file\/d\/file0/.test(P1.imagen_referencia) && P1.imagen_alt==='Foto de una casa' && readSheet(ss,'IMAGENES').length===2,'hacer principal: la anterior pasa a la galería');
_prepararImagenes=()=>({}); 
const pub2=construirDatosPublicos(ss), w1=pub2.palabras.find(p=>p.id_palabra==='P1');
ok(w1.galeria && w1.galeria.length===2 && w1.galeria.some(g=>g.alt==='Casa dibujada con crayón') && w1.imgs===3,'galería publicada con texto alternativo y cuenta de imágenes');
const g=galeriaDatos('P1'); galeriaAccion(g.extra[0].id,'interna'); ok(construirDatosPublicos(ss).palabras.find(p=>p.id_palabra==='P1').galeria.length===1,'imagen solo interna no se publica');
D=propuestaResolver(D.propuestas.find(p=>p.tipo==='fotos').id_propuesta,'transcrita',{borrarFotos:true}); ok(Object.values(MIS).filter(f=>f.trash).length===2,'fotos transcritas a la papelera');
// aportes, novedades, retiro de crédito
ok(construirDatosPublicos(ss).aportes.some(a=>a.nombre==='Mtra. Rosa' && a.n>=4),'«Quiénes han aportado» con permiso');
ok(readSheet(ss,'AVISOS').some(a=>a.id_aviso==='novedades-automatico' && /aportaciones/.test(a.texto)),'aviso automático de novedades');
D=propuestaResolver(id(p=>p.texto==='nawa'),'no',{respuesta:'Ya está como otra palabra'});
est=JSON.parse(doGet({parameter:{action:'propuestas',codigos:r2.codigo}})); ok(est.propuestas[0].items[0].estado==='No incluida' && est.propuestas[0].items[0].respuesta==='Ya está como otra palabra','quien propuso ve la respuesta');
D=propuestasDatos(true); propuestaResolver(D.propuestas.find(p=>p.texto==='maxa').id_propuesta,'retirarCredito',{});
ok(!construirDatosPublicos(ss).fuentes.some(f=>/Mtra. Rosa/.test(f.credito)) && !construirDatosPublicos(ss).aportes.some(a=>a.nombre==='Mtra. Rosa'),'retirar crédito: su nombre deja de mostrarse');
ok(construirDatosPublicos(ss).traducciones.some(t=>t.texto==='maxa'),'…y su aportación sigue (solo sin nombre)');
// panel
const pe=panelEstado(); ok(pe.propuestas.pendientes===0 && pe.sinCorreo===true && pe.drive.pct===80,'Panel: pendientes, correo faltante y espacio de Drive');
// ventanas
uiPropuestas(); uiVerEnlaceApi(); uiPanel(); uiEditar();
// Constancia
{ const PS={}; global.PropertiesService=Object.assign(global.PropertiesService||{},{getDocumentProperties:()=>({getProperty:k=>PS[k]||null,setProperty:(k,v)=>{PS[k]=v;}})});
  const cd=constanciaDatos(); ok(Array.isArray(cd.personas) && cd.personas.some(p=>p.nombre==='Mtra. Rosa' && p.total>=1),'constancia: cuenta lo aceptado de cada persona '+JSON.stringify(cd.personas.map(p=>[p.nombre,p.total])));
  constanciaGuardarFirmas('Autoridad comunitaria'); ok(constanciaDatos().firmantes==='Autoridad comunitaria','constancia: recuerda firmantes');
  uiConstancia(); ok(Object.values(DLG).some(h=>h.includes('constanciaDatos')),'constancia: ventana'); }
for (const [t,h] of Object.entries(DLG)) [...h.matchAll(/<script>([\s\S]*?)<\/script>/g)].forEach((x,i)=>{ try{ new Function(x[1]); }catch(err){ ok(false,'sintaxis '+t+' #'+i+': '+err.message); } });
ok(Object.keys(DLG).length>=4,'ventanas generadas');
console.log(fallos?'FALLOS '+fallos:'TODO OK');
