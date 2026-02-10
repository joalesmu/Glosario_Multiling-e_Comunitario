// =============================================================
// 📘 GLOSARIO MULTIMEDIA COMUNITARIO - V1.0
// Autor: Alejandro Estrada | Versión: FINAL REVISADA
// =============================================================

// --- CONSTANTES DE HOJAS ---
const HOJA_CONFIG = 'CONFIGURACION';
const HOJA_TRADUCCIONES = 'TRADUCCIONES';
const HOJA_PALABRAS = 'PALABRAS';
const HOJA_IDIOMAS = 'IDIOMAS';
const HOJA_CATEGORIAS = 'CATEGORIAS';

// --- COLUMNAS DE CONFIGURACIÓN ---
const COL_AUDIO = 'ID_CARPETA_AUDIOS';
const COL_IMG = 'ID_CARPETA_IMAGENES';
const COL_CAT = 'ID_CARPETA_CATEGORIAS';

// =============================================================
// 🚀 1. API WEB (Conexión con el Visor HTML)
// =============================================================

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = obtenerConfig(ss);
  
  if (!config.audios || !config.imagenes) {
    return json({ status: "error", msg: "⚠️ Faltan configurar las carpetas. Ejecuta la Instalación en el menú Glosario." });
  }

  const p = e.parameter;

  if (p.type === 'audio' && p.file) {
    return servirAudio(p.file, config.audios);
  }

  return servirTodo(ss, config);
}

function servirTodo(ss, config) {
  const i = leerTablaSimple(ss, HOJA_IDIOMAS);
  const t = leerTablaSimple(ss, HOJA_TRADUCCIONES);
  const p = leerTablaSimple(ss, HOJA_PALABRAS);
  const c = leerTablaSimple(ss, HOJA_CATEGORIAS);

  const dicPalabras = {};
  p.forEach(row => dicPalabras[row.id_palabra] = row);

  const dicCategorias = {};
  c.forEach(row => dicCategorias[row.id_categoria] = row);

  const traduccionesFinal = t.map(row => {
    const infoPalabra = dicPalabras[row.id_palabra];
    let catNombre = "General", catImg = "", imgRef = "", vidRef = "";

    if (infoPalabra) {
      const infoCat = dicCategorias[infoPalabra.id_categoria];
      if (infoCat) {
        catNombre = infoCat.nombre_categoria;
        // CORRECCIÓN: Lee 'imagen_portada' (o 'imagen' por compatibilidad)
        catImg = procesarLinkImagen(infoCat.imagen_portada || infoCat.imagen); 
      }
      imgRef = procesarLinkImagen(infoPalabra.imagen_referencia);
      vidRef = infoPalabra.video_referencia;
    }

    return {
      id_palabra: row.id_palabra,
      id_idioma: row.id_idioma,
      texto: row.texto,
      definicion: row.definicion,
      audio: row.audio,
      categoria: catNombre,
      cat_img: catImg,
      imagen: imgRef,
      video: vidRef
    };
  });
  
  const listaCategorias = Object.values(dicCategorias).map(cat => ({
    nombre: cat.nombre_categoria,
    img: procesarLinkImagen(cat.imagen_portada || cat.imagen)
  }));

  return json({
    config: config,
    idiomas: i,
    traducciones: traduccionesFinal,
    categorias: listaCategorias
  });
}

function servirAudio(fileName, folderId) {
  try {
    const cleanName = fileName.split('/').pop(); 
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(cleanName);
    
    if (files.hasNext()) {
      const file = files.next();
      const bytes = file.getBlob().getBytes();
      const base64 = Utilities.base64Encode(bytes);
      return json({ status: 'ok', audioData: `data:${file.getMimeType()};base64,${base64}` });
    }
    return json({ status: 'error', msg: 'Audio no encontrado' });
  } catch (e) { return json({ status: 'error', msg: e.toString() }); }
}

// =============================================================
// 🛠️ 2. MENÚ E INSTALACIÓN
// =============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔹 GLOSARIO 🔹')
    .addItem('🚀 1. Instalación Automática', 'crearEstructuraCarpetas')
    .addSeparator()
    .addItem('📄 Obtener Plantilla de Importación', 'generarPlantilla')
    .addItem('📥 Importar Datos (Fusión Inteligente)', 'interfazImportacion')
    .addSeparator()
    .addItem('📤 Exportar Respaldo (Excel)', 'exportarExcel')
    .addItem('🔗 Obtener URL App Web', 'mostrarUrlWebApp')
    .addToUi();
}

function crearEstructuraCarpetas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConfig = ss.getSheetByName(HOJA_CONFIG); 
  
  const vals = hojaConfig.getDataRange().getValues();
  
  // VERIFICACIÓN INTELIGENTE:
  // Solo se detiene si la Columna D (Índice 3) ya tiene datos.
  // Esto permite tener datos precargados en A, B y C.
  if (vals.length > 1 && vals[1][3] && vals[1][3] !== "") {
    ui.alert('⚠️ Ya existe una configuración de carpetas (Columna D). Si continúas, podrías duplicarlas.');
    return;
  }

  if (ui.alert('🛠️ Instalación', 'Se crearán las carpetas en tu Google Drive.\n\n¿Continuar?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    ss.toast("Creando carpetas...");
    const nombreRaiz = "GLOSARIO_MULTIMEDIA_DATOS"; 
    const carpetaRaiz = DriveApp.createFolder(nombreRaiz);
    carpetaRaiz.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fAudios = carpetaRaiz.createFolder("Glosario_Audios");
    const fImg = carpetaRaiz.createFolder("Glosario_Imagenes");
    const fCat = carpetaRaiz.createFolder("Glosario_Categorias");
    [fAudios, fImg, fCat].forEach(f => f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW));

    const h = vals[0];
    const setID = (col, val) => { const idx = h.indexOf(col); if(idx>-1) hojaConfig.getRange(2, idx+1).setValue(val); };
    setID(COL_AUDIO, fAudios.getId());
    setID(COL_IMG, fImg.getId());
    setID(COL_CAT, fCat.getId());
    
    ui.alert("✅ Instalación Exitosa.\nCarpetas creadas en: " + nombreRaiz);
  } catch (e) { ui.alert("Error: " + e.toString()); }
}

function mostrarUrlWebApp() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (url) {
    ui.alert("🔗 TU LLAVE MAESTRA (URL):\n\n" + url + "\n\nCopia esto y pégalo cuando el glosario te lo pida.");
  } else {
    ui.alert("⚠️ Aún no has implementado la App.\nVe a: Implementar > Nueva Implementación.");
  }
}

// =============================================================
// 🔄 3. IMPORTACIÓN BLINDADA (Fusión de Datos)
// =============================================================

function generarPlantilla() {
  const ss = SpreadsheetApp.create("Plantilla_Importacion_Glosario");
  
  // CORREGIDO: 'nombre_completo' en IDIOMAS y 'imagen_portada' en CATEGORIAS
  const esquemas = [
    { n: HOJA_IDIOMAS, c: ['id_idioma', 'nombre_completo', 'codigo_iso'] },
    { n: HOJA_CATEGORIAS, c: ['id_categoria', 'nombre_categoria', 'imagen_portada'] },
    { n: HOJA_PALABRAS, c: ['id_palabra', 'id_categoria', 'imagen_referencia', 'video_referencia'] },
    { n: HOJA_TRADUCCIONES, c: ['id_traduccion', 'id_palabra', 'id_idioma', 'texto', 'definicion', 'audio', 'nota_variante'] }
  ];

  esquemas.forEach(e => {
    let h = ss.getSheetByName(e.n);
    if (!h) h = ss.insertSheet(e.n);
    h.clear();
    h.getRange(1, 1, 1, e.c.length).setValues([e.c]).setFontWeight("bold").setBackground("#cfe2f3");
  });
  
  const h1 = ss.getSheetByName("Hoja 1"); if(h1) ss.deleteSheet(h1);

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif; text-align:center; padding:20px;">
      <p>✅ Plantilla creada.</p>
      <p>Copia tus datos aquí y luego usa la opción "Importar".</p>
      <a href="${ss.getUrl()}" target="_blank" style="background:#16a34a; color:white; padding:10px; text-decoration:none; border-radius:5px;">📂 Abrir Plantilla</a>
    </div>
  `).setWidth(350).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Plantilla Lista');
}

function interfazImportacion() {
   const html = HtmlService.createHtmlOutput(`
     <div style="font-family:sans-serif; padding:15px;">
       <h3 style="margin-top:0;">🤝 Fusión Inteligente</h3>
       <p style="font-size:12px;">Pega la URL de la Plantilla o Glosario externo:</p>
       <input id="url" placeholder="https://docs.google.com..." style="width:100%;margin-bottom:10px;padding:8px;">
       <button id="btn" onclick="run()" style="background:#2563eb;color:white;border:none;padding:10px;width:100%;cursor:pointer;">INICIAR FUSIÓN</button>
       <p id="st" style="font-size:11px;color:#555;margin-top:10px;"></p>
       <script>
         function run() {
           const u = document.getElementById('url').value;
           if(!u) return alert("Falta la URL");
           document.getElementById('btn').innerText = '⏳ Procesando...';
           document.getElementById('btn').disabled = true;
           google.script.run
             .withSuccessHandler(r => { document.getElementById('btn').innerText='✅ Completado'; document.getElementById('st').innerText=r; })
             .withFailureHandler(e => { document.getElementById('btn').innerText='❌ Error'; document.getElementById('st').innerText=e; document.getElementById('btn').disabled=false; })
             .importarBlindado(u);
         }
       </script></div>`).setWidth(380).setHeight(280);
   SpreadsheetApp.getUi().showModalDialog(html, 'Importar Datos');
}

function importarBlindado(urlOrigen) {
  const ssDst = SpreadsheetApp.getActiveSpreadsheet();
  const ssSrc = SpreadsheetApp.openByUrl(urlOrigen);
  const uid = () => Utilities.getUuid();
  const clean = (t) => String(t || "").toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  
  const leer = (ss, nombre) => {
    const h = ss.getSheetByName(nombre); if (!h) return [];
    const d = h.getDataRange().getValues(); if (d.length < 2) return [];
    const heads = d[0];
    return d.slice(1).map(r => { let o={}; heads.forEach((k,i)=>o[k]=String(r[i])); return o; });
  };

  const mapLang = {}, mapCat = {}, mapConceptos = {};
  
  // CORRECCIÓN: Busca 'nombre_completo'
  leer(ssDst, HOJA_IDIOMAS).forEach(r => mapLang[clean(r.nombre_completo || r.nombre_idioma)] = r.id_idioma);
  leer(ssDst, HOJA_CATEGORIAS).forEach(r => mapCat[clean(r.nombre_categoria)] = r.id_categoria);
  
  let idEsp = "";
  for(let k in mapLang) if(k.includes("espanol") || k.includes("castellano")) idEsp = mapLang[k];
  
  if(idEsp) {
    leer(ssDst, HOJA_TRADUCCIONES).forEach(t => { 
      if(t.id_idioma === idEsp) mapConceptos[clean(t.texto)] = t.id_palabra; 
    });
  }

  const srcLang = leer(ssSrc, HOJA_IDIOMAS);
  const srcCat = leer(ssSrc, HOJA_CATEGORIAS);
  const srcPal = leer(ssSrc, HOJA_PALABRAS);
  const srcTrad = leer(ssSrc, HOJA_TRADUCCIONES);

  const newL=[], newC=[], newP=[], newT=[];
  const mapIdL={}, mapIdC={};
  let skipped = 0;

  // PROCESAR IDIOMAS
  srcLang.forEach(i => {
    // CORRECCIÓN: Prioriza 'nombre_completo'
    const name = i.nombre_completo || i.nombre_idioma;
    if(!name) { skipped++; return; }
    const n = clean(name);
    if(mapLang[n]) mapIdL[i.id_idioma] = mapLang[n];
    else { const id = uid(); mapIdL[i.id_idioma]=id; mapLang[n]=id; newL.push([id, name, i.codigo_iso||""]); }
  });

  // PROCESAR CATEGORIAS
  let idGeneral = mapCat["general"];
  if(!idGeneral) { idGeneral=uid(); mapCat["general"]=idGeneral; newC.push([idGeneral, "General", ""]); }

  srcCat.forEach(c => {
    const name = c.nombre_categoria;
    if(!name) { mapIdC[c.id_categoria] = idGeneral; return; }
    const n = clean(name);
    if(mapCat[n]) mapIdC[c.id_categoria] = mapCat[n];
    else { 
      const id = uid(); 
      mapIdC[c.id_categoria]=id; mapCat[n]=id; 
      // CORRECCIÓN: Prioriza 'imagen_portada'
      const imgVal = c.imagen_portada || c.imagen || "";
      newC.push([id, name, imgVal]); 
    }
  });

  // PROCESAR PALABRAS
  const groups = {};
  srcPal.forEach(p => groups[p.id_palabra] = { p: p, t: [] });
  srcTrad.forEach(t => { if(groups[t.id_palabra]) groups[t.id_palabra].t.push(t); });

  for (let pid in groups) {
    const g = groups[pid];
    if(g.t.length === 0) continue;

    const esTrad = g.t.find(t => {
      const lid = mapIdL[t.id_idioma];
      const srcLName = srcLang.find(l => l.id_idioma == t.id_idioma);
      const n = clean(srcLName ? (srcLName.nombre_completo || srcLName.nombre_idioma) : "");
      return n.includes("espanol") || n.includes("castellano");
    });

    let finalPid = "";

    if (esTrad && mapConceptos[clean(esTrad.texto)]) {
      finalPid = mapConceptos[clean(esTrad.texto)];
    } else {
      finalPid = uid();
      const catId = mapIdC[g.p.id_categoria] || idGeneral;
      newP.push([finalPid, catId, g.p.imagen_referencia||"", g.p.video_referencia||""]);
      if (esTrad) mapConceptos[clean(esTrad.texto)] = finalPid;
    }

    g.t.forEach(t => {
      if (!t.texto) return;
      const lid = mapIdL[t.id_idioma];
      if (lid) newT.push([uid(), finalPid, lid, t.texto, t.definicion||"", t.audio||"", t.nota_variante||""]);
    });
  }

  const append = (n, d) => { if(d.length) ssDst.getSheetByName(n).getRange(ssDst.getSheetByName(n).getLastRow()+1,1,d.length,d[0].length).setValues(d); };
  
  append(HOJA_IDIOMAS, newL);
  append(HOJA_CATEGORIAS, newC);
  append(HOJA_PALABRAS, newP);
  append(HOJA_TRADUCCIONES, newT);

  return `Proceso completado.\n+${newL.length} Idiomas\n+${newP.length} Palabras\n+${newT.length} Traducciones\n(Ignorados: ${skipped})`;
}

// =============================================================
// ⚙️ 4. UTILIDADES GENERALES
// =============================================================

function exportarExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl().replace(/edit$/, 'export?format=xlsx');
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif; text-align:center; padding:20px;">
      <a href="${url}" target="_blank" style="background:#16a34a; color:white; padding:10px; text-decoration:none; border-radius:5px;">📥 Descargar Excel</a>
    </div>
  `).setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Respaldo Portátil');
}

function obtenerConfig(ss) {
  const hoja = ss.getSheetByName(HOJA_CONFIG);
  if (!hoja) return {};
  const data = hoja.getDataRange().getValues();
  if (data.length < 2) return {};
  const h = data[0], v = data[1];
  const getVal = (name) => { const i = h.indexOf(name); return i > -1 ? String(v[i]) : ""; };
  
  return { 
    audios: getVal(COL_AUDIO), 
    imagenes: getVal(COL_IMG), 
    titulo: getVal('NOMBRE_PROYECTO'), 
    subtitulo: getVal('SUBTITULO'), 
    idiomaPrincipal: getVal('IDIOMA_PRINCIPAL')
  };
}

function procesarLinkImagen(valor) {
  if (!valor) return "";
  if (valor.toString().startsWith("http") && valor.includes("drive.google.com")) {
    try { return "https://drive.google.com/thumbnail?id=" + valor.split("/d/")[1].split("/")[0] + "&sz=w400"; } catch(e){}
  }
  try {
    const nombre = valor.split('/').pop(); 
    const archivos = DriveApp.getFilesByName(nombre);
    if (archivos.hasNext()) return "https://drive.google.com/thumbnail?id=" + archivos.next().getId() + "&sz=w400";
  } catch (e) {}
  return ""; 
}

function leerTablaSimple(ss, nombre) {
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) return [];
  const data = hoja.getDataRange().getValues();
  if (data.length < 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = String(row[i]));
    return obj;
  });
}

function json(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }