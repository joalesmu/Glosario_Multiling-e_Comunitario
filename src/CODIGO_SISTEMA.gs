// =============================================================
// 📘 GLOSARIO MULTIMEDIA COMUNITARIO - V1.0
// Autor: Alejandro Estrada | Versión: BETA
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
// 🚀 1. API WEB
// =============================================================

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = obtenerConfig(ss);
  
  if (!config.audios || !config.imagenes) {
    return json({ status: "error", msg: "⚠️ Faltan configurar las carpetas." });
  }

  const p = e.parameter;
  if (p.type === 'audio' && p.file) return servirAudio(p.file, config.audios);
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
      variante: row.nota_variante, 
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
  SpreadsheetApp.getUi().createMenu('🔹 GLOSARIO 🔹')
    .addItem('🚀 1. Instalación Automática', 'crearEstructuraCarpetas')
    .addSeparator()
    .addItem('📄 Obtener Plantilla', 'generarPlantilla')
    .addItem('📥 Importar (Fusión Inteligente)', 'interfazImportacion')
    .addSeparator()
    .addItem('🧹 Corregir Tipografía (+ ➝ ɨ)', 'corregirTipografiaMasiva') // <--- NUEVO BOTÓN
    .addSeparator()
    .addItem('📤 Exportar Excel', 'exportarExcel')
    .addItem('🔗 Ver URL App', 'mostrarUrlWebApp')
    .addToUi();
}

// --- HERRAMIENTA DE CORRECCIÓN ---
function corregirTipografiaMasiva() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(HOJA_TRADUCCIONES);
  
  if (!sheet) return ui.alert("No encuentro la hoja TRADUCCIONES.");
  
  // Confirmación
  if (ui.alert('🧹 Corrector Tipográfico', 'Esto buscará el símbolo "+" en toda la hoja de Traducciones y lo cambiará por "ɨ".\n\n¿Continuar?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return; // Solo encabezados

  const headers = data[0];
  const idxTexto = headers.indexOf('texto');
  const idxDef = headers.indexOf('definicion');

  if (idxTexto === -1) return ui.alert("Error: No encuentro la columna 'texto'");

  let cambios = 0;
  
  // Recorremos y reemplazamos en memoria
  for (let i = 1; i < data.length; i++) {
    // Corregir Texto
    let txt = String(data[i][idxTexto]);
    if (txt.includes('+')) {
      data[i][idxTexto] = txt.replace(/\+/g, 'ɨ');
      cambios++;
    }
    
    // Corregir Definición (si existe)
    if (idxDef > -1) {
      let def = String(data[i][idxDef]);
      if (def.includes('+')) {
        data[i][idxDef] = def.replace(/\+/g, 'ɨ');
        cambios++; // Contamos también estos cambios
      }
    }
  }

  if (cambios > 0) {
    // Guardamos todo de golpe (rápido)
    sheet.getDataRange().setValues(data);
    ui.alert(`✅ ¡Listo! Se corrigieron ${cambios} textos.`);
  } else {
    ui.alert("👍 Todo está limpio. No se encontraron símbolos '+'.");
  }
}

function crearEstructuraCarpetas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConfig = ss.getSheetByName(HOJA_CONFIG); 
  const vals = hojaConfig.getDataRange().getValues();
  
  if (vals.length > 1 && vals[1][3] && vals[1][3] !== "") {
    ui.alert('⚠️ Ya existe configuración. Deteniendo para evitar duplicados.');
    return;
  }
  if (ui.alert('🛠️ Instalación', '¿Crear carpetas en Drive?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
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
    setID(COL_AUDIO, fAudios.getId()); setID(COL_IMG, fImg.getId()); setID(COL_CAT, fCat.getId());
    ui.alert("✅ Instalación Exitosa.");
  } catch (e) { ui.alert("Error: " + e.toString()); }
}

function mostrarUrlWebApp() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? "🔗 URL:\n" + url : "⚠️ Debes implementar la App primero.");
}

// =============================================================
// 🔄 3. IMPORTACIÓN MAESTRA
// =============================================================

function generarPlantilla() {
  const ss = SpreadsheetApp.create("Plantilla_Glosario");
  const esquemas = [
    { n: HOJA_IDIOMAS, c: ['id_idioma', 'nombre_completo', 'codigo_iso'] },
    { n: HOJA_CATEGORIAS, c: ['id_categoria', 'nombre_categoria', 'imagen_portada'] },
    { n: HOJA_PALABRAS, c: ['id_palabra', 'id_categoria', 'imagen_referencia', 'video_referencia'] },
    { n: HOJA_TRADUCCIONES, c: ['id_traduccion', 'id_palabra', 'id_idioma', 'texto', 'definicion', 'audio', 'nota_variante'] }
  ];
  esquemas.forEach(e => {
    let h = ss.getSheetByName(e.n); if (!h) h = ss.insertSheet(e.n);
    h.clear(); h.getRange(1, 1, 1, e.c.length).setValues([e.c]).setFontWeight("bold").setBackground("#cfe2f3");
  });
  const h1 = ss.getSheetByName("Hoja 1"); if(h1) ss.deleteSheet(h1);
  const html = HtmlService.createHtmlOutput(`<div style="font-family:sans-serif;text-align:center;padding:20px;"><a href="${ss.getUrl()}" target="_blank" style="background:#16a34a;color:white;padding:10px;text-decoration:none;border-radius:5px;">📂 Abrir Plantilla</a></div>`).setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Plantilla');
}

function interfazImportacion() {
   const html = HtmlService.createHtmlOutput(`
     <div style="font-family:sans-serif; padding:15px;">
       <h3 style="margin-top:0;">🤝 Fusión Inteligente</h3>
       <p style="font-size:12px;">URL de Plantilla o Glosario:</p>
       <input id="url" placeholder="https://docs.google.com..." style="width:100%;margin-bottom:10px;padding:8px;">
       <button id="btn" onclick="run()" style="background:#2563eb;color:white;border:none;padding:10px;width:100%;cursor:pointer;">INICIAR FUSIÓN</button>
       <p id="st" style="font-size:11px;color:#555;margin-top:10px;"></p>
       <script>
         function run() {
           const u = document.getElementById('url').value;
           if(!u) return alert("Falta URL");
           document.getElementById('btn').innerText = '⏳ Analizando...'; document.getElementById('btn').disabled = true;
           google.script.run.withSuccessHandler(r => { document.getElementById('btn').innerText='✅ Listo'; document.getElementById('st').innerText=r; })
             .withFailureHandler(e => { document.getElementById('btn').innerText='❌ Error'; document.getElementById('st').innerText=e; document.getElementById('btn').disabled=false; })
             .importarBlindado(u);
         }
       </script></div>`).setWidth(380).setHeight(280);
   SpreadsheetApp.getUi().showModalDialog(html, 'Importar');
}

function importarBlindado(urlOrigen) {
  const ssDst = SpreadsheetApp.getActiveSpreadsheet();
  const ssSrc = SpreadsheetApp.openByUrl(urlOrigen);
  const uid = () => Utilities.getUuid();
  // Auto-corrección al importar
  const clean = (t) => String(t || "").toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\+/g, 'i'); 

  const leer = (ss, nombre) => {
    const h = ss.getSheetByName(nombre); if (!h) return [];
    const d = h.getDataRange().getValues(); if (d.length < 2) return [];
    const heads = d[0];
    return d.slice(1).map(r => { let o={}; heads.forEach((k,i)=>o[k]=String(r[i])); return o; });
  };

  const dstLang = leer(ssDst, HOJA_IDIOMAS);
  const dstTrad = leer(ssDst, HOJA_TRADUCCIONES);
  
  const mapLangDst = {};
  let idEspDst = null;
  dstLang.forEach(r => {
    const n = clean(r.nombre_completo || r.nombre_idioma);
    mapLangDst[n] = r.id_idioma;
    if (n.includes('espanol') || n.includes('castellano') || r.codigo_iso === 'es') idEspDst = r.id_idioma;
  });

  const mapaConceptosExistentes = {};
  if (idEspDst) {
    dstTrad.forEach(t => {
      if (t.id_idioma === idEspDst) {
        mapaConceptosExistentes[clean(t.texto)] = t.id_palabra;
      }
    });
  }

  const firmasTraducciones = new Set();
  dstTrad.forEach(t => {
    firmasTraducciones.add(`${t.id_palabra}_${t.id_idioma}_${clean(t.texto)}`);
  });

  const mapaCategoriasDst = {};
  leer(ssDst, HOJA_CATEGORIAS).forEach(r => mapaCategoriasDst[clean(r.nombre_categoria)] = r.id_categoria);

  // --- LEER ORIGEN ---
  const srcLang = leer(ssSrc, HOJA_IDIOMAS);
  const srcCat = leer(ssSrc, HOJA_CATEGORIAS);
  const srcPal = leer(ssSrc, HOJA_PALABRAS);
  const srcTrad = leer(ssSrc, HOJA_TRADUCCIONES);

  const newL=[], newC=[], newP=[], newT=[];
  const mapIdL={}, mapIdC={};

  srcLang.forEach(i => {
    const name = i.nombre_completo || i.nombre_idioma; if(!name) return;
    const n = clean(name);
    if(mapLangDst[n]) {
      mapIdL[i.id_idioma] = mapLangDst[n];
    } else {
      const finalId = uid(); mapIdL[i.id_idioma] = finalId; mapLangDst[n] = finalId;
      newL.push([finalId, name, i.codigo_iso||""]);
    }
  });

  let idGeneral = mapaCategoriasDst["general"];
  if(!idGeneral) { idGeneral="general"; mapaCategoriasDst["general"]=idGeneral; newC.push(["general", "General", ""]); }
  
  srcCat.forEach(c => {
    const name = c.nombre_categoria; if(!name) { mapIdC[c.id_categoria]=idGeneral; return; }
    const n = clean(name);
    if(mapaCategoriasDst[n]) {
      mapIdC[c.id_categoria] = mapaCategoriasDst[n];
    } else {
      const finalId = uid(); mapIdC[c.id_categoria] = finalId; mapaCategoriasDst[n] = finalId;
      newC.push([finalId, name, c.imagen_portada||""]);
    }
  });

  const groups = {};
  srcPal.forEach(p => { if(p.id_palabra && p.id_palabra!=="") groups[p.id_palabra] = { p:p, t:[] }; });
  srcTrad.forEach(t => { if(t.id_palabra && groups[t.id_palabra]) groups[t.id_palabra].t.push(t); });

  for (let pid in groups) {
    const g = groups[pid];
    const tradsValidas = g.t.filter(t => t.texto && String(t.texto).trim() !== "");
    if (tradsValidas.length === 0) continue;

    let palabraEsp = tradsValidas.find(t => {
      const lid = mapIdL[t.id_idioma];
      return lid === idEspDst; 
    });

    let finalPid = null;
    let esNuevaPalabra = true;

    if (palabraEsp && mapaConceptosExistentes[clean(palabraEsp.texto)]) {
      finalPid = mapaConceptosExistentes[clean(palabraEsp.texto)];
      esNuevaPalabra = false;
    } else {
      finalPid = uid();
      esNuevaPalabra = true;
      if (palabraEsp) mapaConceptosExistentes[clean(palabraEsp.texto)] = finalPid;
    }

    if (esNuevaPalabra) {
      const catId = mapIdC[g.p.id_categoria] || idGeneral;
      newP.push([finalPid, catId, g.p.imagen_referencia||"", g.p.video_referencia||""]);
    }

    tradsValidas.forEach(t => {
      const lid = mapIdL[t.id_idioma];
      if (lid) {
        // CORRECCIÓN TIPOGRÁFICA AL IMPORTAR
        const textoReal = String(t.texto).replace(/\+/g, 'ɨ');
        const defReal = String(t.definicion || "").replace(/\+/g, 'ɨ');

        const firma = `${finalPid}_${lid}_${clean(textoReal)}`;
        
        if (!firmasTraducciones.has(firma)) {
          const tradId = uid();
          newT.push([tradId, finalPid, lid, textoReal, defReal, t.audio||"", t.nota_variante||""]);
          firmasTraducciones.add(firma);
        }
      }
    });
  }

  const append = (n, d) => { if(d.length) ssDst.getSheetByName(n).getRange(ssDst.getSheetByName(n).getLastRow()+1,1,d.length,d[0].length).setValues(d); };
  append(HOJA_IDIOMAS, newL); append(HOJA_CATEGORIAS, newC); append(HOJA_PALABRAS, newP); append(HOJA_TRADUCCIONES, newT);

  return `Importación V1.6 Completada.\n+${newL.length} Idiomas\n+${newP.length} Conceptos\n+${newT.length} Traducciones`;
}

// =============================================================
// ⚙️ 4. UTILIDADES
// =============================================================

function exportarExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl().replace(/edit$/, 'export?format=xlsx');
  const html = HtmlService.createHtmlOutput(`<div style="font-family:sans-serif;text-align:center;padding:20px;"><a href="${url}" target="_blank" style="background:#16a34a;color:white;padding:10px;text-decoration:none;border-radius:5px;">📥 Descargar Excel</a></div>`).setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Respaldo');
}

function obtenerConfig(ss) {
  const h = ss.getSheetByName(HOJA_CONFIG).getDataRange().getValues();
  if (h.length < 2) return {};
  const get = (k) => { const i = h[0].indexOf(k); return i > -1 ? String(h[1][i]) : ""; };
  return { audios: get(COL_AUDIO), imagenes: get(COL_IMG), titulo: get('NOMBRE_PROYECTO'), subtitulo: get('SUBTITULO'), idiomaPrincipal: get('IDIOMA_PRINCIPAL') };
}

function procesarLinkImagen(v) {
  if (!v) return "";
  if (String(v).includes("drive.google.com")) try { return "https://drive.google.com/thumbnail?id=" + v.split("/d/")[1].split("/")[0] + "&sz=w400"; } catch(e){}
  return ""; 
}

function leerTablaSimple(ss, nombre) {
  const h = ss.getSheetByName(nombre); if (!h) return [];
  const d = h.getDataRange().getValues(); if (d.length < 1) return [];
  const heads = d[0];
  return d.slice(1).map(r => { let o={}; heads.forEach((k,i)=>o[k]=String(r[i])); return o; });
}

function json(d) { return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); }