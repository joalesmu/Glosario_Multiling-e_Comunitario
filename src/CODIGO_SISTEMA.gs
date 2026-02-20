/**
 * 🏛️ SISTEMA DE GESTIÓN LINGÜÍSTICA
 * ===================================================================
 * @license         Alejandro Estrada - 2026
 * @version         1.0
 * ===================================================================
 */

const CONFIG = {
  SHEETS: {
    CONFIG:       'CONFIGURACION',
    TRADUCCIONES: 'TRADUCCIONES',
    TRAD_CATS:    'TRADUCCIONES_CATS',
    PALABRAS:     'PALABRAS',
    IDIOMAS:      'IDIOMAS',
    CATEGORIAS:   'CATEGORIAS',
    REPORT:       'MESA_DE_TRABAJO',
    AUDIT:        'AUDITORIA_ADMIN'
  },
  COLS: {
    AUDIO_ID:     'ID_CARPETA_AUDIOS',
    IMG_ID:       'ID_CARPETA_IMAGENES',
    CAT_ID:       'ID_CARPETA_CATEGORIAS',
    PROJECT_NAME: 'NOMBRE_PROYECTO',
    APPSHEET_ID:  'APPSHEET_APP_ID',
    MAIN_LANG:    'idioma_activo',
    API_TOKEN:    'API_TOKEN'
  },
  CACHE: { ENABLED: true, TTL: 21600, KEY: 'API_DATA_V1' }
};

// ─────────────────────────────────────────────────────────────────────────────
// 1. MENÚ DE ADMINISTRACIÓN
// ─────────────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('💠 ADMINISTRACIÓN GLOSARIO')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📊 Centro de Auditoría')
      .addItem('🛡️ Dashboard de Imágenes (Admin)',       'uiGenerarDashboardAdmin')
      .addItem('🎛️ Mesa de Trabajo (Faltantes)',        'uiGenerarReporteAvanzado')
      .addItem('📥 Importar desde Mesa',                'servicioImportarDesdeReporte')
      .addItem('🎨 Exportar Lista para Diseñador',      'servicioExportarListaDisenador'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📂 ETL y Colaboración')
      .addItem('📄 Generar Plantilla para Colaborador',   'uiConfigurarPlantillaInteligente')
      .addItem('📥 Importación Inteligente (ID Match)',   'uiImportarDatos')
      .addItem('📤 Descargar Respaldo Excel',             'uiExportarDatos'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚡ Mantenimiento')
      .addItem('🚀 Instalación de Carpetas y Hojas',      'servicioInstalacionCarpetas')
      .addItem('🆔 Reparar IDs Faltantes',                'servicioRepararIDs')
      .addItem('🧹 Normalizar Tipografía (+ ➝ ɨ)',        'servicioNormalizacionTexto')
      .addItem('🗑️ Limpiar Caché API',                    'servicioLimpiarCache'))
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. DETECTOR DINÁMICO DE CONFIGURACIÓN (REPLICABILIDAD)
// ─────────────────────────────────────────────────────────────────────────────

function _getSeparator() {
  const locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  const semicolonLocales = ['es_ES', 'es_MX', 'fr_FR', 'it_IT', 'pt_BR'];
  return semicolonLocales.includes(locale) ? ';' : ',';
}

function getAppConfig(ss) {
  const s = ss.getSheetByName(CONFIG.SHEETS.CONFIG); 
  if (!s) throw new Error('Falta hoja de CONFIGURACIÓN.');
  const d = s.getDataRange().getValues(); 
  const h = d[0].map(x => String(x).toUpperCase().trim()); 
  const v = d[1] || [];
  const get = k => { const i = h.indexOf(k.toUpperCase()); return i > -1 ? String(v[i] || '').trim() : ''; };
  
  return {
    audios: get(CONFIG.COLS.AUDIO_ID), 
    imagenes: get(CONFIG.COLS.IMG_ID), 
    appsheetId: get(CONFIG.COLS.APPSHEET_ID),
    titulo: get(CONFIG.COLS.PROJECT_NAME),
    subtitulo: get('SUBTITULO'),
    idiomaPrincipal: get(CONFIG.COLS.MAIN_LANG),
    apiToken: get(CONFIG.COLS.API_TOKEN)
  };
}

function responseJSON(d) { 
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); 
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. MESA DE TRABAJO (AUDITORÍA DINÁMICA)
// ─────────────────────────────────────────────────────────────────────────────

function uiGenerarReporteAvanzado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS);
  
  const checkboxHtml = idiomas.map(i => {
    let tag = i.nombre_completo || i.nombre || i.id_idioma;
    let iso = (i.codigo_iso || i.id_idioma.substring(0,3)).toUpperCase();
    return `<div style="margin-bottom:8px;">
              <input type="checkbox" id="aud_${iso}" name="idiomas" value="${i.id_idioma}|${iso}|${tag}" checked>
              <label for="aud_${iso}"> ${tag} (${iso})</label>
            </div>`;
  }).join('');

  const html = `
    <div style="font-family:sans-serif;padding:15px;color:#333;">
      <h3 style="color:#2563eb;margin-top:0;">🎛️ Configurar Mesa de Trabajo</h3>
      <p style="font-size:12px;color:#64748b;">Selecciona los idiomas que deseas trabajar hoy:</p>
      
      <div style="max-height:150px;overflow-y:auto;background:#f8fafc;padding:10px;border-radius:6px;border:1px solid #e2e8f0;">
        ${checkboxHtml}
      </div>

      <div style="background:#fff7ed;padding:10px;border-radius:6px;border:1px solid #fdba74;margin-top:15px;font-size:12px;">
        <b>Tareas a detectar:</b><br>
        <label><input type="checkbox" id="chk_trans" checked> 📝 Traducciones faltantes</label><br>
        <label><input type="checkbox" id="chk_audio" checked> 🔊 Audios faltantes</label>
      </div>

      <button id="btn" onclick="run()" style="width:100%;background:#2563eb;color:white;border:none;padding:12px;border-radius:6px;margin-top:15px;cursor:pointer;font-weight:bold;">Generar Mesa ➔</button>
      
      <script>
        function run(){
          const btn = document.getElementById('btn');
          const idiomas = Array.from(document.querySelectorAll('input[name="idiomas"]:checked')).map(c => c.value);
          const opts = { trans: document.getElementById('chk_trans').checked, audio: document.getElementById('chk_audio').checked };
          
          btn.disabled = true; btn.innerText = "⏳ Analizando...";
          google.script.run.withSuccessHandler(() => google.script.host.close()).procesarReporteBack(idiomas, opts);
        }
      </script>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(480), 'Configuración de Auditoría');
}

function procesarReporteBack(seleccionados, opts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const config = getAppConfig(ss);
  const sep = _getSeparator(); 
  
  const palabras = readSheet(ss, CONFIG.SHEETS.PALABRAS); 
  const trads = readSheet(ss, CONFIG.SHEETS.TRADUCCIONES);
  const idIdiomaBase = String(config.idiomaPrincipal || "").trim();

  const mapaReferencia = _mapaTrads(ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES), idIdiomaBase);
  const appSheetUrl = `https://www.appsheet.com/start/${config.appsheetId}`;

  let reporte = [];
  let filasEditables = [];

  palabras.forEach(p => {
    const idPal = String(p.id_palabra).trim();
    const textoBase = mapaReferencia[idPal] || "(Sin traducción base registrada)";

    seleccionados.forEach(sel => {
      let [idLang, iso, nomLang] = sel.split('|');
      if (idLang === idIdiomaBase) return;

      const tRow = trads.find(t => String(t.id_palabra).trim() === idPal && String(t.id_idioma).trim() === idLang);

      let tarea = "";
      if (!tRow && opts.trans) tarea = "📝 FALTA TRADUCCION";
      else if (tRow && opts.audio && (!tRow.audio || tRow.audio.length < 2)) tarea = "🔊 FALTA AUDIO";

      if (tarea !== "") {
        const view = (tarea === "📝 FALTA TRADUCCION") ? "Traducciones_Form" : "Traducciones_Form";
        const rowId = (tRow) ? tRow.id_traduccion : "";
        const deepLink = `${appSheetUrl}#control=${view}${rowId ? '&row=' + rowId : ''}`;
        const formulaLink = `=HYPERLINK("${deepLink}"${sep}"📱 GRABAR / EDITAR")`;

        reporte.push([
          idPal, textoBase, (tRow ? tRow.texto : "---"), idLang, nomLang, tarea, 
          "", "", "PENDIENTE", formulaLink
        ]);
        
        if (tarea === "📝 FALTA TRADUCCION") filasEditables.push(reporte.length + 1);
      }
    });
  });

  let sheet = ss.getSheetByName(CONFIG.SHEETS.REPORT) || ss.insertSheet(CONFIG.SHEETS.REPORT);
  sheet.clear();
  sheet.getDataRange().clearDataValidations();
  sheet.setTabColor("#f39c12");

  if (reporte.length === 0) {
    sheet.getRange(1, 1).setValue("🎉 Todo al día para los idiomas seleccionados.");
    sheet.activate(); return;
  }

  const headers = [["ID_REF", "REFERENCIA", "TEXTO ACTUAL", "ID_LANG", "IDIOMA", "TAREA", "NUEVA TRADUCCIÓN", "NOTA", "ESTADO", "ACCION DIRECTA"]];
  sheet.getRange(1, 1, 1, 10).setValues(headers).setBackground("#2c3e50").setFontColor("white").setFontWeight("bold");
  sheet.getRange(2, 1, reporte.length, 10).setValues(reporte);
  
  if (filasEditables.length > 0) {
    filasEditables.forEach(r => sheet.getRange(r, 7, 1, 2).setBackground("#fffec8").setBorder(true,true,true,true,true,true));
  }
  
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['PENDIENTE', 'IMPORTADO', 'OMITIR']).build();
  sheet.getRange(2, 9, reporte.length, 1).setDataValidation(rule);
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 10, 130);
  sheet.activate();
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. DASHBOARD DE IMÁGENES (ADMIN)
// ─────────────────────────────────────────────────────────────────────────────

function uiGenerarDashboardAdmin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAppConfig(ss);
  const sep = _getSeparator();
  const palabras = readSheet(ss, CONFIG.SHEETS.PALABRAS);
  const idIdiomaBase = String(config.idiomaPrincipal || "").trim();
  
  const mapaRef = _mapaTrads(ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES), idIdiomaBase);
  const appSheetUrl = `https://www.appsheet.com/start/${config.appsheetId}#control=Palabras_Detail&row=`;

  let reporte = [];
  palabras.forEach(p => {
    let id = String(p.id_palabra).trim();
    let txt = mapaRef[id] || "(Sin nombre base)";
    let stImg = (p.imagen_referencia && p.imagen_referencia.length > 5) ? "✅ OK" : "❌ PENDIENTE";
    
    const deepLink = `${appSheetUrl}${id}`;
    const formulaLink = `=HYPERLINK("${deepLink}"${sep}"🖼️ SUBIR IMAGEN")`;
    
    reporte.push([id, txt, p.id_categoria, stImg, formulaLink]);
  });

  let sAudit = ss.getSheetByName(CONFIG.SHEETS.AUDIT) || ss.insertSheet(CONFIG.SHEETS.AUDIT);
  sAudit.clear();
  sAudit.getDataRange().clearDataValidations();
  sAudit.setTabColor("#e11d48");

  sAudit.getRange(1, 1, 1, 5).setValues([["ID_PALABRA", "CONCEPTO", "CATEGORÍA", "ESTADO IMAGEN", "ACCION APPSHEET"]]).setBackground("#1e3a5f").setFontColor("white").setFontWeight("bold");
  if(reporte.length > 0) {
    sAudit.getRange(2, 1, reporte.length, 5).setValues(reporte);
  }
  sAudit.setFrozenRows(1);
  sAudit.activate();
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. GENERADOR DE PLANTILLAS PRO E IMPORTACIÓN (ETL)
// ─────────────────────────────────────────────────────────────────────────────

function uiConfigurarPlantillaInteligente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS);
  
  const checkboxHtml = idiomas.map(i => {
    let tag = i.nombre_completo || i.nombre || i.id_idioma;
    let iso = (i.codigo_iso || i.id_idioma.substring(0,3)).toUpperCase();
    return `<div style="margin-bottom:5px;"><input type="checkbox" name="idiomas" value="${i.id_idioma}|${iso}|${tag}" checked> ${tag} (${iso})</div>`;
  }).join('');

  const html = `
    <div style="font-family:sans-serif; padding:15px; color:#333;">
      <h3 style="margin-top:0; color:#1e3a5f;">📄 Generar Plantilla de Trabajo</h3>
      
      <div style="background:#f1f5f9; padding:10px; border-radius:6px; margin-bottom:15px; font-size:13px;">
        <b>Modo de Plantilla:</b><br>
        <input type="radio" id="modo_a" name="modo" value="NUEVA" checked> <label for="modo_a">Carga Masiva (Nuevas)</label><br>
        <input type="radio" id="modo_b" name="modo" value="FALTANTES"> <label for="modo_b">Completar Pendientes</label>
      </div>

      <div style="max-height:120px; overflow-y:auto; border:1px solid #ddd; padding:8px; border-radius:4px; background:white;">
        ${checkboxHtml}
      </div>

      <div style="margin-top:10px; font-size:12px; border-top:1px solid #eee; padding-top:8px;">
        <b>¿Agregar nuevo idioma en esta carga?</b><br>
        <input type="text" id="n_nom" placeholder="Ej. Cora" style="width:50%;"> 
        <input type="text" id="n_iso" placeholder="ISO" style="width:30%;">
      </div>

      <button id="btn" onclick="run()" style="width:100%; background:#1e3a5f; color:white; border:none; padding:12px; border-radius:6px; margin-top:15px; cursor:pointer; font-weight:bold;">Generar y Analizar Drive ➔</button>
      
      <div id="status" style="margin-top:10px; font-size:11px; color:#d97706; display:none;">
        ⏳ <span id="status_msg">Analizando base de datos...</span>
      </div>

      <script>
        function run(){
          const btn = document.getElementById('btn');
          const idiomas = Array.from(document.querySelectorAll('input[name="idiomas"]:checked')).map(c => c.value);
          const modo = document.querySelector('input[name="modo"]:checked').value;
          
          btn.disabled = true; document.getElementById('status').style.display = 'block';
          
          google.script.run.withSuccessHandler(url => {
               google.script.host.close();
          }).servicioGenerarPlantillaPro(idiomas, modo, document.getElementById('n_nom').value, document.getElementById('n_iso').value);
        }
      </script>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(500), 'Nueva Plantilla');
}

function servicioGenerarPlantillaPro(seleccionados, modo, nNom, nIso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getAppConfig(ss);
  
  const dicAudio = _obtenerDiccionarioDrive(cfg.audios);
  const palabras = readSheet(ss, CONFIG.SHEETS.PALABRAS);
  const trads = readSheet(ss, CONFIG.SHEETS.TRADUCCIONES);
  const idIdiomaBase = String(cfg.idiomaPrincipal || "").trim();
  const mapaReferencia = _mapaTrads(ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES), idIdiomaBase);

  const nueva = SpreadsheetApp.create(`Plantilla_COLABORADOR_${modo}`);
  const h = nueva.getActiveSheet().setName('DATOS_A_IMPORTAR');
  
  let headers = ['ID_PALABRA', 'CONCEPTO_BASE', 'CATEGORIA', 'IMG_STATUS'];
  seleccionados.forEach(s => {
    let [id, iso] = s.split('|');
    headers.push(`TRADUCCION_${iso}`, `AUDIO_${iso}`, `NOTA_${iso}`, `TRAD_CAT_${iso}`);
  });
  if(nNom && nIso) headers.push(`TRADUCCION_${nIso}`, `AUDIO_${nIso}`, `NOTA_${nIso}`, `TRAD_CAT_${nIso}`);

  h.appendRow(headers);
  h.getRange(1, 1, 1, headers.length).setBackground('#1e3a5f').setFontColor('white').setFontWeight('bold');

  if (modo === 'FALTANTES') {
    let rows = [];
    palabras.forEach(p => {
      const idPal = String(p.id_palabra).trim();
      const textoBase = mapaReferencia[idPal] || "(Sin texto base)";
      
      let row = [idPal, textoBase, p.id_categoria, p.imagen_referencia ? "✅" : "❌"];
      let pendiente = false;

      seleccionados.forEach(s => {
        let [uuidL, iso] = s.split('|');
        let t = trads.find(tr => String(tr.id_palabra).trim() === idPal && String(tr.id_idioma).trim() === uuidL);
        
        let sugerido = `${textoBase}_${iso}`.toLowerCase();
        let statusAud = (t && t.audio) ? "✅ OK" : (dicAudio.has(sugerido) ? "🔍 En Drive" : "");
        
        row.push(t ? t.texto : "", statusAud, t ? t.nota_variante : "", "");
        if (!t || !t.audio) pendiente = true;
      });
      if (pendiente) rows.push(row);
    });
    if (rows.length > 0) h.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  h.setFrozenRows(1);
  h.setColumnWidths(1, headers.length, 150);
  DriveApp.getFileById(nueva.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  const url = nueva.getUrl();
  const modalHtml = `<div style="font-family:sans-serif;text-align:center;padding:15px;">
    <h3 style="color:#16a34a;">✅ Plantilla Lista</h3>
    <a href="${url}" target="_blank" style="display:block;background:#2563eb;color:white;padding:12px;text-decoration:none;border-radius:6px;font-weight:bold;margin-bottom:10px;">Abrir Plantilla</a>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(modalHtml).setWidth(350).setHeight(200), 'Resultado');
}

function uiImportarDatos() {
  const html = `<div style="font-family:sans-serif;padding:20px;text-align:center;">
    <h3 style="color:#1e3a5f;">📥 Importar Plantilla</h3>
    <input type="text" id="url" placeholder="Pega el enlace de Google Sheets aquí" style="width:100%;padding:8px;border-radius:4px;border:1px solid #ccc;">
    <button id="btn" onclick="run()" style="width:100%;margin-top:15px;padding:12px;background:#16a34a;color:white;border:none;border-radius:6px;font-weight:bold;cursor:pointer;">🚀 Analizar e Importar</button>
    <div id="res" style="margin-top:10px;font-size:12px;"></div>
    <script>
      function run(){
        const url = document.getElementById('url').value;
        if(!url) return alert("Pega la URL");
        document.getElementById('btn').disabled = true;
        document.getElementById('res').innerText = "⏳ Importando datos...";
        google.script.run.withSuccessHandler(r => document.getElementById('res').innerHTML=r).servicioProcesarImportacionMasiva(url);
      }
    </script>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(280), 'Importación');
}

function servicioProcesarImportacionMasiva(url) {
  try {
    const extSs = SpreadsheetApp.openByUrl(url);
    const data = (extSs.getSheetByName('DATOS_A_IMPORTAR') || extSs.getSheets()[0]).getDataRange().getValues();
    const rawH = data[0].map(h => String(h).toUpperCase().trim());
    
    const idxId = rawH.indexOf('ID_PALABRA'), idxBase = rawH.indexOf('CONCEPTO_BASE'), idxCat = rawH.indexOf('CATEGORIA');
    const langMap = {};
    rawH.forEach((h, i) => {
      let m = h.match(/^(TRADUCCION|TRAD_CAT|AUDIO|NOTA)_([A-Z0-9_]+)$/);
      if(m){
        let iso = m[2]; if(!langMap[iso]) langMap[iso] = {};
        if(m[1]==='TRADUCCION') langMap[iso].tr=i; if(m[1]==='TRAD_CAT') langMap[iso].tc=i;
        if(m[1]==='AUDIO') langMap[iso].au=i; if(m[1]==='NOTA') langMap[iso].no=i;
      }
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = getAppConfig(ss);
    const sIdiomas = ss.getSheetByName(CONFIG.SHEETS.IDIOMAS), sTrads = ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES), sPal = ss.getSheetByName(CONFIG.SHEETS.PALABRAS), sCats = ss.getSheetByName(CONFIG.SHEETS.CATEGORIAS), sTrCats = ss.getSheetByName(CONFIG.SHEETS.TRAD_CATS);

    const dbIdiomas = _mapaIdiomas(sIdiomas);
    const dbCats = _mapaId(sCats, 'id_categoria', ['nombre_categoria']);
    const dbPalById = arrayToMap(readSheet(ss, CONFIG.SHEETS.PALABRAS), 'id_palabra');
    const setTrExist = _setDoble(sTrads, 'id_palabra', 'id_idioma');

    let nIdiomas=[], nPal=[], nCats=[], nTrads=[];
    let stats = { upd:0, ins:0 };

    for (let i = 1; i < data.length; i++) {
      let row = data[i], idPal = _v(row, idxId), base = _v(row, idxBase), cat = _v(row, idxCat);
      if (idPal && dbPalById[idPal]) { stats.upd++; } 
      else if (base) {
        idPal = Utilities.getUuid();
        let idCat = dbCats[cat.toLowerCase()] || 'general';
        if(cat && !dbCats[cat.toLowerCase()]){ idCat=Utilities.getUuid(); nCats.push([idCat, cat, '']); dbCats[cat.toLowerCase()]=idCat; }
        nPal.push([idPal, idCat, '', '']);
        nTrads.push([Utilities.getUuid(), idPal, cfg.idiomaPrincipal, base, '', '', '']);
        stats.ins++;
      }
      if(!idPal) continue;

      Object.keys(langMap).forEach(iso => {
        if(!dbIdiomas[iso]){ let u=Utilities.getUuid(); nIdiomas.push([u, 'Idioma '+iso, iso]); dbIdiomas[iso]=u; }
        let uidL = dbIdiomas[iso], trTxt = _v(row, langMap[iso].tr);
        if (trTxt && !setTrExist.has(`${idPal}||${uidL}`)) {
          nTrads.push([Utilities.getUuid(), idPal, uidL, trTxt, '', _v(row, langMap[iso].au), _v(row, langMap[iso].no)]);
          setTrExist.add(`${idPal}||${uidL}`);
        }
      });
    }
    _batch(sIdiomas, nIdiomas, 3); _batch(sPal, nPal, 4); _batch(sCats, nCats, 3); _batch(sTrads, nTrads, 7);
    servicioLimpiarCache();
    return `<div style="color:#166534;"><b>✅ Éxito:</b> Actualizados: ${stats.upd}, Nuevos: ${stats.ins}</div>`;
  } catch (e) { return _msgError(e.message); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 6. GESTIÓN DE DISEÑO Y EXPORTACIÓN
// ─────────────────────────────────────────────────────────────────────────────

function servicioExportarListaDisenador() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sAudit = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
  if (!sAudit) return SpreadsheetApp.getUi().alert("⚠️ Primero genera el Dashboard de Control.");

  const data = sAudit.getDataRange().getValues();
  let lista = [["CONCEPTOS PENDIENTES DE ILUSTRACIÓN"]];
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).includes("PENDIENTE")) {
      lista.push([data[i][1]]);
    }
  }

  if (lista.length <= 1) return SpreadsheetApp.getUi().alert("🎉 No hay imágenes pendientes.");

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nueva = SpreadsheetApp.create(`LISTA_DISEÑO_PENDIENTES_${ts}`);
  nueva.getActiveSheet().getRange(1, 1, lista.length, 1).setValues(lista).setFontWeight("bold");
  nueva.getActiveSheet().setColumnWidth(1, 400);
  
  const url = nueva.getUrl();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<div style="padding:20px;text-align:center;font-family:sans-serif;">
      <p>Se creó una lista con <b>${lista.length - 1}</b> conceptos.</p>
      <a href="${url}" target="_blank" style="display:block;background:#1e3a5f;color:white;padding:12px;text-decoration:none;border-radius:6px;font-weight:bold;">📥 Abrir Lista para el Diseñador</a>
    </div>`).setWidth(350).setHeight(180), "Exportación Exitosa"
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// 7. HELPERS DE MAPEO E INTELIGENCIA
// ─────────────────────────────────────────────────────────────────────────────

function _obtenerDiccionarioDrive(folderId) {
  const diccionario = new Set();
  if (!folderId || folderId.length < 5) return diccionario;
  try {
    const archivos = DriveApp.getFolderById(folderId).getFiles();
    while (archivos.hasNext()) {
      let f = archivos.next();
      let nom = f.getName().toLowerCase();
      diccionario.add(nom);
      diccionario.add(nom.split('.')[0]);
    }
  } catch (e) { console.warn("Carpeta Drive no encontrada: " + folderId); }
  return diccionario;
}

function _mapaTrads(sheet, idPrincipal) {
  const map = {}; 
  if (!sheet) return map;
  const d = sheet.getDataRange().getValues();
  if (d.length < 2) return map;
  
  const h = d[0].map(x => String(x).toLowerCase().trim());
  const iPal = h.indexOf('id_palabra'), iIdi = h.indexOf('id_idioma'), iTxt = h.indexOf('texto');
  
  if (iPal === -1 || iIdi === -1 || iTxt === -1) return map;

  d.slice(1).forEach(r => { 
    if(String(r[iIdi]).trim() === String(idPrincipal).trim()) {
      map[String(r[iPal]).trim()] = String(r[iTxt]).trim(); 
    }
  });
  return map;
}

function readSheet(ss, name) {
  const s = ss.getSheetByName(name); if (!s) return []; 
  const d = s.getDataRange().getValues(); if (d.length < 2) return [];
  const h = d[0].map(x => String(x).toLowerCase().trim());
  return d.slice(1).map(row => {
    let o = {}; h.forEach((k, i) => o[k] = row[i]); return o;
  });
}

function _mapaIdiomas(sheet) {
  const map = {}; if (!sheet) return map;
  const d = sheet.getDataRange().getValues();
  d.slice(1).forEach(r => {
    let id = String(r[0]).trim();
    let iso = String(r[2] || "").trim().toUpperCase();
    if (id) { map[id.toUpperCase()] = id; if(iso) map[iso] = id; }
  });
  return map;
}

function _mapaId(sheet, colId, colsNombre) {
  const map = {}; if (!sheet) return map;
  const d = sheet.getDataRange().getValues();
  const h = d[0].map(x => String(x).toLowerCase().trim());
  const idxId = h.indexOf(colId.toLowerCase());
  const idxNom = colsNombre.map(c => h.indexOf(c.toLowerCase())).find(i => i > -1);
  if (idxId > -1 && idxNom > -1) {
    d.slice(1).forEach(r => { 
      let id = String(r[idxId]).trim(), nom = String(r[idxNom]).trim().toLowerCase(); 
      if(id && nom) map[nom] = id; 
    });
  }
  return map;
}

function _setDoble(sheet, col1, col2) {
  const s = new Set(); if (!sheet) return s;
  const d = sheet.getDataRange().getValues(); if (d.length < 2) return s;
  const h = d[0].map(x => String(x).toLowerCase().trim());
  const i1 = h.indexOf(col1.toLowerCase()), i2 = h.indexOf(col2.toLowerCase());
  if (i1 > -1 && i2 > -1) d.slice(1).forEach(r => { s.add(`${String(r[i1]).trim()}||${String(r[i2]).trim().toUpperCase()}`); });
  return s;
}

// ─────────────────────────────────────────────────────────────────────────────
// 8. MANTENIMIENTO Y UTILIDADES
// ─────────────────────────────────────────────────────────────────────────────

function servicioRepararIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const defs = [{ n: CONFIG.SHEETS.CATEGORIAS, id: 'id_categoria' }, { n: CONFIG.SHEETS.TRADUCCIONES, id: 'id_traduccion' }, { n: CONFIG.SHEETS.PALABRAS, id: 'id_palabra' }];
  defs.forEach(h => {
    const sheet = ss.getSheetByName(h.n); if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const idIdx = data[0].map(x => String(x).toLowerCase().trim()).indexOf(h.id);
    if (idIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        if (!String(data[i][idIdx]).trim() && data[i].some(v => String(v).trim())) {
          sheet.getRange(i + 1, idIdx + 1).setNumberFormat('@').setValue(Utilities.getUuid());
        }
      }
    }
  });
  SpreadsheetApp.getUi().alert("✅ IDs verificados y reparados.");
}

function servicioInstalacionCarpetas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEETS.IDIOMAS, CONFIG.SHEETS.TRADUCCIONES, CONFIG.SHEETS.TRAD_CATS, CONFIG.SHEETS.PALABRAS, CONFIG.SHEETS.CATEGORIAS, CONFIG.SHEETS.CONFIG].forEach(n => {
    if(!ss.getSheetByName(n)) ss.insertSheet(n);
  });
  SpreadsheetApp.getUi().alert("✅ Estructura de hojas verificada.");
}

function servicioNormalizacionTexto() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TRADUCCIONES);
  if(!s) return;
  const d = s.getDataRange().getValues();
  const n = d.map(r => r.map(c => typeof c === 'string' ? c.replace(/\+/g, 'ɨ') : c));
  s.getDataRange().setValues(n);
  SpreadsheetApp.getUi().alert("✅ Tipografía normalizada (+ ➝ ɨ).");
}

function arrayToMap(arr, key) { return arr.reduce((acc, item) => { acc[item[key]] = item; return acc; }, {}); }
function _v(row, idx) { return (idx === -1 || idx >= row.length || row[idx] == null) ? '' : String(row[idx]).trim(); }
function _batch(sheet, filas, numCols) { if (sheet && filas && filas.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, filas.length, numCols).setValues(filas); }

function servicioLimpiarCache() { 
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CONFIG.CACHE.KEY);
    cache.remove(CONFIG.CACHE.KEY + '_glosario');
    cache.remove(CONFIG.CACHE.KEY + '_idiomas');
    SpreadsheetApp.getUi().alert("✅ Memoria Caché borrada. La web ya puede leer los datos nuevos.");
  } catch(e) {
    SpreadsheetApp.getUi().alert("❌ Error al limpiar: " + e.message);
  }
  return true; 
}

function uiExportarDatos() { SpreadsheetApp.getUi().alert("Usa el botón de respaldo en el modal de Importación Inteligente."); }
function _msgWarn(m) { return `<div style="color:#b45309;background:#fef3c7;padding:8px;border-radius:5px;">⚠️ ${m}</div>`; }
function _msgError(m) { return `<div style="color:red;background:#fee2e2;padding:8px;border-radius:5px;">❌ ${m}</div>`; }

// ─────────────────────────────────────────────────────────────────────────────
// 9. ENDPOINT API REST (WEB APP)
// ─────────────────────────────────────────────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action || 'ping';
  const cacheKey = `${CONFIG.CACHE.KEY}_${action}`;

  if (CONFIG.CACHE.ENABLED && action !== 'ping') {
    const cachedData = CacheService.getScriptCache().get(cacheKey);
    if (cachedData) return ContentService.createTextOutput(cachedData).setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getAppConfig(ss);
  const tokenUser = e.parameter.token || '';

  if (cfg.apiToken && cfg.apiToken.length > 2 && tokenUser !== cfg.apiToken) {
    return responseJSON({ status: 401, error: "Unauthorized: Token inválido o ausente." });
  }

  let responseObj = {};
  
  try {
    if (action === 'idiomas') {
      responseObj = { status: 200, data: readSheet(ss, CONFIG.SHEETS.IDIOMAS) };
      
    } else if (action === 'glosario') {
      responseObj = {
        status: 200,
        data: {
          config: cfg,
          idiomas: readSheet(ss, CONFIG.SHEETS.IDIOMAS),
          categorias: readSheet(ss, CONFIG.SHEETS.CATEGORIAS),
          palabras: readSheet(ss, CONFIG.SHEETS.PALABRAS),
          traducciones: readSheet(ss, CONFIG.SHEETS.TRADUCCIONES)
        }
      };
      
    } else {
      responseObj = { 
        status: 200, 
        message: "API Activa", 
        endpoints_disponibles: ["?action=idiomas", "?action=glosario"] 
      };
    }

    if (CONFIG.CACHE.ENABLED && action !== 'ping') {
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(responseObj), CONFIG.CACHE.TTL);
    }

    return responseJSON(responseObj);

  } catch (err) {
    return responseJSON({ status: 500, error: err.message });
  }
}