/**
 * SISTEMA DE GESTIÓN LINGÜÍSTICA — catálogo multilingüe comunitario
 * ===================================================================
 * Copyright (c) 2025 Alejandro Estrada
 * SPDX-License-Identifier: AGPL-3.0-only
 * @version         1.1
 * -------------------------------------------------------------------
 * Este programa es software libre: puedes redistribuirlo y modificarlo según los
 * términos de la Licencia Pública General Affero de GNU, versión 3, con los
 * términos adicionales de su sección 7 (crédito de autoría visible y marca de
 * versiones modificadas) descritos en NOTICE.md del repositorio original:
 * https://github.com/joalesmu/Glosario_Multiling-e_Comunitario
 * Se distribuye SIN NINGUNA GARANTÍA. Texto completo: https://www.gnu.org/licenses/agpl-3.0.html
 * Los contenidos de cada catálogo pertenecen a sus comunidades y tienen su propia licencia.
 * -------------------------------------------------------------------
 * Principios: ninguna lengua es obligatoria ni pivote; la comunidad decide qué
 * se publica; nadie necesita editar tablas a mano; todo funciona gratis con
 * Google (Sheets, Drive, Apps Script) y GitHub Pages; los datos quedan listos
 * para intercambiarse entre comunidades y para migrar a bases de grafos.
 *
 * Cambios de la 1.0 a la 1.1
 *  INSTALACIÓN Y PANEL
 *  - Funciona en una hoja en blanco: asistente que crea hojas, columnas,
 *    carpetas por lengua, atajos de escritura y ajustes.
 *  - Panel lateral: estado del catálogo, lenguas, ajustes, contacto, estados,
 *    reglas de escritura, publicación, intercambio, respaldo y reparación.
 *  - La estructura se repara sola: agrega lo que falta, nunca borra ni reordena.
 *  DATOS Y ESCRITURA
 *  - Escritura por nombre de columna, índices en lugar de búsquedas repetidas.
 *  - Reglas de escritura ("buscar y reemplazar" permanente) por lengua, para el
 *    equipo o por persona; modo captura (texto_captura → texto) sin pérdidas.
 *  - Alfabeto por lengua (con dígrafos) para juegos de letras.
 *  - Estados como menú desplegable (Borrador, En revisión, Publicado, Retirado)
 *    con gestor en lote; solo lo publicado sale; los estados nunca salen.
 *  - Marca "usar en juegos" por palabra (palabras rituales o delicadas).
 *  - Imágenes de Drive: se resuelven rutas de AppSheet, enlaces, hipervínculos
 *    y fórmulas a un ID; las publicadas se comparten para que el visor las vea.
 *  PUBLICACIÓN Y API
 *  - Publicación previa en Drive + caché; manifiesto de versión; sin token ni
 *    datos internos en la API; audio solo de traducciones publicadas.
 *  - Contacto, créditos y "acerca de" configurables y publicados.
 *  - Paquete de imágenes ligeras para jugar sin internet.
 *  - Lotería en línea: partidas "Nombre + 4 números", rondas, ganadores.
 *  EQUIPO Y CONSENTIMIENTO
 *  - Hoja EQUIPO (correo, nombre, rol): administradores y capturistas. AppSheet
 *    lee los permisos de ahí; el panel 👥 Equipo da o quita acceso a la hoja.
 *  - Cada audio registra quién habla, de dónde y su consentimiento. Solo se
 *    publica el audio con consentimiento público; el nombre de la voz, solo si
 *    la persona aceptó que aparezca.
 *  INTERCAMBIO Y GRAFOS
 *  - Exportar/importar catálogos entre comunidades sin duplicar palabras ni
 *    traducciones, con vista previa; lo importado entra "En revisión".
 *  - Exportación a grafos: N-Triples (OntoLex/SKOS; Apache Jena), CSV de nodos
 *    y aristas (Neo4j, Apache AGE) y GraphML (Apache TinkerPop, Gephi).
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
    REGLAS:       'REGLAS_ESCRITURA',
    RELACIONES:   'RELACIONES',
    EQUIPO:       'EQUIPO',
    REPORT:       'MESA_DE_TRABAJO',
    AUDIT:        'AUDITORIA_ADMIN',
    TEXTOS:       'TEXTOS',
    FUENTES:      'FUENTES',
    VOCES:        'VOCES',
    ACTORES:      'ACTORES',
    AUTORIZACIONES: 'AUTORIZACIONES',
    COMUNIDADES:  'COMUNIDADES',
    HISTORIAL:    'HISTORIAL',
    COPIAS:       'COPIAS_TRABAJO',
    MEDIOS:       'MEDIOS_PROPUESTOS',
    AVISOS:       'AVISOS',
    TIPOS_REL:    'TIPOS_RELACION',
    PROPUESTAS:   'PROPUESTAS',
    IMAGENES:     'IMAGENES'
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
  // Vistas de AppSheet usadas en los enlaces directos
  VISTAS: { TRAD_FORM: 'TRADUCCIONES_Form', PALABRA_DETALLE: 'PALABRAS_Detail' },
  // Estados (columna "estado"): borrador, revisión, publicado, retirado. No son lineales:
  // cualquier registro puede pasar a cualquier estado. SOLO se publica "publicado".
  // Por compatibilidad, una celda vacía (o una hoja sin columna "estado") cuenta como publicada.
  // Columnas que nunca salen en la API
  // Columnas que nunca salen en la API (los estados son solo para administración)
  COLUMNAS_PRIVADAS: ['texto_captura', 'creado_por', 'correo', 'email', 'estado', 'id_voz', 'lote', 'imagen_ref', 'imagen_publica', 'usar_en_juegos', 'origen', 'consentimiento', 'audio_archivo', 'imagen_nota', 'audio_nota'],
  VERSION: '1.1',
  CACHE: { ENABLED: true, TTL: 21600, PREFIX: 'CATALOGO_V2', CHUNK_CHARS: 30000, MAX_CHUNKS: 150 },
  PROPS: { SNAPSHOT_FILE: 'SNAPSHOT_FILE_ID', MANIFIESTO: 'MANIFIESTO', TOKEN: 'API_TOKEN_PUBLICADO' },
  SNAPSHOT_NOMBRE: 'catalogo_publicado.json',
};

// ─────────────────────────────────────────────────────────────────────────────
// 1. MENÚ DE ADMINISTRACIÓN
// ─────────────────────────────────────────────────────────────────────────────
function onOpen() {
  _construirMenu();
  _mensajeHojaEnBlanco();
}

/**
 * Menú ordenado por intención y en el orden del trabajo (ver → trabajar → colaborar → configurar → mantener).
 * Arriba solo lo diario; lo ocasional en submenús; lo que borra, al final y separado; el idioma de los menús
 * siempre en el último lugar y con 🌐, para encontrarlo sin leer. Mientras el catálogo no esté configurado,
 * «Primeros pasos» aparece primero. En el idioma de menús de cada persona.
 */
function _construirMenu() {
  _dicMemo = null;
  const ui = SpreadsheetApp.getUi();
  let nuevo = false;
  try { const ss = SpreadsheetApp.getActiveSpreadsheet(); nuevo = !ss.getSheetByName(CONFIG.SHEETS.CONFIG) || !getAppConfig(ss).titulo; } catch (e) { nuevo = false; }
  const m = ui.createMenu(tx('menu.titulo'));
  if (nuevo) m.addItem(tx('menu.herr.primeros'), 'uiAsistenteInicial').addSeparator();
  m.addItem(tx('menu.panel'),      'uiPanel')
    .addItem(tx('menu.comoVa'),     'uiEstadoCatalogo')
    .addSeparator()
    .addItem(tx('menu.editar'),     'uiEditar')
    .addItem(tx('menu.propuestas'), 'uiPropuestas')
    .addItem(tx('menu.medios'),     'uiMedios')
    .addSubMenu(ui.createMenu(tx('menu.falta'))
      .addItem(tx('menu.falta.lista'),     'uiGenerarReporteAvanzado')
      .addItem(tx('menu.falta.guardar'),   'servicioImportarDesdeReporte')
      .addItem(tx('menu.falta.imagenes'),  'uiGenerarDashboardAdmin'))
    .addItem(tx('menu.estados'),    'uiEstados')
    .addItem(tx('menu.publicar'),   'uiPublicarCatalogo')
    .addSeparator()
    .addSubMenu(ui.createMenu(tx('menu.colaborar'))
      .addItem(tx('menu.compartir.hoja'),        'uiConfigurarPlantillaInteligente')
      .addItem(tx('menu.falta.ilustrar'),        'uiMediosIlustrar')
      .addItem(tx('menu.compartir.traer'),       'uiImportarDatos')
      .addItem(tx('menu.compartir.comunidades'), 'uiIntercambio')
      .addItem(tx('menu.avisos'),                'uiAvisos')
      .addItem(tx('menu.constancia'),            'uiConstancia')
      .addItem(tx('menu.herr.direccion'),        'uiVerEnlaceApi'))
    .addSeparator()
    .addSubMenu(ui.createMenu(tx('menu.configurar'))
      .addItem(tx('menu.lenguas'),    'uiLenguas')
      .addItem(tx('menu.equipo'),     'uiEquipo')
      .addItem(tx('menu.voces'),      'uiVoces')
      .addItem(tx('menu.fuentes'),    'uiFuentes')
      .addItem(tx('menu.atajos'),     'uiReemplazos')
      .addItem(tx('menu.tiposRel'),   'uiTiposRelacion')
      .addItem(tx('menu.herr.datos'), 'uiAjustes')
      .addSeparator()
      .addItem(tx('menu.herr.primeros'), 'uiAsistenteInicial'))
    .addSubMenu(ui.createMenu(tx('menu.mantenimiento'))
      .addItem(tx('menu.coincidencias'),      'uiCoincidencias')
      .addItem(tx('menu.ramas'),              'uiRamas')
      .addItem(tx('menu.herr.reparar'),       'servicioInstalacionCarpetas')
      .addItem(tx('menu.herr.codigos'),       'servicioRepararIDs')
      .addItem(tx('menu.herr.aplicarAtajos'), 'servicioNormalizacionTexto')
      .addItem(tx('menu.herr.modoCaptura'),   'servicioActivarModoCaptura')
      .addItem(tx('menu.herr.visor'),         'servicioLimpiarCache')
      .addItem(tx('menu.compartir.copia'),    'uiExportarDatos')
      .addSeparator()
      .addItem(tx('menu.herr.limpiar'),       'uiLimpiar'))
    .addSeparator()
    .addItem(tx('menu.idioma'),     'uiIdiomaMenus')
    .addToUi();
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. CONFIGURACIÓN Y ACCESO A TABLAS (POR NOMBRE DE COLUMNA)
// ─────────────────────────────────────────────────────────────────────────────

const _s = x => String(x == null ? '' : x).trim();
const _k = x => _s(x).toLowerCase();
const _esc = x => String(x == null ? '' : x).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

const _sinAcentos = x => _k(x).normalize('NFD').replace(/[̀-ͯ]/g, '');

/** Lleva cualquier forma de escribir el estado a uno de: borrador, revisión, publicado, retirado. */
function _estadoCanonico(v) {
  const e = _sinAcentos(v);
  if (!e) return '';
  if (/^(borrador|draft)/.test(e)) return 'borrador';
  if (/(revision|revisar|propuesta|pendiente)/.test(e)) return 'revisión';
  if (/^(publicad|publicar|visible)/.test(e)) return 'publicado';
  if (/^(retirad|oculto|archivad)/.test(e)) return 'retirado';
  return e; // valor desconocido: no se publica
}

/** Etiquetas que se ven en el menú desplegable (hoja y AppSheet). La lógica usa la forma canónica. */
const ESTADOS = [
  { canon: 'borrador',  etiqueta: 'Borrador' },
  { canon: 'revisión',  etiqueta: 'En revisión' },
  { canon: 'publicado', etiqueta: 'Publicado' },
  { canon: 'retirado',  etiqueta: 'Retirado' }
];
function _etiquetaEstado(v) {
  const c = _estadoCanonico(v);
  const e = ESTADOS.find(x => x.canon === c);
  return e ? e.etiqueta : 'Borrador';
}

/** Solo lo publicado (o sin estado, por compatibilidad) sale en el catálogo. */
function _esPublico(o) { const e = _estadoCanonico(o && o.estado); return e === '' || e === 'publicado'; }

// Consentimiento de cada grabación (columna "consentimiento" de TRADUCCIONES).
// Vacío = registros anteriores a la 1.1: el audio se publica, el nombre de la voz no.
const CONSENTIMIENTOS = [
  { canon: 'publico_nombre',  etiqueta: 'Público con nombre' },   // las etiquetas son datos de la hoja: no se traducen
  { canon: 'publico_anonimo', etiqueta: 'Público sin nombre' },
  { canon: 'comunidad',       etiqueta: 'Solo comunidad' },
  { canon: 'pendiente',       etiqueta: 'Pendiente' }
];
function _consentimiento(v) {
  const s = _sinAcentos(v);
  if (!s) return '';
  if (/comunidad/.test(s)) return 'comunidad';
  if (/pendiente|^no$|no public/.test(s)) return 'pendiente';
  if (/sin nombre|anonim/.test(s)) return 'publico_anonimo';
  if (/nombre|credito/.test(s)) return 'publico_nombre';
  return 'publico_anonimo';   // "Público", "Sí"…: se publica sin nombre (la opción más cuidadosa)
}
/** ¿El audio de esta traducción puede salir del catálogo? */
function _audioPublico(t) { const c = _consentimiento(t && t.consentimiento); return c === '' || c === 'publico_nombre' || c === 'publico_anonimo'; }
/** ¿Se puede decir quién habla? Solo con permiso explícito. */
function _vozPublica(t) { return _consentimiento(t && t.consentimiento) === 'publico_nombre'; }

let _cfgMemo = null;
function getAppConfig(ss) {
  if (_cfgMemo) return _cfgMemo;
  const s = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!s) throw new Error(tx('comun.sinPreparar'));
  const d = s.getDataRange().getValues();
  const h = d[0].map(x => _s(x).toUpperCase());
  const v = d[1] || [];
  const get = k => { const i = h.indexOf(k.toUpperCase()); return i > -1 ? _s(v[i]) : ''; };
  _cfgMemo = {
    audios:          get(CONFIG.COLS.AUDIO_ID),
    imagenes:        get(CONFIG.COLS.IMG_ID),
    appsheetId:      get(CONFIG.COLS.APPSHEET_ID),
    titulo:          get(CONFIG.COLS.PROJECT_NAME),
    subtitulo:       get('SUBTITULO'),
    idiomaPrincipal: get(CONFIG.COLS.MAIN_LANG),
    apiToken:        get(CONFIG.COLS.API_TOKEN),
    raiz:            get('ID_CARPETA_RAIZ'),
    idCatalogo:      get('ID_CATALOGO'),
    contacto: {
      nombre: get('CONTACTO_NOMBRE'), correo: get('CONTACTO_CORREO'), telefono: get('CONTACTO_TELEFONO'),
      whatsapp: get('CONTACTO_WHATSAPP'), horario: get('CONTACTO_HORARIO'), texto: get('CONTACTO_TEXTO')
    },
    creditos:        get('CREDITOS'),
    acerca:          get('ACERCA'),
    licencia:        get('LICENCIA'),
    videos:          get('ID_CARPETA_VIDEOS'),
    carpetaCategorias: get(CONFIG.COLS.CAT_ID),
    idiomaInterfaz:  get('IDIOMA_INTERFAZ'),
    solicitarIA:     get('SOLICITAR_IA'),
    urlApi:          get('URL_API'),
    urlVisor:        get('URL_VISOR'),
    claveEquipo:     get('CLAVE_EQUIPO'),
    visorAjustes:    get('VISOR_AJUSTES')
  };
  return _cfgMemo;
}

/** Lee una hoja completa y devuelve sus filas con un índice de columnas por nombre (en minúsculas). */
function _tabla(ss, nombre) {
  const sh = ss.getSheetByName(nombre);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const headers = (data[0] || []).map(h => _k(h));
  const idx = {};
  headers.forEach((h, i) => { if (h && !(h in idx)) idx[h] = i; });
  return { sh, nombre, headers, idx, filas: data.length > 1 ? data.slice(1) : [] };
}

function _col(t, nombre) { return t && (nombre in t.idx) ? t.idx[nombre] : -1; }
function _get(t, fila, nombre) { const i = _col(t, nombre); return i > -1 ? fila[i] : ''; }

function _requerir(t, cols, nombreHoja) {
  if (!t) throw new Error(tx('comun.faltaHoja', { hoja: nombreHoja }));
  const faltan = cols.filter(c => !(c in t.idx));
  if (faltan.length) throw new Error(tx('comun.faltanColumnas', { hoja: t.nombre, columnas: faltan.join(', ') }));
}

/** Agrega filas a partir de objetos {columna: valor}. Las columnas que no existen en la hoja se ignoran. */
function _agregar(t, objetos) {
  if (!t || !objetos || !objetos.length) return;
  const ancho = t.headers.length;
  const filas = objetos.map(o => {
    const f = new Array(ancho).fill('');
    Object.keys(o).forEach(k => { const i = _col(t, k); if (i > -1) f[i] = o[k]; });
    return f;
  });
  const inicio = t.sh.getLastRow() + 1;
  // Columnas de texto en formato plano: "+ki", "-ta" o "=…" no deben volverse fórmulas
  COLS_TEXTO_PLANO.forEach(k => { const i = _col(t, k); if (i > -1) _formatoTexto(t.sh, inicio, i + 1, filas.length); });
  t.sh.getRange(inicio, 1, filas.length, ancho).setValues(filas);
  filas.forEach(f => t.filas.push(f));
}

/** Actualiza celdas puntuales de una fila (índice 0 = primera fila de datos). Solo toca las celdas que cambian. */
function _actualizar(t, filaIdx, cambios) {
  Object.keys(cambios).forEach(k => {
    const c = _col(t, k);
    if (c < 0) return;
    const celda = t.sh.getRange(filaIdx + 2, c + 1);
    if (COLS_TEXTO_PLANO.includes(k)) celda.setNumberFormat('@');
    celda.setValue(cambios[k]);
    t.filas[filaIdx][c] = cambios[k];
  });
}

/** Índice palabra||idioma → fila de TRADUCCIONES. */
function _indiceTrads(t) {
  const m = new Map();
  const iP = _col(t, 'id_palabra'), iL = _col(t, 'id_idioma');
  t.filas.forEach((f, i) => { if (_s(f[iP])) m.set(_k(f[iP]) + '||' + _k(f[iL]), i); });
  return m;
}

/** Mapa id_palabra → texto en la lengua de referencia. */
function _mapaTextoBase(tTr, idIdioma) {
  const m = new Map();
  const ref = _k(idIdioma);
  if (!tTr || !ref) return m;
  tTr.filas.forEach(f => {
    if (_k(_get(tTr, f, 'id_idioma')) === ref) m.set(_k(_get(tTr, f, 'id_palabra')), _s(_get(tTr, f, 'texto')));
  });
  return m;
}

/** Lectura como objetos (se conserva por compatibilidad con la API anterior). */
function readSheet(ss, name) {
  const s = ss.getSheetByName(name); if (!s) return [];
  const d = s.getDataRange().getValues(); if (d.length < 2) return [];
  const h = d[0].map(x => _k(x));
  return d.slice(1).map(row => { const o = {}; h.forEach((k, i) => { if (k) o[k] = row[i]; }); return o; });
}

function responseJSON(d) {
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}

function _lock() { const l = LockService.getScriptLock(); l.waitLock(30000); return l; }

function _rt(texto, url) {
  const b = SpreadsheetApp.newRichTextValue().setText(texto);
  if (url) b.setLinkUrl(url);
  return b.build();
}

function _linkAppSheet(cfg, vista, opts) {
  if (!cfg.appsheetId) return '';
  let url = `https://www.appsheet.com/start/${cfg.appsheetId}#control=${encodeURIComponent(vista)}`;
  if (opts && opts.row) url += `&row=${encodeURIComponent(opts.row)}`;
  if (opts && opts.defaults) url += `&defaults=${encodeURIComponent(JSON.stringify(opts.defaults))}`;
  return url;
}

function _limpiarHoja(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.clear();
}

function _carpetaDelLibro(ss) {
  const padres = DriveApp.getFileById(ss.getId()).getParents();
  return padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
}

/** Las hojas que se comparten con quien colabora se guardan junto al catálogo, en COLABORACIONES (y no sueltas en Mi unidad). */
function _aColaboraciones(ss, nueva) {
  try { DriveApp.getFileById(nueva.getId()).moveTo(_subcarpeta(_carpetaDelLibro(ss), 'COLABORACIONES')); }
  catch (e) { console.warn('No se pudo mover a COLABORACIONES: ' + e.message); } // se queda en Mi unidad: no se pierde nada
  return nueva;
}

function _subcarpeta(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

function _filasIdiomasHtml(ss, nombreInput) {
  return readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma)).map(i => {
    const tag = _s(i.nombre_completo || i.nombre || i.id_idioma);
    const iso = _s(i.codigo_iso || _s(i.id_idioma).substring(0, 3)).toUpperCase();
    return `<div style="margin-bottom:6px;"><label><input type="checkbox" name="${nombreInput}" value="${_esc(i.id_idioma + '|' + iso + '|' + tag)}" checked> ${_esc(tag)} (${_esc(iso)})</label></div>`;
  }).join('');
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. REGLAS DE ESCRITURA ("BUSCAR Y REEMPLAZAR" PERMANENTE Y AUTOMÁTICO)
// ─────────────────────────────────────────────────────────────────────────────
// Hoja REGLAS_ESCRITURA (cada fila es un "Buscar y reemplazar" que se aplica solo):
//   id_idioma | buscar | reemplazar | tipo | orden | activa | aplica_a | coincidir_mayus | nota
//  - id_idioma: ID o código ISO del idioma, o * para todos.
//  - tipo: "texto" (reemplazo literal) o "regex" (expresión regular, admite $1, $2…).
//  - aplica_a: "todos" (o vacío) = todo el equipo; un correo = solo los registros de esa
//    persona (según la columna creado_por de TRADUCCIONES).
//  - coincidir_mayus: VERDADERO por defecto, así "O9 → Ö" y "o9 → ö" son reglas distintas.
// Siempre, antes de las reglas: Unicode NFC y espacios sobrantes.
//
// Dos modos de trabajo:
//  - MODO CAPTURA (si TRADUCCIONES tiene la columna texto_captura): lo que la persona escribe
//    se guarda en texto_captura y "texto" se CALCULA siempre desde ahí. Cambiar o desactivar
//    una regla recalcula todo desde el original: nada se pierde.
//  - MODO DIRECTO (sin texto_captura): las reglas se aplican sobre "texto" y solo se usan las
//    reglas estables (aplicarlas dos veces da lo mismo que una), para no acumular cambios.

const COLS_REGLAS = ['id_idioma', 'buscar', 'reemplazar', 'tipo', 'orden', 'activa', 'aplica_a', 'coincidir_mayus', 'nota'];
const COLS_TEXTO_PLANO = ['texto', 'texto_captura', 'buscar', 'reemplazar', 'nota_variante', 'nombre', 'nombre_categoria', 'descripcion'];

const _verdadero = v => v === true || ['true', 'verdadero', 'sí', 'si', '1', 'x'].includes(_k(v));
const _escRegex = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

function _miCorreo() {
  try { return _k(Session.getActiveUser().getEmail()); } catch (e) { return ''; }
}

function _mapaIsoAId(ss) {
  const m = new Map();
  const tId = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
  if (tId) tId.filas.forEach(f => {
    const id = _s(_get(tId, f, 'id_idioma')), iso = _k(_get(tId, f, 'codigo_iso'));
    if (id && iso) m.set(iso, _k(id));
  });
  return m;
}

function _claveIdioma(valor, isoAId) {
  const k = _k(valor);
  if (!k || k === '*') return k;
  return isoAId.has(k) ? isoAId.get(k) : k;
}

/** Prepara la función que aplica la regla. Lanza un error comprensible si la regla no es válida. */
function _compilarRegla(r) {
  if (!r.buscar) throw new Error(tx('atajos.error.vacio'));
  if (r.buscar === r.reemplazar) throw new Error(tx('atajos.error.iguales'));
  if (r.tipo === 'regex') {
    let re;
    try { re = new RegExp(r.buscar, 'gu' + (r.mayus ? '' : 'i')); }
    catch (e) { throw new Error(tx('atajos.error.patron', { m: e.message })); }
    r.aplicar = t => t.replace(re, r.reemplazar);
  } else if (r.mayus) {
    r.aplicar = t => t.split(r.buscar).join(r.reemplazar);
  } else {
    const re = new RegExp(_escRegex(r.buscar), 'giu');
    r.aplicar = t => t.replace(re, () => r.reemplazar);
  }
  return r;
}

/** Estable = aplicarla dos veces da lo mismo que aplicarla una vez (p. ej. "+ → ɨ" sí; "a → aa" no). */
function _reglaEstable(r, muestras) {
  const pruebas = (muestras || []).concat([r.buscar, `a${r.buscar}a`, r.buscar + r.buscar, ` ${r.buscar} `]);
  return pruebas.every(m => { const una = r.aplicar(m); return r.aplicar(una) === una; });
}

/** Primero las reglas personales (abreviaturas de quien escribe), después las del equipo; dentro de cada grupo, por orden. */
function _ordenarReglas(arr) {
  return arr.sort((a, b) => ((a.aplicaA === 'todos') - (b.aplicaA === 'todos')) || (a.orden - b.orden));
}

function _cargarReglas(ss) {
  const m = new Map();
  const t = _tabla(ss, CONFIG.SHEETS.REGLAS);
  if (!t) return m;
  const isoAId = _mapaIsoAId(ss);
  t.filas.forEach((f, i) => {
    if (!_verdadero(_get(t, f, 'activa'))) return;
    const clave = _claveIdioma(_get(t, f, 'id_idioma'), isoAId);
    if (!clave) return;
    const celdaMayus = _get(t, f, 'coincidir_mayus');
    let aplicaA = _k(_get(t, f, 'aplica_a'));
    if (!aplicaA || aplicaA === 'todo el equipo') aplicaA = 'todos';
    const r = {
      fila: i,
      buscar: String(_get(t, f, 'buscar')),
      reemplazar: String(_get(t, f, 'reemplazar')),
      tipo: _k(_get(t, f, 'tipo')) === 'regex' ? 'regex' : 'texto',
      orden: Number(_get(t, f, 'orden')) || 0,
      mayus: _s(celdaMayus) === '' ? true : _verdadero(celdaMayus),
      aplicaA
    };
    try { _compilarRegla(r); } catch (e) { console.warn(`Regla de la fila ${i + 2} ignorada: ${e.message}`); return; }
    r.estable = _reglaEstable(r);
    if (!m.has(clave)) m.set(clave, []);
    m.get(clave).push(r);
  });
  m.forEach(arr => _ordenarReglas(arr));
  return m;
}

/**
 * Normaliza un texto: NFC + espacios, luego las reglas de "todos los idiomas" (*) y las del idioma.
 * Las reglas personales solo se aplican a los textos de su autor.
 */
function _normalizar(texto, idIdioma, reglas, autor, soloEstables) {
  if (typeof texto !== 'string' || !texto) return texto;
  let t = texto.normalize('NFC').replace(/[ \t ]+/g, ' ').trim();
  const quien = _k(autor);
  [].concat(reglas.get('*') || [], reglas.get(_k(idIdioma)) || [])
    .filter(r => (r.aplicaA === 'todos' || (quien && r.aplicaA === quien)) && (!soloEstables || r.estable))
    .forEach(r => { t = r.aplicar(t); });
  return t.normalize('NFC');
}

/**
 * Recalcula "texto" en TRADUCCIONES.
 * Modo captura: texto = normalizar(texto_captura). Si texto_captura está vacío, se copia el texto actual (migración).
 * Modo directo: texto = normalizar(texto) solo con reglas estables.
 * Devuelve cuántos textos cambiaron.
 */
function aplicarReglasEscritura(ss) {
  const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  const colT = _col(t, 'texto');
  if (!t || colT < 0 || !t.filas.length) return 0;
  const reglas = _cargarReglas(ss);
  const colC = _col(t, 'texto_captura');
  const modoCaptura = colC > -1;

  const nuevosT = t.filas.map(f => [f[colT]]);
  const nuevosC = modoCaptura ? t.filas.map(f => [f[colC]]) : null;
  const cambiosPorFila = [];
  let n = 0, capturasNuevas = 0;
  let ctx = null;
  try { ctx = _contextoPermisos(ss); } catch (e) { ctx = null; }

  t.filas.forEach((f, i) => {
    if (ctx && _s(_get(t, f, 'id_fuente')) && _fuenteFiel(ctx, _get(t, f, 'id_fuente'))) return; // fuente que pide no alterar su contenido
    const actual = f[colT];
    let base = modoCaptura ? f[colC] : actual;
    const cambios = {};
    if (modoCaptura && (base === '' || base == null) && typeof actual === 'string' && actual) {
      base = actual; cambios.texto_captura = actual; nuevosC[i][0] = actual; capturasNuevas++;
    }
    if (typeof base === 'string' && base) {
      const nuevo = _normalizar(base, _get(t, f, 'id_idioma'), reglas, _get(t, f, 'creado_por'), !modoCaptura);
      if (nuevo !== actual) { cambios.texto = nuevo; nuevosT[i][0] = nuevo; n++; }
    }
    if (Object.keys(cambios).length) cambiosPorFila.push([i, cambios]);
  });

  // Pocos cambios: celda por celda (no pisa ediciones simultáneas en otras filas).
  // Muchos cambios (p. ej. la primera migración): una sola escritura por columna.
  if (cambiosPorFila.length > 50) {
    _formatoTexto(t.sh, 2, colT + 1, nuevosT.length);
    t.sh.getRange(2, colT + 1, nuevosT.length, 1).setValues(nuevosT);
    if (capturasNuevas) {
      _formatoTexto(t.sh, 2, colC + 1, nuevosC.length);
      t.sh.getRange(2, colC + 1, nuevosC.length, 1).setValues(nuevosC);
    }
    t.filas.forEach((f, i) => { f[colT] = nuevosT[i][0]; if (modoCaptura) f[colC] = nuevosC[i][0]; });
  } else {
    cambiosPorFila.forEach(([i, c]) => _actualizar(t, i, c));
  }
  return n;
}

/** Formato de texto plano: evita que "+ki", "-ta" o "=..." se interpreten como fórmulas. */
function _formatoTexto(sh, fila, col, filas) {
  if (filas > 0) sh.getRange(fila, col, filas, 1).setNumberFormat('@');
}

/** Agrega al final de la fila de encabezados las columnas que falten. Devuelve las que agregó. */
function _asegurarColumnas(sh, columnas) {
  const ultima = Math.max(sh.getLastColumn(), 1);
  const actuales = sh.getRange(1, 1, 1, ultima).getValues()[0].map(h => _k(h));
  const faltan = columnas.filter(c => !actuales.includes(_k(c)));
  if (!faltan.length) return [];
  const libre = actuales.every(h => !h) ? 1 : ultima + 1;
  const necesarias = libre + faltan.length - 1;
  if (necesarias > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), necesarias - sh.getMaxColumns());
  sh.getRange(1, libre, 1, faltan.length).setValues([faltan]);
  return faltan;
}

function _hojaReglas(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEETS.REGLAS);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.REGLAS);
    sh.getRange(1, 1, 1, COLS_REGLAS.length).setValues([COLS_REGLAS]);
    sh.getRange(1, 1, 1, COLS_REGLAS.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('white');
    sh.setFrozenRows(1);
  } else {
    _asegurarColumnas(sh, COLS_REGLAS);
  }
  // "buscar" y "reemplazar" siempre como texto plano ("+" o "=" no deben volverse fórmulas)
  const t = _tabla(ss, CONFIG.SHEETS.REGLAS);
  ['buscar', 'reemplazar'].forEach(c => _formatoTexto(sh, 2, _col(t, c) + 1, sh.getMaxRows() - 1));
  return t;
}

// ── Diálogo "Buscar y reemplazar" ────────────────────────────────────────────

function uiReemplazos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _hojaReglas(ss);
  const opcionesIdioma = [`<option value="*">${txh('atajos.todas')}</option>`].concat(
    readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma)).map(i =>
      `<option value="${_esc(i.id_idioma)}">${_esc(i.nombre_propio || i.nombre_completo || i.nombre || i.id_idioma)}${i.codigo_iso ? ' (' + _esc(i.codigo_iso) + ')' : ''}</option>`)
  ).join('');

  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}
    .caja{padding:14px}
    label{display:block;font-weight:bold;margin-top:10px}
    input[type=text],select{width:100%;box-sizing:border-box;padding:9px;font-size:16px;border:1px solid #cbd5e1;border-radius:6px}
    .fila{display:flex;gap:10px}.fila>div{flex:1}
    .op{font-weight:normal;display:block;margin:4px 0}
    button{padding:11px 14px;border:none;border-radius:6px;font-weight:bold;font-size:14px;cursor:pointer}
    .prim{background:#1e3a5f;color:#fff}.sec{background:#e2e8f0;color:#1f2937}
    #res{margin-top:10px;font-size:13px}
    .ej{font-family:monospace;background:#f8fafc;padding:3px 6px;border-radius:4px;display:block;margin:2px 0}
    .aviso{background:#fef3c7;color:#92400e;padding:8px;border-radius:6px;margin-top:6px}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:6px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:6px}
    table{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}
    td,th{border-bottom:1px solid #e2e8f0;padding:5px;text-align:left}
    .apagada{color:#94a3b8;text-decoration:line-through}
    small{color:#64748b;font-weight:normal}
  </style>
  ${_tScript()}
  <div class="caja">
    <p style="margin-top:0">${txh('atajos.intro')}</p>
    <div class="fila">
      <div><label for="buscar">${txh('atajos.buscar')}</label><input type="text" id="buscar" placeholder="+" autocomplete="off"></div>
      <div><label for="reemplazar">${txh('atajos.reemplazar')}</label><input type="text" id="reemplazar" placeholder="ɨ" autocomplete="off"></div>
    </div>
    <label for="idioma">${txh('atajos.lengua')}</label>
    <select id="idioma">${opcionesIdioma}</select>
    <label>${txh('atajos.paraQuien')}</label>
    <span class="op"><input type="radio" name="alcance" value="todos" id="a_todos" checked> <label for="a_todos" style="display:inline;font-weight:normal">${txh('atajos.todos')}</label></span>
    <span class="op"><input type="radio" name="alcance" value="mio" id="a_mio"> <label for="a_mio" style="display:inline;font-weight:normal">${txh('atajos.mio')} <small id="yo"></small></label></span>
    <span class="op"><input type="checkbox" id="mayus" checked> <label for="mayus" style="display:inline;font-weight:normal">${txh('atajos.mayus')} <small>${txh('atajos.mayus.d')}</small></label></span>
    <span class="op"><input type="checkbox" id="regex"> <label for="regex" style="display:inline;font-weight:normal">${txh('atajos.regex')}</label></span>
    <div style="margin-top:12px;display:flex;gap:8px">
      <button class="sec" id="bProbar" onclick="probar()">${txh('atajos.probar')}</button>
      <button class="prim" id="bGuardar" onclick="guardar()">${txh('atajos.guardar')}</button>
    </div>
    <div id="res" role="status" aria-live="polite"></div>
    <h4 style="margin:16px 0 0">${txh('atajos.guardados')}</h4>
    <div id="lista">${txh('comun.cargando')}</div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    const vis = s => esc(s).replace(/ /g, '␣') || '<i>' + th('atajos.vacio') + '</i>';
    function regla(){
      return { buscar: $('buscar').value, reemplazar: $('reemplazar').value, idioma: $('idioma').value,
               alcance: document.querySelector('input[name=alcance]:checked').value,
               mayus: $('mayus').checked, tipo: $('regex').checked ? 'regex' : 'texto' };
    }
    function ocupado(si){ $('bProbar').disabled = si; $('bGuardar').disabled = si; }
    function fallo(e){ ocupado(false); $('res').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; }
    function avisos(a){ return (a || []).map(x => '<div class="aviso">⚠️ ' + esc(x) + '</div>').join(''); }
    function probar(){
      ocupado(true); $('res').textContent = tt('atajos.revisando');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(r => {
        ocupado(false);
        const ej = r.ejemplos.map(e => '<span class="ej">' + vis(e.antes) + ' → ' + vis(e.despues) + '</span>').join('');
        $('res').innerHTML = '<div class="ok">' + th('atajos.cambiaria', {n: r.afectados}) + '</div>' + ej + avisos(r.avisos);
      }).reemplazoProbar(regla());
    }
    function guardar(){
      ocupado(true); $('res').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(r => {
        ocupado(false);
        $('res').innerHTML = '<div class="ok">✅ ' + esc(r.mensaje) + '</div>' + avisos(r.avisos);
        $('buscar').value = ''; $('reemplazar').value = '';
        pintar(r.lista);
      }).reemplazoGuardar(regla());
    }
    function cambiar(fila, activa){
      $('lista').textContent = tt('atajos.aplicando');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(pintar).reemplazoCambiarEstado(fila, activa);
    }
    function pintar(d){
      $('yo').textContent = d.yo ? '(' + d.yo + ')' : tt('atajos.sinCorreo');
      if (!d.yo) $('a_mio').disabled = true;
      if (!d.reglas.length) { $('lista').innerHTML = '<p><small>' + th('atajos.ninguno') + '</small></p>'; return; }
      $('lista').innerHTML = '<table><tr><th>' + th('atajos.col.escribe') + '</th><th>' + th('atajos.col.queda') + '</th><th>' + th('atajos.col.lengua') + '</th><th>' + th('atajos.col.para') + '</th><th></th></tr>' +
        d.reglas.map(r => '<tr class="' + (r.activa ? '' : 'apagada') + '"><td>' + vis(r.buscar) + (r.tipo === 'regex' ? ' <small>' + th('atajos.patron') + '</small>' : '') +
          '</td><td>' + vis(r.reemplazar) + '</td><td>' + esc(r.idioma) + '</td><td>' + esc(r.para) +
          '</td><td><button class="sec" style="padding:5px 8px" onclick="cambiar(' + r.fila + ',' + !r.activa + ')">' + th(r.activa ? 'atajos.apagar' : 'atajos.encender') + '</button></td></tr>').join('') +
        '</table>';
    }
    google.script.run.withFailureHandler(fallo).withSuccessHandler(pintar).reemplazosListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(660), tx('atajos.titulo'));
}

function _reglaDesdeDialogo(d, yo) {
  if (!d || !d.idioma) throw new Error(tx('atajos.error.lengua'));
  if (d.alcance === 'mio' && !yo) throw new Error(tx('atajos.error.correo'));
  return _compilarRegla({
    buscar: String(d.buscar || ''),
    reemplazar: String(d.reemplazar || ''),
    tipo: d.tipo === 'regex' ? 'regex' : 'texto',
    mayus: d.mayus !== false,
    aplicaA: d.alcance === 'mio' ? yo : 'todos',
    orden: 1e9
  });
}

/** Muestra qué cambiaría la regla sin guardar nada. */
function reemplazoProbar(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yo = _miCorreo();
  const r = _reglaDesdeDialogo(d, yo);
  const isoAId = _mapaIsoAId(ss);
  const clave = _claveIdioma(d.idioma, isoAId);
  const reglas = _cargarReglas(ss);
  const conNueva = new Map();
  reglas.forEach((v, k) => conNueva.set(k, v.slice()));
  if (!conNueva.has(clave)) conNueva.set(clave, []);
  conNueva.get(clave).push(r);
  _ordenarReglas(conNueva.get(clave));

  const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  _requerir(t, ['id_idioma', 'texto'], CONFIG.SHEETS.TRADUCCIONES);
  const modoCaptura = _col(t, 'texto_captura') > -1;
  const hayAutor = _col(t, 'creado_por') > -1;
  let afectados = 0;
  const ejemplos = [], muestras = [];
  t.filas.forEach(f => {
    const idL = _k(_get(t, f, 'id_idioma'));
    if (clave !== '*' && idL !== clave) return;
    const cap = modoCaptura ? _get(t, f, 'texto_captura') : '';
    const base = (typeof cap === 'string' && cap) ? cap : _get(t, f, 'texto');
    if (typeof base !== 'string' || !base) return;
    const autor = _get(t, f, 'creado_por');
    const antes = _normalizar(base, idL, reglas, autor, !modoCaptura);
    const despues = _normalizar(base, idL, conNueva, autor, false);
    if (antes === despues) return;
    afectados++;
    if (ejemplos.length < 8) ejemplos.push({ antes, despues });
    if (muestras.length < 300) muestras.push(base);
  });

  r.estable = _reglaEstable(r, muestras);
  const avisos = [];
  if (!r.estable) {
    avisos.push(tx(modoCaptura ? 'atajos.aviso.inestableCaptura' : 'atajos.aviso.inestable'));
  }
  if (r.aplicaA !== 'todos' && !hayAutor) avisos.push(tx('atajos.aviso.sinAutor'));
  const nada = !afectados ? tx('atajos.aviso.nada') : '';
  if (nada) avisos.push(nada);
  return { afectados, ejemplos, estable: r.estable, modoCaptura, avisos, nada };
}

function reemplazoGuardar(d) {
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const yo = _miCorreo();
    const prueba = reemplazoProbar(d);
    if (!prueba.estable && !prueba.modoCaptura) throw new Error(prueba.avisos[0]);
    const r = _reglaDesdeDialogo(d, yo);
    const tR = _hojaReglas(ss);
    const orden = tR.filas.reduce((max, f) => Math.max(max, Number(_get(tR, f, 'orden')) || 0), 0) + 1;
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    _agregar(tR, [{
      id_idioma: d.idioma, buscar: r.buscar, reemplazar: r.reemplazar, tipo: r.tipo, orden,
      activa: true, aplica_a: r.aplicaA, coincidir_mayus: r.mayus,
      nota: tx('atajos.nota', { quien: yo || '—', fecha })
    }]);
    const n = aplicarReglasEscritura(ss);
    publicarCatalogo(ss);
    return { mensaje: tx('atajos.guardado', { n }), avisos: prueba.avisos.filter(a => a !== prueba.nada), lista: reemplazosListar() };
  } finally { lock.releaseLock(); }
}

function reemplazoCambiarEstado(fila, activa) {
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const t = _tabla(ss, CONFIG.SHEETS.REGLAS);
    if (!t || !t.filas[fila]) throw new Error(tx('atajos.error.noExiste'));
    if (activa && _col(_tabla(ss, CONFIG.SHEETS.TRADUCCIONES), 'texto_captura') < 0) {
      const f = t.filas[fila];
      const r = _compilarRegla({ buscar: String(_get(t, f, 'buscar')), reemplazar: String(_get(t, f, 'reemplazar')),
        tipo: _k(_get(t, f, 'tipo')) === 'regex' ? 'regex' : 'texto', mayus: _s(_get(t, f, 'coincidir_mayus')) === '' ? true : _verdadero(_get(t, f, 'coincidir_mayus')) });
      if (!_reglaEstable(r)) throw new Error(tx('atajos.error.inestable'));
    }
    _actualizar(t, fila, { activa: !!activa });
    aplicarReglasEscritura(ss);
    publicarCatalogo(ss);
    return reemplazosListar();
  } finally { lock.releaseLock(); }
}

function reemplazosListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yo = _miCorreo();
  const t = _tabla(ss, CONFIG.SHEETS.REGLAS);
  if (!t) return { yo, reglas: [] };
  const nombres = new Map();
  readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => {
    const nom = _s(i.nombre_completo || i.nombre || i.id_idioma);
    if (_s(i.id_idioma)) nombres.set(_k(i.id_idioma), nom);
    if (_s(i.codigo_iso)) nombres.set(_k(i.codigo_iso), nom);
  });
  const reglas = [];
  t.filas.forEach((f, i) => {
    const buscar = String(_get(t, f, 'buscar'));
    if (!buscar) return;
    const idL = _s(_get(t, f, 'id_idioma'));
    const aplica = _k(_get(t, f, 'aplica_a'));
    reglas.push({
      fila: i, buscar, reemplazar: String(_get(t, f, 'reemplazar')),
      tipo: _k(_get(t, f, 'tipo')) === 'regex' ? 'regex' : 'texto',
      idioma: idL === '*' ? tx('atajos.todasCorto') : (nombres.get(_k(idL)) || idL || tx('atajos.sinLengua')),
      para: !aplica || aplica === 'todos' ? tx('atajos.paraTodos') : (aplica === yo ? tx('atajos.soloYo') : aplica),
      activa: _verdadero(_get(t, f, 'activa'))
    });
  });
  return { yo, reglas };
}

/** Agrega texto_captura y creado_por a TRADUCCIONES y migra los textos actuales. */
function servicioActivarModoCaptura() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ui.alert(tx('captura.titulo'), tx('captura.texto'), ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const lock = _lock();
  try {
    const sh = ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES);
    if (!sh) throw new Error(tx('comun.faltaHoja', { hoja: CONFIG.SHEETS.TRADUCCIONES }));
    const agregadas = _asegurarColumnas(sh, ['texto_captura', 'creado_por']);
    const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
    _formatoTexto(sh, 2, _col(t, 'texto_captura') + 1, sh.getMaxRows() - 1);
    const n = aplicarReglasEscritura(ss);
    publicarCatalogo(ss);
    ui.alert(tx('captura.listo', { columnas: agregadas.length ? agregadas.join(', ') : tx('captura.yaEstaban'), n }));
  } finally { lock.releaseLock(); }
}

function servicioNormalizacionTexto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const n = aplicarReglasEscritura(ss);
    publicarCatalogo(ss);
    SpreadsheetApp.getUi().alert(tx('atajos.aplicados', { n }));
  } finally { lock.releaseLock(); }
}

/** Activador horario: normaliza y publica si hubo cambios. */
function tareaAutomatica() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    aplicarReglasEscritura(ss);
    publicarCatalogo(ss);
  } finally { lock.releaseLock(); }
  try { _revisarMantenimientoProgramado(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { console.warn('Mantenimiento programado: ' + e.message); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. MESA DE TRABAJO (AUDITORÍA DINÁMICA)
// ─────────────────────────────────────────────────────────────────────────────

function uiGenerarReporteAvanzado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const html = `
    ${_tScript()}
    <div style="font-family:sans-serif;padding:15px;color:#333;">
      <h3 style="color:#2563eb;margin-top:0;">${txh('mesa.encabezado')}</h3>
      <p style="font-size:13px;color:#475569;">${txh('mesa.intro')}</p>
      <p style="font-size:12px;color:#64748b;">${txh('mesa.lenguas')}</p>
      <div style="max-height:150px;overflow-y:auto;background:#f8fafc;padding:10px;border-radius:6px;border:1px solid #e2e8f0;">
        ${_filasIdiomasHtml(ss, 'idiomas')}
      </div>
      <div style="background:#fff7ed;padding:10px;border-radius:6px;border:1px solid #fdba74;margin-top:15px;font-size:12px;">
        <b>${txh('mesa.buscar')}</b><br>
        <label><input type="checkbox" id="chk_trans" checked> ${txh('mesa.sinTraducir')}</label><br>
        <label><input type="checkbox" id="chk_audio" checked> ${txh('mesa.sinAudio')}</label>
      </div>
      <button id="btn" onclick="run()" style="width:100%;background:#2563eb;color:white;border:none;padding:12px;border-radius:6px;margin-top:15px;cursor:pointer;font-weight:bold;">${txh('mesa.armar')}</button>
      <div id="res" style="margin-top:10px;font-size:12px;"></div>
      <script>
        function run(){
          const btn = document.getElementById('btn');
          const idiomas = Array.from(document.querySelectorAll('input[name="idiomas"]:checked')).map(c => c.value);
          const opts = { trans: document.getElementById('chk_trans').checked, audio: document.getElementById('chk_audio').checked };
          btn.disabled = true; btn.innerText = tt('mesa.buscando');
          google.script.run
            .withSuccessHandler(() => google.script.host.close())
            .withFailureHandler(e => { btn.disabled = false; btn.innerText = tt('mesa.armar'); document.getElementById('res').innerText = tt('comun.error', {m: e.message}); })
            .procesarReporteBack(idiomas, opts);
        }
      </script>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(540), tx('mesa.titulo'));
}

/** Columnas de la lista de lo que falta (hoja MESA_DE_TRABAJO). Se aceptan también los nombres de la versión anterior. */
// Los títulos se escriben en el idioma de menús de quien arma la lista; al leer se aceptan todos.
const MESA_CLAVES = { id: 'mesa.col.codigo', ref: 'mesa.col.referencia', actual: 'mesa.col.actual', lengua: 'mesa.col.codigoLengua',
  nueva: 'mesa.col.nueva', nota: 'mesa.col.nota', estado: 'mesa.col.estado' };
const MESA_ANTERIOR = { id: 'id_ref', lengua: 'id_lang', nueva: 'nueva traducción', nota: 'nota', estado: 'estado' };
function _colMesa(t, clave) {
  for (const forma of _todasLasFormas(MESA_CLAVES[clave])) { const a = _col(t, forma); if (a > -1) return a; }
  return _col(t, MESA_ANTERIOR[clave] || '');
}

function procesarReporteBack(seleccionados, opts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getAppConfig(ss);
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  _requerir(tPal, ['id_palabra'], CONFIG.SHEETS.PALABRAS);
  _requerir(tTr, ['id_traduccion', 'id_palabra', 'id_idioma', 'texto'], CONFIG.SHEETS.TRADUCCIONES);

  const idBase = _k(cfg.idiomaPrincipal);
  const base = _mapaTextoBase(tTr, idBase);
  const idx = _indiceTrads(tTr);
  const langs = seleccionados
    .map(s => { const [id, iso, nom] = s.split('|'); return { id, iso, nom }; })
    .filter(l => _k(l.id) !== idBase);

  const filas = [], links = [], editables = [];
  const TAREA_TRAD = tx('mesa.tarea.traduccion'), TAREA_AUDIO = tx('mesa.tarea.audio'), PENDIENTE = tx('mesa.valor.pendiente');
  tPal.filas.forEach(p => {
    const idPal = _s(_get(tPal, p, 'id_palabra'));
    if (!idPal) return;
    const ref = base.get(_k(idPal)) || tx('mesa.sinReferencia');
    langs.forEach(l => {
      const i = idx.get(_k(idPal) + '||' + _k(l.id));
      const tr = i === undefined ? null : tTr.filas[i];
      let tarea = '';
      if (!tr && opts.trans) tarea = TAREA_TRAD;
      else if (tr && opts.audio && _s(_get(tTr, tr, 'audio')).length < 2) tarea = TAREA_AUDIO;
      if (!tarea) return;
      const url = tr
        ? _linkAppSheet(cfg, CONFIG.VISTAS.TRAD_FORM, { row: _s(_get(tTr, tr, 'id_traduccion')) })
        : _linkAppSheet(cfg, CONFIG.VISTAS.TRAD_FORM, { defaults: { id_palabra: idPal, id_idioma: l.id } });
      filas.push([idPal, ref, tr ? _s(_get(tTr, tr, 'texto')) : '—', l.id, l.nom, tarea, '', '', PENDIENTE, '']);
      links.push(url);
      if (tarea === TAREA_TRAD) editables.push(filas.length + 1);
    });
  });

  const sheet = ss.getSheetByName(CONFIG.SHEETS.REPORT) || ss.insertSheet(CONFIG.SHEETS.REPORT);
  _limpiarHoja(sheet);
  sheet.setTabColor('#f39c12');

  if (!filas.length) {
    sheet.getRange(1, 1).setValue(tx('mesa.nadaFalta'));
    sheet.activate(); return;
  }

  // Los títulos los lee "Guardar lo que se llenó en la lista" (ver MESA_CLAVES)
  const headers = [['mesa.col.codigo', 'mesa.col.referencia', 'mesa.col.actual', 'mesa.col.codigoLengua', 'mesa.col.lengua', 'mesa.col.falta',
    'mesa.col.nueva', 'mesa.col.nota', 'mesa.col.estado', 'mesa.col.abrir'].map(k => tx(k))];
  sheet.getRange(1, 1, 1, 10).setValues(headers).setBackground('#2c3e50').setFontColor('white').setFontWeight('bold');
  sheet.getRange(2, 1, filas.length, 10).setValues(filas);
  sheet.getRange(2, 10, links.length, 1).setRichTextValues(links.map(u => [_rt(u ? tx('mesa.abrirApp') : tx('mesa.sinApp'), u)]));

  if (editables.length) {
    sheet.getRangeList(editables.map(r => `G${r}:H${r}`))
      .setBackground('#fffec8').setBorder(true, true, true, true, true, true);
  }
  const rule = SpreadsheetApp.newDataValidation().requireValueInList([PENDIENTE, tx('mesa.valor.guardado'), tx('mesa.valor.omitir')]).build();
  sheet.getRange(2, 9, filas.length, 1).setDataValidation(rule);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 10, 130);
  sheet.hideColumns(1); sheet.hideColumns(4);   // códigos internos: no se tocan
  sheet.activate();
}

/** Importa a TRADUCCIONES lo escrito en "NUEVA TRADUCCIÓN" de la Mesa (filas en estado PENDIENTE). */
function servicioImportarDesdeReporte() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tM0 = _tabla(ss, CONFIG.SHEETS.REPORT);
  if (!tM0 || _colMesa(tM0, 'id') < 0) return ui.alert(tx('mesa.primeroLista'));
  const resp = ui.alert(tx('mesa.preguntaTitulo'), tx('mesa.preguntaTexto'), ui.ButtonSet.YES_NO_CANCEL);
  if (resp === ui.Button.CANCEL || resp === ui.Button.CLOSE) return;
  const estado = _etiquetaEstado(resp === ui.Button.YES ? 'publicado' : 'revisión');
  const lock = _lock();
  try {
    const tM = _tabla(ss, CONFIG.SHEETS.REPORT);
    const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
    _requerir(tTr, ['id_traduccion', 'id_palabra', 'id_idioma', 'texto'], CONFIG.SHEETS.TRADUCCIONES);
    const idx = _indiceTrads(tTr);
    const reglas = _cargarReglas(ss);
    const modoCaptura = _col(tTr, 'texto_captura') > -1;
    const yo = _miCorreo();
    const nuevas = [], enEstaCarga = new Set();
    let insertadas = 0, actualizadas = 0;
    const colEstado = _colMesa(tM, 'estado');
    const pendientes = _todasLasFormas('mesa.valor.pendiente').concat(['pendiente']);
    const GUARDADO = tx('mesa.valor.guardado');

    tM.filas.forEach((f, i) => {
      if (!pendientes.includes(_k(f[colEstado]))) return;
      const val = c => { const i = _colMesa(tM, c); return i > -1 ? _s(f[i]) : ''; };
      const texto = val('nueva');
      if (!texto) return;
      const idPal = val('id'), idLang = val('lengua');
      if (!idPal || !idLang) return;
      const nota = val('nota');
      const clave = _k(idPal) + '||' + _k(idLang);

      if (idx.has(clave)) {
        const fi = idx.get(clave), fila = tTr.filas[fi], cambios = {};
        const norm = _normalizar(texto, idLang, reglas, _get(tTr, fila, 'creado_por') || yo, !modoCaptura);
        if (_s(_get(tTr, fila, 'texto')) !== norm) { cambios.texto = norm; cambios.texto_captura = texto; }
        if (nota) cambios.nota_variante = nota;
        if (Object.keys(cambios).length) { _actualizar(tTr, fi, cambios); actualizadas++; }
      } else if (!enEstaCarga.has(clave)) {
        const norm = _normalizar(texto, idLang, reglas, yo, !modoCaptura);
        nuevas.push({ id_traduccion: Utilities.getUuid(), id_palabra: idPal, id_idioma: idLang, texto: norm, texto_captura: texto, nota_variante: nota, creado_por: yo, estado });
        enEstaCarga.add(clave);
        insertadas++;
      }
      tM.sh.getRange(i + 2, colEstado + 1).setValue(GUARDADO);
    });

    _agregar(tTr, nuevas);
    if (insertadas || actualizadas) publicarCatalogo(ss);
    ui.alert(tx('mesa.guardadoListo', { estado, nuevas: insertadas, corregidas: actualizadas }));
  } finally { lock.releaseLock(); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. DASHBOARD DE IMÁGENES (ADMIN)
// ─────────────────────────────────────────────────────────────────────────────

function uiGenerarDashboardAdmin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getAppConfig(ss);
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  _requerir(tPal, ['id_palabra'], CONFIG.SHEETS.PALABRAS);
  const base = _mapaTextoBase(_tabla(ss, CONFIG.SHEETS.TRADUCCIONES), cfg.idiomaPrincipal);
  const nomCat = new Map();
  const tCat = _tabla(ss, CONFIG.SHEETS.CATEGORIAS);
  if (tCat) tCat.filas.forEach(f => nomCat.set(_k(_get(tCat, f, 'id_categoria')), _s(_get(tCat, f, 'nombre_categoria'))));

  const filas = [], links = [];
  tPal.filas.forEach(p => {
    const id = _s(_get(tPal, p, 'id_palabra'));
    if (!id) return;
    const img = _s(_get(tPal, p, 'imagen_referencia'));
    filas.push([id, base.get(_k(id)) || tx('mesa.sinReferencia'), nomCat.get(_k(_get(tPal, p, 'id_categoria'))) || '', tx(img.length > 5 ? 'imagenes.si' : 'imagenes.falta'), '']);
    links.push(_linkAppSheet(cfg, CONFIG.VISTAS.PALABRA_DETALLE, { row: id }));
  });

  const sAudit = ss.getSheetByName(CONFIG.SHEETS.AUDIT) || ss.insertSheet(CONFIG.SHEETS.AUDIT);
  _limpiarHoja(sAudit);
  sAudit.setTabColor('#e11d48');
  sAudit.getRange(1, 1, 1, 5).setValues([['imagenes.col.codigo', 'imagenes.col.palabra', 'imagenes.col.categoria', 'imagenes.col.tiene', 'imagenes.col.agregar'].map(k => tx(k))])
    .setBackground('#1e3a5f').setFontColor('white').setFontWeight('bold');
  if (filas.length) {
    sAudit.getRange(2, 1, filas.length, 5).setValues(filas);
    sAudit.getRange(2, 5, links.length, 1).setRichTextValues(links.map(u => [_rt(u ? tx('mesa.abrirApp') : tx('mesa.sinApp'), u)]));
  }
  sAudit.setFrozenRows(1);
  sAudit.hideColumns(1);
  sAudit.activate();
}

// ─────────────────────────────────────────────────────────────────────────────
// 6. PLANTILLAS PARA COLABORADORES E IMPORTACIÓN (ETL)
// ─────────────────────────────────────────────────────────────────────────────

function uiConfigurarPlantillaInteligente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nombres = JSON.stringify({ NUEVA: tx('plantilla.archivo.nuevas', { fecha: hoy }), FALTANTES: tx('plantilla.archivo.faltantes', { fecha: hoy }) }).replace(/</g, '\\u003c');
  const html = `
    <style>
      body{margin:0}.caja{font-family:sans-serif;padding:15px;color:#333;font-size:13px}
      .gris{background:#f1f5f9;padding:10px;border-radius:6px;margin-bottom:12px}
      input[type=text]{padding:7px;border:1px solid #cbd5e1;border-radius:5px;box-sizing:border-box;font-size:13px}
      .grande{display:block;width:100%;border:none;padding:12px;border-radius:6px;margin-top:8px;cursor:pointer;font-weight:bold;font-size:14px;text-align:center;text-decoration:none;box-sizing:border-box}
      .azul{background:#2563eb;color:#fff}.oscuro{background:#1e3a5f;color:#fff}.verde{background:#16a34a;color:#fff}.claro{background:#e2e8f0;color:#1f2937}
      .fila{display:flex;gap:6px}.fila>*{flex:1}
      small{color:#64748b}.ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:6px}
      .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:6px}
    </style>
    ${_tScript()}
    <div class="caja">
     <div id="paso1">
      <h3 style="margin-top:0; color:#1e3a5f;">${txh('plantilla.encabezado')}</h3>
      <p style="color:#475569; margin-top:0;">${txh('plantilla.intro')}</p>
      <div class="gris">
        <b>${txh('plantilla.paraQue')}</b><br>
        <label><input type="radio" name="modo" value="NUEVA" checked onchange="sugerir()"> ${txh('plantilla.nuevas')}</label><br>
        <label><input type="radio" name="modo" value="FALTANTES" onchange="sugerir()"> ${txh('plantilla.faltantes')}</label>
      </div>
      <div style="margin-bottom:4px;"><b>${txh('plantilla.lenguas')}</b></div>
      <div style="max-height:110px; overflow-y:auto; border:1px solid #ddd; padding:8px; border-radius:4px; background:white;">
        ${_filasIdiomasHtml(ss, 'idiomas')}
      </div>
      <div style="margin-top:10px; border-top:1px solid #eee; padding-top:8px;">
        <b>${txh('plantilla.otraLengua')}</b><br>
        <input type="text" id="n_nom" placeholder="${txh('plantilla.otraNombre')}" style="width:55%;">
        <input type="text" id="n_iso" placeholder="${txh('plantilla.otraCodigo')}" title="${txh('plantilla.otraCodigo.d')}" style="width:30%;">
      </div>
      <div style="margin-top:10px;">
        <label for="nombre"><b>${txh('plantilla.nombre')}</b></label><br>
        <input type="text" id="nombre" maxlength="100" style="width:100%;" oninput="this.dataset.tocado=1">
        <small>${txh('plantilla.nombre.d')}</small>
      </div>
      <button id="btn" class="grande oscuro" onclick="run()">${txh('plantilla.crear')}</button>
     </div>
     <div id="res" role="status" aria-live="polite" style="margin-top:10px;"></div>
     <div id="paso2" style="display:none">
      <h3 style="margin-top:0; color:#166534;">${txh('plantilla.lista')}</h3>
      <a id="abrir" class="grande azul" target="_blank"></a>
      <p><small>${txh('plantilla.carpeta')}</small></p>
      <div class="gris">
        <b>${txh('plantilla.quien')}</b><br>
        <label><input type="radio" name="quien" value="correos" checked onchange="verQuien()"> ${txh('plantilla.quien.correos')}</label><br>
        <label><input type="radio" name="quien" value="enlace" onchange="verQuien()"> ${txh('plantilla.quien.enlace')}</label>
        <div id="qCorreos" style="margin-top:6px">
          <input type="text" id="correos" placeholder="persona@ejemplo.com" style="width:100%;" aria-label="${txh('plantilla.correos')}">
          <small>${txh('plantilla.correos.d')}</small>
          <button class="grande claro" onclick="permiso()">${txh('plantilla.darPermiso')}</button>
        </div>
        <div id="qEnlace" style="display:none;margin-top:6px">
          <small>${txh('plantilla.quien.enlace.d')}</small>
          <button class="grande claro" onclick="permiso()">${txh('plantilla.abrirEnlace')}</button>
        </div>
        <div id="mPermiso"></div>
      </div>
      <b>${txh('plantilla.enviar')}</b>
      <div class="fila">
        <a id="wa" class="grande verde" target="_blank">${txh('plantilla.whatsapp')}</a>
        <a id="mail" class="grande claro" target="_blank">${txh('plantilla.correo')}</a>
      </div>
      <button class="grande claro" onclick="copiar()">${txh('plantilla.copiar')}</button>
      <input type="text" id="enlace" readonly style="width:100%;margin-top:6px;font-size:12px" aria-label="${txh('plantilla.enlace')}">
      <div id="mCopia" role="status" aria-live="polite"></div>
     </div>
      <script>
        const $ = id => document.getElementById(id);
        const NOMBRES = ${nombres};
        let hoja = null;
        function sugerir(){ const n = $('nombre'); if (!n.dataset.tocado) n.value = NOMBRES[document.querySelector('input[name="modo"]:checked').value]; }
        sugerir();
        function run(){
          const btn = $('btn'), res = $('res');
          const idiomas = Array.from(document.querySelectorAll('input[name="idiomas"]:checked')).map(c => c.value);
          const modo = document.querySelector('input[name="modo"]:checked').value;
          btn.disabled = true; res.innerText = tt('plantilla.preparando');
          google.script.run
            .withSuccessHandler(r => {
              hoja = r; res.innerHTML = '';
              $('paso1').style.display = 'none'; $('paso2').style.display = 'block';
              $('abrir').href = r.url; $('abrir').textContent = r.filas ? tt('plantilla.abrirN', {n: r.filas}) : tt('plantilla.abrir');
              $('enlace').value = r.url;
              const msj = tt('plantilla.mensaje', {nombre: r.nombre, url: r.url});
              $('wa').href = 'https://wa.me/?text=' + encodeURIComponent(msj);
              $('mail').href = 'mailto:?subject=' + encodeURIComponent(r.nombre) + '&body=' + encodeURIComponent(msj);
            })
            .withFailureHandler(e => { btn.disabled = false; res.innerText = tt('comun.error', {m: e.message}); })
            .servicioGenerarPlantillaPro(idiomas, modo, $('n_nom').value, $('n_iso').value, '', $('nombre').value);
        }
        function verQuien(){ const e = document.querySelector('input[name="quien"]:checked').value === 'enlace'; $('qEnlace').style.display = e ? 'block' : 'none'; $('qCorreos').style.display = e ? 'none' : 'block'; $('mPermiso').innerHTML = ''; }
        function permiso(){
          const enlace = document.querySelector('input[name="quien"]:checked').value === 'enlace';
          $('mPermiso').textContent = tt('comun.guardando');
          google.script.run.withSuccessHandler(r => { $('mPermiso').innerHTML = '<div class="ok">' + (r.enlace ? th('plantilla.permisoEnlace') : th('plantilla.permisoCorreos', {n: r.n})) + '</div>'; })
            .withFailureHandler(e => { $('mPermiso').innerHTML = '<div class="error">' + _escT(tt('comun.error', {m: e.message})) + '</div>'; })
            .plantillaCompartir(hoja.id, enlace ? '' : $('correos').value, enlace);
        }
        function copiar(){
          const i = $('enlace'); i.focus(); i.select();
          let listo = false;
          try { listo = document.execCommand('copy'); } catch (e) {}
          if (!listo && navigator.clipboard) navigator.clipboard.writeText(i.value).then(() => $('mCopia').innerHTML = '<div class="ok">' + th('plantilla.copiado') + '</div>').catch(() => $('mCopia').innerHTML = '<small>' + th('plantilla.copiaManual') + '</small>');
          else $('mCopia').innerHTML = listo ? '<div class="ok">' + th('plantilla.copiado') + '</div>' : '<small>' + th('plantilla.copiaManual') + '</small>';
        }
      </script>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(660), tx('plantilla.titulo'));
}

/** Da permiso de escribir en una hoja de COLABORACIONES: a ciertos correos o a cualquiera con el enlace. */
function plantillaCompartir(id, correos, cualquiera) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let f;
  try { f = DriveApp.getFileById(_s(id)); } catch (e) { throw new Error(tx('plantilla.error.noEsta')); }
  // Solo hojas de la carpeta COLABORACIONES de este catálogo
  const carpeta = _subcarpeta(_carpetaDelLibro(ss), 'COLABORACIONES').getId();
  const padres = f.getParents(); let dentro = false;
  while (padres.hasNext()) if (padres.next().getId() === carpeta) dentro = true;
  if (!dentro) throw new Error(tx('plantilla.error.noEsta'));
  if (cualquiera) {
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT); }
    catch (e) { throw new Error(tx('plantilla.error.enlace')); }
    return { enlace: true };
  }
  const lista = _s(correos).split(/[,;\s]+/).map(c => c.toLowerCase()).filter(c => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c));
  if (!lista.length) throw new Error(tx('plantilla.error.correos'));
  f.addEditors(lista);
  return { n: lista.length };
}

function servicioGenerarPlantillaPro(seleccionados, modo, nNom, nIso, correos, nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getAppConfig(ss);
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  _requerir(tPal, ['id_palabra'], CONFIG.SHEETS.PALABRAS);
  _requerir(tTr, ['id_palabra', 'id_idioma', 'texto'], CONFIG.SHEETS.TRADUCCIONES);

  const base = _mapaTextoBase(tTr, cfg.idiomaPrincipal);
  const idx = _indiceTrads(tTr);
  const langs = seleccionados.map(s => { const [id, iso] = s.split('|'); return { id, iso: _s(iso).toUpperCase() }; });
  const isoNuevo = _s(nIso).toUpperCase().replace(/[^A-Z0-9_]/g, '');

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nombreHoja = _s(nombre).slice(0, 100) || tx(modo === 'FALTANTES' ? 'plantilla.archivo.faltantes' : 'plantilla.archivo.nuevas', { fecha });
  const nueva = _aColaboraciones(ss, SpreadsheetApp.create(nombreHoja));
  const h = nueva.getActiveSheet().setName('DATOS_A_IMPORTAR');
  // Instrucciones en su propia pestaña (los títulos de DATOS_A_IMPORTAR no se traducen: los lee el catálogo)
  const inst = nueva.insertSheet(tx('plantilla.hojaInstrucciones'), 0);
  inst.getRange(1, 1, 9, 1).setValues([1, 2, 3, 4, 5, 6, 7, 8, 9].map(i => [tx('plantilla.inst.' + i)]));
  inst.setColumnWidth(1, 700).getRange(1, 1).setFontWeight('bold').setFontSize(14);
  const headers = ['ID_PALABRA', 'CONCEPTO_BASE', 'CATEGORIA', 'IMG_STATUS'];
  langs.forEach(l => headers.push(`TRADUCCION_${l.iso}`, `AUDIO_${l.iso}`, `NOTA_${l.iso}`, `DESCRIPCION_${l.iso}`, `TRAD_CAT_${l.iso}`));
  if (_s(nNom) && isoNuevo) headers.push(`TRADUCCION_${isoNuevo}`, `AUDIO_${isoNuevo}`, `NOTA_${isoNuevo}`, `DESCRIPCION_${isoNuevo}`, `TRAD_CAT_${isoNuevo}`);
  h.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#1e3a5f').setFontColor('white').setFontWeight('bold');

  // Nombre de idiomas nuevos: la importación lo lee de esta hoja
  if (_s(nNom) && isoNuevo) {
    nueva.insertSheet('META').getRange(1, 1, 2, 2).setValues([['CODIGO_ISO', 'NOMBRE'], [isoNuevo, _s(nNom)]]);
  }

  let filas = [];
  if (modo === 'FALTANTES') {
    const dicAudio = _obtenerDiccionarioDrive(cfg.audios);
    tPal.filas.forEach(p => {
      const idPal = _s(_get(tPal, p, 'id_palabra'));
      if (!idPal) return;
      const textoBase = base.get(_k(idPal)) || '';
      const row = [idPal, textoBase || tx('plantilla.sinBase'), _s(_get(tPal, p, 'id_categoria')), _s(_get(tPal, p, 'imagen_referencia')) ? '✅' : '❌'];
      let pendiente = false;
      langs.forEach(l => {
        const i = idx.get(_k(idPal) + '||' + _k(l.id));
        const t = i === undefined ? null : tTr.filas[i];
        const sugerido = `${textoBase}_${l.iso}`.toLowerCase();
        const audio = t ? _s(_get(tTr, t, 'audio')) : '';
        const statusAud = audio ? tx('plantilla.audioOk') : (textoBase && dicAudio.has(sugerido) ? tx('plantilla.audioDrive') : '');
        row.push(t ? _s(_get(tTr, t, 'texto')) : '', statusAud, t ? _s(_get(tTr, t, 'nota_variante')) : '', t ? _s(_get(tTr, t, 'descripcion')) : '', '');
        if (!t || !audio) pendiente = true;
      });
      if (_s(nNom) && isoNuevo) row.push('', '', '', '', '');
      if (pendiente) filas.push(row);
    });
    if (filas.length) h.getRange(2, 1, filas.length, headers.length).setValues(filas);
  }

  h.setFrozenRows(1);
  h.setColumnWidths(1, headers.length, 150);

  const lista = _s(correos).split(/[,;\s]+/).filter(c => /.+@.+\..+/.test(c));
  if (lista.length) DriveApp.getFileById(nueva.getId()).addEditors(lista);
  return { url: nueva.getUrl(), id: nueva.getId(), nombre: nombreHoja, filas: filas.length };
}

function uiImportarDatos() {
  const html = `${_tScript()}<div style="font-family:sans-serif;padding:20px;text-align:center;">
    <h3 style="color:#1e3a5f;">${txh('traer.titulo')}</h3>
    <p style="font-size:13px;color:#475569;text-align:left;">${txh('traer.intro')}</p>
    <input type="text" id="url" placeholder="${txh('traer.enlace')}" style="width:100%;padding:8px;border-radius:4px;border:1px solid #ccc;">
    <label for="estado" style="display:block;text-align:left;margin-top:12px;font-size:13px;"><b>${txh('traer.comoQuedan')}</b></label>
    <select id="estado" style="width:100%;padding:8px;border-radius:4px;border:1px solid #ccc;">
      <option value="revisión" selected>${txh('traer.revision')}</option>
      <option value="borrador">${txh('traer.borrador')}</option>
      <option value="publicado">${txh('traer.publicado')}</option>
    </select>
    <div style="text-align:left;font-size:11px;color:#64748b;margin-top:4px;">${txh('traer.nota')}</div>
    <button type="button" onclick="verLenguas()" style="width:100%;margin-top:10px;padding:9px;background:#e2e8f0;border:none;border-radius:6px;font-weight:bold;cursor:pointer;">${txh('traer.verLenguas')}</button>
    <div id="lenguas" style="text-align:left;font-size:13px;margin-top:6px;"></div>
    <button id="btn" onclick="run()" style="width:100%;margin-top:15px;padding:12px;background:#16a34a;color:white;border:none;border-radius:6px;font-weight:bold;cursor:pointer;">${txh('traer.boton')}</button>
    <div id="res" style="margin-top:10px;font-size:12px;text-align:left;"></div>
    <script>
      function run(){
        const url = document.getElementById('url').value;
        if(!url) return alert(tt('traer.faltaEnlace'));
        const btn = document.getElementById('btn'), res = document.getElementById('res');
        btn.disabled = true; res.innerText = tt('traer.trayendo');
        google.script.run
          .withSuccessHandler(r => { res.innerHTML = r; btn.disabled = false; })
          .withFailureHandler(e => { res.innerText = tt('comun.error', {m: e.message}); btn.disabled = false; })
          .servicioProcesarImportacionMasiva(url, document.getElementById('estado').value,
            Array.from(document.querySelectorAll('#lenguas input:checked')).map(c => c.value));
      }
      function verLenguas(){
        const url = document.getElementById('url').value, box = document.getElementById('lenguas');
        if (!url) { box.innerText = tt('traer.faltaEnlace'); return; }
        box.innerText = tt('comun.cargando');
        google.script.run.withFailureHandler(e => { box.innerText = tt('comun.error', {m: e.message}); }).withSuccessHandler(l => {
          box.innerHTML = '<b>' + th('traer.cualesLenguas') + '</b><br>' + (l.length ? l.map(x => '<label style="display:block"><input type="checkbox" value="' + _escT(x.codigo) + '" checked> ' +
            _escT(x.nombre) + ' (' + _escT(x.codigo) + ')' + (x.existe ? '' : ' <small style="color:#b45309">' + th('traer.lenguaNuevaAviso') + '</small>') + '</label>').join('') : th('traer.sinColumnas'));
        }).traerVerLenguas(url);
      }
    </script>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(500), tx('traer.titulo'));
}

function servicioProcesarImportacionMasiva(url, estadoNuevos, soloLenguas) {
  let lock;
  try {
    lock = _lock();
    const estado = _etiquetaEstado(estadoNuevos || 'revisión');
    const ext = SpreadsheetApp.openByUrl(url);
    const data = (ext.getSheetByName('DATOS_A_IMPORTAR') || ext.getSheets()[0]).getDataRange().getValues();
    if (data.length < 2) return _msgWarn(txh('traer.vacia'));

    const H = data[0].map(x => _s(x).toUpperCase());
    const iId = H.indexOf('ID_PALABRA'), iBase = H.indexOf('CONCEPTO_BASE'), iCat = H.indexOf('CATEGORIA');
    const langMap = {};
    H.forEach((hh, i) => {
      const m = hh.match(/^(TRADUCCION|TRAD_CAT|AUDIO|NOTA|DESCRIPCION|TIPO|ORIGINAL|FUENTE|UBICACION)_([A-Z0-9_]+)$/);
      if (!m) return;
      const iso = m[2];
      langMap[iso] = langMap[iso] || { tr: -1, tc: -1, au: -1, no: -1, de: -1, ti: -1, or: -1, fu: -1, ub: -1 };
      langMap[iso][{ TRADUCCION: 'tr', TRAD_CAT: 'tc', AUDIO: 'au', NOTA: 'no', DESCRIPCION: 'de', TIPO: 'ti', ORIGINAL: 'or', FUENTE: 'fu', UBICACION: 'ub' }[m[1]]] = i;
    });
    // Traer solo algunas lenguas (las demás columnas se ignoran)
    if (soloLenguas && soloLenguas.length) {
      const quiero = soloLenguas.map(x => _s(x).toUpperCase());
      Object.keys(langMap).forEach(iso => { if (!quiero.includes(iso)) delete langMap[iso]; });
      if (!Object.keys(langMap).length) return _msgWarn(txh('traer.sinLenguas'));
    }
    // Opcionales: de qué fuente viene cada fila y en qué parte de ella (página, entrada, sección)
    const iFuente = H.indexOf('FUENTE'), iUbic = H.indexOf('UBICACION');
    let nombreArchivo = '';
    try { nombreArchivo = ext.getName(); } catch (e) {}
    const lote = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + ' · ' + (nombreArchivo || tx('traer.loteSinNombre'));
    const nombresNuevos = {};
    const meta = ext.getSheetByName('META');
    if (meta) meta.getDataRange().getValues().slice(1).forEach(r => { if (_s(r[0])) nombresNuevos[_s(r[0]).toUpperCase()] = _s(r[1]); });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = getAppConfig(ss);
    if (Object.values(langMap).some(m => m.de > -1)) { const shT = ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES); if (shT) _asegurarColumnas(shT, ['descripcion']); }
    const tId = _tabla(ss, CONFIG.SHEETS.IDIOMAS), tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
    const tCat = _tabla(ss, CONFIG.SHEETS.CATEGORIAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
    const tTc = _tabla(ss, CONFIG.SHEETS.TRAD_CATS);
    _requerir(tId, ['id_idioma'], CONFIG.SHEETS.IDIOMAS);
    _requerir(tPal, ['id_palabra'], CONFIG.SHEETS.PALABRAS);
    _requerir(tTr, ['id_traduccion', 'id_palabra', 'id_idioma', 'texto'], CONFIG.SHEETS.TRADUCCIONES);
    const reglas = _cargarReglas(ss);
    const modoCaptura = _col(tTr, 'texto_captura') > -1;
    const yo = _miCorreo();
    const st = { palabras: 0, trNuevas: 0, trAct: 0, idiomas: 0, cats: 0, tcNuevas: 0, tcAct: 0, fuentes: 0, avisos: [] };
    let ctxImp = null;
    try { ctxImp = _contextoPermisos(ss); } catch (e) { ctxImp = null; }
    // Fuentes por ID o por título; una que no existe se crea con permisos Pendiente y sin crédito público (se revisa en 📚 Fuentes)
    const tFu = _tabla(ss, CONFIG.SHEETS.FUENTES), fuentePorClave = new Map(), fuentesNuevas = [];
    if (tFu) tFu.filas.forEach(f => { const id = _s(_get(tFu, f, 'id_fuente')); if (id) { fuentePorClave.set(_k(id), id); fuentePorClave.set(_k(_get(tFu, f, 'titulo')), id); } });
    const resolverFuente = v => {
      v = _s(v); if (!v || !tFu) return '';
      if (fuentePorClave.has(_k(v))) return fuentePorClave.get(_k(v));
      const id = Utilities.getUuid();
      fuentesNuevas.push({ id_fuente: id, titulo: v, tipo: 'Otro', base_uso: 'Autorización expresa', publicar_credito: 'Sí', notas: tx('traer.fuenteNueva'), estado: _etiquetaEstado('publicado') });
      fuentePorClave.set(_k(v), id); st.fuentes++;
      return id;
    };
    const nuevos = { idiomas: [], palabras: [], cats: [], trads: [], tcats: [] };

    // Idiomas: por ID, por el código de la comunidad (wix) o por el código internacional (hch), sin distinguir mayúsculas
    const idiomaPorClave = new Map();
    tId.filas.forEach(f => {
      const id = _s(_get(tId, f, 'id_idioma')); if (!id) return;
      idiomaPorClave.set(_k(id), id);
      const intl = _k(_get(tId, f, 'codigo_internacional')); if (intl && !idiomaPorClave.has(intl)) idiomaPorClave.set(intl, id);
      const iso = _k(_get(tId, f, 'codigo_iso')); if (iso) idiomaPorClave.set(iso, id);   // el propio manda sobre el internacional
    });
    const resolverIdioma = iso => {
      const k = _k(iso);
      if (idiomaPorClave.has(k)) return idiomaPorClave.get(k);
      const id = Utilities.getUuid(), nombre = nombresNuevos[iso] || tx('traer.lenguaNueva', { codigo: iso });
      nuevos.idiomas.push({ id_idioma: id, nombre: nombre, nombre_completo: nombre, codigo_iso: iso });
      idiomaPorClave.set(k, id); st.idiomas++; (st.creadas = st.creadas || []).push(_s(iso).toUpperCase());
      return id;
    };

    // Categorías: por ID o por nombre; si no existe y hay nombre, se crea. Sin categoría → vacío.
    const catPorClave = new Map();
    if (tCat) tCat.filas.forEach(f => {
      const id = _s(_get(tCat, f, 'id_categoria')); if (!id) return;
      catPorClave.set(_k(id), id);
      const nom = _k(_get(tCat, f, 'nombre_categoria')); if (nom) catPorClave.set(nom, id);
    });
    const resolverCategoria = valor => {
      const v = _s(valor); if (!v) return '';
      if (catPorClave.has(_k(v))) return catPorClave.get(_k(v));
      if (!tCat) return '';
      const id = Utilities.getUuid();
      nuevos.cats.push({ id_categoria: id, nombre_categoria: v });
      catPorClave.set(_k(v), id); catPorClave.set(_k(id), id); st.cats++;
      return id;
    };

    const catDePalabra = new Map();
    tPal.filas.forEach(f => catDePalabra.set(_k(_get(tPal, f, 'id_palabra')), _s(_get(tPal, f, 'id_categoria'))));

    // Traducciones: actualizar si existen, insertar si no (sin duplicar)
    const idxTr = _indiceTrads(tTr), pendTr = new Map();
    const fielDe = extra => !!(extra.id_fuente && ctxImp && _fuenteFiel(ctxImp, extra.id_fuente));
    const upsertTr = (idPal, idL, texto, audio, nota, extra, desc) => {
      extra = extra || {};
      const descN = desc ? _s(fielDe(extra) ? _s(desc).normalize('NFC') : _normalizar(desc, idL, reglas, yo, !modoCaptura)).slice(0, LARGO_DESCRIPCION) : '';
      const clave = _k(idPal) + '||' + _k(idL);
      const fiel = extra.id_fuente && ctxImp && _fuenteFiel(ctxImp, extra.id_fuente);
      const fiExistente = idxTr.has(clave) ? idxTr.get(clave) : -1;
      const autor = (fiExistente > -1 && _s(_get(tTr, tTr.filas[fiExistente], 'creado_por'))) || yo;
      const textoN = texto ? (fiel ? _s(texto).normalize('NFC') : _normalizar(texto, idL, reglas, autor, !modoCaptura)) : '';
      if (fiExistente > -1) {
        const fi = fiExistente, f = tTr.filas[fi], c = {};
        if (textoN && _s(_get(tTr, f, 'texto')) !== textoN) { c.texto = textoN; c.texto_captura = texto; }
        if (audio && _s(_get(tTr, f, 'audio')) !== audio) c.audio = audio;
        if (nota && _s(_get(tTr, f, 'nota_variante')) !== nota) c.nota_variante = nota;
        if (descN && _s(_get(tTr, f, 'descripcion')) !== descN) c.descripcion = descN;
        Object.keys(extra).forEach(k => { if (extra[k] && !_s(_get(tTr, f, k))) c[k] = extra[k]; }); // solo llena lo vacío
        if (Object.keys(c).length) { _actualizar(tTr, fi, c); st.trAct++; }
      } else if (pendTr.has(clave)) {
        const o = pendTr.get(clave);
        if (audio && !o.audio) o.audio = audio;
        if (nota && !o.nota_variante) o.nota_variante = nota;
        if (descN && !o.descripcion) o.descripcion = descN;
      } else if (textoN) {
        const o = Object.assign({ id_traduccion: Utilities.getUuid(), id_palabra: idPal, id_idioma: idL, texto: textoN, texto_captura: texto, audio: audio, nota_variante: nota, descripcion: descN, creado_por: yo, estado }, extra);
        nuevos.trads.push(o); pendTr.set(clave, o); st.trNuevas++;
      }
    };

    // Traducciones de categorías (si la hoja TRADUCCIONES_CATS tiene una estructura reconocible)
    let tcTexto = null, tcId = null;
    const idxTc = new Map(), pendTc = new Set();
    if (tTc && _col(tTc, 'id_categoria') > -1 && _col(tTc, 'id_idioma') > -1) {
      tcTexto = ['texto', 'nombre', 'traduccion', 'nombre_traducido'].find(c => _col(tTc, c) > -1) || null;
      tcId = tTc.headers.find(hh => hh.startsWith('id_') && hh !== 'id_categoria' && hh !== 'id_idioma') || null;
      tTc.filas.forEach((f, i) => idxTc.set(_k(_get(tTc, f, 'id_categoria')) + '||' + _k(_get(tTc, f, 'id_idioma')), i));
    }
    const upsertTc = (idCat, idL, texto) => {
      if (!idCat) return;
      if (!tcTexto) { if (!st.avisos.includes('tc')) st.avisos.push('tc'); return; }
      const clave = _k(idCat) + '||' + _k(idL);
      const textoN = _normalizar(texto, idL, reglas, yo, true);
      if (idxTc.has(clave)) {
        const fi = idxTc.get(clave);
        if (_s(_get(tTc, tTc.filas[fi], tcTexto)) !== textoN) { _actualizar(tTc, fi, { [tcTexto]: textoN }); st.tcAct++; }
      } else if (!pendTc.has(clave)) {
        const o = { id_categoria: idCat, id_idioma: idL, [tcTexto]: textoN };
        if (tcId) o[tcId] = Utilities.getUuid();
        nuevos.tcats.push(o); pendTc.add(clave); st.tcNuevas++;
      }
    };

    const idBase = cfg.idiomaPrincipal;
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      let idPal = _v(row, iId);
      const baseTxt = _v(row, iBase), cat = _v(row, iCat);
      const tieneDatos = Object.keys(langMap).some(iso => _v(row, langMap[iso].tr));

      if (idPal && catDePalabra.has(_k(idPal))) {
        // palabra existente
      } else if (baseTxt || tieneDatos) {
        idPal = Utilities.getUuid();
        const idCat = resolverCategoria(cat);
        nuevos.palabras.push({ id_palabra: idPal, id_categoria: idCat, estado, creado_por: yo });
        catDePalabra.set(_k(idPal), idCat);
        st.palabras++;
        if (baseTxt && baseTxt.indexOf('(Sin') !== 0) {
          if (idBase) upsertTr(idPal, idBase, baseTxt, '', '');
          else if (!st.avisos.includes('base')) st.avisos.push('base');
        }
      } else continue;

      Object.keys(langMap).forEach(iso => {
        const m = langMap[iso];
        const texto = _v(row, m.tr), audio = _limpiarAudio(_v(row, m.au)), nota = _v(row, m.no), tc = _v(row, m.tc), desc = _v(row, m.de);
        if (!texto && !audio && !nota && !tc && !desc) return;
        const idL = resolverIdioma(iso);
        const tipo = TIPOS_REGISTRO.find(x => _sinAcentos(x) === _sinAcentos(_v(row, m.ti))) || '';
        const orig = ORIGINALIDAD.find(x => _sinAcentos(x) === _sinAcentos(_v(row, m.or))) || '';
        const extra = { id_fuente: resolverFuente(_v(row, m.fu) || _v(row, iFuente)), ubicacion_fuente: _v(row, m.ub) || _v(row, iUbic), tipo_registro: tipo, original: orig, lote };
        if (texto || audio || nota || desc) upsertTr(idPal, idL, texto, audio, nota, extra, desc);
        if (tc) upsertTc(catDePalabra.get(_k(idPal)), idL, tc);
      });
    }

    if (tFu) _agregar(tFu, fuentesNuevas);
    _agregar(tId, nuevos.idiomas);
    _agregar(tCat, nuevos.cats);
    _agregar(tPal, nuevos.palabras);
    _agregar(tTr, nuevos.trads);
    if (tcTexto) _agregar(tTc, nuevos.tcats);
    publicarCatalogo(ss);

    const avisos = [];
    if (st.avisos.includes('base')) avisos.push(txh('traer.avisoBase'));
    if (st.avisos.includes('tc')) avisos.push(txh('traer.avisoCat'));
    if (st.creadas) avisos.push(txh('traer.avisoLenguaNueva', { codigos: st.creadas.join(', ') }));
    if (st.fuentes) avisos.push(txh('traer.avisoFuente', { n: st.fuentes }));
    return `<div style="color:#166534;"><b>${txh('traer.listo')}</b><br>` +
      txh('traer.resumen', { palabras: st.palabras, nuevas: st.trNuevas, corregidas: st.trAct, categorias: st.cats, nombresCat: st.tcNuevas + st.tcAct, lenguas: st.idiomas }) + '</div>' + (avisos.length ? _msgWarn(avisos.join('<br>')) : '');
  } catch (e) {
    return _msgError(e.message);
  } finally {
    if (lock) lock.releaseLock();
  }
}

/** Las plantillas muestran estados ("✅ OK", "🔍 En Drive"); no deben guardarse como nombre de archivo de audio. */
function _limpiarAudio(v) {
  const s = _s(v);
  if (!s || /^(✅|🔍|❌)/.test(s) || ['ok', 'en drive', 'pendiente'].includes(_k(s))) return '';
  return s;
}

// ─────────────────────────────────────────────────────────────────────────────
// 7. PUBLICACIÓN PREVIA (SNAPSHOT) — lo que sirve la API
// ─────────────────────────────────────────────────────────────────────────────

function construirDatosPublicos(ss) {
  const cfg = getAppConfig(ss);
  const oculto = o => !_esPublico(o);
  const limpiar = arr => arr.map(o => {
    const c = {};
    Object.keys(o).forEach(k => { if (k && !CONFIG.COLUMNAS_PRIVADAS.includes(k)) c[k] = o[k]; });
    return c;
  });
  const palabrasRaw = readSheet(ss, CONFIG.SHEETS.PALABRAS).filter(p => _s(p.id_palabra) && !oculto(p));
  const visibles = new Set(palabrasRaw.map(p => _k(p.id_palabra)));
  const tradCats = limpiar(readSheet(ss, CONFIG.SHEETS.TRAD_CATS).filter(o => !oculto(o)));

  // Categorías con sus traducciones { id_idioma: texto } (el visor lee c.traducciones)
  const campoTextoCat = ['texto', 'nombre', 'traduccion', 'nombre_traducido'];
  const porCategoria = {};
  tradCats.forEach(tc => {
    const txt = _s(campoTextoCat.map(k => tc[k]).find(v => _s(v)));
    if (!txt) return;
    const idc = _s(tc.id_categoria);
    (porCategoria[idc] = porCategoria[idc] || {})[_s(tc.id_idioma)] = txt;
  });
  const categorias = limpiar(readSheet(ss, CONFIG.SHEETS.CATEGORIAS).filter(o => _s(o.id_categoria) && !oculto(o)))
    .map(c => Object.assign(c, { traducciones: porCategoria[_s(c.id_categoria)] || {} }));

  // Palabras: imagen como enlace directo de Drive, marca de juegos y crédito de la imagen (nombre público, no el código)
  const actoresPub = _actores(ss);
  const galerias = _galeriasPublicas(ss, visibles, actoresPub), cuenta = _conteoImagenes(ss);
  const palabras = palabrasRaw.map(p => {
    const id = _s(p.imagen_id);
    const img = id ? `https://lh3.googleusercontent.com/d/${id}` : _imagenDirecta(p.imagen_referencia);
    const o = Object.assign(limpiar([p])[0], { imagen_referencia: img, imagen_id: id, juegos: _k(p.usar_en_juegos) !== 'no',
      imagen_credito: _s(p.imagen_credito) ? _creditoVisible(actoresPub, p.imagen_credito) : '', imgs: cuenta[_k(p.id_palabra)] || 0 });
    const gal = galerias[_k(p.id_palabra)]; if (gal && gal.length) o.galeria = gal;
    if (_s(p.imagenes_max) === '') delete o.imagenes_max;
    return o;
  });
  // Audio solo con consentimiento público; nombre y lugar de la voz solo si la persona lo aceptó.
  // Apóstrofos unificados según la lengua y, si el texto trae varias formas ("el rancho, la casa"), la lista `formas`.
  const ap = _mapaApostrofos(ss), ctx = _contextoPermisos(ss), voces = ctx.voces;
  // Fuentes: solo se publica el crédito de las que lo permiten (o lo exigen por fidelidad)
  const fuentes = Object.values(ctx.fuentes).filter(f => _k(f.id_fuente) !== FUENTE_EQUIPO && !oculto(f) && (_sinAcentos(f.publicar_credito) !== 'no' || _fuenteFiel(ctx, f.id_fuente)))
    .map(f => ({ id_fuente: _s(f.id_fuente), titulo: _s(f.titulo), credito: _s(f.credito) || _s(f.autoria), anio: _s(f.anio), licencia: _s(f.licencia),
      enlace: _baseUso(f) === 'publico' ? _s(f.enlace) : '', sistema_escritura: _s(f.sistema_escritura) }));
  const conCredito = new Set(fuentes.map(f => _k(f.id_fuente)));
  // Una traducción solo sale si su fuente permite publicarla (las del equipo, siempre; las demás, según sus autorizaciones)
  const traducciones = limpiar(readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).filter(t => !oculto(t) && visibles.has(_k(t.id_palabra)) && _usoFuente(ctx, t.id_fuente, 'publicar') === 'si')
    .map(t => {
      _conVoz(t, voces);
      if (_usoFuente(ctx, t.id_fuente, 'didactico') !== 'si') t.didactico = false; // el visor no la usa en juegos ni actividades
      t.tipo_registro = _tipoRegistro(t);
      if (!conCredito.has(_k(t.id_fuente))) { delete t.id_fuente; delete t.ubicacion_fuente; }
      if (!_audioPublico(t)) t.audio = '';
      if (!_audioPublico(t) || !_vozPublica(t) || !_s(t.audio)) { delete t.hablante; delete t.lugar_hablante; }
      return t;
    }))
    .map(t => {
      t.texto = _unificarApostrofos(t.texto, ap[_k(t.id_idioma)]);
      if (_s(t.descripcion)) t.descripcion = _unificarApostrofos(_s(t.descripcion), ap[_k(t.id_idioma)]).slice(0, LARGO_DESCRIPCION); else delete t.descripcion;
      const formas = _formas(t.texto);
      if (formas.length > 1) t.formas = formas;
      return Object.assign(t, { variante: _s(t.nota_variante) });
    });

  // Lenguas: nombre visible, alfabeto y atajos de escritura del equipo (para escribir en los juegos)
  const reglas = _cargarReglas(ss);
  const idiomas = limpiar(readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(o => _s(o.id_idioma) && !oculto(o)))
    .map(i => {
      const atajos = [].concat(reglas.get('*') || [], reglas.get(_k(i.id_idioma)) || [])
        .filter(r => r.aplicaA === 'todos' && r.tipo !== 'regex').map(r => [r.buscar, r.reemplazar, r.mayus ? 1 : 0]);
      return Object.assign(i, {
        nombre_completo: _s(i.nombre_completo || i.nombre || i.nombre_propio),
        alfabeto: _normalizarAlfabeto(i.alfabeto), atajos   // con espacios, aunque en la hoja se haya escrito con comas
      });
    });

  // Relaciones: solo las publicadas, entre palabras visibles y de tipos que el catálogo decidió mostrar
  const tiposPub = new Map(_tiposRelacion(ss).filter(t => t.activo && t.publica).map(t => [_k(t.id), t]));
  const relaciones = readSheet(ss, CONFIG.SHEETS.RELACIONES)
    .filter(r => _s(r.id_relacion) && !oculto(r) && tiposPub.has(_k(r.tipo)) && _k(r.id_origen) !== _k(r.id_destino) &&
      visibles.has(_k(r.id_origen)) && visibles.has(_k(r.id_destino)))
    .map(r => ({ id_relacion: _s(r.id_relacion), id_origen: _s(r.id_origen), tipo: _s(tiposPub.get(_k(r.tipo)).id), id_destino: _s(r.id_destino), nota: _s(r.nota) }));
  const tiposUsados = new Set(relaciones.map(r => _k(r.tipo)));
  const tiposRelacion = [...tiposPub.values()].filter(t => tiposUsados.has(_k(t.id)))
    .map(t => ({ id: t.id, clave: t.clave, nombre: t.nombre, inverso: t.inverso, simetrica: t.simetrica, orden: t.orden }));

  return {
    // Solo información pública: sin token, sin ID de AppSheet, sin estados ni correos de quien captura
    config: {
      version: CONFIG.VERSION, idCatalogo: cfg.idCatalogo,
      titulo: cfg.titulo, subtitulo: cfg.subtitulo, idiomaPrincipal: cfg.idiomaPrincipal,
      contacto: cfg.contacto, creditos: cfg.creditos, acerca: cfg.acerca, licencia: cfg.licencia
    },
    idiomas, categorias, palabras, traducciones, relaciones, tiposRelacion, fuentes, gobernanza: _gobernanzaPublica(ss),
    traducciones_categorias: tradCats,
    aportes: _aportesPublicos(ss),
    interfaz: _interfazPublica(ss, cfg)
  };
}

/**
 * Traducciones del visor que viajan con el catálogo: solo los textos "visor.*" ya traducidos en TEXTOS.
 * El visor las ofrece en su botón 🌐; lo que falte se ve en el idioma de respaldo.
 */
function _interfazPublica(ss, cfg) {
  let tabla = null;
  try { tabla = _leerTextos(ss); } catch (e) { tabla = null; }
  const idiomas = [];
  if (tabla) tabla.codigos.filter(c => c !== 'es').forEach(c => {
    const textos = {};
    Object.keys(T_VISOR).forEach(k => { if (tabla.cols[c] && tabla.cols[c][k]) textos[k] = tabla.cols[c][k]; });
    if (Object.keys(textos).length) idiomas.push({ codigo: c, nombre: tabla.nombres[c] || c, textos });
  });
  const pred = _k(cfg && cfg.idiomaInterfaz);
  return { predeterminado: idiomas.some(i => i.codigo === pred) ? pred : '', total: Object.keys(T_VISOR).length, idiomas };
}

/** Si la referencia ya es una imagen web (https://…png), se usa tal cual; si es de Drive, se convierte. */
function _imagenDirecta(ref) {
  const v = _s(ref);
  if (!v) return '';
  const id = _idDrive(v);
  if (id) return `https://lh3.googleusercontent.com/d/${id}`;
  return /^https?:\/\//i.test(v) ? v : '';
}

/** Extrae un ID de archivo de Drive de un enlace o de un ID suelto. */
function _idDrive(v) {
  const s = _s(v);
  const m = s.match(/\/d\/([-\w]{20,})/) || s.match(/[?&]id=([-\w]{20,})/) || s.match(/^([-\w]{25,})$/);
  return m ? m[1] : '';
}

// ── Imágenes de Drive ────────────────────────────────────────────────────────
// AppSheet guarda rutas ("Palabras_Images/abc.jpg"); a mano se pegan enlaces de Drive,
// hipervínculos (Ctrl+K) o fórmulas HIPERVINCULO. Todo se lleva a un ID de archivo, se guarda
// en imagen_id (con la referencia que se resolvió en imagen_ref, para notar cambios) y las
// imágenes de palabras PUBLICADAS se comparten como "cualquiera con el enlace" para que el
// visor pueda mostrarlas. Las de borrador o revisión siguen privadas.

/** Referencia real de cada celda de imagen_referencia: enlace del formato, fórmula o valor. */
function _referenciasImagen(t) {
  const c = _col(t, 'imagen_referencia');
  if (c < 0 || !t.filas.length) return [];
  const rango = t.sh.getRange(2, c + 1, t.filas.length, 1);
  const formulas = rango.getFormulas();
  let ricos = null;
  try { ricos = rango.getRichTextValues(); } catch (e) { ricos = null; }
  return t.filas.map((f, i) => {
    const fx = _s(formulas[i][0]);
    const mf = fx.match(/HYPERLINK\(\s*"([^"]+)"/i) || fx.match(/HIPERVINCULO\(\s*"([^"]+)"/i);
    if (mf) return mf[1];
    const rt = ricos && ricos[i][0];
    if (rt) {
      const url = rt.getLinkUrl() || (rt.getRuns().map(r => r.getLinkUrl()).find(Boolean));
      if (url) return url;
    }
    return _s(f[c]);
  });
}

/** Busca el archivo de una referencia: ID/enlace de Drive o ruta de AppSheet (por nombre de archivo). */
function _resolverArchivo(ref) {
  const id = _idDrive(ref);
  if (id) { try { const f = DriveApp.getFileById(id); if (!f.isTrashed()) return f; } catch (e) { /* sin acceso */ } }
  if (/^https?:\/\//i.test(_s(ref)) && !/drive\.google|docs\.google|googleusercontent/i.test(ref)) return null; // imagen web externa
  const nombre = _s(ref).split('/').pop();
  if (!nombre || nombre.length < 3) return null;
  const it = DriveApp.getFilesByName(nombre);
  while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) return f; }
  return null;
}

/**
 * Resuelve imágenes nuevas o cambiadas y comparte las publicadas. Trabaja por lotes
 * (máximo 120 archivos por ejecución) para no exceder el tiempo; lo demás sigue en la próxima hora.
 */
function _prepararImagenes(ss) {
  const t = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  const rep = { resueltas: 0, sinArchivo: 0, compartidas: 0, errores: 0, pendientes: 0 };
  if (!t || _col(t, 'imagen_referencia') < 0 || _col(t, 'imagen_id') < 0) return rep;
  const refs = _referenciasImagen(t);
  const inicio = Date.now();
  let presupuesto = 120;
  t.filas.forEach((f, i) => {
    const ref = refs[i];
    const idActual = _s(_get(t, f, 'imagen_id')), refPrevia = _s(_get(t, f, 'imagen_ref'));
    if (!ref) { if (idActual || refPrevia) _actualizar(t, i, { imagen_id: '', imagen_ref: '' }); return; }
    if (idActual && refPrevia === ref) return;
    if (presupuesto <= 0 || Date.now() - inicio > 120000) { rep.pendientes++; return; }
    presupuesto--;
    const arch = _resolverArchivo(ref);
    if (arch) { _actualizar(t, i, { imagen_id: arch.getId(), imagen_ref: ref }); rep.resueltas++; }
    else { _actualizar(t, i, { imagen_id: '', imagen_ref: ref }); rep.sinArchivo++; }
  });
  // Compartir solo las publicadas que aún no se comparten
  if (_col(t, 'imagen_publica') > -1) {
    t.filas.forEach((f, i) => {
      const id = _s(_get(t, f, 'imagen_id'));
      if (!id || !_esPublico({ estado: _get(t, f, 'estado') }) || _s(_get(t, f, 'imagen_publica')) === id) return;
      if (presupuesto <= 0 || Date.now() - inicio > 200000) { rep.pendientes++; return; }
      presupuesto--;
      try {
        const f = DriveApp.getFileById(id);
        const acceso = f.getSharingAccess();
        if (acceso !== DriveApp.Access.ANYONE_WITH_LINK && acceso !== DriveApp.Access.ANYONE)
          f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        _actualizar(t, i, { imagen_publica: id }); rep.compartidas++;
      } catch (e) { rep.errores++; } // p. ej., imagen de otra cuenta que no se puede compartir, o dominio que lo prohíbe
    });
  }
  return rep;
}

// ── Audio para el visor (?type=audio&file=…) ─────────────────────────────────
// El visor pide cada audio por el valor de la columna "audio". Por seguridad, solo se entregan
// audios que aparecen en traducciones PUBLICADAS (nadie puede pedir otros archivos del Drive).

function _audiosPublicados(texto) {
  const o = JSON.parse(texto || _leerSnapshot());
  return new Set((o.data.traducciones || []).map(t => _s(t.audio)).filter(Boolean));
}

/** Encuentra el archivo de Drive a partir de un ID, una URL de Drive o una ruta de AppSheet ("Carpeta/archivo.mp3"). */
function _buscarArchivoAudio(valor) {
  const v = _s(valor);
  const id = (v.match(/[-\w]{25,}/) || [])[0];
  if (id) { try { const f = DriveApp.getFileById(id); if (!f.isTrashed()) return f; } catch (e) { /* no es un ID */ } }
  const nombre = v.split('/').pop();
  if (!nombre) return null;
  const cfg = getAppConfig(SpreadsheetApp.getActiveSpreadsheet());
  const carpetas = [];
  const raiz = _carpetaValida(cfg.audios);
  if (raiz) { carpetas.push(raiz); const subs = raiz.getFolders(); while (subs.hasNext()) carpetas.push(subs.next()); }
  for (const c of carpetas) { const it = c.getFilesByName(nombre); if (it.hasNext()) return it.next(); }
  const todos = DriveApp.getFilesByName(nombre); // archivos que AppSheet guardó en su propia carpeta
  while (todos.hasNext()) { const f = todos.next(); if (!f.isTrashed()) return f; }
  return null;
}

function _servirAudio(file, congelado) {
  // Primero, que esté publicado (en la versión que ve quien pregunta); después, la caché
  if (!_audiosPublicados(congelado ? _leerSnapshotCongelado() : null).has(_s(file))) return responseJSON({ status: 404, error: 'Audio no publicado.' });
  const clave = 'AUD_' + _hash(_s(file));
  const cache = CacheService.getScriptCache();
  const enCache = cache.get(clave);
  if (enCache) return responseJSON({ status: 'ok', audioData: enCache });
  const f = _buscarArchivoAudio(file);
  if (!f) return responseJSON({ status: 404, error: 'No se encontró el archivo de audio.' });
  const blob = f.getBlob();
  if (blob.getBytes().length > 5 * 1024 * 1024) return responseJSON({ status: 413, error: 'El audio supera 5 MB.' });
  const tipo = blob.getContentType() && blob.getContentType().indexOf('audio') === 0 ? blob.getContentType() : 'audio/mpeg';
  const datos = `data:${tipo};base64,${Utilities.base64Encode(blob.getBytes())}`;
  if (datos.length < 95000) { try { cache.put(clave, datos, CONFIG.CACHE.TTL); } catch (e) { /* audio grande: sin caché */ } }
  return responseJSON({ status: 'ok', audioData: datos });
}

/** Arma el catálogo público y lo guarda en Drive + caché. Si no hubo cambios, no reescribe nada (salvo `forzar`). */
function publicarCatalogo(ss, forzar) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  // Una copia de trabajo nunca publica: compartiría la publicación del catálogo principal
  if (_esRama(ss)) return { cambio: false, copiaDeTrabajo: true, version: '', conteos: {} };
  const cfg = getAppConfig(ss);
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG.PROPS.TOKEN, cfg.apiToken || '');
  try { props.setProperty('IMAGENES_ULTIMO_REPORTE', JSON.stringify(_prepararImagenes(ss))); }
  catch (e) { console.warn('Imágenes: ' + e.message); }

  const data = construirDatosPublicos(ss);
  const dataJson = JSON.stringify(data);
  const version = _hash(dataJson);
  const previo = _leerManifiesto();
  if (!forzar && previo && previo.version === version && props.getProperty(CONFIG.PROPS.SNAPSHOT_FILE)) {
    try { _publicarEstadoVisor(ss); } catch (e) { console.warn('Estado del visor: ' + e.message); }   // avisos que empiezan o vencen
    return Object.assign({ cambio: false }, previo);
  }

  const generado = new Date().toISOString();
  const texto = `{"status":200,"version":"${version}","generado":"${generado}","data":${dataJson}}`;
  _guardarSnapshotDrive(ss, texto);
  _guardarEnCache(texto);

  const manifiesto = {
    version, generado, bytes: texto.length, idCatalogo: cfg.idCatalogo || "",
    conteos: { idiomas: data.idiomas.length, categorias: data.categorias.length, palabras: data.palabras.length, traducciones: data.traducciones.length }
  };
  props.setProperty(CONFIG.PROPS.MANIFIESTO, JSON.stringify(manifiesto));
  try { _publicarEstadoVisor(ss); } catch (e) { console.warn('Estado del visor: ' + e.message); }
  return Object.assign({ cambio: true }, manifiesto);
}

function _leerManifiesto() {
  const m = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPS.MANIFIESTO);
  return m ? JSON.parse(m) : null;
}

function _guardarSnapshotDrive(ss, texto) {
  const props = PropertiesService.getScriptProperties();
  let file = null;
  const id = props.getProperty(CONFIG.PROPS.SNAPSHOT_FILE);
  if (id) { try { file = DriveApp.getFileById(id); if (file.isTrashed()) file = null; } catch (e) { file = null; } }
  if (file) file.setContent(texto);
  else {
    file = _carpetaDelLibro(ss).createFile(CONFIG.SNAPSHOT_NOMBRE, texto, 'application/json');
    props.setProperty(CONFIG.PROPS.SNAPSHOT_FILE, file.getId());
  }
}

/** Guarda el texto en fragmentos (cada valor de caché admite como máximo ~100 KB). */
function _guardarEnCache(texto, prefijo) {
  if (!CONFIG.CACHE.ENABLED) return;
  const cache = CacheService.getScriptCache();
  const P = prefijo || CONFIG.CACHE.PREFIX, N = CONFIG.CACHE.CHUNK_CHARS;
  const partes = {};
  let pos = 0, n = 0;
  while (pos < texto.length) {
    let fin = Math.min(pos + N, texto.length);
    const c = texto.charCodeAt(fin - 1);
    if (fin < texto.length && c >= 0xD800 && c <= 0xDBFF) fin--; // no partir caracteres compuestos (pares sustitutos)
    partes[`${P}_${n}`] = texto.substring(pos, fin);
    pos = fin; n++;
  }
  if (n > CONFIG.CACHE.MAX_CHUNKS) { cache.remove(`${P}_meta`); return; }
  try {
    cache.putAll(partes, CONFIG.CACHE.TTL);
    cache.put(`${P}_meta`, JSON.stringify({ n }), CONFIG.CACHE.TTL);
  } catch (e) {
    console.warn('No se pudo guardar en caché: ' + e.message);
    cache.remove(`${P}_meta`);
  }
}

function _leerDeCache(prefijo) {
  if (!CONFIG.CACHE.ENABLED) return null;
  const cache = CacheService.getScriptCache();
  const P = prefijo || CONFIG.CACHE.PREFIX;
  const meta = cache.get(`${P}_meta`);
  if (!meta) return null;
  const keys = Array.from({ length: JSON.parse(meta).n }, (_, i) => `${P}_${i}`);
  const vals = cache.getAll(keys);
  const partes = keys.map(k => vals[k]);
  return partes.some(p => p == null) ? null : partes.join('');
}

/** Caché → archivo en Drive → (último recurso) publicar en este momento. */
function _leerSnapshot() {
  const enCache = _leerDeCache();
  if (enCache) return enCache;
  const id = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPS.SNAPSHOT_FILE);
  if (id) {
    try {
      const t = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8');
      _guardarEnCache(t);
      return t;
    } catch (e) { /* se vuelve a publicar abajo */ }
  }
  publicarCatalogo(SpreadsheetApp.getActiveSpreadsheet(), true);
  const id2 = PropertiesService.getScriptProperties().getProperty(CONFIG.PROPS.SNAPSHOT_FILE);
  return _leerDeCache() || DriveApp.getFileById(id2).getBlob().getDataAsString('UTF-8');
}

function _hash(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 12);
}

function uiPublicarCatalogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const r = publicarCatalogo(ss, true);
  const d = _porQueNoSeVen(ss);
  SpreadsheetApp.getUi().alert(tx('publicar.listo', { palabras: r.conteos.palabras, traducciones: r.conteos.traducciones, lenguas: r.conteos.idiomas }) + (d ? '\n\n' + d : ''));
}

/** Explica, en palabras, por qué hay traducciones que el visor no muestra (estado, palabra o permiso de la fuente). */
function _porQueNoSeVen(ss) {
  const palabras = readSheet(ss, CONFIG.SHEETS.PALABRAS), trads = readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).filter(t => _s(t.id_palabra) && _s(t.texto));
  if (!trads.length) return '';
  const visibles = new Set(palabras.filter(p => _s(p.id_palabra) && _esPublico(p)).map(p => _k(p.id_palabra)));
  const ctx = _contextoPermisos(ss);
  let estado = 0, palabra = 0, sinFuente = 0; const fuentes = {};
  trads.forEach(t => {
    if (!_esPublico(t)) { estado++; return; }
    if (!visibles.has(_k(t.id_palabra))) { palabra++; return; }
    if (_usoFuente(ctx, t.id_fuente, 'publicar') !== 'si') {
      const f = ctx.fuentes[_k(t.id_fuente)];
      if (!f) { sinFuente++; return; }
      const n = _s(f.titulo) || _s(t.id_fuente);
      fuentes[n] = (fuentes[n] || 0) + 1;
    }
  });
  const l = [];
  if (estado) l.push(tx('publicar.no.estado', { n: estado }));
  if (palabra) l.push(tx('publicar.no.palabra', { n: palabra }));
  Object.keys(fuentes).forEach(f => l.push(tx('publicar.no.fuente', { n: fuentes[f], fuente: f })));
  if (sinFuente) l.push(tx('publicar.no.fuenteBorrada', { n: sinFuente }));
  return l.length ? tx('publicar.no.titulo') + '\n• ' + l.join('\n• ') : '';
}

/**
 * 🔗 Dirección del catálogo. Google a veces da con getUrl() una dirección que no es la de la aplicación web
 * (una implementación vieja o la de "biblioteca"), así que la dirección buena se guarda en CONFIG (URL_API)
 * y se puede probar desde aquí antes de dársela a nadie.
 */
const _RE_EXEC = /^https:\/\/script\.google\.com\/(a\/[^/]+\/)?macros\/s\/[\w-]+\/exec$/;

function _direccionValida(url) {
  url = _s(url).split('?')[0].split('#')[0];
  if (/\/macros\/library\//.test(url)) throw new Error(tx('direccion.error.biblioteca'));
  if (/\/dev$/.test(url)) throw new Error(tx('direccion.error.dev'));
  if (!_RE_EXEC.test(url)) throw new Error(tx('direccion.error.formato'));
  return url;
}

/** Llama a la dirección como lo haría el visor y explica lo que encontró. */
function direccionProbar(url) {
  url = _direccionValida(url);
  let r;
  try { r = UrlFetchApp.fetch(url + '?action=manifiesto&token=' + encodeURIComponent(_s(getAppConfig(SpreadsheetApp.getActiveSpreadsheet()).apiToken)), { muteHttpExceptions: true, followRedirects: true }); }
  catch (e) { return { ok: false, mensaje: tx('direccion.prueba.sinRespuesta') }; }
  const texto = r.getContentText() || '';
  let j = null; try { j = JSON.parse(texto); } catch (e) { j = null; }
  if (!j) { let cod = 0; try { cod = r.getResponseCode(); } catch (e) {} return _diagnosticoNoJSON(texto, cod); }
  if (j.status === 401) return { ok: false, mensaje: tx('direccion.prueba.otraClave') };
  if (j.status === 200) {
    const mio = _s(getAppConfig(SpreadsheetApp.getActiveSpreadsheet()).idCatalogo);
    return { ok: true, mensaje: tx('direccion.prueba.ok', { palabras: (j.conteos || {}).palabras || 0 }) + (j.idCatalogo && mio && j.idCatalogo !== mio ? ' ' + tx('direccion.prueba.otroCatalogo') : '') };
  }
  return { ok: false, mensaje: tx('direccion.prueba.error', { error: j.error || j.status }) };
}

/** Explica en palabras sencillas por qué la dirección devolvió una página y no datos, y muestra lo que dijo Google. */
function _diagnosticoNoJSON(texto, codigo) {
  const plano = _s(texto).replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/&nbsp;|&#160;/g, ' ')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
  let clave = 'direccion.prueba.html';
  if (/accounts\.google|ServiceLogin|signin/i.test(texto)) clave = 'direccion.prueba.acceso';
  else if (/unable to open the file|no se puede abrir el archivo|no es posible abrir el archivo/i.test(plano) || codigo === 404) clave = 'direccion.prueba.noExiste';
  else if (/function not found|no se encontr[oó] la funci[oó]n|doGet/i.test(plano)) clave = 'direccion.prueba.sinDoGet';
  else if (/authoriz|autoriza|permis/i.test(plano)) clave = 'direccion.prueba.autorizar';
  const detalle = plano ? ' ' + tx('direccion.prueba.detalle', { codigo: codigo || '?', texto: plano.slice(0, 220) }) : '';
  return { ok: false, mensaje: tx(clave) + (clave === 'direccion.prueba.acceso' ? '' : detalle) };
}

/** Lo que necesita alguien para abrir el catálogo en el visor: dirección, clave y, si se guardó, la dirección del visor. */
function conexionDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  const cfg = getAppConfig(ss);
  let sugerida = ''; try { sugerida = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return {
    url: cfg.urlApi || (_RE_EXEC.test(sugerida) ? sugerida : ''), guardada: !!cfg.urlApi,
    clave: _s(cfg.apiToken).length > 2 ? _s(cfg.apiToken) : '',   // así la trata el catálogo
    visor: cfg.urlVisor, titulo: cfg.titulo || ss.getName(), soyAdmin: _soyAdmin(ss),
    claveEquipo: _s(cfg.claveEquipo).length > 2 ? _s(cfg.claveEquipo) : '', ajustes: _visorAjustes(ss),
    contacto: CAMPOS_CONTACTO.filter(k => _s((cfg.contacto || {})[k])), hoy: _hoyISO(),
    propuestas: _soyAdmin(ss) ? _propuestasAjustes(ss) : null
  };
}

function _visorValido(u) {
  u = _s(u);
  if (!u) return '';
  if (!/^https:\/\/[^\s?#]+$/.test(u)) throw new Error(tx('conexion.error.visor'));
  return u;
}

/** Guarda la dirección del catálogo y la del visor (las dos opcionales por separado). */
function conexionGuardar(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const v = {};
  if (_s(d.url)) v.URL_API = _direccionValida(d.url);
  v.URL_VISOR = _visorValido(d.visor);
  _escribirConfig(ss, v);
  const r = conexionDatos();
  r.prueba = v.URL_API ? direccionProbar(v.URL_API) : null;
  return r;
}

/** accion: 'nueva' | 'quitar'. Rige de inmediato: quien tenga la clave anterior ya no entra. */
function claveCambiar(accion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lock = _lock();
  try {
    const antes = _s(getAppConfig(ss).apiToken);
    const nueva = accion === 'nueva' ? _nuevaClave() : accion === 'quitar' ? '' : null;
    if (nueva === null) throw new Error(tx('conexion.error.accion'));
    _escribirConfig(ss, { API_TOKEN: nueva });
    PropertiesService.getScriptProperties().setProperty(CONFIG.PROPS.TOKEN, nueva);
    _historial(ss, 'CONFIG', 'API_TOKEN', { clave: antes ? '•••' : '' }, { clave: nueva ? '•••' : '' });
    return conexionDatos();
  } finally { lock.releaseLock(); }
}

function uiVerEnlaceApi() {
  const html = `<style>body{font-family:sans-serif;font-size:13px;color:#1f2937;margin:0}.c{padding:12px;line-height:1.5}
      input:not([type=checkbox]):not([type=radio]),textarea{width:100%;font-size:13px;padding:7px;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:6px;font-family:inherit}
      button,.btn{display:inline-block;padding:8px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#e2e8f0;margin:6px 4px 0 0;color:#1f2937;text-decoration:none;font-size:13px}
      .prim{background:#475569;color:#fff}.peligro{background:#fee2e2;color:#b91c1c}
      .sec{border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:10px;background:#f8fafc}h3{margin:0 0 6px;font-size:14px}
      .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:8px}.error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
      small{color:#64748b}.dos{display:grid;grid-template-columns:1fr 1fr;gap:10px}code{font-size:15px;background:#fff;border:1px solid #cbd5e1;border-radius:6px;padding:4px 8px;display:inline-block}</style>
    ${_tScript()}
    <style>.tabs{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:10px}.tabs button{margin:0}.tabs button.on{background:#475569;color:#fff}
      .op{display:flex;gap:8px;align-items:flex-start;border:1px solid #e2e8f0;border-radius:8px;padding:8px;margin:4px 0;background:#fff;font-weight:normal}
      .op small{display:block}.estadoM{padding:8px;border-radius:6px;margin-bottom:8px}.on1{background:#fef3c7;color:#92400e}.off1{background:#dcfce7;color:#166534}
      label{display:block;margin-top:6px}.inl{display:inline-flex;gap:6px;align-items:center;font-weight:normal;margin-right:10px}.op input{margin-top:3px}</style>
    <div class="c">
      <div class="tabs" role="tablist"><button data-tab="pConexion" class="on">🔗 ${txh('vtab.conexion')}</button><button data-tab="pMant">🚧 ${txh('vtab.mant')}</button><button data-tab="pVisor">⚙️ ${txh('vtab.visor')}</button><button data-tab="pPart">📥 ${txh('vtab.part')}</button></div>
      <div id="pConexion">
      <p style="margin-top:0">${txh('conexion.intro')}</p>
      <div class="sec">
        <h3>1. ${txh('direccion.etiqueta')}</h3>
        <input id="u" placeholder="https://script.google.com/macros/s/…/exec" onclick="this.select()" autocomplete="off" spellcheck="false">
        <div id="uAviso"></div>
        <button onclick="probar()">${txh('direccion.probar')}</button>
        <details style="margin-top:6px"><summary><small>${txh('direccion.dondeEsta')}</small></summary>
          <ol>${[1, 2, 3, 4, 5].map(i => '<li>' + txh('direccion.paso' + i) + '</li>').join('')}</ol><p><small>${txh('direccion.actualizar')}</small></p></details>
      </div>
      <div class="sec">
        <h3>2. ${txh('direccion.claveTitulo')}</h3>
        <div id="clave"></div>
        <div id="claveBotones"></div>
        <h3 style="margin-top:12px">👥 ${txh('equipo.claveTitulo')}</h3>
        <small>${txh('equipo.claveD')}</small>
        <div id="claveEq"></div>
      </div>
      <div class="sec">
        <h3>3. ${txh('conexion.visor')} <small>${txh('comun.opcional')}</small></h3>
        <input id="v" placeholder="https://…github.io/…/" autocomplete="off" spellcheck="false">
        <small>${txh('conexion.visor.d')}</small>
      </div>
      <button class="prim" onclick="guardar()">${txh('conexion.guardar')}</button>
      <div id="m" role="status" aria-live="polite"></div>
      <div class="sec" style="margin-top:12px">
        <h3>📤 ${txh('conexion.compartir')}</h3>
        <small>${txh('conexion.compartir.d')}</small>
        <div><label class="inl"><input type="radio" name="para" value="publico" checked> ${txh('conexion.para.publico')}</label> &nbsp;
          <label class="inl"><input type="radio" name="para" value="equipo"> ${txh('conexion.para.equipo')}</label></div>
        <textarea id="msg" style="height:120px;margin-top:6px" readonly onclick="this.select()"></textarea>
        <div id="enlace"></div>
        <div><button onclick="copiar()">📋 ${txh('conexion.copiar')}</button><a class="btn" id="wa" target="_blank" rel="noopener">💬 WhatsApp</a><a class="btn" id="mail" target="_blank" rel="noopener">✉️ ${txh('conexion.correo')}</a></div>
        <p id="avisoClave"><small>${txh('conexion.avisoClave')}</small></p>
      </div>
      </div>
      <div id="pMant" hidden>
        <div id="mEstado"></div>
        <div class="sec"><h3>${txh('mant.como')}</h3>
          <label class="op"><input type="radio" name="modo" value="congelar" checked><span><b>❄️ ${txh('mant.congelar')}</b><small>${txh('mant.congelar.d')}</small></span></label>
          <label class="op"><input type="radio" name="modo" value="cerrar"><span><b>🔒 ${txh('mant.cerrar')}</b><small>${txh('mant.cerrar.d')}</small></span></label>
          <div id="opCongelar"><label class="inl"><input type="checkbox" id="mAvisoPub" checked> ${txh('mant.avisoPublico')}</label></div>
          <div id="opCerrar" hidden><label class="inl"><input type="checkbox" id="mPurgar" checked> ${txh('mant.purgar')}</label><br><small>${txh('mant.purgar.d')}</small></div>
          <label for="mMensaje">${txh('mant.mensaje')} <small>${txh('comun.opcional')}</small></label><textarea id="mMensaje" rows="3" maxlength="800" placeholder="${txh('mant.mensaje.ph')}"></textarea>
          <label for="mHasta">${txh('mant.hasta')} <small>${txh('comun.opcional')}</small></label><input id="mHasta" maxlength="40" placeholder="${txh('mant.hasta.ph')}">
          <details><summary>${txh('mant.imagen')}</summary><small>${txh('mant.imagen.d')}</small><input id="mImagen" placeholder="${txh('mant.imagen.ph')}"><input id="mAlt" placeholder="${txh('visorAj.alt')}" style="margin-top:4px"></details>
          <div id="mContacto"></div>
          <div id="mEqAviso"></div>
          <button class="prim" data-acc="encender">🚧 ${txh('mant.encender')}</button>
        </div>
        <div class="sec"><h3>📅 ${txh('mant.programar')}</h3><small>${txh('mant.programar.d')}</small>
          <div class="dos"><div><label for="pIni">${txh('mant.inicio')}</label><input type="datetime-local" id="pIni"></div><div><label for="pFin">${txh('mant.fin')} <small>${txh('comun.opcional')}</small></label><input type="datetime-local" id="pFin"></div></div>
          <label class="inl"><input type="checkbox" id="pAviso" checked> ${txh('mant.crearAviso')}</label>
          <textarea id="pTexto" rows="3" maxlength="500"></textarea>
          <button data-acc="programar">📅 ${txh('mant.programarBoton')}</button>
        </div>
      </div>
      <div id="pVisor" hidden>
        <div class="sec"><h3>🔄 ${txh('visorAj.revision')}</h3><small>${txh('visorAj.revision.d')}</small>
          <select id="vRev"><option value="5">${txh('visorAj.min', {n: 5})}</option><option value="10">${txh('visorAj.min', {n: 10})}</option><option value="30">${txh('visorAj.min', {n: 30})}</option><option value="0">${txh('visorAj.nunca')}</option></select></div>
        <div class="sec"><h3>🎨 ${txh('visorAj.identidad')}</h3><small>${txh('visorAj.identidad.d')}</small>
          <input id="vImg" placeholder="${txh('mant.imagen.ph')}"><input id="vAlt" placeholder="${txh('visorAj.alt')}" style="margin-top:4px">
          <label class="inl"><input type="checkbox" id="vTitulo"> ${txh('visorAj.enTitulo')}</label></div>
        <div class="sec"><h3>🖨️ ${txh('visorAj.leyenda')}</h3><small>${txh('visorAj.leyenda.d')}</small>
          <input id="vLeyenda" maxlength="160" placeholder="${txh('visorAj.leyenda.ph')}"></div>
        <button class="prim" data-acc="visorGuardar">💾 ${txh('conexion.guardar')}</button>
        <div class="sec" style="margin-top:10px"><h3>📣 ${txh('avisos.titulo')}</h3><small>${txh('visorAj.avisos.d')}</small><br><button data-acc="avisos">📣 ${txh('visorAj.abrirAvisos')}</button></div>
        <div class="sec"><h3>🧽 ${txh('visorAj.borrar')}</h3><small>${txh('visorAj.borrar.d')}</small><div id="vBorrar"></div></div>
      </div>
      <div id="pPart" hidden>
        <div class="sec"><h3>📥 ${txh('part.titulo')}</h3><small>${txh('part.d')}</small>
          <label class="inl" style="margin-top:8px"><input type="checkbox" id="prActivo"> <b>${txh('part.activo')}</b></label>
          <div id="prSinCorreo"></div></div>
        <div class="sec"><h3>🔑 ${txh('part.quien')}</h3>
          <label class="inl"><input type="radio" name="prQuien" value="todos"> ${txh('part.quien.todos')}</label><br>
          <label class="inl"><input type="radio" name="prQuien" value="clave"> ${txh('part.quien.clave')}</label>
          <p><small>${txh('part.clave.d')}</small></p><div id="prClave"></div></div>
        <div class="sec"><h3>⚖️ ${txh('part.limites')}</h3>
          <div class="dos"><div><label for="prImg">${txh('part.maxImagenes')}</label><input id="prImg" type="number" min="1" max="30"></div>
          <div><label for="prDia">${txh('part.porDia')}</label><input id="prDia" type="number" min="1" max="200"></div></div>
          <div class="dos"><div><label for="prDesc">${txh('part.maxPendDesc')}</label><input id="prDesc" type="number" min="1" max="20"></div>
          <div><label for="prCorr">${txh('part.maxPendCorr')}</label><input id="prCorr" type="number" min="1" max="20"></div></div>
          <small>${txh('part.limites.d')}</small></div>
        <div class="sec"><h3>📜 ${txh('part.criterios')}</h3><small>${txh('part.criterios.d')}</small><textarea id="prCrit" style="height:80px" maxlength="1500" placeholder="${txh('part.criterios.ph')}"></textarea></div>
        <div class="sec"><label class="inl"><input type="checkbox" id="prAportes"> ${txh('part.aportes')}</label><br>
          <label class="inl"><input type="checkbox" id="prNov"> ${txh('part.novedades')}</label></div>
        <button class="prim" data-acc="partGuardar">💾 ${txh('conexion.guardar')}</button> <button data-acc="abrirPropuestas">📥 ${txh('part.abrirPropuestas')}</button>
        <p><small>💡 ${txh('part.piloto')}</small></p>
      </div>
    </div>
    <script>
      const $ = id => document.getElementById(id), esc = _escT;
      let D = null, confirmar = '';
      const pinta = r => { $('m').innerHTML = '<div class="' + (r.ok ? 'ok' : 'error') + '">' + esc(r.mensaje) + '</div>'; };
      const fallo = e => pinta({ ok: false, mensaje: e.message || e });
      let confEq = '', confBorrar = false;
      const para = () => (document.querySelector('input[name="para"]:checked') || {}).value === 'equipo' ? 'equipo' : 'publico';
      const claveDe = () => para() === 'equipo' ? D.claveEquipo : D.clave;
      document.querySelector('.tabs').addEventListener('click', e => { const b = e.target.closest('[data-tab]'); if (!b) return;
        document.querySelectorAll('.tabs button').forEach(x => x.classList.toggle('on', x === b)); ['pConexion', 'pMant', 'pVisor', 'pPart'].forEach(i => $(i).hidden = i !== b.dataset.tab); });
      document.querySelectorAll('input[name="para"]').forEach(r => r.onchange = pintar);
      document.querySelectorAll('input[name="modo"]').forEach(r => r.onchange = modoUI);
      function modoUI() { const c = (document.querySelector('input[name="modo"]:checked') || {}).value === 'cerrar'; $('opCerrar').hidden = !c; $('opCongelar').hidden = c; }
      function enlaceDirecto() {
        if (!D.visor || !D.url) return '';
        const k = claveDe();
        return D.visor + (D.visor.includes('?') ? '&' : '?') + 'api=' + encodeURIComponent(D.url) + (k ? '&token=' + encodeURIComponent(k) : '');
      }
      function mensaje() {
        const e = enlaceDirecto();
        if (!D.url) return tt('conexion.sinDireccion');
        if (para() === 'equipo' && !D.claveEquipo) return tt('equipo.sinClave');
        return tt(para() === 'equipo' ? 'conexion.mensajeEquipo' : 'conexion.mensaje', { catalogo: D.titulo }) + '\\n\\n' +
          (e ? tt('conexion.mensaje.enlace', { enlace: e }) + '\\n\\n' + tt('conexion.mensaje.oManual', { visor: D.visor }) + '\\n' : tt('conexion.mensaje.abreVisor') + '\\n') +
          tt('conexion.mensaje.direccion', { url: D.url }) + '\\n' +
          (claveDe() ? tt('conexion.mensaje.clave', { clave: claveDe() }) : tt('conexion.mensaje.sinClave'));
      }
      function pintarEquipo() {
        let h = D.claveEquipo ? '<code>' + esc(D.claveEquipo) + '</code><br>' : '<p><small>' + th('equipo.sinClave') + '</small></p>';
        if (D.soyAdmin) h += confEq ? '<div class="error">' + th(confEq === 'nueva' ? 'equipo.confirmaNueva' : 'equipo.confirmaQuitar') + '</div><button class="peligro" data-acc="eqSi">' + th('conexion.siCambiar') + '</button><button data-acc="eqNo">' + th('conexion.cancelar') + '</button>'
          : '<button data-acc="eqNueva">🔑 ' + th(D.claveEquipo ? 'conexion.claveNueva' : 'conexion.claveCrear') + '</button>' + (D.claveEquipo ? '<button data-acc="eqQuitar">' + th('conexion.claveQuitar') + '</button>' : '');
        $('claveEq').innerHTML = h;
      }
      function pintarMant() {
        const a = D.ajustes, mt = a.mant, pr = a.programa;
        $('mEstado').innerHTML = (mt.activo ? '<div class="estadoM on1">🚧 ' + th('mant.activo', {modo: tt('mant.m.' + mt.modo), desde: mt.desde, quien: mt.quien}) + '<br><button class="prim" data-acc="terminar">✅ ' + th('mant.terminar') + '</button></div>'
          : '<div class="estadoM off1">✅ ' + th('mant.enLinea') + '</div>') +
          (pr ? '<div class="estadoM on1">📅 ' + th('mant.programado', {inicio: String(pr.inicio).replace('T', ' '), fin: pr.fin ? String(pr.fin).replace('T', ' ') : '—', modo: tt('mant.m.' + (pr.modo || 'congelar'))}) + ' <button data-acc="cancelarProg">' + th('conexion.cancelar') + '</button></div>' : '');
        $('mContacto').innerHTML = '<label>' + th('mant.contacto') + '</label>' + (D.contacto.length ? D.contacto.map(k => '<label class="inl"><input type="checkbox" name="mC" value="' + k + '"' + ((mt.contacto || []).includes(k) ? ' checked' : '') + '> ' + th('mant.c.' + k) + '</label> ').join('') : '<small>' + th('mant.sinContacto') + '</small>');
        $('mEqAviso').innerHTML = D.claveEquipo ? '' : '<div class="error">' + th('mant.sinClaveEquipo') + '</div>';
        if (!$('pTexto').value) $('pTexto').value = tt('mant.textoAviso');
      }
      function datosMant() {
        return { modo: (document.querySelector('input[name="modo"]:checked') || {}).value || 'congelar', mensaje: $('mMensaje').value, hasta: $('mHasta').value, imagen: $('mImagen').value, alt: $('mAlt').value,
          contacto: [...document.querySelectorAll('input[name="mC"]:checked')].map(x => x.value), purgar: $('mPurgar').checked, avisoPublico: $('mAvisoPub').checked };
      }
      function pintarVisor() {
        const a = D.ajustes;
        $('vRev').value = String(a.revision); $('vImg').value = a.identidad.imagen || ''; $('vAlt').value = a.identidad.alt || ''; $('vTitulo').checked = !!a.identidad.enTitulo; $('vLeyenda').value = (a.impresion || {}).leyenda || '';
        $('vBorrar').innerHTML = confBorrar ? '<div class="error">' + th('visorAj.borrar.confirma') + '</div><button class="peligro" data-acc="borrarSi">' + th('conexion.siCambiar') + '</button><button data-acc="borrarNo">' + th('conexion.cancelar') + '</button>'
          : '<button data-acc="borrar">🧽 ' + th('visorAj.borrarBoton') + '</button>';
      }
      function pintarPart() {
        const p = D.propuestas; if (!p) return;
        $('prActivo').checked = !!p.activo; document.querySelectorAll('input[name="prQuien"]').forEach(r => r.checked = r.value === (p.quien === 'clave' ? 'clave' : 'todos'));
        $('prImg').value = p.maxImagenes; $('prDia').value = p.porDia; $('prDesc').value = p.maxPendDesc; $('prCorr').value = p.maxPendCorr;
        $('prCrit').value = p.criterios || ''; $('prAportes').checked = !!p.mostrarAportes; $('prNov').checked = !!p.novedades;
        $('prClave').innerHTML = (p.clave ? '<p>' + th('part.claveEs') + ' <code>' + esc(p.clave) + '</code></p>' : '<p><small>' + th('part.sinClave') + '</small></p>') +
          '<button data-acc="prClaveNueva">🔑 ' + th(p.clave ? 'conexion.claveNueva' : 'conexion.claveCrear') + '</button>' + (p.clave ? '<button data-acc="prClaveQuitar">' + th('conexion.claveQuitar') + '</button>' : '');
        $('prSinCorreo').innerHTML = D.contacto.includes('correo') ? '' : '<div class="error">' + th('part.sinCorreo') + '</div>';
      }
      const datosPart = () => ({ activo: $('prActivo').checked, quien: (document.querySelector('input[name="prQuien"]:checked') || {}).value || 'todos',
        maxImagenes: $('prImg').value, porDia: $('prDia').value, maxPendDesc: $('prDesc').value, maxPendCorr: $('prCorr').value, criterios: $('prCrit').value,
        mostrarAportes: $('prAportes').checked, novedades: $('prNov').checked });
      const listo = k => r => { D = r; pintar(); pinta({ ok: true, mensaje: tt(k) }); };
      document.body.addEventListener('click', e => {
        const b = e.target.closest('[data-acc]'); if (!b) return;
        const acc = b.dataset.acc, run = () => { $('m').textContent = tt('comun.guardando'); return google.script.run.withFailureHandler(fallo); };
        if (acc === 'eqNueva') { confEq = 'nueva'; pintarEquipo(); }
        if (acc === 'eqQuitar') { confEq = 'quitar'; pintarEquipo(); }
        if (acc === 'eqNo') { confEq = ''; pintarEquipo(); }
        if (acc === 'eqSi') { const a = confEq; confEq = ''; run().withSuccessHandler(listo(a === 'nueva' ? 'equipo.lista' : 'equipo.quitada')).claveEquipoCambiar(a); }
        if (acc === 'encender') run().withSuccessHandler(listo('mant.encendido')).mantenimientoEncender(datosMant());
        if (acc === 'terminar') run().withSuccessHandler(listo('mant.terminado')).mantenimientoTerminar();
        if (acc === 'programar') run().withSuccessHandler(listo('mant.programadoOk')).mantenimientoProgramar(Object.assign(datosMant(), { inicio: $('pIni').value, fin: $('pFin').value, crearAviso: $('pAviso').checked, textoAviso: $('pTexto').value }));
        if (acc === 'cancelarProg') run().withSuccessHandler(listo('mant.programaCancelado')).mantenimientoCancelarPrograma();
        if (acc === 'visorGuardar') run().withSuccessHandler(listo('conexion.guardado')).visorAjustesGuardar({ revision: $('vRev').value, identidad: { imagen: $('vImg').value, alt: $('vAlt').value, enTitulo: $('vTitulo').checked }, impresion: { leyenda: $('vLeyenda').value } });
        if (acc === 'borrar') { confBorrar = true; pintarVisor(); }
        if (acc === 'borrarNo') { confBorrar = false; pintarVisor(); }
        if (acc === 'borrarSi') { confBorrar = false; run().withSuccessHandler(listo('visorAj.borrado')).visorBorrarCopias(); }
        if (acc === 'avisos') google.script.run.uiAvisos();
        if (acc === 'partGuardar') run().withSuccessHandler(listo('conexion.guardado')).propuestasAjustesGuardar(datosPart());
        if (acc === 'prClaveNueva') { if (confirm(tt('part.confirmaClave'))) run().withSuccessHandler(listo('conexion.guardado')).propuestasAjustesGuardar(Object.assign(datosPart(), { clave: 'nueva' })); }
        if (acc === 'prClaveQuitar') run().withSuccessHandler(listo('conexion.guardado')).propuestasAjustesGuardar(Object.assign(datosPart(), { clave: 'quitar', quien: 'todos' }));
        if (acc === 'abrirPropuestas') google.script.run.uiPropuestas();
      });
      function pintar() {
        $('u').value = D.url; $('v').value = D.visor || '';
        $('uAviso').innerHTML = D.guardada ? '' : '<p><small>' + th(D.url ? 'direccion.sinComprobar' : 'direccion.pega') + '</small></p>';
        $('clave').innerHTML = D.clave ? '<p>' + th('conexion.claveEs') + '</p><code>' + esc(D.clave) + '</code>' : '<p><small>' + th('direccion.sinClave') + '</small></p>';
        let b = '';
        if (D.soyAdmin) {
          if (confirmar) b = '<div class="error">' + th(confirmar === 'nueva' ? 'conexion.confirmaNueva' : 'conexion.confirmaQuitar') + '</div><button class="peligro" onclick="cambiarClave()">' + th('conexion.siCambiar') + '</button><button onclick="confirmar=\\'\\';pintar()">' + th('conexion.cancelar') + '</button>';
          else b = '<button onclick="confirmar=\\'nueva\\';pintar()">🔑 ' + th(D.clave ? 'conexion.claveNueva' : 'conexion.claveCrear') + '</button>' + (D.clave ? '<button onclick="confirmar=\\'quitar\\';pintar()">🔓 ' + th('conexion.claveQuitar') + '</button>' : '');
        }
        $('claveBotones').innerHTML = b;
        const m = mensaje(), e = enlaceDirecto();
        $('msg').value = m;
        $('enlace').innerHTML = e ? '<p><b>' + th('conexion.enlaceDirecto') + '</b><br><a href="' + esc(e) + '" target="_blank" rel="noopener" style="word-break:break-all">' + esc(e) + '</a></p>' : '';
        $('wa').href = 'https://wa.me/?text=' + encodeURIComponent(m);
        $('mail').href = 'mailto:?subject=' + encodeURIComponent(tt('conexion.asunto', { catalogo: D.titulo })) + '&body=' + encodeURIComponent(m);
        $('avisoClave').style.display = claveDe() ? '' : 'none';
        pintarEquipo(); pintarMant(); pintarVisor(); pintarPart();
      }
      function probar() { $('m').textContent = tt('comun.cargando'); google.script.run.withFailureHandler(fallo).withSuccessHandler(pinta).direccionProbar($('u').value); }
      function guardar() {
        $('m').textContent = tt('comun.guardando');
        google.script.run.withFailureHandler(fallo).withSuccessHandler(r => { D = r; pintar(); if (r.prueba) pinta(r.prueba.ok ? r.prueba : { ok: false, mensaje: tt('conexion.guardadoPeroPrueba') + ' ' + r.prueba.mensaje }); else pinta({ ok: true, mensaje: tt('conexion.guardado') }); })
          .conexionGuardar({ url: $('u').value, visor: $('v').value });
      }
      function cambiarClave() {
        const a = confirmar; confirmar = '';
        google.script.run.withFailureHandler(fallo).withSuccessHandler(r => { D = r; pintar(); pinta({ ok: true, mensaje: tt(a === 'nueva' ? 'conexion.claveLista' : 'conexion.claveQuitada') }); }).claveCambiar(a);
      }
      function copiar() { $('msg').select(); try { navigator.clipboard.writeText($('msg').value).then(() => pinta({ ok: true, mensaje: tt('conexion.copiado') }), () => document.execCommand('copy')); } catch (e) { document.execCommand('copy'); } }
      google.script.run.withFailureHandler(fallo).withSuccessHandler(r => { D = r; pintar(); const mt = r.ajustes.mant;
        document.querySelector('input[name="modo"][value="' + (mt.modo === 'cerrar' ? 'cerrar' : 'congelar') + '"]').checked = true; modoUI();
        $('mMensaje').value = mt.mensaje || ''; $('mHasta').value = mt.hasta || ''; $('mImagen').value = mt.imagen || ''; $('mAlt').value = mt.alt || ''; $('mPurgar').checked = mt.purgar !== false; $('mAvisoPub').checked = mt.avisoPublico !== false;
      }).conexionDatos();
    </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(740), tx('direccion.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 8. MANTENIMIENTO Y UTILIDADES
// ─────────────────────────────────────────────────────────────────────────────

function servicioRepararIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const defs = [
    { n: CONFIG.SHEETS.CATEGORIAS, id: 'id_categoria' },
    { n: CONFIG.SHEETS.TRADUCCIONES, id: 'id_traduccion' },
    { n: CONFIG.SHEETS.PALABRAS, id: 'id_palabra' }
  ];
  let n = 0;
  defs.forEach(d => {
    const t = _tabla(ss, d.n);
    const c = _col(t, d.id);
    if (c < 0) return;
    t.filas.forEach((f, i) => {
      if (!_s(f[c]) && f.some(v => _s(v))) {
        t.sh.getRange(i + 2, c + 1).setNumberFormat('@').setValue(Utilities.getUuid());
        n++;
      }
    });
  });
  SpreadsheetApp.getUi().alert(n ? tx('codigos.listo', { n }) : tx('codigos.bien'));
}

/** Verifica y repara la estructura sin borrar datos (menú y panel). */
function servicioInstalacionCarpetas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  let rep, r;
  try {
    rep = asegurarEstructura(ss);
    _activarAutomatizacion();
    r = publicarCatalogo(ss, true);
  } finally { lock.releaseLock(); }
  const lineas = [tx('reparar.listo'), ''];
  if (!rep.hojasCreadas.length && !rep.columnas.length) lineas.push(tx('reparar.enOrden'));
  if (rep.hojasCreadas.length) lineas.push(tx('reparar.hojas', { hojas: rep.hojasCreadas.join(', ') }));
  if (rep.columnas.length) lineas.push(tx('reparar.columnas', { columnas: rep.columnas.join(' · ') }), '', tx('reparar.appsheet'));
  if (rep.equipo) lineas.push(tx('reparar.equipo', { correo: rep.equipo }));
  lineas.push('', tx('reparar.cadaHora'));
  SpreadsheetApp.getUi().alert(lineas.join('\n'));
}

function _activarAutomatizacion() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tareaAutomatica').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tareaAutomatica').timeBased().everyHours(1).create();
}

function servicioLimpiarCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(`${CONFIG.CACHE.PREFIX}_meta`);
    ['API_DATA_V1', 'API_DATA_V1_glosario', 'API_DATA_V1_idiomas'].forEach(k => cache.remove(k)); // claves de la versión 1.0
    const r = publicarCatalogo(SpreadsheetApp.getActiveSpreadsheet(), true);
    SpreadsheetApp.getUi().alert(tx('reparar.visorListo'));
  } catch (e) {
    SpreadsheetApp.getUi().alert(tx('comun.noTermino', { m: e.message }));
  }
  return true;
}

/** Respaldo .xlsx en la carpeta RESPALDOS, junto a la hoja de cálculo. */
function uiExportarDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const resp = UrlFetchApp.fetch(`https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) return SpreadsheetApp.getUi().alert(tx('copia.error'));
  const file = _subcarpeta(_carpetaDelLibro(ss), 'RESPALDOS').createFile(resp.getBlob().setName(`RESPALDO_${ss.getName()}_${ts}.xlsx`));
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<div style="padding:20px;text-align:center;font-family:sans-serif;">
      <p>${txh('copia.texto')}</p>
      <a href="${_esc(file.getUrl())}" target="_blank" style="display:block;background:#1e3a5f;color:white;padding:12px;text-decoration:none;border-radius:6px;font-weight:bold;">${txh('copia.abrir')}</a>
    </div>`).setWidth(380).setHeight(220), tx('copia.titulo'));
}

function _obtenerDiccionarioDrive(folderId) {
  const diccionario = new Set();
  if (!folderId || folderId.length < 5) return diccionario;
  const leer = carpeta => {
    const archivos = carpeta.getFiles();
    while (archivos.hasNext()) {
      const nom = archivos.next().getName().toLowerCase();
      diccionario.add(nom);
      diccionario.add(nom.replace(/\.[^.]+$/, ''));
    }
  };
  try {
    const raiz = DriveApp.getFolderById(folderId);
    leer(raiz);
    const subs = raiz.getFolders(); // subcarpetas por lengua
    while (subs.hasNext()) leer(subs.next());
  } catch (e) { console.warn('Carpeta Drive no encontrada: ' + folderId); }
  return diccionario;
}

function _v(row, idx) { return (idx == null || idx < 0 || idx >= row.length || row[idx] == null) ? '' : String(row[idx]).trim(); }
function _msgWarn(m) { return `<div style="color:#b45309;background:#fef3c7;padding:8px;border-radius:5px;margin-top:6px;">⚠️ ${m}</div>`; }
function _msgError(m) { return `<div style="color:#b91c1c;background:#fee2e2;padding:8px;border-radius:5px;">❌ ${_esc(m)}</div>`; }

// ─────────────────────────────────────────────────────────────────────────────
// 9. ENDPOINT API (WEB APP)
// ─────────────────────────────────────────────────────────────────────────────
// ?action=manifiesto  → versión y conteos (≈1 KB): el visor lo consulta primero
// ?action=glosario    → catálogo completo (misma forma que la versión 1.0)
// ?action=idiomas     → solo idiomas
// Ninguna respuesta lee las hojas: todo sale de la publicación previa.

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = _s(p.action) || 'ping';
  const info = {
    status: 200, message: 'API Activa',
    version: CONFIG.VERSION,
    endpoints_disponibles: ['?action=manifiesto', '?action=idiomas', '?action=glosario', '?action=paquete&ids=…', '?type=audio&file=…']
  };
  try {
    if (action === 'ping' && p.type !== 'audio') return responseJSON(info);
    const rol = _rolToken(p.token);
    const enMant = _respuestaMantenimiento(rol || 'publico');   // «cerrado»: pantalla de mantenimiento, aunque la clave falte
    if (enMant) return responseJSON(enMant);
    if (!rol) return responseJSON({ status: 401, error: 'Unauthorized: Token inválido o ausente.' });
    const cong = _publicoCongelado(rol);                        // «congelado»: el público ve la versión anterior
    if (p.type === 'audio') return _servirAudio(p.file, cong);
    if (action === 'lot_info' || action === 'lot_estado') return responseJSON(_lotGet(action, p));
    if (action === 'paquete') return responseJSON(_paqueteImagenes(_s(p.ids).split(',').filter(Boolean), Number(p.ancho) || 400));
    if (action === 'propuestas') return responseJSON(_estadoPropuestas(p.codigos));

    if (action === 'manifiesto') {
      const m = (cong && _manifiestoCongelado()) || _leerManifiesto() || publicarCatalogo(SpreadsheetApp.getActiveSpreadsheet(), true);
      return responseJSON(Object.assign({ status: 200 }, m, { cambio: undefined, rol, estado: _estadoParaVisor(_estadoVisor(), rol) }));
    }
    const snap = () => (cong && _leerSnapshotCongelado()) || _leerSnapshot();
    if (action === 'glosario') {
      return ContentService.createTextOutput(snap()).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'idiomas') {
      const o = JSON.parse(snap());
      return responseJSON({ status: 200, version: o.version, data: o.data.idiomas });
    }
    return responseJSON(info);
  } catch (err) {
    return responseJSON({ status: 500, error: err.message });
  }
}

function _tokenValido(t) {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty(CONFIG.PROPS.TOKEN);
  if (token === null) {
    token = getAppConfig(SpreadsheetApp.getActiveSpreadsheet()).apiToken || '';
    props.setProperty(CONFIG.PROPS.TOKEN, token);
  }
  if (!token || token.length <= 2) return true;
  return String(t || '') === token;
}

// ─────────────────────────────────────────────────────────────────────────────
// 10. ESTRUCTURA AUTOMÁTICA, ASISTENTE INICIAL Y PANEL DE CONTROL
// ─────────────────────────────────────────────────────────────────────────────
// Principio: nadie debería tener que crear hojas, columnas ni editar CONFIGURACION a mano.
// El script crea lo que falta, NUNCA borra ni reordena columnas, y todo ajuste se hace
// desde el panel. Los encabezados quedan protegidos con advertencia.

const VERSION_ESTRUCTURA = '1.3';
const ESTADOS_REGISTRO = ESTADOS.map(e => e.etiqueta);
const DESC_PROTECCION = 'Encabezados del catálogo (no editar a mano; usa el panel)';

/** Estructura esperada de cada hoja. */
function _esquema() {
  const S = CONFIG.SHEETS;
  return [
    { hoja: S.CONFIG, color: '#64748b', tecnica: true,
      cols: ['NOMBRE_PROYECTO', 'SUBTITULO', 'idioma_activo', 'ID_CATALOGO', 'ID_CARPETA_RAIZ', 'ID_CARPETA_AUDIOS', 'ID_CARPETA_IMAGENES',
             'ID_CARPETA_VIDEOS', 'ID_CARPETA_CATEGORIAS', 'APPSHEET_APP_ID', 'API_TOKEN', 'VERSION_ESTRUCTURA',
             'CONTACTO_NOMBRE', 'CONTACTO_CORREO', 'CONTACTO_TELEFONO', 'CONTACTO_WHATSAPP', 'CONTACTO_HORARIO', 'CONTACTO_TEXTO',
             'CREDITOS', 'ACERCA', 'LICENCIA', 'IDIOMA_INTERFAZ', 'SOLICITAR_IA', 'URL_API', 'URL_VISOR', 'CLAVE_EQUIPO', 'VISOR_AJUSTES', 'RAMA_DE', 'RAMA_URL', 'RAMA_NOMBRE', 'RAMA_FECHA', 'RAMA_AUTOR'] },
    { hoja: S.IDIOMAS, color: '#7c3aed',
      cols: ['id_idioma', 'nombre_completo', 'nombre_propio', 'codigo_iso', 'glottocode', 'variante', 'region',
             'alfabeto', 'carpeta_audios', 'carpeta_videos', 'estado', 'notas', 'codigo_internacional', 'id_fuente_alfabeto'],
      alias: { nombre_completo: ['nombre'] } },
    { hoja: S.CATEGORIAS, color: '#0891b2', cols: ['id_categoria', 'nombre_categoria', 'imagen_categoria', 'estado'] },
    { hoja: S.TRAD_CATS, color: '#0891b2', cols: ['id_trad_cat', 'id_categoria', 'id_idioma', 'texto'],
      alias: { id_trad_cat: ['id_traduccion_cat', 'id_traduccion_categoria', 'id'], texto: ['nombre', 'traduccion', 'nombre_traducido'] } },
    { hoja: S.PALABRAS, color: '#16a34a',
      cols: ['id_palabra', 'id_categoria', 'imagen_referencia', 'video', 'estado', 'usar_en_juegos', 'creado_por', 'origen',
             'imagen_id', 'imagen_ref', 'imagen_publica', 'imagen_credito', 'imagen_nota', 'imagen_alt', 'imagenes_max'] },
    { hoja: S.TRADUCCIONES, color: '#2563eb',
      cols: ['id_traduccion', 'id_palabra', 'id_idioma', 'texto', 'texto_captura', 'audio', 'hablante', 'lugar_hablante', 'consentimiento',
             'video', 'nota_variante', 'silabas', 'estado', 'creado_por', 'origen',
             'id_voz', 'tipo_registro', 'original', 'id_fuente', 'ubicacion_fuente', 'lote', 'audio_archivo', 'audio_nota', 'descripcion'],
      // En hojas con datos, texto_captura cambia el modo de trabajo: solo se agrega con "Activar modo captura"
      optIn: ['texto_captura'] },
    // Relaciones entre palabras (variante de, parte de, sinónimo…): la base para migrar a grafos
    { hoja: S.RELACIONES, color: '#be185d',
      cols: ['id_relacion', 'id_origen', 'tipo', 'id_destino', 'nota', 'estado', 'creado_por'] },
    { hoja: S.REGLAS, color: '#b45309', tecnica: true, cols: COLS_REGLAS },
    // Quién administra y quién captura. AppSheet lee los permisos de aquí; se edita con el panel 👥 Equipo
    { hoja: S.EQUIPO, color: '#0f766e', tecnica: true, cols: COLS_EQUIPO },
    // De dónde viene cada texto y quién presta su voz (sin datos de contacto: esos van en el directorio privado)
    { hoja: S.FUENTES, color: '#92400e', tecnica: true, cols: COLS_FUENTES },
    { hoja: S.ACTORES, color: '#92400e', tecnica: true, cols: COLS_ACTORES },
    { hoja: S.AUTORIZACIONES, color: '#92400e', tecnica: true, cols: COLS_AUTORIZACIONES },
    { hoja: S.COMUNIDADES, color: '#7f1d1d', tecnica: true, cols: COLS_COMUNIDADES },
    { hoja: S.VOCES, color: '#7f1d1d', tecnica: true, cols: COLS_VOCES },
    // Cambios de permisos y acuerdos: quién, cuándo, de qué a qué (las ventanas solo agregan)
    { hoja: S.HISTORIAL, color: '#334155', tecnica: true, cols: COLS_HISTORIAL },
    // Copias de trabajo creadas desde este catálogo (solo en el principal)
    { hoja: S.COPIAS, color: '#15803d', tecnica: true, cols: COLS_COPIAS },
    // Imágenes y sonidos propuestos que esperan revisión (🖼️ Imágenes y sonidos)
    { hoja: S.MEDIOS, color: '#a16207', tecnica: true, cols: COLS_MEDIOS },
    // Avisos que se muestran arriba del visor (📣 Avisos en el visor)
    { hoja: S.AVISOS, color: '#0369a1', tecnica: true, cols: COLS_AVISOS },
    // Vocabulario de tipos de relación entre palabras (🕸️ Tipos de relación)
    { hoja: S.TIPOS_REL, color: '#be185d', tecnica: true, cols: COLS_TIPOS_REL },
    // Lo que proponen docentes, hablantes y familias desde el visor (📥 Propuestas). Nada entra sin revisión
    { hoja: S.PROPUESTAS, color: '#0f766e', tecnica: true, cols: COLS_PROPUESTAS },
    // Galería: imágenes extra de cada palabra (la principal sigue en PALABRAS)
    { hoja: S.IMAGENES, color: '#16a34a', tecnica: true, cols: COLS_IMAGENES }
  ];
}

function _encabezados(sh) {
  const lc = sh.getLastColumn();
  return lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(h => _k(h)) : [];
}

/** Crea hojas y columnas faltantes, da formato, protege encabezados y valida. Idempotente. */
function asegurarEstructura(ss) {
  const rep = { hojasCreadas: [], columnas: [], omitidas: [] };
  _esquema().forEach(e => {
    let sh = ss.getSheetByName(e.hoja), nueva = false;
    if (!sh) { sh = ss.insertSheet(e.hoja); nueva = true; rep.hojasCreadas.push(e.hoja); }
    const sinDatos = sh.getLastRow() <= 1;
    const actuales = _encabezados(sh);
    const tiene = c => actuales.includes(_k(c)) || ((e.alias || {})[c] || []).some(a => actuales.includes(_k(a)));
    let faltan = e.cols.filter(c => !tiene(c));
    if (!sinDatos && e.optIn) {
      const omit = faltan.filter(c => e.optIn.includes(c));
      if (omit.length) rep.omitidas.push(`${e.hoja}: ${omit.join(', ')}`);
      faltan = faltan.filter(c => !e.optIn.includes(c));
    }
    if (faltan.length) {
      _asegurarColumnas(sh, faltan);
      if (!nueva) rep.columnas.push(`${e.hoja}: ${faltan.join(', ')}`);
    }
    _darFormato(sh, e);
  });
  _sincronizarTextos(ss);
  _asegurarFuenteEquipo(ss);
  try { _asegurarTiposRelacion(ss); } catch (e) { /* no detiene la reparación */ }
  _registrarHoja(ss);
  _aplicarValidaciones(ss);
  // Identificador permanente del catálogo (para intercambio y grafos)
  const shCfg = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  _cfgMemo = null;
  if (shCfg && !getAppConfig(ss).idCatalogo) _escribirConfig(ss, { ID_CATALOGO: Utilities.getUuid() });
  // Quien instala queda como primera persona administradora
  const tEq = _tabla(ss, CONFIG.SHEETS.EQUIPO);
  const yo = _correoActual();
  if (tEq && yo && !tEq.filas.some(f => _s(_get(tEq, f, 'correo')))) {
    _agregar(tEq, [{ correo: yo, nombre: '', rol: 'Administrador', activo: 'Sí', notas: tx('equipo.notaInstalo') }]);
    rep.equipo = yo;
  }
  return rep;
}

function _darFormato(sh, e) {
  const lc = sh.getLastColumn();
  if (!lc) return;
  sh.getRange(1, 1, 1, lc).setFontWeight('bold').setBackground(e.color).setFontColor('white');
  sh.setFrozenRows(1);
  sh.setTabColor(e.color);
  // Columnas de texto e identificadores en formato de texto plano
  const filas = sh.getMaxRows() - 1;
  _encabezados(sh).forEach((h, i) => {
    if (COLS_TEXTO_PLANO.includes(h) || h.startsWith('id_') || h === 'codigo_iso' || h === 'glottocode') _formatoTexto(sh, 2, i + 1, filas);
  });
  // Encabezados protegidos con advertencia (no bloquea a AppSheet ni al script)
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => p.getDescription() === DESC_PROTECCION).forEach(p => p.remove());
  sh.getRange(1, 1, 1, lc).protect().setDescription(DESC_PROTECCION).setWarningOnly(true);
}

function _aplicarValidaciones(ss) {
  const S = CONFIG.SHEETS;
  // Menú desplegable estricto: solo se pueden elegir los cuatro estados
  const estados = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_REGISTRO, true)
    .setAllowInvalid(false).setHelpText(tx('estados.ayudaValidacion')).build();
  [S.IDIOMAS, S.CATEGORIAS, S.PALABRAS, S.TRADUCCIONES].forEach(n => {
    _normalizarEstados(ss, n);
    _validarColumna(ss, n, 'estado', estados);
  });
  const siNo = SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).setAllowInvalid(true)
    .setHelpText(tx('ayuda.juegos')).build();
  _validarColumna(ss, S.PALABRAS, 'usar_en_juegos', siNo);
  _validarColumna(ss, S.RELACIONES, 'estado', estados);
  const lista = (valores, ayuda) => SpreadsheetApp.newDataValidation().requireValueInList(valores, true).setAllowInvalid(true).setHelpText(ayuda).build();
  // Los valores de las listas no se traducen (son datos que leen AppSheet y el visor); las ayudas sí
  _validarColumna(ss, S.TRADUCCIONES, 'consentimiento', lista(CONSENTIMIENTOS.map(c => c.etiqueta), tx('ayuda.consentimiento')));
  _validarColumna(ss, S.EQUIPO, 'rol', lista(ROLES_EQUIPO, tx('ayuda.rol')));
  _validarColumna(ss, S.EQUIPO, 'activo', lista(['Sí', 'No'], tx('ayuda.activo')));
  _validarColumna(ss, S.TRADUCCIONES, 'tipo_registro', lista(TIPOS_REGISTRO, tx('ayuda.tipoRegistro')));
  _validarColumna(ss, S.TRADUCCIONES, 'original', lista(ORIGINALIDAD, tx('ayuda.original')));
  _validarColumna(ss, S.VOCES, 'consentimiento', lista(CONSENTIMIENTOS.map(c => c.etiqueta), tx('ayuda.consentimiento')));
  _validarColumna(ss, S.VOCES, 'permiso_ia', lista(PERMISOS_IA.map(p => p.etiqueta), tx('ayuda.permisoIA')));
  _validarColumna(ss, S.VOCES, 'forma_permiso', lista(FORMAS_PERMISO, tx('ayuda.formaPermiso')));
  _validarColumna(ss, S.VOCES, 'activo', lista(['Sí', 'No'], tx('ayuda.activo')));
  _validarColumna(ss, S.FUENTES, 'tipo', lista(TIPOS_FUENTE, tx('ayuda.tipoFuente')));
  _validarColumna(ss, S.FUENTES, 'base_uso', lista(BASES_USO, tx('ayuda.baseUso')));
  ['publicar_credito', 'fidelidad'].forEach(c => _validarColumna(ss, S.FUENTES, c, lista(['Sí', 'No'], tx('ayuda.siNo'))));
  _validarColumna(ss, S.FUENTES, 'estado', estados);
  _validarColumna(ss, S.ACTORES, 'tipo', lista(TIPOS_ACTOR, tx('ayuda.tipoActor')));
  _validarColumna(ss, S.AUTORIZACIONES, 'rol', lista(ROLES_AUTORIZACION, tx('ayuda.rolAutorizacion')));
  _validarColumna(ss, S.AUTORIZACIONES, 'tipo_permiso', lista(TIPOS_PERMISO, tx('ayuda.tipoPermiso')));
  ['vigente', 'impugnada'].forEach(c => _validarColumna(ss, S.AUTORIZACIONES, c, lista(['Sí', 'No'], tx('ayuda.siNo'))));
  _validarColumna(ss, S.COMUNIDADES, 'en_disputa', lista(['Sí', 'No'], tx('ayuda.disputa')));
  _validarColumna(ss, S.COMUNIDADES, 'estado', estados);
  _validarColumna(ss, S.COPIAS, 'estado', lista(ESTADOS_COPIA, tx('ayuda.estadoCopia')));
  _validarColumna(ss, S.PROPUESTAS, 'estado', lista(ESTADOS_PROPUESTA, tx('ayuda.estadoPropuesta')));
  _validarColumna(ss, S.IMAGENES, 'estado', estados);
  _validarColumna(ss, S.IMAGENES, 'solo_interna', lista(['Sí', 'No'], tx('ayuda.siNo')));
  _validarColumna(ss, S.IMAGENES, 'tipo', lista(TIPOS_IMAGEN, tx('ayuda.tipoImagen')));
  USOS.forEach(u => { _validarColumna(ss, S.FUENTES, 'uso_' + u, lista(VALORES_USO, tx('ayuda.uso'))); _validarColumna(ss, S.AUTORIZACIONES, 'uso_' + u, lista(VALORES_USO, tx('ayuda.uso'))); });
  // Las columnas id_* no llevan lista desplegable: AppSheet las interpreta mal (Enum o Color) en vez de Ref
  [[S.TRADUCCIONES, 'id_idioma'], [S.TRAD_CATS, 'id_idioma'], [S.TRADUCCIONES, 'id_palabra'], [S.PALABRAS, 'id_categoria'], [S.TRAD_CATS, 'id_categoria'],
   [S.TRADUCCIONES, 'id_voz'], [S.TRADUCCIONES, 'id_fuente'], [S.VOCES, 'id_idioma'], [S.VOCES, 'id_comunidad'], [S.FUENTES, 'id_comunidad'],
   [S.AUTORIZACIONES, 'id_actor'], [S.AUTORIZACIONES, 'id_fuente'], [S.AUTORIZACIONES, 'id_comunidad'], [S.COMUNIDADES, 'id_idioma'], [S.ACTORES, 'id_comunidad']]
    .forEach(([h, c]) => _quitarValidacion(ss, h, c));
}

function _quitarValidacion(ss, hoja, col) {
  const sh = ss.getSheetByName(hoja);
  if (!sh || sh.getMaxRows() < 2) return;
  const c = _encabezados(sh).indexOf(col);
  if (c > -1) sh.getRange(2, c + 1, sh.getMaxRows() - 1, 1).clearDataValidations();
}

/** Lleva los estados escritos a mano ("publicada", "revision"…) a la etiqueta del menú. No cambia su significado. */
function _normalizarEstados(ss, hoja) {
  const t = _tabla(ss, hoja);
  const c = _col(t, 'estado');
  if (c < 0 || !t.filas.length) return 0;
  let n = 0;
  const valores = t.filas.map(f => {
    const v = _s(f[c]);
    if (!v) return [f[c]];
    const canon = _estadoCanonico(v);
    if (!ESTADOS.some(e => e.canon === canon)) return [f[c]]; // desconocido: se deja para revisión humana
    const et = _etiquetaEstado(canon);
    if (et !== v) n++;
    return [et];
  });
  if (n) t.sh.getRange(2, c + 1, valores.length, 1).setValues(valores);
  return n;
}

function _validarColumna(ss, hoja, col, regla) {
  if (!regla) return;
  const sh = ss.getSheetByName(hoja);
  if (!sh || sh.getMaxRows() < 2) return;
  const c = _encabezados(sh).indexOf(col);
  if (c > -1) sh.getRange(2, c + 1, sh.getMaxRows() - 1, 1).setDataValidation(regla);
}

/** En una hoja de cálculo nueva y vacía, deja una indicación de cómo empezar. */
function _mensajeHojaEnBlanco() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss.getSheetByName(CONFIG.SHEETS.CONFIG)) return;
    const hojas = ss.getSheets();
    if (hojas.length === 1 && hojas[0].getLastRow() === 0) {
      hojas[0].getRange('A1').setValue(tx('inicio.mensaje')).setFontWeight('bold').setFontSize(14);
    }
  } catch (e) { /* sin permisos en el disparador simple: no pasa nada */ }
}

function _eliminarHojaInicial(ss) {
  const nombres = _esquema().map(e => e.hoja).concat([CONFIG.SHEETS.REPORT, CONFIG.SHEETS.AUDIT, CONFIG.SHEETS.TEXTOS]);
  ss.getSheets().forEach(sh => {
    if (nombres.includes(sh.getName()) || ss.getSheets().length < 2) return;
    const vacia = sh.getLastRow() === 0 || (sh.getLastRow() === 1 && sh.getLastColumn() === 1 && String(sh.getRange('A1').getValue()).indexOf('👋') === 0);
    if (vacia) ss.deleteSheet(sh);
  });
}

// ── Configuración y carpetas ─────────────────────────────────────────────────

/** Escribe valores en CONFIGURACION (fila 2) por nombre de columna; crea la columna si falta. */
function _escribirConfig(ss, valores) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!sh) throw new Error(tx('comun.sinPreparar'));
  _asegurarColumnas(sh, Object.keys(valores));
  const h = _encabezados(sh);
  Object.keys(valores).forEach(k => {
    const i = h.indexOf(_k(k));
    if (i > -1) sh.getRange(2, i + 1).setNumberFormat('@').setValue(valores[k] == null ? '' : String(valores[k]));
  });
  _cfgMemo = null;
}

function _nuevaClave() { return Utilities.getUuid().replace(/-/g, '').slice(0, 20); }

function _carpetaValida(id) {
  if (!_s(id)) return null;
  try { const f = DriveApp.getFolderById(_s(id)); return f.isTrashed() ? null : f; } catch (e) { return null; }
}

/** Crea (o recupera) la carpeta del proyecto con Audios, Imágenes, Videos y Categorías. */
function _asegurarCarpetas(ss) {
  const cfg = getAppConfig(ss);
  const raiz = _carpetaValida(cfg.raiz) || _subcarpeta(_carpetaDelLibro(ss), `${cfg.titulo || ss.getName()} – Archivos`);
  const sub = (id, nombre) => (_carpetaValida(id) || _subcarpeta(raiz, nombre)).getId();
  const ids = {
    ID_CARPETA_RAIZ: raiz.getId(),
    ID_CARPETA_AUDIOS: sub(cfg.audios, 'Audios'),
    ID_CARPETA_IMAGENES: sub(cfg.imagenes, 'Imágenes'),
    ID_CARPETA_VIDEOS: sub(cfg.videos, 'Videos'),
    ID_CARPETA_CATEGORIAS: sub(cfg.carpetaCategorias, 'Categorías')
  };
  _escribirConfig(ss, ids);
  return ids;
}

function _etiquetaIdioma(nombrePropio, nombre, iso, variante) {
  const base = _s(nombrePropio) || _s(nombre) || 'Lengua';
  const extra = [_s(variante), _s(iso)].filter(Boolean).join(', ');
  return (extra ? `${base} (${extra})` : base).replace(/[\/\\:*?"<>|]/g, '-');
}

/** Subcarpetas de audio y video para una lengua (fila filaIdx de IDIOMAS). */
function _asegurarCarpetasIdioma(ss, t, filaIdx) {
  const cfg = getAppConfig(ss);
  const f = t.filas[filaIdx];
  const etiqueta = _etiquetaIdioma(_get(t, f, 'nombre_propio'), _get(t, f, 'nombre_completo') || _get(t, f, 'nombre'), _get(t, f, 'codigo_iso'), _get(t, f, 'variante'));
  const cambios = {};
  const audios = _carpetaValida(cfg.audios), videos = _carpetaValida(cfg.videos);
  if (audios && !_carpetaValida(_get(t, f, 'carpeta_audios'))) cambios.carpeta_audios = _subcarpeta(audios, etiqueta).getId();
  if (videos && !_carpetaValida(_get(t, f, 'carpeta_videos'))) cambios.carpeta_videos = _subcarpeta(videos, etiqueta).getId();
  if (Object.keys(cambios).length) _actualizar(t, filaIdx, cambios);
}

// ── Lenguas ──────────────────────────────────────────────────────────────────

function _datosIdioma(d) {
  const datos = {
    nombre_propio: _s(d.nombre_propio), nombre_completo: _s(d.nombre_completo),
    codigo_iso: _k(d.codigo_iso), glottocode: _k(d.glottocode),
    variante: _s(d.variante), region: _s(d.region), notas: _s(d.notas),
    codigo_internacional: _k(d.codigo_internacional), id_fuente_alfabeto: _s(d.id_fuente_alfabeto),
    alfabeto: _normalizarAlfabeto(d.alfabeto),
    estado: _etiquetaEstado(_estadoCanonico(d.estado) || 'publicado')
  };
  const quien = datos.nombre_propio || datos.nombre_completo || tx('lengua.laLengua');
  if (!datos.nombre_propio && !datos.nombre_completo) throw new Error(tx('lengua.error.nombre'));
  if (datos.codigo_iso && !/^[a-z]{3}$/.test(datos.codigo_iso)) throw new Error(tx('lengua.error.codigo', { lengua: quien }));
  if (datos.glottocode && !/^[a-z0-9]{4}\d{4}$/.test(datos.glottocode)) throw new Error(tx('lengua.error.glotto', { lengua: quien }));
  if (datos.codigo_internacional && !/^[a-z]{3}$/.test(datos.codigo_internacional)) throw new Error(tx('lengua.error.codigoIntl', { lengua: quien }));
  return datos;
}

/**
 * Crea o actualiza una lengua. No duplica: misma lengua = mismo nombre + misma variante. Devuelve su ID.
 * fusionar = true (asistente): en una lengua existente solo completa campos vacíos y no cambia su estado.
 */
function _upsertIdioma(ss, d, crearCarpetas, fusionar) {
  const t = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
  _requerir(t, ['id_idioma'], CONFIG.SHEETS.IDIOMAS);
  const datos = _datosIdioma(d);
  const nombreClave = _k(datos.nombre_propio || datos.nombre_completo) + '|' + _k(datos.variante);
  let fi = d.id ? t.filas.findIndex(f => _k(_get(t, f, 'id_idioma')) === _k(d.id)) : -1;
  if (fi < 0 && !d.id) {
    fi = t.filas.findIndex(f => (_k(_get(t, f, 'nombre_propio') || _get(t, f, 'nombre_completo') || _get(t, f, 'nombre')) + '|' + _k(_get(t, f, 'variante'))) === nombreClave);
  }
  const fila = Object.assign({ nombre: datos.nombre_completo || datos.nombre_propio }, datos);
  let id;
  if (fi > -1) {
    id = _s(_get(t, t.filas[fi], 'id_idioma'));
    if (fusionar) {
      const actual = t.filas[fi], soloNuevos = {};
      Object.keys(fila).forEach(k => { if (k !== 'estado' && _s(fila[k]) && !_s(_get(t, actual, k))) soloNuevos[k] = fila[k]; });
      _actualizar(t, fi, soloNuevos);
    } else {
      _actualizar(t, fi, fila);
    }
  } else {
    id = Utilities.getUuid();
    _agregar(t, [Object.assign({ id_idioma: id }, fila)]);
    fi = t.filas.length - 1;
  }
  if (crearCarpetas) _asegurarCarpetasIdioma(ss, t, fi);
  return id;
}

/** Alfabeto escrito con espacios o comas: "a ch e i ɨ k ꞌ". Sin repetidos, en el orden dado. */
function _normalizarAlfabeto(v) {
  const vistos = new Set();
  return _s(v).normalize('NFC').split(/[\s,;]+/).map(x => x.trim().toLowerCase()).filter(x => x && !vistos.has(x) && vistos.add(x)).join(' ');
}

/** Sugiere un alfabeto a partir de las palabras registradas en una lengua (letras simples, por frecuencia).
 *  Los dígrafos (ch, tz, kw…) los agrega quien administra, idealmente consultando a la comunidad. */
function sugerirAlfabeto(idIdioma) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  if (!t) return '';
  const cuenta = {};
  t.filas.forEach(f => {
    if (_k(_get(t, f, 'id_idioma')) !== _k(idIdioma)) return;
    (_s(_get(t, f, 'texto')).normalize('NFC').toLowerCase().match(/[\p{L}\uA78B\uA78C\u02BC'’][\p{M}]*/gu) || [])
      .forEach(l => { cuenta[l] = (cuenta[l] || 0) + 1; });
  });
  return Object.keys(cuenta).sort((a, b) => a.localeCompare(b, 'es')).join(' ');
}

/** Atajos de escritura en texto: "+=ɨ; O9=Ö" (también acepta → o -> y saltos de línea). */
function _parsearAtajos(texto) {
  return _s(texto).split(/[;\n]+/).map(x => x.trim()).filter(Boolean).map(par => {
    const m = par.match(/^(.+?)\s*(?:→|->|=)\s*(.+)$/);
    if (!m) throw new Error(tx('atajos.error.formato', { atajo: par }));
    return { buscar: m[1].trim(), reemplazar: m[2].trim() };
  });
}

function _agregarAtajos(ss, idIdioma, atajos) {
  if (!atajos.length) return 0;
  const tR = _hojaReglas(ss);
  const existentes = new Set(tR.filas.map(f => _k(_get(tR, f, 'id_idioma')) + '|' + String(_get(tR, f, 'buscar'))));
  let orden = tR.filas.reduce((m, f) => Math.max(m, Number(_get(tR, f, 'orden')) || 0), 0);
  const nuevas = [];
  atajos.forEach(a => {
    _compilarRegla({ buscar: a.buscar, reemplazar: a.reemplazar, tipo: 'texto', mayus: true });
    if (existentes.has(_k(idIdioma) + '|' + a.buscar)) return;
    nuevas.push({ id_idioma: idIdioma, buscar: a.buscar, reemplazar: a.reemplazar, tipo: 'texto', orden: ++orden,
      activa: true, aplica_a: 'todos', coincidir_mayus: true, nota: tx('atajos.notaAsistente') });
  });
  _agregar(tR, nuevas);
  return nuevas.length;
}

function listarIdiomas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
  if (!t) return [];
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  const conteo = {};
  if (tTr) tTr.filas.forEach(f => {
    const k = _k(_get(tTr, f, 'id_idioma'));
    conteo[k] = conteo[k] || { textos: 0, audios: 0 };
    if (_s(_get(tTr, f, 'texto'))) conteo[k].textos++;
    if (_s(_get(tTr, f, 'audio')).length > 1) conteo[k].audios++;
  });
  return t.filas.filter(f => _s(_get(t, f, 'id_idioma'))).map(f => {
    const id = _s(_get(t, f, 'id_idioma'));
    const c = conteo[_k(id)] || { textos: 0, audios: 0 };
    const url = cid => _s(cid) ? `https://drive.google.com/drive/folders/${_s(cid)}` : '';
    return {
      id, nombre_propio: _s(_get(t, f, 'nombre_propio')), nombre_completo: _s(_get(t, f, 'nombre_completo') || _get(t, f, 'nombre')),
      codigo_iso: _s(_get(t, f, 'codigo_iso')), glottocode: _s(_get(t, f, 'glottocode')), variante: _s(_get(t, f, 'variante')),
      region: _s(_get(t, f, 'region')), notas: _s(_get(t, f, 'notas')), estado: _estadoCanonico(_get(t, f, 'estado')) || 'publicado',
      alfabeto: _s(_get(t, f, 'alfabeto')),
      codigo_internacional: _s(_get(t, f, 'codigo_internacional')), id_fuente_alfabeto: _s(_get(t, f, 'id_fuente_alfabeto')),
      textos: c.textos, audios: c.audios,
      urlAudios: url(_get(t, f, 'carpeta_audios')), urlVideos: url(_get(t, f, 'carpeta_videos'))
    };
  });
}

function guardarIdioma(d) {
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    asegurarEstructura(ss);
    const cfg = getAppConfig(ss);
    const crear = !!(_carpetaValida(cfg.audios) || _carpetaValida(cfg.videos));
    const id = _upsertIdioma(ss, d, crear);
    const n = _agregarAtajos(ss, id, _parsearAtajos(d.atajos));
    if (n) aplicarReglasEscritura(ss);
    publicarCatalogo(ss);
    return listarIdiomas();
  } finally { lock.releaseLock(); }
}

// ── Asistente inicial ────────────────────────────────────────────────────────

function uiAsistenteInicial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg = {};
  try { if (ss.getSheetByName(CONFIG.SHEETS.CONFIG)) cfg = getAppConfig(ss); } catch (e) { cfg = {}; }
  const existentes = ss.getSheetByName(CONFIG.SHEETS.IDIOMAS) ? listarIdiomas().length : 0;
  const op = txh('comun.opcional');
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}
    .caja{padding:14px}
    label{display:block;font-weight:bold;margin-top:10px}
    input[type=text],input[type=number],select,textarea{width:100%;box-sizing:border-box;padding:9px;font-size:16px;border:1px solid #cbd5e1;border-radius:6px}
    .lengua{border:1px solid #c4b5fd;background:#faf5ff;border-radius:8px;padding:10px;margin-top:10px}
    .dos{display:flex;gap:8px}.dos>div{flex:1}
    small{color:#64748b;font-weight:normal}
    .op{display:block;margin:6px 0}
    button{padding:12px;border:none;border-radius:6px;font-weight:bold;font-size:15px;cursor:pointer;width:100%;background:#1e3a5f;color:#fff;margin-top:14px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
    .ok{background:#dcfce7;color:#166534;padding:10px;border-radius:6px;margin-top:8px}
    ol{padding-left:20px}
  </style>
  ${_tScript()}
  <div class="caja" id="form">
    <p style="margin-top:0">${txh('asis.intro')}
    ${existentes ? `<br><small>${txh('asis.yaHay', { n: existentes })}</small>` : ''}</p>
    <label for="proyecto">${txh('asis.nombre')}</label>
    <input type="text" id="proyecto" value="${_esc(cfg.titulo || '')}" placeholder="${txh('asis.nombre.ej')}">
    <label for="subtitulo">${txh('asis.frase')} <small>${op}</small></label>
    <input type="text" id="subtitulo" value="${_esc(cfg.subtitulo || '')}">
    <label for="n">${txh('asis.cuantas')} <small>${txh('asis.cuantas.d')}</small></label>
    <input type="number" id="n" min="1" max="30" value="${existentes ? 1 : 2}" onchange="pintar()">
    <div id="lenguas"></div>
    <label for="ref">${txh('asis.referencia')} <small>${txh('asis.referencia.d')}</small></label>
    <select id="ref"></select>
    <label for="cats">${txh('asis.grupos')} <small>${txh('asis.grupos.d')}</small></label>
    <textarea id="cats" rows="2"></textarea>
    <label>${txh('asis.opciones')}</label>
    <span class="op"><input type="checkbox" id="carpetas" checked> <label for="carpetas" style="display:inline;font-weight:normal">${txh('asis.carpetas')}</label></span>
    <span class="op"><input type="checkbox" id="clave"> <label for="clave" style="display:inline;font-weight:normal">${txh('asis.clave')} <small>${txh('asis.clave.d')}</small></label></span>
    <span class="op"><input type="checkbox" id="ocultar" checked> <label for="ocultar" style="display:inline;font-weight:normal">${txh('asis.ocultar')}</label></span>
    <label for="appsheet">${txh('asis.appsheet')} <small>${txh('asis.appsheet.d')}</small></label>
    <input type="text" id="appsheet" value="${_esc(cfg.appsheetId || '')}">
    <button id="btn" onclick="instalar()">${txh('asis.preparar')}</button>
    <div id="res" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    let previos = [];
    function leer(){ return Array.from(document.querySelectorAll('.lengua')).map(b => ({
      nombre_propio: b.querySelector('.np').value, nombre_completo: b.querySelector('.nc').value,
      codigo_iso: b.querySelector('.iso').value, variante: b.querySelector('.var').value, atajos: b.querySelector('.at').value })); }
    function pintar(){
      previos = leer();
      const n = Math.max(1, Math.min(30, Number($('n').value) || 1));
      let h = '';
      for (let i = 0; i < n; i++) {
        const p = previos[i] || {};
        h += '<div class="lengua"><b>' + th('asis.lengua', {n: i + 1}) + '</b>' +
          '<label>' + th('lengua.nombrePropio') + ' <small>' + th('lengua.nombrePropio.d') + '</small></label><input type="text" class="np" oninput="refs()" value="' + esc(p.nombre_propio) + '">' +
          '<label>' + th('lengua.nombreOtra') + ' <small>' + th('lengua.nombreOtra.d') + '</small></label><input type="text" class="nc" oninput="refs()" value="' + esc(p.nombre_completo) + '">' +
          '<div class="dos"><div><label>' + th('lengua.codigo') + ' <small>' + th('lengua.codigo.d') + '</small></label><input type="text" class="iso" maxlength="3" value="' + esc(p.codigo_iso) + '"></div>' +
          '<div><label>' + th('lengua.variante') + ' <small>' + th('comun.opcional') + '</small></label><input type="text" class="var" value="' + esc(p.variante) + '"></div></div>' +
          '<label>' + th('lengua.atajos') + ' <small>' + th('lengua.atajos.d') + '</small></label><input type="text" class="at" value="' + esc(p.atajos) + '">' +
          '</div>';
      }
      $('lenguas').innerHTML = h; refs();
    }
    function refs(){
      const sel = $('ref'), actual = sel.value;
      sel.innerHTML = '<option value="">' + th('comun.ninguna') + '</option>' + leer().map((l, i) => {
        const nom = l.nombre_propio || l.nombre_completo; return nom ? '<option value="' + i + '">' + esc(nom) + '</option>' : '';
      }).join('');
      sel.value = actual;
    }
    function instalar(){
      const d = { proyecto: $('proyecto').value, subtitulo: $('subtitulo').value, lenguas: leer(), referencia: $('ref').value,
                  categorias: $('cats').value, carpetas: $('carpetas').checked, clave: $('clave').checked,
                  ocultar: $('ocultar').checked, appsheet: $('appsheet').value };
      $('btn').disabled = true; $('res').textContent = tt('asis.preparando');
      google.script.run.withFailureHandler(e => { $('btn').disabled = false; $('res').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; })
        .withSuccessHandler(r => {
          $('form').innerHTML = '<div class="ok"><b>' + th('asis.listo', {nombre: r.proyecto}) + '</b><br>' + esc(r.resumen.join(' · ')) + '</div>' +
            (r.carpeta ? '<p><a href="' + esc(r.carpeta) + '" target="_blank">' + th('asis.abrirCarpeta') + '</a></p>' : '') +
            '<p><b>' + th('asis.siguientes') + '</b></p><ol>' + r.pasos.map(p => '<li>' + esc(p) + '</li>').join('') + '</ol>' +
            '<p><small>' + th('asis.demas') + '</small></p>';
        }).instalarCatalogo(d);
    }
    pintar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(680), tx('asis.titulo'));
}

function instalarCatalogo(d) {
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!_s(d.proyecto)) throw new Error(tx('asis.error.nombre'));
    const lenguas = (d.lenguas || []).filter(l => _s(l.nombre_propio) || _s(l.nombre_completo));
    if (!lenguas.length) throw new Error(tx('asis.error.lengua'));
    lenguas.forEach(l => { _datosIdioma(l); _parsearAtajos(l.atajos); }); // valida todo antes de tocar nada

    asegurarEstructura(ss);
    _eliminarHojaInicial(ss);
    _escribirConfig(ss, { NOMBRE_PROYECTO: _s(d.proyecto), SUBTITULO: _s(d.subtitulo), VERSION_ESTRUCTURA });
    if (_s(d.appsheet)) _escribirConfig(ss, { APPSHEET_APP_ID: _s(d.appsheet) });
    if (d.clave && !getAppConfig(ss).apiToken) _escribirConfig(ss, { API_TOKEN: _nuevaClave() });

    let carpeta = '';
    if (d.carpetas) carpeta = `https://drive.google.com/drive/folders/${_asegurarCarpetas(ss).ID_CARPETA_RAIZ}`;

    let atajos = 0;
    const ids = lenguas.map(l => {
      const id = _upsertIdioma(ss, l, !!d.carpetas, true);
      atajos += _agregarAtajos(ss, id, _parsearAtajos(l.atajos));
      return id;
    });
    if (_s(d.referencia) !== '' && ids[Number(d.referencia)]) _escribirConfig(ss, { idioma_activo: ids[Number(d.referencia)] });

    // Categorías iniciales (sin duplicar)
    const tCat = _tabla(ss, CONFIG.SHEETS.CATEGORIAS);
    const yaHay = new Set(tCat.filas.map(f => _k(_get(tCat, f, 'nombre_categoria'))));
    const cats = [...new Set(_s(d.categorias).split(/[\n,;]+/).map(_s).filter(c => c && !yaHay.has(_k(c))))];
    _agregar(tCat, cats.map(c => ({ id_categoria: Utilities.getUuid(), nombre_categoria: c, estado: _etiquetaEstado('publicado') })));

    _activarAutomatizacion();
    if (d.ocultar) _hojasTecnicas(ss, true);
    aplicarReglasEscritura(ss);
    publicarCatalogo(ss, true);

    const cfg = getAppConfig(ss);
    const pasos = [
      tx('asis.paso.direccion'), tx('asis.paso.appsheet'), tx('asis.paso.equipo')
    ];
    if (cfg.apiToken) pasos.push(tx('asis.paso.clave', { clave: cfg.apiToken }));
    return {
      proyecto: cfg.titulo, carpeta, pasos,
      resumen: [tx('asis.resumen.lenguas', { n: lenguas.length }), tx('asis.resumen.categorias', { n: cats.length }), tx('asis.resumen.atajos', { n: atajos }),
        tx(d.carpetas ? 'asis.resumen.carpetas' : 'asis.resumen.sinCarpetas')]
    };
  } finally { lock.releaseLock(); }
}

// ── Ajustes del proyecto ─────────────────────────────────────────────────────

function obtenerAjustes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  const cfg = getAppConfig(ss);
  const url = id => _carpetaValida(id) ? `https://drive.google.com/drive/folders/${id}` : '';
  return {
    titulo: cfg.titulo, subtitulo: cfg.subtitulo, idiomaPrincipal: cfg.idiomaPrincipal,
    appsheetId: cfg.appsheetId, apiToken: cfg.apiToken,
    contacto: cfg.contacto, creditos: cfg.creditos, acerca: cfg.acerca, licencia: cfg.licencia,
    idiomas: listarIdiomas().map(i => ({ id: i.id, nombre: i.nombre_propio || i.nombre_completo })),
    carpetas: [['archivos', cfg.raiz], ['audios', cfg.audios], ['imagenes', cfg.imagenes], ['videos', cfg.videos], ['categorias', cfg.carpetaCategorias]]
      .map(([k, id]) => ({ nombre: tx('ajustes.carpeta.' + k), url: url(id) }))
  };
}

function guardarAjustes(d) {
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!_s(d.titulo)) throw new Error(tx('ajustes.error.nombre'));
    const c = d.contacto || {};
    if (_s(c.correo) && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(_s(c.correo))) throw new Error(tx('ajustes.error.correo'));
    const valores = {
      NOMBRE_PROYECTO: _s(d.titulo), SUBTITULO: _s(d.subtitulo), idioma_activo: _s(d.idiomaPrincipal), APPSHEET_APP_ID: _s(d.appsheetId),
      CONTACTO_NOMBRE: _s(c.nombre), CONTACTO_CORREO: _s(c.correo), CONTACTO_TELEFONO: _s(c.telefono),
      CONTACTO_WHATSAPP: _s(c.whatsapp).replace(/[^\d+]/g, ''), CONTACTO_HORARIO: _s(c.horario), CONTACTO_TEXTO: _s(c.texto).slice(0, 600),
      CREDITOS: _s(d.creditos).slice(0, 3000), ACERCA: _s(d.acerca).slice(0, 3000), LICENCIA: _s(d.licencia).slice(0, 300)
    };
    if (d.clave === 'nueva') valores.API_TOKEN = _nuevaClave();
    if (d.clave === 'quitar') valores.API_TOKEN = '';
    _escribirConfig(ss, valores);
    publicarCatalogo(ss, true);
    return obtenerAjustes();
  } finally { lock.releaseLock(); }
}

function crearCarpetasDesdePanel() {
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    _asegurarCarpetas(ss);
    const t = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
    if (t) t.filas.forEach((f, i) => { if (_s(_get(t, f, 'id_idioma'))) _asegurarCarpetasIdioma(ss, t, i); });
    return obtenerAjustes();
  } finally { lock.releaseLock(); }
}

function uiAjustes() {
  const op = txh('comun.opcional');
  const area = 'style="width:100%;box-sizing:border-box;padding:9px;font-size:15px;border:1px solid #cbd5e1;border-radius:6px"';
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:14px}
    label{display:block;font-weight:bold;margin-top:10px}
    input[type=text],select{width:100%;box-sizing:border-box;padding:9px;font-size:16px;border:1px solid #cbd5e1;border-radius:6px}
    small{color:#64748b;font-weight:normal}
    button{padding:10px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;margin-top:10px}
    .prim{background:#1e3a5f;color:#fff;width:100%}.sec{background:#e2e8f0}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:8px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
    code{background:#f1f5f9;padding:2px 5px;border-radius:4px;word-break:break-all}
  </style>
  ${_tScript()}
  <div class="caja">
    <label for="titulo">${txh('asis.nombre')}</label><input type="text" id="titulo">
    <label for="subtitulo">${txh('ajustes.frase')}</label><input type="text" id="subtitulo">
    <label for="ref">${txh('asis.referencia')} <small>${txh('ajustes.referencia.d')}</small></label><select id="ref"></select>
    <label for="appsheet">${txh('asis.appsheet')} <small>${txh('ajustes.appsheet.d')}</small></label><input type="text" id="appsheet">
    <h3 style="margin:16px 0 0;font-size:15px">${txh('ajustes.contacto')} <small>${txh('ajustes.contacto.d')}</small></h3>
    <label for="c_nombre">${txh('ajustes.responsable')}</label><input type="text" id="c_nombre">
    <label for="c_correo">${txh('ajustes.correo')}</label><input type="text" id="c_correo" placeholder="proyecto@ejemplo.org">
    <label for="c_tel">${txh('ajustes.telefono')}</label><input type="text" id="c_tel">
    <label for="c_wa">${txh('ajustes.whatsapp')} <small>${txh('ajustes.whatsapp.d')}</small></label><input type="text" id="c_wa">
    <label for="c_hor">${txh('ajustes.horario')} <small>${op}</small></label><input type="text" id="c_hor">
    <label for="c_txt">${txh('ajustes.invitacion')}</label><textarea id="c_txt" rows="3" ${area} placeholder="${txh('ajustes.invitacion.ej')}"></textarea>
    <h3 style="margin:16px 0 0;font-size:15px">${txh('ajustes.acerca')}</h3>
    <label for="acerca">${txh('ajustes.acerca.d')}</label><textarea id="acerca" rows="4" ${area}></textarea>
    <label for="creditos">${txh('ajustes.creditos')} <small>${txh('ajustes.creditos.d')}</small></label><textarea id="creditos" rows="3" ${area}></textarea>
    <label for="licencia">${txh('ajustes.licencia')} <small>${txh('ajustes.licencia.d')}</small></label><input type="text" id="licencia">
    <label>${txh('ajustes.clave')}</label>
    <div id="clave"></div>
    <select id="accionClave" style="margin-top:6px">
      <option value="">${txh('ajustes.clave.igual')}</option><option value="nueva">${txh('ajustes.clave.nueva')}</option><option value="quitar">${txh('ajustes.clave.quitar')}</option>
    </select>
    <button class="prim" id="btn" onclick="guardar()">${txh('ajustes.guardar')}</button>
    <label>${txh('ajustes.carpetas')}</label>
    <div id="carpetas"></div>
    <button class="sec" onclick="carpetas()">${txh('ajustes.carpetas.crear')}</button>
    <div id="res" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    const fallo = e => { $('btn').disabled = false; $('res').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; };
    function pintar(a, msg){
      $('titulo').value = a.titulo; $('subtitulo').value = a.subtitulo; $('appsheet').value = a.appsheetId;
      const c = a.contacto || {};
      $('c_nombre').value = c.nombre || ''; $('c_correo').value = c.correo || ''; $('c_tel').value = c.telefono || '';
      $('c_wa').value = c.whatsapp || ''; $('c_hor').value = c.horario || ''; $('c_txt').value = c.texto || '';
      $('acerca').value = a.acerca || ''; $('creditos').value = a.creditos || ''; $('licencia').value = a.licencia || '';
      $('ref').innerHTML = '<option value="">' + th('comun.ninguna') + '</option>' + a.idiomas.map(i => '<option value="' + esc(i.id) + '">' + esc(i.nombre) + '</option>').join('');
      $('ref').value = a.idiomaPrincipal || '';
      $('clave').innerHTML = a.apiToken ? th('ajustes.clave.es', {clave: a.apiToken}).replace(esc(a.apiToken), '<code>' + esc(a.apiToken) + '</code>') : '<small>' + th('ajustes.clave.sin') + '</small>';
      $('accionClave').value = '';
      $('carpetas').innerHTML = a.carpetas.map(c => c.url ? '<a href="' + esc(c.url) + '" target="_blank">' + esc(c.nombre) + '</a>' : '<small>' + th('ajustes.carpeta.noEsta', {carpeta: c.nombre}) + '</small>').join(' · ');
      $('btn').disabled = false; $('res').innerHTML = msg ? '<div class="ok">✅ ' + esc(msg) + '</div>' : '';
    }
    function guardar(){
      $('btn').disabled = true;
      google.script.run.withFailureHandler(fallo).withSuccessHandler(a => pintar(a, tt('ajustes.guardado')))
        .guardarAjustes({ titulo: $('titulo').value, subtitulo: $('subtitulo').value, idiomaPrincipal: $('ref').value, appsheetId: $('appsheet').value, clave: $('accionClave').value,
          contacto: { nombre: $('c_nombre').value, correo: $('c_correo').value, telefono: $('c_tel').value, whatsapp: $('c_wa').value, horario: $('c_hor').value, texto: $('c_txt').value },
          acerca: $('acerca').value, creditos: $('creditos').value, licencia: $('licencia').value });
    }
    function carpetas(){ $('res').textContent = tt('ajustes.revisandoCarpetas');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(a => pintar(a, tt('ajustes.carpetasListas'))).crearCarpetasDesdePanel(); }
    google.script.run.withFailureHandler(fallo).withSuccessHandler(a => pintar(a)).obtenerAjustes();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(700), tx('ajustes.titulo'));
}

// ── Lenguas (panel) ──────────────────────────────────────────────────────────

function uiLenguas() {
  const op = txh('comun.opcional');
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:14px}
    label{display:block;font-weight:bold;margin-top:8px}
    input[type=text],select{width:100%;box-sizing:border-box;padding:8px;font-size:15px;border:1px solid #cbd5e1;border-radius:6px}
    .dos{display:flex;gap:8px}.dos>div{flex:1}
    small{color:#64748b;font-weight:normal}
    button{padding:9px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer}
    .prim{background:#7c3aed;color:#fff}.sec{background:#e2e8f0}
    table{width:100%;border-collapse:collapse;font-size:13px}td,th{border-bottom:1px solid #e2e8f0;padding:6px;text-align:left;vertical-align:top}
    .retirado{color:#94a3b8}
    .form{background:#faf5ff;border:1px solid #c4b5fd;border-radius:8px;padding:10px;margin-top:12px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:8px}
  </style>
  ${_tScript()}
  <div class="caja">
    <div id="lista">${txh('comun.cargando')}</div>
    <div class="form">
      <b id="titulo">${txh('lenguas.agregar')}</b>
      <input type="hidden" id="id">
      <div class="dos"><div><label>${txh('lengua.nombrePropio')} <small>${txh('lengua.nombrePropio.d')}</small></label><input type="text" id="nombre_propio"></div>
        <div><label>${txh('lengua.nombreOtra')} <small>${op}</small></label><input type="text" id="nombre_completo"></div></div>
      <div class="dos"><div><label>${txh('lengua.codigo')} <small>${txh('lengua.codigo.d')}</small></label><input type="text" id="codigo_iso" maxlength="3"></div>
        <div><label>${txh('lengua.codigoIntl')} <small>${txh('lengua.codigoIntl.d')}</small></label><input type="text" id="codigo_internacional" maxlength="3"></div>
        <div><label>${txh('lenguas.glotto')} <small>${txh('lenguas.glotto.d')}</small></label><input type="text" id="glottocode" maxlength="8"></div></div>
      <div class="dos"><div><label>${txh('lengua.variante')}</label><input type="text" id="variante"></div>
        <div><label>${txh('lenguas.region')}</label><input type="text" id="region"></div></div>
      <label>${txh('lenguas.letras')} <small>${txh('lenguas.letras.d')}</small></label>
      <div style="display:flex;gap:6px"><input type="text" id="alfabeto" style="flex:1"><button class="sec" type="button" onclick="sugerir()" title="${txh('lenguas.sugerir.d')}">${txh('lenguas.sugerir')}</button></div>
      <label>${txh('lenguas.alfabetoFuente')} <small>${txh('lenguas.alfabetoFuente.d')}</small></label><select id="id_fuente_alfabeto"><option value="">—</option></select>
      <label>${txh('lenguas.atajosNuevos')} <small>${txh('lenguas.atajosNuevos.d')}</small></label><input type="text" id="atajos">
      <label>${txh('lenguas.notas')}</label><input type="text" id="notas">
      <label>${txh('lenguas.estado')}</label>
      <select id="estado"><option value="publicado">${txh('lenguas.estado.publicado')}</option><option value="borrador">${txh('lenguas.estado.borrador')}</option><option value="revisión">${txh('lenguas.estado.revision')}</option><option value="retirado">${txh('lenguas.estado.retirado')}</option></select>
      <div style="margin-top:10px;display:flex;gap:8px">
        <button class="prim" id="btn" onclick="guardar()">${txh('lenguas.guardar')}</button>
        <button class="sec" onclick="limpiar()">${txh('lenguas.nueva')}</button>
      </div>
      <div id="res" role="status" aria-live="polite"></div>
    </div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    const campos = ['id','nombre_propio','nombre_completo','codigo_iso','codigo_internacional','glottocode','variante','region','alfabeto','id_fuente_alfabeto','notas','atajos','estado'];
    const ESTADO = { publicado: '', borrador: tt('lenguas.estado.borrador'), 'revisión': tt('lenguas.estado.revision'), retirado: tt('lenguas.estado.retirado') };
    function sugerir(){
      const id = $('id').value;
      if (!id) { $('res').innerHTML = '<div class="error">' + th('lenguas.primeroGuardar') + '</div>'; return; }
      $('res').textContent = tt('lenguas.revisandoPalabras');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(a => {
        $('alfabeto').value = a; $('res').innerHTML = '<div class="ok">' + th('lenguas.revisaPropuesta') + '</div>';
      }).sugerirAlfabeto(id);
    }
    let datos = [];
    const fallo = e => { $('btn').disabled = false; $('res').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; };
    function pintar(l, msg){
      datos = l; $('btn').disabled = false;
      $('lista').innerHTML = !l.length ? '<p><small>' + th('lenguas.ninguna') + '</small></p>' :
        '<table><tr><th>' + th('lenguas.col.lengua') + '</th><th>' + th('lenguas.col.palabras') + '</th><th>' + th('lenguas.col.audios') + '</th><th>' + th('lenguas.col.carpetas') + '</th><th></th></tr>' + l.map((x, i) =>
          '<tr class="' + (x.estado !== 'publicado' ? 'retirado' : '') + '"><td><b>' + esc(x.nombre_propio || x.nombre_completo) + '</b>' +
          (x.nombre_propio && x.nombre_completo ? '<br><small>' + esc(x.nombre_completo) + '</small>' : '') +
          '<br><small>' + esc([x.codigo_iso, x.variante, ESTADO[x.estado] || ''].filter(Boolean).join(' · ')) + '</small></td>' +
          '<td>' + x.textos + '</td><td>' + x.audios + '</td><td>' +
          (x.urlAudios ? '<a href="' + esc(x.urlAudios) + '" target="_blank">' + th('lenguas.audios') + '</a> ' : '') + (x.urlVideos ? '<a href="' + esc(x.urlVideos) + '" target="_blank">' + th('lenguas.videos') + '</a>' : '') +
          '</td><td><button class="sec" onclick="editar(' + i + ')">' + th('lenguas.editarBoton') + '</button></td></tr>').join('') + '</table>';
      if (msg) { $('res').innerHTML = '<div class="ok">✅ ' + esc(msg) + '</div>'; limpiar(true); }
    }
    function editar(i){ const x = datos[i]; campos.forEach(c => { if ($(c)) $(c).value = x[c] || ''; }); $('atajos').value = ''; $('estado').value = x.estado || 'publicado'; $('titulo').textContent = tt('lenguas.editar', {lengua: x.nombre_propio || x.nombre_completo}); }
    function limpiar(conservarMsg){ campos.forEach(c => { if ($(c)) $(c).value = ''; }); $('estado').value = 'publicado'; $('titulo').textContent = tt('lenguas.agregar'); if (!conservarMsg) $('res').innerHTML = ''; }
    function guardar(){
      const d = {}; campos.forEach(c => d[c] = $(c).value);
      $('btn').disabled = true; $('res').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(l => pintar(l, tt('lenguas.guardada'))).guardarIdioma(d);
    }
    google.script.run.withSuccessHandler(l => { $('id_fuente_alfabeto').innerHTML = '<option value="">—</option>' + l.map(f => '<option value="' + esc(f.id) + '">' + esc(f.titulo) + '</option>').join(''); }).fuentesNormas();
    google.script.run.withFailureHandler(fallo).withSuccessHandler(l => pintar(l)).listarIdiomas();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(680), tx('lenguas.titulo'));
}

/** Fuentes que pueden servir de alfabeto de referencia (normas oficiales, gramáticas). */
function fuentesNormas() {
  return readSheet(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.FUENTES).filter(f => _s(f.id_fuente) && /alfabeto|norma|gram/i.test(_sinAcentos(f.tipo)))
    .map(f => ({ id: _s(f.id_fuente), titulo: _s(f.titulo) }));
}

// ── Hojas técnicas y panel lateral ───────────────────────────────────────────

function _hojasTecnicas(ss, ocultar) {
  _esquema().filter(e => e.tecnica).forEach(e => {
    const sh = ss.getSheetByName(e.hoja);
    if (sh) ocultar ? sh.hideSheet() : sh.showSheet();
  });
  const tx_ = ss.getSheetByName(CONFIG.SHEETS.TEXTOS);
  if (tx_) ocultar ? tx_.hideSheet() : tx_.showSheet();
}

function uiAlternarHojasTecnicas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const ocultar = sh ? !sh.isSheetHidden() : true;
  _hojasTecnicas(ss, ocultar);
  return tx(ocultar ? 'panel.hojas.ocultas' : 'panel.hojas.visibles');
}

function panelEstado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(CONFIG.SHEETS.CONFIG)) return { configurado: false };
  _cfgMemo = null;
  const cfg = getAppConfig(ss);
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES), tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  const m = _leerManifiesto();
  return {
    configurado: !!cfg.titulo, titulo: cfg.titulo || tx('panel.sinNombre'),
    lenguas: listarIdiomas().filter(i => i.estado === 'publicado').length,
    palabras: tPal ? tPal.filas.filter(f => _s(_get(tPal, f, 'id_palabra'))).length : 0,
    porEstado: (() => {
      const c = { publicado: 0, 'revisión': 0, borrador: 0, retirado: 0 };
      if (tPal) tPal.filas.forEach(f => {
        if (!_s(_get(tPal, f, 'id_palabra'))) return;
        const e = _estadoCanonico(_get(tPal, f, 'estado')) || 'publicado';
        if (e in c) c[e]++;
      });
      return c;
    })(),
    traducciones: tTr ? tTr.filas.length : 0,
    modoCaptura: _col(tTr, 'texto_captura') > -1,
    clave: !!cfg.apiToken,
    sinDireccion: !!cfg.titulo && !cfg.urlApi,
    propuestas: (() => { try { const pa = _propuestasAjustes(ss); return { activo: !!pa.activo, pendientes: ss.getSheetByName(CONFIG.SHEETS.PROPUESTAS) ? readSheet(ss, CONFIG.SHEETS.PROPUESTAS).filter(x => _s(x.estado) === 'Pendiente').length : 0 }; } catch (e) { return null; } })(),
    sinCorreo: !_s((cfg.contacto || {}).correo),
    drive: _espacioDrive(),
    publicado: m ? Utilities.formatDate(new Date(m.generado), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : ''
  };
}

function uiPanel() {
  const b = (fn, k, d) => `<button onclick="abrir('${fn}')">${txh(k)}${d ? '<br><small>' + txh(d) + '</small>' : ''}</button>`;
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0;padding:12px}
    h3{margin:0 0 4px}small{color:#64748b}
    .estado{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin:10px 0;line-height:1.6}
    button{display:block;width:100%;text-align:left;padding:11px 12px;margin:6px 0;border:1px solid #cbd5e1;background:#fff;border-radius:8px;font-size:14px;cursor:pointer}
    button:hover,button:focus{background:#eef2ff;outline:2px solid #6366f1}
    .grupo{font-weight:bold;margin-top:14px;color:#475569;font-size:12px;text-transform:uppercase;letter-spacing:.04em}
    #msg{margin-top:8px}
  </style>
  ${_tScript()}
  <h3 id="t">${txh('panel.catalogo')}</h3>
  <div class="estado" id="estado" role="status">${txh('comun.cargando')}</div>
  <div id="falta"></div>
  ${b('uiEstadoCatalogo', 'panel.comoVa', 'panel.comoVa.d')}
  <div class="grupo">${txh('panel.grupo.diario')}</div>
  ${b('uiEditar', 'panel.editar', 'panel.editar.d')}
  ${b('uiPropuestas', 'panel.propuestas', 'panel.propuestas.d')}
  ${b('uiMedios', 'panel.medios', 'panel.medios.d')}
  ${b('uiGenerarReporteAvanzado', 'panel.falta')}
  ${b('uiEstados', 'panel.estados', 'panel.estados.d')}
  ${b('uiPublicarCatalogo', 'panel.publicar', 'panel.publicar.d')}
  <div class="grupo">${txh('panel.grupo.colaborar')}</div>
  ${b('uiConfigurarPlantillaInteligente', 'panel.hojaAyuda')}
  ${b('uiMediosIlustrar', 'panel.hojaIlustrar')}
  ${b('uiImportarDatos', 'panel.traer')}
  ${b('uiIntercambio', 'panel.comunidades')}
  ${b('uiAvisos', 'panel.avisos')}
  ${b('uiVerEnlaceApi', 'panel.direccion')}
  <div class="grupo">${txh('panel.grupo.configurar')}</div>
  ${b('uiLenguas', 'panel.lenguas', 'panel.lenguas.d')}
  ${b('uiEquipo', 'panel.equipo', 'panel.equipo.d')}
  ${b('uiVoces', 'panel.voces')}
  ${b('uiFuentes', 'panel.fuentes')}
  ${b('uiReemplazos', 'panel.atajos', 'panel.atajos.d')}
  ${b('uiTiposRelacion', 'panel.tiposRel', 'panel.tiposRel.d')}
  ${b('uiAjustes', 'panel.datos', 'panel.datos.d')}
  ${b('uiAsistenteInicial', 'panel.primeros')}
  <div class="grupo">${txh('panel.grupo.mantenimiento')}</div>
  ${b('uiCoincidencias', 'panel.coincidencias')}
  ${b('uiRamas', 'panel.ramas')}
  ${b('servicioInstalacionCarpetas', 'panel.reparar', 'panel.reparar.d')}
  ${b('uiExportarDatos', 'panel.copia')}
  ${b('uiLimpiar', 'panel.limpiar', 'panel.limpiar.d')}
  <div class="grupo"></div>
  ${b('uiIdiomaMenus', 'panel.idioma', 'panel.idioma.d')}
  <button onclick="alternar()">${txh('panel.hojasInternas')}</button>
  <div id="msg" role="status" aria-live="polite"></div>
  <p style="margin-top:16px;font-size:11px;color:#94a3b8;text-align:center">${txh('panel.credito', { autor: AUTOR })}</p>
  <script>
    const $ = id => document.getElementById(id);
    function abrir(fn){ $('msg').textContent = tt('comun.abriendo'); google.script.run.withFailureHandler(e => $('msg').textContent = tt('comun.error', {m: e.message})).withSuccessHandler(() => { $('msg').textContent = ''; estado(); })[fn](); }
    function alternar(){ google.script.run.withSuccessHandler(m => $('msg').textContent = m).uiAlternarHojasTecnicas(); }
    function estado(){
      google.script.run.withFailureHandler(e => $('estado').textContent = tt('comun.error', {m: e.message})).withSuccessHandler(s => {
        if (!s.configurado) { $('t').textContent = tt('panel.nuevo.titulo'); $('estado').innerHTML = th('panel.nuevo.texto'); return; }
        $('t').textContent = s.titulo;
        const e = s.porEstado;
        $('estado').innerHTML = th('panel.resumen.cuentas', {lenguas: s.lenguas, palabras: s.palabras, traducciones: s.traducciones}) + '<br>' +
          th('panel.resumen.estados', {publicadas: e.publicado, revision: e['revisión'], borrador: e.borrador}) +
          (e.retirado ? th('panel.resumen.retiradas', {n: e.retirado}) : '') + '<br>' +
          th('panel.resumen.publicado', {fecha: s.publicado || tt('panel.resumen.nunca')}) + '<br>' + th(s.clave ? 'panel.resumen.conClave' : 'panel.resumen.sinClave') +
          (s.modoCaptura ? '<br>' + th('panel.resumen.captura') : '');
        $('falta').innerHTML = s.sinDireccion ? '<div class="estado" style="background:#fffbeb;border-color:#fde68a"><b>' + th('panel.sinDireccion.t') + '</b><ol style="margin:4px 0 0 18px;padding:0">' +
          ['panel.sinDireccion.1', 'panel.sinDireccion.2', 'panel.sinDireccion.3'].map(k => '<li>' + th(k) + '</li>').join('') + '</ol>' +
          '<button onclick="abrir(\\'uiVerEnlaceApi\\')">' + th('panel.direccion') + '</button></div>' : '';
        const pr = s.propuestas, av = [];
        if (pr && pr.pendientes) av.push('<button onclick="abrir(\\'uiPropuestas\\')">📥 ' + th('panel.propuestasPend', {n: pr.pendientes}) + '</button>');
        if (s.sinCorreo) av.push('<div>✉️ ' + th('panel.sinCorreo') + '</div>');
        if (s.drive && s.drive.pct >= 70) av.push('<div>💾 ' + th('panel.drive', {pct: s.drive.pct}) + '</div>');
        if (av.length) $('falta').innerHTML += '<div class="estado" style="background:#f0fdfa;border-color:#99f6e4">' + av.join('') + '</div>';
      }).panelEstado();
    }
    estado();
  </script>`;
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle(tx('panel.titulo')));
}

// ── Estados de las palabras (gestor tipo WooCommerce) ────────────────────────
// Solo para administración: el público nunca ve estados; solo ve lo "Publicado".

/** Cambia la columna "estado" de las filas que cumplen el filtro. Escribe en bloque si son muchas. */
function _cambiarEstadoFilas(t, filtro, etiqueta) {
  const c = _col(t, 'estado');
  if (c < 0) return 0;
  const cambiar = [];
  t.filas.forEach((f, i) => { if (filtro(f) && _s(f[c]) !== etiqueta) cambiar.push(i); });
  if (cambiar.length > 50) {
    cambiar.forEach(i => { t.filas[i][c] = etiqueta; });
    t.sh.getRange(2, c + 1, t.filas.length, 1).setValues(t.filas.map(f => [f[c]]));
  } else {
    cambiar.forEach(i => _actualizar(t, i, { estado: etiqueta }));
  }
  return cambiar.length;
}

function _tablaConEstado(ss, hoja) {
  const sh = ss.getSheetByName(hoja);
  if (!sh) return null;
  if (_asegurarColumnas(sh, ['estado']).length) _aplicarValidaciones(ss);
  return _tabla(ss, hoja);
}

function estadosListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  const cfg = getAppConfig(ss);
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  _requerir(tPal, ['id_palabra'], CONFIG.SHEETS.PALABRAS);
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES), tCat = _tabla(ss, CONFIG.SHEETS.CATEGORIAS), tId = _tabla(ss, CONFIG.SHEETS.IDIOMAS);

  const catNom = new Map();
  if (tCat) tCat.filas.forEach(f => catNom.set(_k(_get(tCat, f, 'id_categoria')), _s(_get(tCat, f, 'nombre_categoria'))));
  const idiomaNom = new Map();
  if (tId) tId.filas.forEach(f => idiomaNom.set(_k(_get(tId, f, 'id_idioma')), _s(_get(tId, f, 'nombre_propio') || _get(tId, f, 'nombre_completo') || _get(tId, f, 'nombre'))));

  // Etiqueta de cada palabra: su texto en la lengua de referencia o, si no hay, el primero que exista
  const ref = _k(cfg.idiomaPrincipal), info = new Map();
  if (tTr) tTr.filas.forEach(f => {
    const id = _k(_get(tTr, f, 'id_palabra'));
    const o = info.get(id) || { ref: '', otra: '', n: 0, sinPublicar: 0 };
    const txt = _s(_get(tTr, f, 'texto')), lang = _k(_get(tTr, f, 'id_idioma'));
    if (txt) {
      o.n++;
      if (ref && lang === ref && !o.ref) o.ref = txt;
      else if (!o.otra) o.otra = txt + (idiomaNom.get(lang) ? ` (${idiomaNom.get(lang)})` : '');
    }
    if (!_esPublico({ estado: _get(tTr, f, 'estado') })) o.sinPublicar++;
    info.set(id, o);
  });

  return tPal.filas.map(f => {
    const id = _s(_get(tPal, f, 'id_palabra'));
    if (!id) return null;
    const o = info.get(_k(id)) || {};
    return {
      id, etiqueta: o.ref || o.otra || tx('estados.sinTexto'),
      categoria: catNom.get(_k(_get(tPal, f, 'id_categoria'))) || '',
      estado: _estadoCanonico(_get(tPal, f, 'estado')) || 'publicado',
      textos: o.n || 0, sinPublicar: o.sinPublicar || 0,
      imagen: _s(_get(tPal, f, 'imagen_referencia')).length > 5,
      juegos: _k(_get(tPal, f, 'usar_en_juegos')) !== 'no'
    };
  }).filter(Boolean);
}

/** Cambia el estado de una o varias palabras (y, si se pide, de todas sus traducciones). */
function estadosCambiar(ids, estado, incluirTraducciones) {
  const canon = _estadoCanonico(estado);
  if (!ESTADOS.some(e => e.canon === canon)) throw new Error(tx('estados.error.estado'));
  if (!ids || !ids.length) throw new Error(tx('estados.error.ninguna'));
  const etiqueta = _etiquetaEstado(canon);
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sel = new Set(ids.map(_k));
    const tPal = _tablaConEstado(ss, CONFIG.SHEETS.PALABRAS);
    const n = _cambiarEstadoFilas(tPal, f => sel.has(_k(_get(tPal, f, 'id_palabra'))), etiqueta);
    let nt = 0;
    if (incluirTraducciones) {
      const tTr = _tablaConEstado(ss, CONFIG.SHEETS.TRADUCCIONES);
      if (tTr) nt = _cambiarEstadoFilas(tTr, f => sel.has(_k(_get(tTr, f, 'id_palabra'))), etiqueta);
    }
    publicarCatalogo(ss);
    return {
      mensaje: tx(incluirTraducciones ? 'estados.listoTrads' : 'estados.listo', { n, nt, estado: tx('estado.' + canon) }),
      lista: estadosListar()
    };
  } finally { lock.releaseLock(); }
}

/** Marca si una o varias palabras pueden aparecer en juegos (Sí/No). */
function estadosJuegos(ids, usar) {
  if (!ids || !ids.length) throw new Error(tx('estados.error.ninguna'));
  const lock = _lock();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG.SHEETS.PALABRAS);
    if (_asegurarColumnas(sh, ['usar_en_juegos']).length) _aplicarValidaciones(ss);
    const t = _tabla(ss, CONFIG.SHEETS.PALABRAS);
    const sel = new Set(ids.map(_k)), valor = usar ? 'Sí' : 'No';
    let n = 0;
    t.filas.forEach((f, i) => { if (sel.has(_k(_get(t, f, 'id_palabra'))) && _s(_get(t, f, 'usar_en_juegos')) !== valor) { _actualizar(t, i, { usar_en_juegos: valor }); n++; } });
    publicarCatalogo(ss);
    return { mensaje: tx(usar ? 'estados.juegosAhora' : 'estados.juegosYaNo', { n }), lista: estadosListar() };
  } finally { lock.releaseLock(); }
}

function uiEstados() {
  const opciones = ESTADOS.map(e => `<option value="${_esc(e.canon)}">${txh('estado.' + e.canon)}</option>`).join('');
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:12px}
    .tabs{display:flex;flex-wrap:wrap;gap:4px 12px;margin-bottom:8px}
    .tabs a{cursor:pointer;color:#2563eb;text-decoration:none}.tabs a.activa{color:#111;font-weight:bold}
    .barra{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:8px}
    input[type=search],select{padding:7px;font-size:14px;border:1px solid #cbd5e1;border-radius:6px}
    input[type=search]{flex:1;min-width:160px}
    button{padding:8px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#1e3a5f;color:#fff}
    .sec{background:#e2e8f0;color:#1f2937}
    table{width:100%;border-collapse:collapse;font-size:13px}
    th,td{border-bottom:1px solid #e2e8f0;padding:6px;text-align:left;vertical-align:middle}
    th{background:#f8fafc;position:sticky;top:0}
    .scroll{max-height:430px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:6px}
    small{color:#64748b}
    .e-borrador{border-left:4px solid #94a3b8}.e-revisión{border-left:4px solid #f59e0b}.e-publicado{border-left:4px solid #16a34a}.e-retirado{border-left:4px solid #dc2626}
    #msg{margin-top:8px;min-height:20px}
    .ok{background:#dcfce7;color:#166534;padding:6px 8px;border-radius:6px}
    .error{background:#fee2e2;color:#b91c1c;padding:6px 8px;border-radius:6px}
  </style>
  ${_tScript()}
  <div class="caja">
    <p style="margin:0 0 8px">${txh('estados.intro')}</p>
    <div class="tabs" id="tabs" role="tablist"></div>
    <div class="barra">
      <input type="search" id="buscar" placeholder="${txh('estados.buscar')}" oninput="pintar()" aria-label="Buscar">
    </div>
    <div class="barra">
      <label for="lote"><b>${txh('estados.conMarcadas')}</b></label>
      <select id="lote"><option value="">${txh('estados.elegir')}</option>${opciones}<option value="juegos:si">${txh('estados.juegosSi')}</option><option value="juegos:no">${txh('estados.juegosNo')}</option></select>
      <label><input type="checkbox" id="trads" checked> ${txh('estados.tambienTrads')}</label>
      <button onclick="lote()">${txh('estados.aplicar')}</button>
    </div>
    <div class="scroll">
      <table>
        <thead><tr><th><input type="checkbox" id="todas" onclick="marcarTodas(this.checked)" aria-label="${txh('estados.selTodas')}"></th>
          <th>${txh('estados.col.palabra')}</th><th>${txh('estados.col.categoria')}</th><th>${txh('estados.col.traducciones')}</th><th>${txh('estados.col.juegos')}</th><th>${txh('estados.col.estado')}</th></tr></thead>
        <tbody id="filas"><tr><td colspan="6">${txh('comun.cargando')}</td></tr></tbody>
      </table>
    </div>
    <div style="margin-top:6px"><small id="cuenta"></small> <button class="sec" id="mas" style="display:none" onclick="limite+=300;pintar()">${txh('estados.mas')}</button></div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    const ESTADOS = ${JSON.stringify(ESTADOS.map(e => ({ canon: e.canon, etiqueta: tx('estado.' + e.canon) }))).replace(/</g, '\\u003c')};
    let datos = [], filtro = 'todos', limite = 300, marcadas = new Set();
    function ok(m){ $('msg').innerHTML = '<div class="ok">✅ ' + esc(m) + '</div>'; }
    function fallo(e){ $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; }
    function tabs(){
      const n = c => datos.filter(d => c === 'todos' || d.estado === c).length;
      $('tabs').innerHTML = [{canon:'todos', etiqueta: tt('estado.todas')}].concat(ESTADOS).map(e =>
        '<a role="tab" class="' + (filtro === e.canon ? 'activa' : '') + '" onclick="filtro=\\'' + e.canon + '\\';limite=300;pintar()">' + esc(e.etiqueta) + ' (' + n(e.canon) + ')</a>').join(' | ');
    }
    function visibles(){
      const q = $('buscar').value.toLowerCase().trim();
      return datos.filter(d => (filtro === 'todos' || d.estado === filtro) && (!q || (d.etiqueta + ' ' + d.categoria).toLowerCase().includes(q)));
    }
    function pintar(){
      tabs();
      const v = visibles();
      $('filas').innerHTML = v.slice(0, limite).map(d =>
        '<tr class="e-' + d.estado + '"><td><input type="checkbox" ' + (marcadas.has(d.id) ? 'checked' : '') + ' onclick="marcar(\\'' + d.id + '\\',this.checked)" aria-label="' + th('estados.sel') + '"></td>' +
        '<td>' + esc(d.etiqueta) + (d.imagen ? ' 🖼️' : '') + (d.sinPublicar ? ' <small>' + th('estados.noSeVen', {n: d.sinPublicar}) + '</small>' : '') + '</td>' +
        '<td>' + esc(d.categoria) + '</td><td>' + d.textos + '</td>' +
        '<td><input type="checkbox" ' + (d.juegos ? 'checked' : '') + ' title="' + th('estados.col.juegos') + '" aria-label="' + th('estados.col.juegos') + '" onchange="juegos(\\'' + d.id + '\\',this.checked)"></td>' +
        '<td><select aria-label="' + th('estados.col.estado') + '" onchange="uno(\\'' + d.id + '\\',this.value)">' +
          ESTADOS.map(e => '<option value="' + e.canon + '"' + (e.canon === d.estado ? ' selected' : '') + '>' + esc(e.etiqueta) + '</option>').join('') +
        '</select></td></tr>').join('') || '<tr><td colspan="6"><small>' + th('estados.vacia') + '</small></td></tr>';
      $('cuenta').textContent = tt('estados.cuenta', {n: Math.min(v.length, limite), total: v.length, sel: marcadas.size});
      $('mas').style.display = v.length > limite ? 'inline-block' : 'none';
    }
    function marcar(id, si){ si ? marcadas.add(id) : marcadas.delete(id); pintar(); }
    function marcarTodas(si){ visibles().forEach(d => si ? marcadas.add(d.id) : marcadas.delete(d.id)); pintar(); }   // todas las de la pestaña y búsqueda, también las que aún no se muestran
    function aplicar(ids, estado, trads){
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(r => { datos = r.lista; marcadas.clear(); $('todas').checked = false; pintar(); ok(r.mensaje); })
        .estadosCambiar(ids, estado, trads);
    }
    function uno(id, estado){ aplicar([id], estado, $('trads').checked); }
    function juegos(ids, si){
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(fallo).withSuccessHandler(r => { datos = r.lista; marcadas.clear(); pintar(); ok(r.mensaje); })
        .estadosJuegos(Array.isArray(ids) ? ids : [ids], si);
    }
    function lote(){
      const e = $('lote').value;
      if (!e) return fallo(tt('estados.error.elegir'));
      if (!marcadas.size) return fallo(tt('estados.error.ninguna'));
      if (e.indexOf('juegos:') === 0) return juegos(Array.from(marcadas), e === 'juegos:si');
      aplicar(Array.from(marcadas), e, $('trads').checked);
    }
    google.script.run.withFailureHandler(fallo).withSuccessHandler(l => { datos = l; pintar(); }).estadosListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(780).setHeight(680), tx('estados.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 11. LOTERÍA EN LÍNEA (SINCRONIZACIÓN LIGERA)
// ─────────────────────────────────────────────────────────────────────────────
// Cada partida vive 6 horas en la caché del script. Solo se guarda lo mínimo:
// la configuración (palabras, lengua, variantes), cuántas cartas han salido,
// quién se unió y quién ganó. El orden de las cartas y las tablas NO viajan por
// la red: cada dispositivo los calcula igual a partir de la semilla de la partida.

const LOT = { TTL: 21600, MAX_CONFIG: 20000, MAX_JUGADORES: 300 };
const LOT_VARIANTES = ['llena', 'linea', 'esquinas', 'centro'];

/** Las partidas se identifican como "Nombre + 4 números" (ej. Juguemos 4320). Los números no dependen del alfabeto de ninguna lengua. */
function _lotSlug(nombre) { return _sinAcentos(nombre).replace(/[^a-z0-9]+/g, '').slice(0, 20); }
function _lotClave(nombre, numero) {
  const slug = _lotSlug(nombre), num = _s(numero).replace(/\D/g, '');
  if (!slug || num.length !== 4) return '';
  return `LOT_${slug}_${num}`;
}

function _lotLeer(nombre, numero) {
  const k = _lotClave(nombre, numero);
  if (!k) return null;
  const v = CacheService.getScriptCache().getAll([k + '_cfg', k + '_est']);
  if (!v[k + '_cfg'] || !v[k + '_est']) return null;
  return { k, cfg: JSON.parse(v[k + '_cfg']), est: JSON.parse(v[k + '_est']) };
}

/** Guarda ambas partes: así cada movimiento renueva las 6 horas de vida de la partida. */
function _lotGuardar(k, cfg, est) {
  est.v = (est.v || 0) + 1;
  est.actualizado = Date.now();
  CacheService.getScriptCache().putAll({ [k + '_cfg']: JSON.stringify(cfg), [k + '_est']: JSON.stringify(est) }, LOT.TTL);
}

function _lotPublica(cfg) { const c = Object.assign({}, cfg); delete c.secreto; return c; }
function _cuatroDigitos() { return String(Math.floor(Math.random() * 10000)).padStart(4, '0'); }

function _lotValidarConfig(cfg) {
  if (!cfg || typeof cfg !== 'object') throw new Error('Configuración vacía.');
  if (JSON.stringify(cfg).length > LOT.MAX_CONFIG) throw new Error('La configuración es demasiado grande.');
  if (!Array.isArray(cfg.ids) || cfg.ids.length < 32 || cfg.ids.length > 120) throw new Error('La lotería necesita entre 32 y 120 cartas.');
  if (!_s(cfg.semilla) || !_s(cfg.lengua)) throw new Error('Faltan la semilla o la lengua de la partida.');
  if (!_lotSlug(cfg.nombre)) throw new Error('La partida necesita un nombre (con letras o números).');
  cfg.nombre = _s(cfg.nombre).slice(0, 20);
  cfg.variantes = (cfg.variantes || []).filter(v => LOT_VARIANTES.includes(v));
  if (!cfg.variantes.length) cfg.variantes = ['llena'];
  return cfg;
}

/** Acciones de lotería por GET (lectura). */
function _lotGet(action, p) {
  const r = _lotLeer(p.nombre, p.numero);
  if (!r) return { status: 404, error: 'No existe esa partida o ya terminó su tiempo (6 horas). Revisa el nombre y los 4 números.' };
  if (action === 'lot_info') return { status: 200, nombre: r.cfg.nombre, numero: r.k.slice(-4), config: _lotPublica(r.cfg), estado: r.est };
  return { status: 200, estado: r.est };
}

/** Acciones de lotería por POST (escritura). */
function _lotPost(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (d.accion === 'lot_crear') {
      if (!_s(d.secreto) || _s(d.secreto).length < 12) throw new Error('Falta la clave del anfitrión.');
      const cfg = _lotValidarConfig(d.config);
      cfg.secreto = _s(d.secreto);
      let numero = '';
      for (let i = 0; i < 30 && !numero; i++) { const n = _cuatroDigitos(); if (!_lotLeer(cfg.nombre, n)) numero = n; }
      if (!numero) throw new Error('Hay demasiadas partidas con ese nombre. Usa otro nombre.');
      const ronda = /^\d{4}$/.test(_s(d.ronda)) ? _s(d.ronda) : _cuatroDigitos();
      _lotGuardar(_lotClave(cfg.nombre, numero), cfg, { n: 0, ronda, fin: false, ganadores: [], jugadores: [], creado: Date.now() });
      return { status: 200, nombre: cfg.nombre, numero, ronda };
    }

    const r = _lotLeer(d.nombre, d.numero);
    if (!r) return { status: 404, error: 'No existe esa partida o ya terminó su tiempo (6 horas). Revisa el nombre y los 4 números.' };
    const { k, cfg, est } = r;

    if (d.accion === 'lot_unirse') {
      if (est.jugadores.length >= LOT.MAX_JUGADORES) throw new Error('La partida ya está llena.');
      // Los números de tabla empiezan después de las tablas impresas o descargadas: así no se repiten
      const tabla = (Number(cfg.impresas) || 0) + est.jugadores.length + 1;
      const nombre = _s(d.jugador).slice(0, 40) || `Jugador ${tabla}`;
      est.jugadores.push({ tabla, nombre });
      _lotGuardar(k, cfg, est);
      return { status: 200, tabla, jugador: nombre, numero: k.slice(-4), config: _lotPublica(cfg), estado: est };
    }

    if (d.accion === 'lot_actualizar') {
      if (_s(d.secreto) !== cfg.secreto) return { status: 403, error: 'Solo quien creó la partida puede cantar.' };
      if (/^\d{4}$/.test(_s(d.ronda)) && _s(d.ronda) !== est.ronda) {
        // Nueva ronda: mismas tablas, nuevo orden, sin ganadores
        est.ronda = _s(d.ronda); est.n = 0; est.ganadores = []; est.fin = false;
      }
      if (d.n != null) est.n = Math.max(0, Math.min(cfg.ids.length, Math.floor(Number(d.n)) || 0));
      if (d.fin != null) est.fin = !!d.fin;
      _lotGuardar(k, cfg, est);
      return { status: 200, estado: est };
    }

    if (d.accion === 'lot_ganar') {
      const variante = _s(d.variante);
      if (!cfg.variantes.includes(variante)) return { status: 400, error: 'Esa forma de ganar no está en juego.' };
      if (_s(d.ronda) && _s(d.ronda) !== est.ronda) return { status: 409, error: 'Esa ronda ya terminó.', estado: est };
      if (est.fin || est.ganadores.some(g => g.variante === variante)) return { status: 200, nuevo: false, estado: est };
      est.ganadores.push({ variante, tabla: Number(d.tabla) || 0, nombre: _s(d.jugador).slice(0, 40) || 'Alguien', carta: est.n, t: Date.now() });
      if (cfg.variantes.every(v => est.ganadores.some(g => g.variante === v))) est.fin = true;
      _lotGuardar(k, cfg, est);
      return { status: 200, nuevo: true, estado: est };
    }

    return { status: 400, error: 'Acción de lotería no reconocida.' };
  } finally { lock.releaseLock(); }
}

// ── Paquete de imágenes ligeras para jugar sin internet ──────────────────────
// Devuelve las imágenes de palabras PUBLICADAS como datos incrustables (base64), en tamaño reducido,
// para que el visor arme un archivo que funciona sin conexión. Si el tiempo no alcanza, devuelve
// las que pudo y la lista "faltan" para pedirlas en una segunda llamada.
function _paqueteImagenes(ids, ancho) {
  if (ids.length > 120) return { status: 400, error: 'Máximo 120 imágenes por paquete.' };
  ancho = Math.max(120, Math.min(1000, Math.round(ancho / 50) * 50));
  const o = JSON.parse(_leerSnapshot());
  const porId = {};
  (o.data.palabras || []).forEach(p => { if (p.imagen_id) porId[p.id_palabra] = p.imagen_id; });
  const cache = CacheService.getScriptCache(), inicio = Date.now(), imagenes = {}, faltan = [];
  const token = ScriptApp.getOAuthToken();
  ids.forEach(id => {
    const archivo = porId[id];
    if (!archivo) return;                       // no publicada o sin imagen de Drive
    const clave = `IMG_${archivo}_${ancho}`;
    const enCache = cache.get(clave);
    if (enCache) { imagenes[id] = enCache; return; }
    if (Date.now() - inicio > 200000) { faltan.push(id); return; }
    let datos = '';
    try {
      const r = UrlFetchApp.fetch(`https://drive.google.com/thumbnail?id=${archivo}&sz=w${ancho}`,
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true, followRedirects: true });
      const tipo = String(r.getHeaders()['Content-Type'] || r.getHeaders()['content-type'] || '');
      if (r.getResponseCode() === 200 && tipo.indexOf('image') === 0) datos = `data:${tipo.split(';')[0]};base64,${Utilities.base64Encode(r.getContent())}`;
    } catch (e) { /* se intenta con la miniatura de Drive */ }
    if (!datos) {
      try {
        const f = DriveApp.getFileById(archivo);
        const b = f.getThumbnail() || (f.getSize() < 400 * 1024 ? f.getBlob() : null);
        if (b) datos = `data:${b.getContentType() || 'image/jpeg'};base64,${Utilities.base64Encode(b.getBytes())}`;
      } catch (e) { /* sin acceso */ }
    }
    if (datos) { imagenes[id] = datos; if (datos.length < 95000) { try { cache.put(clave, datos, 21600); } catch (e) {} } }
  });
  return { status: 200, ancho, imagenes, faltan };
}

/** Entrada POST de la aplicación web (por ahora solo la usa la lotería). */
function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const rol = _rolToken(d.token);
    const enMant = _respuestaMantenimiento(rol || 'publico');
    if (enMant) return responseJSON(enMant);
    if (!rol) return responseJSON({ status: 401, error: 'Unauthorized: Token inválido o ausente.' });
    if (String(d.accion || '').indexOf('lot_') === 0) return responseJSON(_lotPost(d));
    if (d.accion === 'proponer') return responseJSON(_recibirPropuestas(d));
    return responseJSON({ status: 400, error: 'Acción no reconocida.' });
  } catch (err) {
    return responseJSON({ status: 500, error: err.message });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// 12. ESTADO DEL CATÁLOGO (¿QUÉ JUEGOS ESTÁN LISTOS EN CADA LENGUA?)
// ─────────────────────────────────────────────────────────────────────────────
// Los juegos dependen de la calidad de los datos: imagen, texto en la lengua, audio,
// alfabeto. Este panel dice qué falta para habilitar cada actividad, por lengua.

const REQUISITOS_JUEGOS = [
  { id: 'loteria',   nombre: 'Lotería',               min: 32, campo: 'img' },
  { id: 'loteriaA',  nombre: 'Lotería con audio',     min: 32, campo: 'imgAudio' },
  { id: 'memorama',  nombre: 'Memorama',              min: 6,  campo: 'img' },
  { id: 'charadas',  nombre: 'Charadas',              min: 10, campo: 'img' },
  { id: 'escucha',   nombre: '¿Cuál escuchaste?',     min: 4,  campo: 'imgAudio' },
  { id: 'dictado',   nombre: 'Dictado',               min: 5,  campo: 'audio' },
  { id: 'letras',    nombre: 'Ahorcado, sopa y arma la palabra', min: 6, campo: 'largas' },
  { id: 'basta',     nombre: 'Basta',                 min: 3,  campo: 'categorias', alfabeto: true }
];

function estadoCatalogo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES), tId = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
  if (!tPal || !tTr || !tId) return { lenguas: [], global: {}, juegos: REQUISITOS_JUEGOS };

  const pal = new Map();
  const global = { palabras: 0, publicadas: 0, 'revisión': 0, borrador: 0, retirado: 0, sinImagen: 0, imagenSinArchivo: 0, fueraDeJuegos: 0,
    audiosSinConsentimiento: 0, audiosNoPublicos: 0, audiosSinHablante: 0, coincidencias: 0 };
  try { global.coincidencias = _coincidencias(ss).length; } catch (e) { /* sin hoja de relaciones */ }
  let ctxP = null;
  try { ctxP = _contextoPermisos(ss); } catch (e) { ctxP = null; }
  global.esperanPermiso = 0;
  tPal.filas.forEach(f => {
    const id = _k(_get(tPal, f, 'id_palabra'));
    if (!id) return;
    global.palabras++;
    const e = _estadoCanonico(_get(tPal, f, 'estado')) || 'publicado';
    if (e in global) global[e]++;
    const ref = _s(_get(tPal, f, 'imagen_referencia')), imgId = _s(_get(tPal, f, 'imagen_id'));
    const conImg = !!(imgId || _imagenDirecta(ref));
    if (ref && !imgId && !_imagenDirecta(ref)) global.imagenSinArchivo++;
    const publica = _esPublico({ estado: _get(tPal, f, 'estado') });
    const juegos = _k(_get(tPal, f, 'usar_en_juegos')) !== 'no';
    if (publica) { global.publicadas++; if (!conImg) global.sinImagen++; if (!juegos) global.fueraDeJuegos++; }
    pal.set(id, { publica, conImg, juegos, cat: _k(_get(tPal, f, 'id_categoria')) });
  });

  const porLengua = {}, voces = _mapaVoces(ss);
  tTr.filas.forEach(f => {
    const p = pal.get(_k(_get(tTr, f, 'id_palabra')));
    if (!p || !p.publica || !_esPublico({ estado: _get(tTr, f, 'estado') })) return;
    if (ctxP && _usoFuente(ctxP, _get(tTr, f, 'id_fuente'), 'publicar') !== 'si') { global.esperanPermiso++; return; } // su fuente aún no autoriza publicar
    const txt = _s(_get(tTr, f, 'texto'));
    if (!txt) return;
    const L = _k(_get(tTr, f, 'id_idioma'));
    const c = porLengua[L] = porLengua[L] || { textos: 0, audio: 0, img: 0, imgAudio: 0, largas: 0, sinAudio: 0, cats: new Set() };
    const cons = _conVoz({ consentimiento: _get(tTr, f, 'consentimiento'), id_voz: _get(tTr, f, 'id_voz'), hablante: _get(tTr, f, 'hablante') }, voces);
    const hayAudio = _s(_get(tTr, f, 'audio')).length > 1;
    if (hayAudio && !_consentimiento(cons.consentimiento)) global.audiosSinConsentimiento++;
    if (hayAudio && !_audioPublico(cons)) global.audiosNoPublicos++;
    if (hayAudio && !_s(cons.hablante) && !_s(cons.id_voz)) global.audiosSinHablante++;
    const audio = hayAudio && _audioPublico(cons);   // solo cuenta lo que el visor puede reproducir
    c.textos++;
    if (audio) c.audio++; else c.sinAudio++;
    if (p.juegos) {
      if (p.conImg) c.img++;
      if (p.conImg && audio) c.imgAudio++;
      if (txt.replace(/[^\p{L}]/gu, '').length >= 3) c.largas++;
      if (p.cat) c.cats.add(p.cat);
    }
  });

  const lenguas = tId.filas.filter(f => _s(_get(tId, f, 'id_idioma')) && _esPublico({ estado: _get(tId, f, 'estado') })).map(f => {
    const id = _s(_get(tId, f, 'id_idioma'));
    const c = porLengua[_k(id)] || { textos: 0, audio: 0, img: 0, imgAudio: 0, largas: 0, sinAudio: 0, cats: new Set() };
    const alfabeto = _s(_get(tId, f, 'alfabeto'));
    const valores = Object.assign({}, c, { categorias: c.cats.size });
    const juegos = REQUISITOS_JUEGOS.map(j => {
      const tiene = valores[j.campo] || 0;
      const ok = tiene >= j.min && (!j.alfabeto || !!alfabeto);
      return { id: j.id, ok, tiene, min: j.min, faltaAlfabeto: !!j.alfabeto && !alfabeto };
    });
    return {
      id, nombre: _s(_get(tId, f, 'nombre_propio') || _get(tId, f, 'nombre_completo') || _get(tId, f, 'nombre')),
      textos: c.textos, audio: c.audio, img: c.img, sinAudio: c.sinAudio, alfabeto: !!alfabeto, juegos
    };
  });

  let imagenes = {};
  try { imagenes = JSON.parse(PropertiesService.getScriptProperties().getProperty('IMAGENES_ULTIMO_REPORTE') || '{}'); } catch (e) {}
  return { lenguas, global, imagenes, juegos: REQUISITOS_JUEGOS, publicado: _leerManifiesto() };
}

/** Lista de palabras publicadas sin imagen, lista para enviar a quien ilustra (sin generar hojas intermedias). */
/** Desde ¿Cómo va?: crea la hoja para quien ilustra (la misma de 🖼️ Imágenes y sonidos, que después se trae). */
function exportarFaltantesImagen() {
  try { const r = ilustracionCrear({}); return { n: r.filas, url: r.url }; }
  catch (e) { if (e.message === tx('ilustrar.todas')) return { n: 0 }; throw e; }
}

function uiEstadoCatalogo() {
  const nombresJuegos = JSON.stringify(REQUISITOS_JUEGOS.map(j => [j.id, tx('juego.' + j.id)])).replace(/</g, '\\u003c');
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:12px}
    table{width:100%;border-collapse:collapse;font-size:13px}th,td{border-bottom:1px solid #e2e8f0;padding:6px;text-align:left;vertical-align:top}
    th{background:#f8fafc}
    .chip{display:inline-block;margin:2px;padding:3px 8px;border-radius:12px;font-size:12px;border:1px solid #cbd5e1}
    .si{background:#dcfce7;border-color:#16a34a;color:#166534}.no{background:#fff7ed;border-color:#f59e0b;color:#92400e}
    .tarjetas{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px}
    .t{flex:1;min-width:120px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:8px}
    .t b{display:block;font-size:20px}
    button{padding:9px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#e2e8f0;margin:3px 3px 0 0}
    .prim{background:#1e3a5f;color:#fff}
    .aviso{background:#fef3c7;color:#92400e;padding:8px;border-radius:6px;margin:8px 0}
    small{color:#64748b}
  </style>
  ${_tScript()}
  <div class="caja">
    <div id="contenido">${txh('comoVa.calculando')}</div>
    <div style="margin-top:12px">
      <button class="prim" onclick="faltantes()">${txh('comoVa.ilustrar')}</button>
      <button onclick="abrir('uiGenerarReporteAvanzado')">${txh('comoVa.falta')}</button>
      <button onclick="abrir('uiEstados')">${txh('comoVa.estados')}</button>
      <button onclick="abrir('uiLenguas')">${txh('comoVa.lenguas')}</button>
    </div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    const NOMBRES = Object.fromEntries(${nombresJuegos});
    function abrir(fn){ google.script.run[fn](); }
    function faltantes(){
      $('msg').textContent = tt('comoVa.preparandoLista');
      google.script.run.withSuccessHandler(r => {
        $('msg').innerHTML = r.n ? '<p>' + th('comoVa.listaHecha', {n: r.n}) + ' <a href="' + esc(r.url) + '" target="_blank">' + th('comoVa.abrirLista') + '</a> ' + th('comoVa.compartirLista') + '</p>' : '<p>' + th('comoVa.todasImagen') + '</p>';
      }).withFailureHandler(e => $('msg').textContent = tt('comun.error', {m: e.message})).exportarFaltantesImagen();
    }
    google.script.run.withSuccessHandler(r => {
      const g = r.global || {}, im = r.imagenes || {};
      const aviso = (k, n) => '<div class="aviso">' + th(k, {n}) + '</div>', nota = (k, n) => '<p><small>' + th(k, {n}) + '</small></p>';
      let h = '<div class="tarjetas">' +
        '<div class="t"><b>' + (g.publicadas || 0) + '</b>' + th('comoVa.seVen', {total: g.palabras || 0}) + '</div>' +
        '<div class="t"><b>' + (g['revisión'] || 0) + '</b>' + th('comoVa.porRevisar') + '</div>' +
        '<div class="t"><b>' + (g.borrador || 0) + '</b>' + th('comoVa.borrador') + '</div>' +
        '<div class="t"><b>' + (g.sinImagen || 0) + '</b>' + th('comoVa.sinImagen') + '</div></div>';
      if (g.imagenSinArchivo) h += aviso('comoVa.imagenPerdida', g.imagenSinArchivo);
      if (im.errores) h += aviso('comoVa.imagenNoPublica', im.errores);
      if (im.pendientes) h += aviso('comoVa.imagenPendiente', im.pendientes);
      if (g.audiosSinConsentimiento) h += aviso('comoVa.sinConsentimiento', g.audiosSinConsentimiento);
      if (g.audiosNoPublicos) h += nota('comoVa.noPublicos', g.audiosNoPublicos);
      if (g.audiosSinHablante) h += nota('comoVa.sinHablante', g.audiosSinHablante);
      if (g.esperanPermiso) h += aviso('comoVa.esperanPermiso', g.esperanPermiso);
      if (g.coincidencias) h += '<div class="aviso">' + th('comoVa.coincidencias', {n: g.coincidencias}) + ' <button onclick="abrir(\'uiCoincidencias\')">' + th('comoVa.coincBoton') + '</button></div>';
      if (!r.lenguas.length) h += '<p>' + th('comoVa.sinLenguas') + '</p>';
      else {
        h += '<table><tr><th>' + th('comoVa.col.lengua') + '</th><th>' + th('comoVa.col.palabras') + '</th><th>' + th('comoVa.col.audio') + '</th><th>' + th('comoVa.col.imagen') + '</th><th>' + th('comoVa.col.juegos') + '</th></tr>' + r.lenguas.map(l =>
          '<tr><td><b>' + esc(l.nombre) + '</b>' + (l.alfabeto ? '' : '<br><small>' + th('comoVa.faltanLetras') + '</small>') + '</td><td>' + l.textos + '</td><td>' + l.audio +
          (l.sinAudio ? '<br><small>' + th('comoVa.faltan', {n: l.sinAudio}) + '</small>' : '') + '</td><td>' + l.img + '</td><td>' +
          l.juegos.map(j => '<span class="chip ' + (j.ok ? 'si' : 'no') + '" title="' + esc(j.ok ? tt('comoVa.listo') : (j.faltaAlfabeto ? tt('comoVa.faltanLetrasAyuda') : tt('comoVa.tieneNecesita', {tiene: j.tiene, min: j.min}))) + '">' +
            (j.ok ? '✅ ' : '⏳ ') + esc(NOMBRES[j.id]) + (j.ok ? '' : ' (' + esc(j.faltaAlfabeto ? tt('comoVa.faltanLetrasCorto') : tt('comoVa.deCuantos', {tiene: j.tiene, min: j.min})) + ')') + '</span>').join('') +
          '</td></tr>').join('') + '</table>' +
          '<p><small>' + th('comoVa.nota') + '</small></p>';
      }
      $('contenido').innerHTML = h;
    }).withFailureHandler(e => $('contenido').textContent = tt('comun.error', {m: e.message})).estadoCatalogo();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(860).setHeight(640), tx('comoVa.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 13. INTERCAMBIO ENTRE COMUNIDADES (SIN DUPLICADOS)
// ─────────────────────────────────────────────────────────────────────────────
// Un catálogo se exporta como archivo JSON. Otro catálogo lo importa con vista previa:
//  - lenguas: se reconocen por ID, por código ISO + variante, por Glottocode o por nombre + variante;
//  - categorías: por ID o por nombre;
//  - palabras: por ID o (si se elige) porque ya existe la misma escritura en la misma lengua;
//  - traducciones: nunca se duplica la misma palabra en la misma lengua; si la escritura es
//    distinta, se informa como variante y no se sobrescribe nada.
// Todo lo nuevo entra "En revisión" con su origen: cada comunidad decide qué publica.
// Los identificadores se conservan, así las siguientes importaciones se reconocen solas.

const FORMATO_INTERCAMBIO = 'catalogo-lenguas-intercambio';

function _carpetaExportaciones(ss) { return _subcarpeta(_carpetaDelLibro(ss), 'EXPORTACIONES'); }

/** Comparación tolerante: minúsculas, sin acentos ni espacios dobles, apóstrofos unificados. */
function _pliegue(t) {
  return _sinAcentos(_s(t).normalize('NFC')).replace(/[’'ʼꞌ´`]/g, "'").replace(/\s+/g, ' ').trim();
}

function servicioExportarIntercambio(op) {
  op = op || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  const cfg = getAppConfig(ss);
  const soloPublicado = op.alcance !== 'todo';
  const vale = o => !soloPublicado || _esPublico(o);
  const inicio = Date.now();
  const rep = { palabras: 0, traducciones: 0, imagenes: 0, audios: 0, audiosSinArchivo: 0, audiosPorTiempo: 0 };

  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma) && vale(i)).map(i => ({
    id_idioma: _s(i.id_idioma), nombre_completo: _s(i.nombre_completo || i.nombre), nombre_propio: _s(i.nombre_propio),
    codigo_iso: _s(i.codigo_iso), glottocode: _s(i.glottocode), variante: _s(i.variante), region: _s(i.region), alfabeto: _s(i.alfabeto)
  }));
  const tradCat = {};
  readSheet(ss, CONFIG.SHEETS.TRAD_CATS).forEach(tc => {
    const txt = _s(tc.texto || tc.nombre || tc.traduccion || tc.nombre_traducido);
    if (txt) (tradCat[_s(tc.id_categoria)] = tradCat[_s(tc.id_categoria)] || {})[_s(tc.id_idioma)] = txt;
  });
  const categorias = readSheet(ss, CONFIG.SHEETS.CATEGORIAS).filter(c => _s(c.id_categoria) && vale(c)).map(c => ({
    id_categoria: _s(c.id_categoria), nombre_categoria: _s(c.nombre_categoria), traducciones: tradCat[_s(c.id_categoria)] || {}
  }));

  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  const palabras = [], incluidas = new Set();
  tPal.filas.forEach(f => {
    const id = _s(_get(tPal, f, 'id_palabra'));
    if (!id || !vale({ estado: _get(tPal, f, 'estado') })) return;
    const imgId = _s(_get(tPal, f, 'imagen_id'));
    let imagen = '';
    if (imgId) {
      let publica = _s(_get(tPal, f, 'imagen_publica')) === imgId;
      if (!publica && op.compartirMedios) {
        try { DriveApp.getFileById(imgId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); publica = true; } catch (e) {}
      }
      if (publica) { imagen = `https://lh3.googleusercontent.com/d/${imgId}`; rep.imagenes++; }
    } else imagen = _imagenDirecta(_get(tPal, f, 'imagen_referencia'));
    palabras.push({ id_palabra: id, id_categoria: _s(_get(tPal, f, 'id_categoria')), imagen, usar_en_juegos: _s(_get(tPal, f, 'usar_en_juegos')) });
    incluidas.add(_k(id));
  });
  rep.palabras = palabras.length;

  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  const traducciones = [];
  const ctx = _contextoPermisos(ss);
  rep.noCompartibles = 0;
  tTr.filas.forEach(f => {
    if (!incluidas.has(_k(_get(tTr, f, 'id_palabra'))) || !vale({ estado: _get(tTr, f, 'estado') })) return;
    const texto = _s(_get(tTr, f, 'texto'));
    if (!texto) return;
    // Solo viaja lo que su fuente permite compartir con otros catálogos
    if (op.lenguas && op.lenguas.length && !op.lenguas.map(_k).includes(_k(_get(tTr, f, 'id_idioma')))) return;
    if (_usoFuente(ctx, _get(tTr, f, 'id_fuente'), 'compartir') !== 'si') { rep.noCompartibles++; return; }
    const t = {
      id_traduccion: _s(_get(tTr, f, 'id_traduccion')), id_palabra: _s(_get(tTr, f, 'id_palabra')), id_idioma: _s(_get(tTr, f, 'id_idioma')),
      texto, nota_variante: _s(_get(tTr, f, 'nota_variante')), silabas: _s(_get(tTr, f, 'silabas')), descripcion: _s(_get(tTr, f, 'descripcion'))
    };
    if (!t.descripcion) delete t.descripcion;
    // El consentimiento viaja con el audio; sin consentimiento público, el audio no se comparte
    const fila = _conVoz({ consentimiento: _get(tTr, f, 'consentimiento'), id_voz: _get(tTr, f, 'id_voz'), hablante: _get(tTr, f, 'hablante'), lugar_hablante: _get(tTr, f, 'lugar_hablante') }, ctx.voces);
    const audio = _audioPublico(fila) ? _s(_get(tTr, f, 'audio')) : '';
    if (audio) {
      t.consentimiento = _consentimiento(fila.consentimiento) || 'publico_anonimo';
      if (_vozPublica(fila)) { t.hablante = _s(fila.hablante); t.lugar_hablante = _s(fila.lugar_hablante); }
    }
    if (audio && op.compartirMedios) {
      if (Date.now() - inicio > 240000) rep.audiosPorTiempo++;
      else {
        const arch = _buscarArchivoAudio(audio);
        if (arch) {
          try { arch.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); t.audio = arch.getId(); rep.audios++; }
          catch (e) { rep.audiosSinArchivo++; }
        } else rep.audiosSinArchivo++;
      }
    }
    traducciones.push(t);
  });
  rep.traducciones = traducciones.length;

  const relaciones = readSheet(ss, CONFIG.SHEETS.RELACIONES)
    .filter(r => _s(r.id_relacion) && vale(r) && incluidas.has(_k(r.id_origen)) && incluidas.has(_k(r.id_destino)))
    .map(r => ({ id_relacion: _s(r.id_relacion), id_origen: _s(r.id_origen), tipo: _s(r.tipo), id_destino: _s(r.id_destino), nota: _s(r.nota) }));

  _filtrarPaqueteLenguas({ idiomas, categorias, palabras, traducciones, relaciones }, op.lenguas, true);
  rep.palabras = palabras.length; rep.traducciones = traducciones.length;
  const paquete = {
    formato: FORMATO_INTERCAMBIO, version: CONFIG.VERSION, creado: new Date().toISOString(),
    origen: { idCatalogo: cfg.idCatalogo, titulo: cfg.titulo, licencia: cfg.licencia, contacto: { nombre: cfg.contacto.nombre, correo: cfg.contacto.correo } },
    alcance: soloPublicado ? 'publicado' : 'todo',
    idiomas, categorias, palabras, traducciones, relaciones,
    tiposRelacion: (() => { const u = new Set(relaciones.map(r => _k(r.tipo))); return _tiposRelacion(ss).filter(t => u.has(_k(t.id)))
      .map(t => ({ id: t.id, nombre: _etiquetaTipoRel(t, false), inverso: t.simetrica ? '' : _etiquetaTipoRel(t, true), simetrica: t.simetrica })); })()
  };
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const nombre = `Intercambio_${(cfg.titulo || 'catalogo').replace(/[^\p{L}\p{N}]+/gu, '_')}_${ts}.json`;
  const archivo = _carpetaExportaciones(ss).createFile(nombre, JSON.stringify(paquete), 'application/json');
  return { url: archivo.getUrl(), nombre, rep };
}

/**
 * Deja en un paquete solo las lenguas elegidas (ids). Las palabras que se quedan sin traducciones salen;
 * los nombres de categoría en otras lenguas también. Modifica los arreglos en su lugar. Sin lista = todas.
 */
function _filtrarPaqueteLenguas(p, lenguas, enSitio) {
  if (!lenguas || !lenguas.length) return p;
  const L = new Set(lenguas.map(_k));
  const corta = (arr, f) => { const q = arr.filter(f); arr.length = 0; q.forEach(x => arr.push(x)); };
  corta(p.idiomas, i => L.has(_k(i.id_idioma)));
  corta(p.traducciones, t => L.has(_k(t.id_idioma)));
  const conTr = new Set(p.traducciones.map(t => _k(t.id_palabra)));
  corta(p.palabras, x => conTr.has(_k(x.id_palabra)));
  const idsP = new Set(p.palabras.map(x => _k(x.id_palabra))), idsT = new Set(p.traducciones.map(t => _k(t.id_traduccion)));
  corta(p.relaciones || [], r => [r.id_origen, r.id_destino].every(x => idsP.has(_k(x)) || idsT.has(_k(x))));
  (p.categorias || []).forEach(c => { if (c.traducciones) Object.keys(c.traducciones).forEach(k => { if (!L.has(_k(k))) delete c.traducciones[k]; }); });
  return p;
}

/**
 * Importa (o solo previsualiza, si aplicar = false) un paquete de intercambio.
 * opciones: { unirPorTexto: true|false, estado: 'revisión'|'borrador'|'publicado' }
 */
function servicioImportarIntercambio(texto, opciones, aplicar) {
  let paquete;
  try { paquete = JSON.parse(texto); } catch (e) { throw new Error(tx('inter.error.leer')); }
  if (!paquete || paquete.formato !== FORMATO_INTERCAMBIO) throw new Error(tx('inter.error.formato'));
  opciones = opciones || {};
  // Solo las lenguas elegidas del paquete (ids del paquete); sin lista = todas
  paquete.relaciones = paquete.relaciones || []; paquete.categorias = paquete.categorias || [];
  paquete.idiomas = paquete.idiomas || []; paquete.palabras = paquete.palabras || []; paquete.traducciones = paquete.traducciones || [];
  _filtrarPaqueteLenguas(paquete, opciones.lenguas, true);
  const unirPorTexto = opciones.unirPorTexto !== false;
  const estado = _etiquetaEstado(opciones.estado || 'revisión');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = aplicar ? _lock() : null;
  try {
    if (aplicar) asegurarEstructura(ss);
    const cfg = getAppConfig(ss);
    const origen = _s(paquete.origen && paquete.origen.idCatalogo) || 'importado';
    const nombreOrigen = _s(paquete.origen && paquete.origen.titulo) || tx('inter.otroCatalogo');
    if (origen && origen === cfg.idCatalogo) throw new Error(tx('inter.error.mismo'));
    const reglas = _cargarReglas(ss);
    const rep = {
      origen: nombreOrigen,
      idiomas: { nuevos: [], reconocidos: 0 }, categorias: { nuevas: 0, reconocidas: 0 },
      palabras: { nuevas: 0, porId: 0, unidasPorTexto: 0, ambiguas: [] },
      traducciones: { nuevas: 0, iguales: 0, variantes: [] }, relaciones: { nuevas: 0, existentes: 0 }
    };

    // ── Lenguas
    const tId = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
    const idxIdioma = new Map();
    const claveNombre = (n, v) => 'n|' + _pliegue(n) + '|' + _pliegue(v);
    tId.filas.forEach(f => {
      const id = _s(_get(tId, f, 'id_idioma')); if (!id) return;
      const v = _get(tId, f, 'variante');
      idxIdioma.set('id|' + _k(id), id);
      if (_s(_get(tId, f, 'codigo_iso'))) idxIdioma.set('iso|' + _k(_get(tId, f, 'codigo_iso')) + '|' + _pliegue(v), id);
      if (_s(_get(tId, f, 'glottocode'))) idxIdioma.set('g|' + _k(_get(tId, f, 'glottocode')) + '|' + _pliegue(v), id);
      [_get(tId, f, 'nombre_propio'), _get(tId, f, 'nombre_completo'), _get(tId, f, 'nombre')].forEach(n => { if (_s(n)) idxIdioma.set(claveNombre(n, v), id); });
    });
    const mapaIdioma = {}, nuevosIdiomas = [];
    (paquete.idiomas || []).forEach(i => {
      const v = i.variante;
      const encontrado = idxIdioma.get('id|' + _k(i.id_idioma)) ||
        (_s(i.codigo_iso) && idxIdioma.get('iso|' + _k(i.codigo_iso) + '|' + _pliegue(v))) ||
        (_s(i.glottocode) && idxIdioma.get('g|' + _k(i.glottocode) + '|' + _pliegue(v))) ||
        [i.nombre_propio, i.nombre_completo].map(n => _s(n) && idxIdioma.get(claveNombre(n, v))).find(Boolean);
      if (encontrado) { mapaIdioma[_k(i.id_idioma)] = encontrado; rep.idiomas.reconocidos++; return; }
      const id = idxIdioma.has('id|' + _k(i.id_idioma)) ? Utilities.getUuid() : _s(i.id_idioma);
      mapaIdioma[_k(i.id_idioma)] = id;
      nuevosIdiomas.push({ id_idioma: id, nombre_completo: _s(i.nombre_completo), nombre: _s(i.nombre_completo || i.nombre_propio), nombre_propio: _s(i.nombre_propio),
        codigo_iso: _s(i.codigo_iso), glottocode: _s(i.glottocode), variante: _s(i.variante), region: _s(i.region), alfabeto: _s(i.alfabeto),
        estado, notas: `Importada de ${nombreOrigen}` });
      rep.idiomas.nuevos.push(_s(i.nombre_propio || i.nombre_completo));
      idxIdioma.set('id|' + _k(id), id);
    });

    // ── Categorías
    const tCat = _tabla(ss, CONFIG.SHEETS.CATEGORIAS);
    const idxCat = new Map();
    if (tCat) tCat.filas.forEach(f => {
      const id = _s(_get(tCat, f, 'id_categoria')); if (!id) return;
      idxCat.set('id|' + _k(id), id);
      if (_s(_get(tCat, f, 'nombre_categoria'))) idxCat.set('n|' + _pliegue(_get(tCat, f, 'nombre_categoria')), id);
    });
    const mapaCat = {}, nuevasCats = [], nuevasTradCats = [];
    (paquete.categorias || []).forEach(c => {
      const enc = idxCat.get('id|' + _k(c.id_categoria)) || idxCat.get('n|' + _pliegue(c.nombre_categoria));
      if (enc) { mapaCat[_k(c.id_categoria)] = enc; rep.categorias.reconocidas++; return; }
      const id = _s(c.id_categoria) || Utilities.getUuid();
      mapaCat[_k(c.id_categoria)] = id;
      nuevasCats.push({ id_categoria: id, nombre_categoria: _s(c.nombre_categoria), estado });
      Object.keys(c.traducciones || {}).forEach(L => {
        const lid = mapaIdioma[_k(L)];
        if (lid) nuevasTradCats.push({ id_trad_cat: Utilities.getUuid(), id_categoria: id, id_idioma: lid, texto: _s(c.traducciones[L]) });
      });
      idxCat.set('id|' + _k(id), id); idxCat.set('n|' + _pliegue(c.nombre_categoria), id);
      rep.categorias.nuevas++;
    });

    // ── Palabras
    const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
    const palabrasLocales = new Set(tPal.filas.map(f => _k(_get(tPal, f, 'id_palabra'))).filter(Boolean));
    const idxTexto = new Map();       // idioma|texto plegado → Set(id_palabra)
    const idxTrad = new Map();        // id_palabra|idioma → texto plegado
    tTr.filas.forEach(f => {
      const pid = _k(_get(tTr, f, 'id_palabra')), L = _k(_get(tTr, f, 'id_idioma')), t = _pliegue(_get(tTr, f, 'texto'));
      if (!pid || !t) return;
      const k = L + '|' + t;
      if (!idxTexto.has(k)) idxTexto.set(k, new Set());
      idxTexto.get(k).add(pid);
      idxTrad.set(pid + '|' + L, t);
    });
    const tradsPorPalabra = {};
    (paquete.traducciones || []).forEach(t => { (tradsPorPalabra[_k(t.id_palabra)] = tradsPorPalabra[_k(t.id_palabra)] || []).push(t); });
    const mapaPal = {}, nuevasPal = [];
    (paquete.palabras || []).forEach(p => {
      const pid = _k(p.id_palabra);
      if (palabrasLocales.has(pid)) { mapaPal[pid] = _s(p.id_palabra); rep.palabras.porId++; return; }
      if (unirPorTexto) {
        const candidatos = new Set();
        (tradsPorPalabra[pid] || []).forEach(t => {
          const L = mapaIdioma[_k(t.id_idioma)];
          if (!L) return;
          (idxTexto.get(_k(L) + '|' + _pliegue(t.texto)) || new Set()).forEach(c => candidatos.add(c));
        });
        if (candidatos.size === 1) {
          const local = [...candidatos][0];
          const original = tPal.filas.find(f => _k(_get(tPal, f, 'id_palabra')) === local);
          mapaPal[pid] = _s(_get(tPal, original, 'id_palabra'));
          rep.palabras.unidasPorTexto++;
          return;
        }
        if (candidatos.size > 1) rep.palabras.ambiguas.push((tradsPorPalabra[pid] || []).map(t => t.texto).join(' / '));
      }
      mapaPal[pid] = _s(p.id_palabra) || Utilities.getUuid();
      nuevasPal.push({ id_palabra: mapaPal[pid], id_categoria: mapaCat[_k(p.id_categoria)] || '', imagen_referencia: _s(p.imagen),
        usar_en_juegos: _s(p.usar_en_juegos), estado, origen, creado_por: 'importación: ' + nombreOrigen });
      palabrasLocales.add(_k(mapaPal[pid]));
      rep.palabras.nuevas++;
    });

    // ── Traducciones (nunca se duplica palabra + lengua)
    const modoCaptura = _col(tTr, 'texto_captura') > -1;
    const nuevasTr = [];
    (paquete.traducciones || []).forEach(t => {
      const pid = mapaPal[_k(t.id_palabra)], L = mapaIdioma[_k(t.id_idioma)];
      if (!pid || !L || !_s(t.texto)) return;
      const clave = _k(pid) + '|' + _k(L);
      const previa = idxTrad.get(clave), plegado = _pliegue(t.texto);
      if (previa != null) {
        if (previa === plegado) rep.traducciones.iguales++;
        else if (rep.traducciones.variantes.length < 200) rep.traducciones.variantes.push({ texto: _s(t.texto), lengua: L });
        return;
      }
      idxTrad.set(clave, plegado);
      nuevasTr.push({ id_traduccion: Utilities.getUuid(), id_palabra: pid, id_idioma: L,
        texto: _normalizar(_s(t.texto), L, reglas, '', !modoCaptura), texto_captura: _s(t.texto),
        nota_variante: _s(t.nota_variante), silabas: _s(t.silabas), audio: _s(t.audio), descripcion: _s(t.descripcion).slice(0, LARGO_DESCRIPCION),
        hablante: _s(t.hablante), lugar_hablante: _s(t.lugar_hablante),
        consentimiento: _s(t.audio) ? ((CONSENTIMIENTOS.find(c => c.canon === _consentimiento(t.consentimiento)) || CONSENTIMIENTOS[1]).etiqueta) : '',
        estado, origen, creado_por: 'importación: ' + nombreOrigen });
      rep.traducciones.nuevas++;
    });

    // ── Relaciones
    const tRel = _tabla(ss, CONFIG.SHEETS.RELACIONES);
    const relExist = new Set(tRel ? tRel.filas.map(f => [_k(_get(tRel, f, 'id_origen')), _k(_get(tRel, f, 'tipo')), _k(_get(tRel, f, 'id_destino'))].join('|')) : []);
    const nuevasRel = [];
    (paquete.relaciones || []).forEach(r => {
      const a = mapaPal[_k(r.id_origen)], b = mapaPal[_k(r.id_destino)];
      if (!a || !b) return;
      const k = [_k(a), _k(r.tipo), _k(b)].join('|');
      if (relExist.has(k)) { rep.relaciones.existentes++; return; }
      relExist.add(k);
      nuevasRel.push({ id_relacion: Utilities.getUuid(), id_origen: a, tipo: _s(r.tipo), id_destino: b, nota: _s(r.nota), estado, creado_por: 'importación: ' + nombreOrigen });
      rep.relaciones.nuevas++;
    });

    if (aplicar) {
      _agregar(tId, nuevosIdiomas);
      if (tCat) _agregar(tCat, nuevasCats);
      const tTc = _tabla(ss, CONFIG.SHEETS.TRAD_CATS);
      if (tTc) _agregar(tTc, nuevasTradCats);
      _agregar(tPal, nuevasPal);
      _agregar(tTr, nuevasTr);
      if (tRel) _agregar(tRel, nuevasRel);
      if (tRel && nuevasRel.length) rep.relaciones.tiposNuevos = _tiposDesdePaquete(ss, paquete.tiposRelacion, nuevasRel.map(r => r.tipo));
      publicarCatalogo(ss, true);
    }
    rep.aplicado = !!aplicar;
    rep.estado = estado;
    return rep;
  } finally { if (lock) lock.releaseLock(); }
}

// ─────────────────────────────────────────────────────────────────────────────
// 14. EXPORTACIÓN A BASES DE GRAFOS
// ─────────────────────────────────────────────────────────────────────────────
// Un solo modelo de nodos y aristas, varios formatos:
//  - N-Triples con OntoLex-Lemon, SKOS y Dublin Core → Apache Jena / Fuseki, Oxigraph, GraphDB
//  - CSV de nodos y aristas → Neo4j (neo4j-admin import / LOAD CSV) y Apache AGE (PostgreSQL)
//  - GraphML → Apache TinkerPop (Gremlin), Gephi, Cytoscape, NetworkX
// Ninguna lengua es "la principal": un Concepto no tiene nombre propio; los Textos lo expresan.

function _construirGrafo(ss, soloPublicado) {
  _cfgMemo = null;
  const cfg = getAppConfig(ss);
  const vale = o => !soloPublicado || _esPublico(o);
  const nodos = [], aristas = [], vistos = new Set();
  const nodo = (id, tipo, props) => { if (vistos.has(id)) return; vistos.add(id); nodos.push({ id, tipo, props: props || {} }); };
  const arista = (de, a, tipo, props) => aristas.push({ de, a, tipo, props: props || {} });
  const unicas = new Set();
  const aristaUnica = (de, a, tipo) => { const k = de + '>' + tipo + '>' + a; if (!unicas.has(k)) { unicas.add(k); arista(de, a, tipo); } };
  const ap = _mapaApostrofos(ss);
  const catId = `catalogo:${cfg.idCatalogo || 'sin-id'}`;
  nodo(catId, 'Catalogo', { nombre: cfg.titulo, licencia: cfg.licencia });

  const idiomas = {};
  readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma) && vale(i)).forEach(i => {
    const id = `idioma:${_s(i.id_idioma)}`;
    // codigo = el que usa la comunidad (ej. wix); iso = código internacional ISO 639-3 (ej. hch), solo si se registró
    idiomas[_k(i.id_idioma)] = { id, iso: _k(i.codigo_internacional), codigo: _k(i.codigo_iso) };
    nodo(id, 'Idioma', { nombre: _s(i.nombre_propio || i.nombre_completo || i.nombre), nombre_otra: _s(i.nombre_completo || i.nombre),
      codigo: _k(i.codigo_iso), iso: _k(i.codigo_internacional), glottocode: _k(i.glottocode), variante: _s(i.variante) });
  });
  const cats = {};
  readSheet(ss, CONFIG.SHEETS.CATEGORIAS).filter(c => _s(c.id_categoria) && vale(c)).forEach(c => {
    const id = `categoria:${_s(c.id_categoria)}`;
    cats[_k(c.id_categoria)] = id;
    nodo(id, 'Categoria', { nombre: _s(c.nombre_categoria) });
  });
  const conceptos = {};
  readSheet(ss, CONFIG.SHEETS.PALABRAS).filter(p => _s(p.id_palabra) && vale(p)).forEach(p => {
    const id = `concepto:${_s(p.id_palabra)}`;
    conceptos[_k(p.id_palabra)] = id;
    nodo(id, 'Concepto', { uuid: _s(p.id_palabra) });
    arista(id, catId, 'DEL_CATALOGO');
    if (cats[_k(p.id_categoria)]) arista(id, cats[_k(p.id_categoria)], 'EN_CATEGORIA');
    const imgId = _s(p.imagen_id);
    if (imgId) {
      const m = `medio:${imgId}`;
      nodo(m, 'Medio', { tipo: 'imagen', url: `https://lh3.googleusercontent.com/d/${imgId}` });
      arista(id, m, 'TIENE_IMAGEN');
    }
  });
  // Galería: imágenes extra publicadas (no las internas)
  if (ss.getSheetByName(CONFIG.SHEETS.IMAGENES)) readSheet(ss, CONFIG.SHEETS.IMAGENES).filter(i => _s(i.archivo) && vale(i) && _s(i.solo_interna) !== 'Sí' && conceptos[_k(i.id_palabra)]).forEach(i => {
    const fid = _idDrive(i.archivo) || _s(i.archivo), m = `medio:${fid}`;
    nodo(m, 'Medio', { tipo: 'imagen', url: `https://lh3.googleusercontent.com/d/${fid}`, nota: _s(i.alt) });
    arista(conceptos[_k(i.id_palabra)], m, 'TIENE_IMAGEN');
  });
  const textos = {};
  const ctx = _contextoPermisos(ss), voces = ctx.voces;
  // Procedencia: comunidades, fuentes y quienes las autorizan (sin datos de contacto)
  const comunidades = {};
  Object.values(ctx.comunidades).filter(c => vale(c)).forEach(c => {
    const id = `comunidad:${_s(c.id_comunidad)}`; comunidades[_k(c.id_comunidad)] = id;
    nodo(id, 'Comunidad', { nombre: _s(c.nombre_propio), nombre_otra: _s(c.nombre_otro), variante: _s(c.variante), region: _s(c.region) });
    if (idiomas[_k(c.id_idioma)]) arista(id, idiomas[_k(c.id_idioma)].id, 'HABLA');
  });
  const fuentes = {};
  Object.values(ctx.fuentes).filter(f => _k(f.id_fuente) !== FUENTE_EQUIPO && vale(f) && (!soloPublicado || _sinAcentos(f.publicar_credito) !== 'no' || _fuenteFiel(ctx, f.id_fuente))).forEach(f => {
    const id = `fuente:${_s(f.id_fuente)}`; fuentes[_k(f.id_fuente)] = id;
    nodo(id, 'Fuente', { nombre: _s(f.titulo), tipo: _s(f.tipo), credito: _s(f.credito) || _s(f.autoria), licencia: _s(f.licencia), escritura: _s(f.sistema_escritura) });
    if (comunidades[_k(f.id_comunidad)]) arista(id, comunidades[_k(f.id_comunidad)], 'DE_COMUNIDAD');
    (ctx.porFuente[_k(f.id_fuente)] || []).forEach(a => {
      const ac = ctx.actores[_k(a.id_actor)]; if (!ac) return;
      const aid = `actor:${_s(ac.id_actor)}`;
      nodo(aid, 'Actor', { nombre: _s(ac.nombre), tipo: _s(ac.tipo), credito: _s(ac.credito) });
      arista(aid, id, 'AUTORIZA', { tipo: _s(a.rol), nota: _s(a.tipo_permiso) });
    });
  });
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).filter(t => vale(t) && conceptos[_k(t.id_palabra)] && idiomas[_k(t.id_idioma)] && _s(t.texto) &&
    (!soloPublicado || _usoFuente(ctx, t.id_fuente, 'publicar') === 'si')).forEach(t => {
    _conVoz(t, voces);
    const id = `texto:${_s(t.id_traduccion) || Utilities.getUuid()}`;
    const L = idiomas[_k(t.id_idioma)];
    textos[_k(t.id_traduccion)] = id;
    const texto = _unificarApostrofos(t.texto, ap[_k(t.id_idioma)]);
    nodo(id, 'Texto', { texto, lengua: L.codigo, iso: L.iso, nota: _s(t.nota_variante), silabas: _s(t.silabas), tipo: _tipoRegistro(t),
      descripcion: _s(t.descripcion) ? _unificarApostrofos(_s(t.descripcion), ap[_k(t.id_idioma)]) : '' });
    arista(id, conceptos[_k(t.id_palabra)], 'EXPRESA');
    arista(id, L.id, 'EN_LENGUA');
    if (fuentes[_k(t.id_fuente)]) arista(id, fuentes[_k(t.id_fuente)], 'VIENE_DE', { nota: _s(t.ubicacion_fuente) });
    // Cada forma escrita es un nodo propio, compartido entre palabras: así se ve qué forma nombra varias ideas
    _formas(texto).forEach(fm => {
      const fid = `forma:${L.id.replace(/^idioma:/, '')}:${_claveForma(fm)}`;
      if (!vistos.has(fid)) { nodo(fid, 'Forma', { texto: fm, lengua: L.codigo, iso: L.iso }); arista(fid, L.id, 'EN_LENGUA'); }
      arista(id, fid, 'SE_ESCRIBE');
      aristaUnica(fid, conceptos[_k(t.id_palabra)], 'NOMBRA');
    });
    const audioId = (!soloPublicado || _audioPublico(t)) ? _idDrive(t.audio) : '';
    if (audioId) {
      const m = `medio:${audioId}`, voz = _vozPublica(t) || !soloPublicado;
      nodo(m, 'Medio', { tipo: 'audio', drive: audioId, consentimiento: _consentimiento(t.consentimiento) || 'sin_registro',
        hablante: voz ? _s(t.hablante) : '', lugar: voz ? _s(t.lugar_hablante) : '' });
      arista(id, m, 'TIENE_AUDIO');
      const com = comunidades[_k(_comunidadDe(ctx, t))];
      if (com) aristaUnica(m, com, 'DE_COMUNIDAD');
    }
  });
  // RELACION: el tipo (código estable), su nombre y el de la dirección contraria. En lo publicado, solo los tipos públicos.
  const tiposRel = new Map(_tiposRelacion(ss).map(t => [_k(t.id), t]));
  readSheet(ss, CONFIG.SHEETS.RELACIONES).filter(r => _s(r.id_relacion) && vale(r)).forEach(r => {
    const t = tiposRel.get(_k(r.tipo));
    if (soloPublicado && !(t && t.publica && t.activo)) return;
    const a = conceptos[_k(r.id_origen)] || textos[_k(r.id_origen)], b = conceptos[_k(r.id_destino)] || textos[_k(r.id_destino)];
    if (a && b) arista(a, b, 'RELACION', { tipo: _s(r.tipo), nombre: _etiquetaTipoRel(t, false) || _s(r.tipo),
      inverso: t && !t.simetrica ? _etiquetaTipoRel(t, true) : '', simetrica: t && t.simetrica ? 'si' : '', nota: _s(r.nota) });
  });
  return { cfg, nodos, aristas };
}

/** N-Triples (RDF 1.1) con OntoLex-Lemon. IRIs estables: urn:catalogo-lenguas:<catálogo>:<tipo>:<id>. */
function _grafoNTriples(g) {
  const base = `urn:catalogo-lenguas:${encodeURIComponent(g.cfg.idCatalogo || 'sin-id')}:`;
  const iri = id => `<${base}${id.split(':').map(encodeURIComponent).join(':')}>`;
  const V = {
    type: '<http://www.w3.org/1999/02/22-rdf-syntax-ns#type>', label: '<http://www.w3.org/2000/01/rdf-schema#label>',
    comment: '<http://www.w3.org/2000/01/rdf-schema#comment>', sameAs: '<http://www.w3.org/2002/07/owl#sameAs>',
    LexicalConcept: '<http://www.w3.org/ns/lemon/ontolex#LexicalConcept>', LexicalEntry: '<http://www.w3.org/ns/lemon/ontolex#LexicalEntry>',
    Form: '<http://www.w3.org/ns/lemon/ontolex#Form>', evokes: '<http://www.w3.org/ns/lemon/ontolex#evokes>',
    canonicalForm: '<http://www.w3.org/ns/lemon/ontolex#canonicalForm>', writtenRep: '<http://www.w3.org/ns/lemon/ontolex#writtenRep>',
    skosConcept: '<http://www.w3.org/2004/02/skos/core#Concept>', prefLabel: '<http://www.w3.org/2004/02/skos/core#prefLabel>',
    language: '<http://purl.org/dc/terms/language>', subject: '<http://purl.org/dc/terms/subject>', identifier: '<http://purl.org/dc/terms/identifier>',
    LinguisticSystem: '<http://purl.org/dc/terms/LinguisticSystem>', isPartOf: '<http://purl.org/dc/terms/isPartOf>',
    title: '<http://purl.org/dc/terms/title>', license: '<http://purl.org/dc/terms/license>', depiction: '<http://xmlns.com/foaf/0.1/depiction>',
    audio: '<http://schema.org/audio>', LexicalSense: '<http://www.w3.org/ns/lemon/ontolex#LexicalSense>',
    sense: '<http://www.w3.org/ns/lemon/ontolex#sense>', isLexicalizedSenseOf: '<http://www.w3.org/ns/lemon/ontolex#isLexicalizedSenseOf>', Dataset: '<http://www.w3.org/ns/dcat#Dataset>'
  };
  const lit = (v, tag) => '"' + String(v).replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n').replace(/\r/g, '\\r').replace(/\t/g, '\\t') + '"' + (tag ? '@' + tag : '');
  const lineas = [];
  const t = (s, p, o) => lineas.push(`${s} ${p} ${o} .`);
  const tipos = {}, relDecl = new Set();
  g.nodos.forEach(n => { tipos[n.id] = n; });
  g.nodos.forEach(n => {
    const s = iri(n.id), p = n.props;
    if (n.tipo === 'Catalogo') { t(s, V.type, V.Dataset); if (p.nombre) t(s, V.title, lit(p.nombre)); if (p.licencia) t(s, V.license, lit(p.licencia)); }
    if (n.tipo === 'Idioma') {
      t(s, V.type, V.LinguisticSystem); if (p.nombre) t(s, V.label, lit(p.nombre));
      if (p.nombre_otra && p.nombre_otra !== p.nombre) t(s, V.label, lit(p.nombre_otra));
      if (p.codigo) t(s, V.identifier, lit(p.codigo));
      if (/^[a-z]{3}$/.test(p.iso)) { t(s, V.identifier, lit(p.iso)); t(s, V.sameAs, `<http://lexvo.org/id/iso639-3/${p.iso}>`); }
      if (p.glottocode) t(s, V.sameAs, `<https://glottolog.org/resource/languoid/id/${encodeURIComponent(p.glottocode)}>`);
      if (p.variante) t(s, V.comment, lit('Variante: ' + p.variante));
    }
    if (n.tipo === 'Categoria') { t(s, V.type, V.skosConcept); if (p.nombre) t(s, V.prefLabel, lit(p.nombre)); }
    if (n.tipo === 'Concepto') t(s, V.type, V.LexicalConcept);
    if (n.tipo === 'Fuente') { t(s, V.type, `<http://purl.org/dc/terms/BibliographicResource>`); if (p.nombre) t(s, V.title, lit(p.nombre)); if (p.licencia) t(s, V.license, lit(p.licencia)); }
    if (n.tipo === 'Actor' || n.tipo === 'Comunidad') { t(s, V.type, n.tipo === 'Actor' ? '<http://xmlns.com/foaf/0.1/Agent>' : '<http://xmlns.com/foaf/0.1/Group>'); if (p.nombre) t(s, V.label, lit(p.nombre)); }
    const tag = /^[a-z]{3}$/.test(p.iso || '') ? p.iso : 'und'; // solo el código internacional sirve como etiqueta de idioma
    // Forma = entrada léxica (una por escritura y lengua); Texto = el registro de una traducción (un sentido de esa entrada)
    if (n.tipo === 'Forma') {
      const escrita = iri(n.id + ':escrita');
      t(s, V.type, V.LexicalEntry); t(s, V.canonicalForm, escrita); t(escrita, V.type, V.Form); t(escrita, V.writtenRep, lit(p.texto, tag));
    }
    if (n.tipo === 'Texto') {
      t(s, V.type, V.LexicalSense); t(s, V.label, lit(p.texto, tag));
      if (p.nota) t(s, V.comment, lit(p.nota));
      // Descripción escrita por la comunidad en la misma lengua
      if (p.descripcion) t(s, '<http://www.w3.org/2004/02/skos/core#definition>', lit(p.descripcion, tag));
    }
  });
  const pred = { DEL_CATALOGO: V.isPartOf, EN_CATEGORIA: V.subject, EXPRESA: V.isLexicalizedSenseOf, NOMBRA: V.evokes, EN_LENGUA: V.language,
    VIENE_DE: '<http://purl.org/dc/terms/source>', DE_COMUNIDAD: '<http://purl.org/dc/terms/spatial>', HABLA: V.language };
  g.aristas.forEach(a => {
    if (pred[a.tipo]) return t(iri(a.de), pred[a.tipo], iri(a.a));
    if (a.tipo === 'SE_ESCRIBE') return t(iri(a.a), V.sense, iri(a.de));
    if (a.tipo === 'AUTORIZA') return t(iri(a.a), '<http://purl.org/dc/terms/rightsHolder>', iri(a.de));
    const destino = tipos[a.a];
    if (a.tipo === 'TIENE_IMAGEN' && destino) return t(iri(a.de), V.depiction, `<${destino.props.url}>`);
    if (a.tipo === 'TIENE_AUDIO' && destino) return t(iri(a.de), V.audio, `<https://drive.google.com/file/d/${destino.props.drive}>`);
    if (a.tipo === 'RELACION') {
      const p = `<urn:catalogo-lenguas:relacion:${encodeURIComponent(_sinAcentos(a.props.tipo || 'relacionado').replace(/\s+/g, '_'))}>`;
      t(iri(a.de), p, iri(a.a));
      if (!relDecl.has(p)) { relDecl.add(p); t(p, V.label, lit(a.props.nombre || a.props.tipo)); if (a.props.simetrica) t(p, V.type, '<http://www.w3.org/2002/07/owl#SymmetricProperty>'); }
      // Equivalencias SKOS para que otras herramientas entiendan los tipos básicos
      if (a.props.tipo === 'tipo_de') t(iri(a.de), '<http://www.w3.org/2004/02/skos/core#broader>', iri(a.a));
      if (a.props.tipo === 'relacionada') t(iri(a.de), '<http://www.w3.org/2004/02/skos/core#related>', iri(a.a));
      return;
    }
  });
  return lineas.join('\n') + '\n';
}

function _csv(filas) {
  return filas.map(f => f.map(v => { const s = String(v == null ? '' : v); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(',')).join('\n') + '\n';
}

/** CSV para Neo4j (un archivo de nodos y uno de aristas) y para Apache AGE (un archivo por etiqueta y por tipo). */
function _grafoCSV(g) {
  const props = ['nombre', 'nombre_otra', 'texto', 'lengua', 'codigo', 'iso', 'glottocode', 'variante', 'region', 'nota', 'silabas', 'descripcion', 'url', 'tipo', 'uuid', 'licencia', 'credito', 'escritura'];
  const archivos = [];
  archivos.push(Utilities.newBlob(_csv([['id:ID', ':LABEL'].concat(props)].concat(g.nodos.map(n => [n.id, n.tipo].concat(props.map(p => n.props[p] || '')))) ), 'text/csv', 'neo4j/nodos.csv'));
  archivos.push(Utilities.newBlob(_csv([[':START_ID', ':END_ID', ':TYPE', 'tipo', 'nombre', 'inverso', 'nota']].concat(g.aristas.map(a => [a.de, a.a, a.tipo, a.props.tipo || '', a.props.nombre || '', a.props.inverso || '', a.props.nota || '']))), 'text/csv', 'neo4j/aristas.csv'));
  // Apache AGE: identificadores numéricos por etiqueta y aristas con tipo de vértice
  const num = {}, porTipo = {};
  g.nodos.forEach(n => { porTipo[n.tipo] = porTipo[n.tipo] || []; porTipo[n.tipo].push(n); num[n.id] = porTipo[n.tipo].length; });
  Object.keys(porTipo).forEach(tipo => {
    archivos.push(Utilities.newBlob(_csv([['id', 'clave'].concat(props)].concat(porTipo[tipo].map(n => [num[n.id], n.id].concat(props.map(p => n.props[p] || ''))))), 'text/csv', `age/vertices_${tipo}.csv`));
  });
  const tipoDe = {};
  g.nodos.forEach(n => { tipoDe[n.id] = n.tipo; });
  const aristasPorTipo = {};
  g.aristas.forEach(a => { (aristasPorTipo[a.tipo] = aristasPorTipo[a.tipo] || []).push(a); });
  Object.keys(aristasPorTipo).forEach(tipo => {
    archivos.push(Utilities.newBlob(_csv([['start_id', 'start_vertex_type', 'end_id', 'end_vertex_type', 'tipo', 'nombre', 'inverso', 'nota']].concat(aristasPorTipo[tipo].map(a =>
      [num[a.de], tipoDe[a.de], num[a.a], tipoDe[a.a], a.props.tipo || '', a.props.nombre || '', a.props.inverso || '', a.props.nota || '']))), 'text/csv', `age/aristas_${tipo}.csv`));
  });
  archivos.push(Utilities.newBlob([
    'CATÁLOGO DE LENGUAS — EXPORTACIÓN A GRAFOS', '',
    'NEO4J', '  neo4j-admin database import full --nodes=neo4j/nodos.csv --relationships=neo4j/aristas.csv catalogo', '  (o LOAD CSV desde Cypher).', '',
    'APACHE AGE (PostgreSQL)', "  SELECT create_graph('catalogo');",
    "  SELECT create_vlabel('catalogo','Concepto');  -- y así para cada archivo age/vertices_<Etiqueta>.csv",
    "  SELECT load_labels_from_file('catalogo','Concepto','/ruta/age/vertices_Concepto.csv');",
    "  SELECT create_elabel('catalogo','EXPRESA');   -- y así para cada archivo age/aristas_<TIPO>.csv",
    "  SELECT load_edges_from_file('catalogo','EXPRESA','/ruta/age/aristas_EXPRESA.csv');", '',
    'MODELO', '  Concepto: la idea (sin lengua principal). Texto: cómo la expresa una lengua (EXPRESA → Concepto, EN_LENGUA → Idioma).',
    '  Forma: cada escritura de una lengua, una sola vez aunque aparezca en varias palabras (Texto -SE_ESCRIBE-> Forma -NOMBRA-> Concepto).',
    '  Una Forma con varias aristas NOMBRA es una palabra con varios sentidos (o un duplicado por revisar).',
    '  Procedencia: Texto -VIENE_DE-> Fuente <-AUTORIZA- Actor; Medio (audio) -DE_COMUNIDAD-> Comunidad -HABLA-> Idioma.',
    '  lengua = código que usa la comunidad; iso = código internacional ISO 639-3 (si se registró).',
    '  Categoria, Idioma, Medio (imagen/audio), Catalogo (procedencia). RELACION: vínculos registrados por la comunidad.'
  ].join('\n'), 'text/plain', 'LEEME.txt'));
  return Utilities.zip(archivos, 'grafo_csv.zip');
}

/** GraphML: lo leen Apache TinkerPop, Gephi, Cytoscape, yEd y NetworkX. */
function _grafoGraphML(g) {
  const x = v => String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
  const props = ['nombre', 'texto', 'lengua', 'iso', 'glottocode', 'variante', 'nota', 'descripcion', 'url', 'tipo'];
  const claves = [`<key id="labelV" for="node" attr.name="labelV" attr.type="string"/>`, `<key id="labelE" for="edge" attr.name="labelE" attr.type="string"/>`]
    .concat(props.map(p => `<key id="${p}" for="all" attr.name="${p}" attr.type="string"/>`), [`<key id="inverso" for="edge" attr.name="inverso" attr.type="string"/>`]);
  const nodos = g.nodos.map(n => `<node id="${x(n.id)}"><data key="labelV">${x(n.tipo)}</data>` +
    props.filter(p => n.props[p]).map(p => `<data key="${p}">${x(n.props[p])}</data>`).join('') + `</node>`);
  const aristas = g.aristas.map((a, i) => `<edge id="e${i}" source="${x(a.de)}" target="${x(a.a)}"><data key="labelE">${x(a.tipo)}</data>` +
    ['tipo', 'nombre', 'nota'].filter(p => a.props[p]).map(p => `<data key="${p}">${x(a.props[p])}</data>`).join('') +
    (a.props.inverso ? `<data key="inverso">${x(a.props.inverso)}</data>` : '') + `</edge>`);
  return `<?xml version="1.0" encoding="UTF-8"?>\n<graphml xmlns="http://graphml.graphdrawing.org/xmlns">\n${claves.join('\n')}\n<graph id="catalogo" edgedefault="directed">\n${nodos.join('\n')}\n${aristas.join('\n')}\n</graph>\n</graphml>\n`;
}

function servicioExportarGrafo(formato, alcance) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const g = _construirGrafo(ss, alcance !== 'todo');
  const carpeta = _carpetaExportaciones(ss);
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const base = `Grafo_${(g.cfg.titulo || 'catalogo').replace(/[^\p{L}\p{N}]+/gu, '_')}_${ts}`;
  const archivos = [];
  if (formato === 'nt' || formato === 'todos') archivos.push(carpeta.createFile(Utilities.newBlob(_grafoNTriples(g), 'application/n-triples', base + '.nt')));
  if (formato === 'csv' || formato === 'todos') archivos.push(carpeta.createFile(_grafoCSV(g).setName(base + '_csv.zip')));
  if (formato === 'graphml' || formato === 'todos') archivos.push(carpeta.createFile(Utilities.newBlob(_grafoGraphML(g), 'application/xml', base + '.graphml')));
  return { nodos: g.nodos.length, aristas: g.aristas.length, archivos: archivos.map(f => ({ nombre: f.getName(), url: f.getUrl() })) };
}

function uiIntercambio() {
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:14px}
    h3{margin:14px 0 6px;font-size:15px}
    .sec{border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:10px;background:#f8fafc}
    label{display:block;margin:4px 0}
    button{padding:9px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#e2e8f0;margin:6px 4px 0 0}
    .prim{background:#1e3a5f;color:#fff}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:8px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
    small{color:#64748b} ul{margin:4px 0 4px 18px;padding:0}
  </style>
  ${_tScript()}
  <div class="caja">
    <div class="sec">
      <h3 style="margin-top:0">${txh('inter.enviar')}</h3>
      <p style="margin:0 0 6px"><small>${txh('inter.enviar.d')}</small></p>
      <label><input type="radio" name="alc" value="publicado" checked> ${txh('inter.soloVisibles')}</label>
      <label><input type="radio" name="alc" value="todo"> ${txh('inter.todas')}</label>
      <label><input type="checkbox" id="medios"> ${txh('inter.medios')} <small>${txh('inter.medios.d')}</small></label>
      <div><b>${txh('inter.queLenguas')}</b><div id="lenguasExp"></div></div>
      <button class="prim" onclick="exportar()">${txh('inter.crear')}</button>
      <div id="r1" role="status" aria-live="polite"></div>
    </div>
    <div class="sec">
      <h3 style="margin-top:0">${txh('exportarHoja.titulo')}</h3>
      <p style="margin:0 0 6px"><small>${txh('exportarHoja.d')}</small></p>
      <div id="lenguasHoja"></div>
      <label><input type="radio" name="halc" value="todo" checked> ${txh('exportarHoja.todo')}</label>
      <label><input type="radio" name="halc" value="publicado"> ${txh('inter.soloVisibles')}</label>
      <label><input type="checkbox" id="hCompartir" checked> ${txh('exportarHoja.compartibles')}</label>
      <button class="prim" onclick="hoja()">${txh('exportarHoja.crear')}</button>
      <div id="r4" role="status" aria-live="polite"></div>
    </div>
    <div class="sec">
      <h3 style="margin-top:0">${txh('inter.recibir')}</h3>
      <p style="margin:0 0 6px"><small>${txh('inter.recibir.d')}</small></p>
      <input type="file" id="archivo" accept=".json,application/json">
      <div id="lenguasImp"></div>
      <label><input type="checkbox" id="unir" checked> ${txh('inter.unir')}</label>
      <label>${txh('inter.quedan')} <select id="estado"><option value="revisión" selected>${txh('inter.quedan.revision')}</option><option value="borrador">${txh('inter.quedan.borrador')}</option><option value="publicado">${txh('inter.quedan.publicado')}</option></select></label>
      <button onclick="importar(false)">${txh('inter.ver')}</button> <button class="prim" id="bImp" onclick="importar(true)" disabled>${txh('inter.agregar')}</button>
      <div id="r2" role="status" aria-live="polite"></div>
    </div>
    <details class="sec"><summary><b>${txh('inter.investigacion')}</b> <small>${txh('inter.investigacion.d')}</small></summary>
      <p><small>${txh('inter.investigacion.texto')}</small></p>
      <label><input type="radio" name="galc" value="publicado" checked> ${txh('inter.soloPublicado')}</label>
      <label><input type="radio" name="galc" value="todo"> ${txh('inter.todo')}</label>
      <button onclick="grafo('nt')">N-Triples (Apache Jena, Oxigraph)</button>
      <button onclick="grafo('csv')">CSV (Neo4j, Apache AGE)</button>
      <button onclick="grafo('graphml')">GraphML (TinkerPop, Gephi)</button>
      <button class="prim" onclick="grafo('todos')">${txh('inter.todos')}</button>
      <div id="r3" role="status" aria-live="polite"></div>
    </details>
    <p><small>${txh('inter.carpeta')}</small></p>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    const err = (id) => e => { $(id).innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; };
    const val = n => document.querySelector('input[name="' + n + '"]:checked').value;
    let texto = '';
    const casillas = l => l.map(x => '<label><input type="checkbox" value="' + esc(x.v) + '" checked> ' + esc(x.t) + '</label>').join('');
    const marcadas = id => Array.from(document.querySelectorAll('#' + id + ' input:checked')).map(c => c.value);
    google.script.run.withSuccessHandler(l => { const ls = l.map(i => ({ v: i.id, t: (i.nombre_propio || i.nombre_completo) + (i.codigo_iso ? ' (' + i.codigo_iso + ')' : '') }));
      $('lenguasExp').innerHTML = casillas(ls); $('lenguasHoja').innerHTML = casillas(ls); }).listarIdiomas();
    function hoja(){
      $('r4').textContent = tt('inter.preparando');
      google.script.run.withFailureHandler(err('r4')).withSuccessHandler(r => {
        $('r4').innerHTML = '<div class="ok">✅ <a href="' + esc(r.url) + '" target="_blank">' + esc(r.nombre) + '</a><br>' + th('exportarHoja.listo', {n: r.palabras}) +
          (r.omitidas ? '<br>' + th('inter.noCompartibles', {n: r.omitidas}) : '') + '</div>';
      }).exportarHojaLenguas({ lenguas: marcadas('lenguasHoja'), alcance: val('halc'), soloCompartibles: $('hCompartir').checked });
    }
    function exportar(){
      $('r1').textContent = tt('inter.preparando');
      google.script.run.withFailureHandler(err('r1')).withSuccessHandler(r => {
        $('r1').innerHTML = '<div class="ok">✅ <a href="' + esc(r.url) + '" target="_blank">' + esc(r.nombre) + '</a><br>' +
          th('inter.resumenArchivo', {palabras: r.rep.palabras, traducciones: r.rep.traducciones, imagenes: r.rep.imagenes}) + (r.rep.audios ? th('inter.resumenAudios', {n: r.rep.audios}) : '') + '.' +
          (r.rep.audiosPorTiempo ? '<br>' + th('inter.audiosTiempo', {n: r.rep.audiosPorTiempo}) : '') +
          (r.rep.noCompartibles ? '<br>' + th('inter.noCompartibles', {n: r.rep.noCompartibles}) : '') +
          '<br><small>' + th('inter.comparteArchivo') + '</small></div>';
      }).servicioExportarIntercambio({ alcance: val('alc'), compartirMedios: $('medios').checked, lenguas: marcadas('lenguasExp') });
    }
    $('archivo').addEventListener('change', e => {
      const f = e.target.files[0]; $('bImp').disabled = true; texto = '';
      if (!f) return;
      const lector = new FileReader(); lector.onload = () => { texto = lector.result; $('r2').textContent = tt('inter.archivoListo');
        let ids = [];
        try { ids = (JSON.parse(texto).idiomas || []).map(i => ({ v: i.id_idioma, t: (i.nombre_propio || i.nombre_completo || i.id_idioma) + (i.codigo_iso ? ' (' + i.codigo_iso + ')' : '') })); } catch (x) {}
        $('lenguasImp').innerHTML = ids.length ? '<b>' + th('inter.queLenguasTraer') + '</b>' + casillas(ids) : '';
      }; lector.readAsText(f);
    });
    function importar(aplicar){
      if (!texto) { $('r2').innerHTML = '<div class="error">' + th('inter.eligeArchivo') + '</div>'; return; }
      $('r2').textContent = tt(aplicar ? 'inter.agregando' : 'inter.revisando');
      google.script.run.withFailureHandler(err('r2')).withSuccessHandler(r => {
        const v = r.traducciones.variantes;
        $('r2').innerHTML = '<div class="ok"><b>' + th(r.aplicado ? 'inter.seAgregaron' : 'inter.seAgregaria', {origen: r.origen}) + '</b><ul>' +
          '<li>' + th('inter.lenguas', {n: r.idiomas.reconocidos}) + (r.idiomas.nuevos.length ? th('inter.lenguasNuevas', {n: r.idiomas.nuevos.length, nombres: r.idiomas.nuevos.join(', ')}) : '') + '</li>' +
          '<li>' + th('inter.categorias', {n: r.categorias.nuevas}) + '</li>' +
          '<li>' + th('inter.palabras', {n: r.palabras.nuevas, ya: r.palabras.porId + r.palabras.unidasPorTexto}) + '</li>' +
          '<li>' + th('inter.traducciones', {n: r.traducciones.nuevas, iguales: r.traducciones.iguales}) + '</li>' +
          (r.relaciones.nuevas ? '<li>' + th('inter.relaciones', {n: r.relaciones.nuevas}) + '</li>' : '') + '</ul>' +
          (v.length ? '<small>' + th('inter.variantes', {n: v.length, lista: v.slice(0, 12).map(x => x.texto).join(' · ')}) + '</small><br>' : '') +
          (r.palabras.ambiguas.length ? '<small>' + th('inter.ambiguas', {n: r.palabras.ambiguas.length, lista: r.palabras.ambiguas.slice(0, 8).join(' · ')}) + '</small><br>' : '') +
          '<small>' + th('inter.quedanComo', {estado: r.estado}) + '</small></div>';
        $('bImp').disabled = !!r.aplicado;
      }).servicioImportarIntercambio(texto, { unirPorTexto: $('unir').checked, estado: $('estado').value, lenguas: marcadas('lenguasImp') }, aplicar);
    }
    function grafo(f){
      $('r3').textContent = tt('inter.generando');
      google.script.run.withFailureHandler(err('r3')).withSuccessHandler(r => {
        $('r3').innerHTML = '<div class="ok">' + th('inter.grafoListo', {nodos: r.nodos, aristas: r.aristas}) + '<br>' + r.archivos.map(a => '<a href="' + esc(a.url) + '" target="_blank">' + esc(a.nombre) + '</a>').join('<br>') + '</div>';
      }).servicioExportarGrafo(f, val('galc'));
    }
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(680), tx('inter.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 15. EQUIPO: ADMINISTRADORES Y CAPTURISTAS
// ─────────────────────────────────────────────────────────────────────────────
// La hoja EQUIPO dice quién es quién. AppSheet la lee para dar permisos (fórmulas abajo) y
// el panel 👥 Equipo la mantiene: nadie la edita a mano.
//  - Capturista: agrega y edita palabras en AppSheet. No necesita acceso a la hoja de cálculo.
//  - Administrador: además borra, cambia estados y publica. El panel le da acceso de editor a la
//    hoja y a la carpeta del catálogo (para publicar imágenes y audios).
// El acceso a la app de AppSheet se da en AppSheet → Share (no se puede automatizar sin un plan de pago).

const COLS_EQUIPO = ['correo', 'nombre', 'rol', 'activo', 'notas'];
const ROLES_EQUIPO = ['Administrador', 'Capturista'];
const LIMITE_APPSHEET_GRATIS = 10;

/** Fórmulas que usa la app plantilla de AppSheet (se muestran en el panel para quien la prepara). */
const FORMULAS_APPSHEET = {
  esAdmin: 'IN(LOWER(USEREMAIL()), SELECT(EQUIPO[correo], AND([rol] = "Administrador", [activo] = "Sí")))',
  // Quien no está activo en EQUIPO solo puede ver, aunque tenga acceso a la app desde AppSheet → Share
  permisos: 'IF(IN(LOWER(USEREMAIL()), SELECT(EQUIPO[correo], AND([rol] = "Administrador", [activo] = "Sí"))), "ALL_CHANGES", ' +
            'IF(IN(LOWER(USEREMAIL()), SELECT(EQUIPO[correo], [activo] = "Sí")), "ADDS_AND_UPDATES", "READ_ONLY"))'
};

function _correoActual() {
  let c = '';
  try { c = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!c) { try { c = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  return _k(c);
}

function _rolEquipo(v) { return /^admin/.test(_sinAcentos(v)) ? 'Administrador' : 'Capturista'; }
function _activo(v) { return _sinAcentos(v) !== 'no'; }

function _hojaEquipo(ss) {
  let t = _tabla(ss, CONFIG.SHEETS.EQUIPO);
  if (!t) { asegurarEstructura(ss); t = _tabla(ss, CONFIG.SHEETS.EQUIPO); }
  return t;
}

function equipoListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = _hojaEquipo(ss);
  let duenio = '';
  try { duenio = _k(ss.getOwner() && ss.getOwner().getEmail()); } catch (e) {}
  const lista = t.filas.filter(f => _s(_get(t, f, 'correo'))).map(f => ({
    correo: _k(_get(t, f, 'correo')), nombre: _s(_get(t, f, 'nombre')), rol: _rolEquipo(_get(t, f, 'rol')),
    activo: _activo(_get(t, f, 'activo')), notas: _s(_get(t, f, 'notas'))
  }));
  return { lista, yo: _correoActual(), duenio, formulas: FORMULAS_APPSHEET };
}

function _adminsActivos(t, excepto) {
  return t.filas.filter(f => _k(_get(t, f, 'correo')) && _k(_get(t, f, 'correo')) !== excepto &&
    _rolEquipo(_get(t, f, 'rol')) === 'Administrador' && _activo(_get(t, f, 'activo'))).length;
}

/** Da o quita acceso de editor a la hoja y a la carpeta del catálogo. Devuelve avisos legibles. */
function _accesoEditor(ss, correo, dar) {
  const avisos = [];
  let duenio = '';
  try { duenio = _k(ss.getOwner() && ss.getOwner().getEmail()); } catch (e) {}
  if (!dar && (correo === duenio || correo === _correoActual())) return avisos; // nunca se quita a quien es dueño ni a uno mismo
  try { dar ? ss.addEditor(correo) : ss.removeEditor(correo); avisos.push(tx(dar ? 'equipo.aviso.compartida' : 'equipo.aviso.retirada')); }
  catch (e) { avisos.push(tx(dar ? 'equipo.aviso.noCompartir' : 'equipo.aviso.noDejar', { m: e.message })); }
  _cfgMemo = null;
  const raiz = getAppConfig(ss).raiz;
  if (raiz) {
    try { const c = DriveApp.getFolderById(raiz); dar ? c.addEditor(correo) : c.removeEditor(correo); }
    catch (e) { avisos.push(tx(dar ? 'equipo.aviso.noCarpeta' : 'equipo.aviso.noDejarCarpeta', { m: e.message })); }
  }
  return avisos;
}

/** Agrega o actualiza a una persona: { correo, nombre, rol, activo }. */
function equipoGuardar(p) {
  const correo = _k(p && p.correo);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) throw new Error(tx('equipo.error.correo'));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const t = _hojaEquipo(ss);
    const i = t.filas.findIndex(f => _k(_get(t, f, 'correo')) === correo);
    const rol = _rolEquipo(p.rol), activo = p.activo === false || p.activo === 'No' ? 'No' : 'Sí';
    const antes = i > -1 ? { admin: _rolEquipo(_get(t, t.filas[i], 'rol')) === 'Administrador' && _activo(_get(t, t.filas[i], 'activo')) } : { admin: false };
    const ahoraAdmin = rol === 'Administrador' && activo === 'Sí';
    if (antes.admin && !ahoraAdmin && _adminsActivos(t, correo) === 0) throw new Error(tx('equipo.error.ultimoCambiar'));
    const datos = { correo, nombre: _s(p.nombre), rol, activo };
    if (i > -1) _actualizar(t, i, datos);
    else _agregar(t, [Object.assign(datos, { notas: tx('equipo.notaAlta', { fecha: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') }) })]);
    let avisos = [];
    const activos = t.filas.filter(f => _k(_get(t, f, 'correo')) && _activo(_get(t, f, 'activo'))).length;
    if (activos > LIMITE_APPSHEET_GRATIS) avisos.push(tx('equipo.aviso.limite', { n: activos, limite: LIMITE_APPSHEET_GRATIS }));
    if (ahoraAdmin && !antes.admin) avisos = avisos.concat(_accesoEditor(ss, correo, true));
    if (!ahoraAdmin && antes.admin) avisos = avisos.concat(_accesoEditor(ss, correo, false));
    if (ahoraAdmin !== antes.admin) { try { _sincronizarAccesoDirectorio(ss); } catch (e) {} }
    return Object.assign(equipoListar(), { avisos, guardado: correo, nuevo: i < 0 });
  } finally { lock.releaseLock(); }
}

/** Quita a una persona del equipo (y su acceso de editor, si era administradora). */
function equipoQuitar(correo) {
  correo = _k(correo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const t = _hojaEquipo(ss);
    const i = t.filas.findIndex(f => _k(_get(t, f, 'correo')) === correo);
    if (i < 0) return equipoListar();
    const eraAdmin = _rolEquipo(_get(t, t.filas[i], 'rol')) === 'Administrador' && _activo(_get(t, t.filas[i], 'activo'));
    if (eraAdmin && _adminsActivos(t, correo) === 0) throw new Error(tx('equipo.error.ultimoQuitar'));
    _espacioParaBorrar(t.sh, 1);
    t.sh.deleteRow(i + 2);
    const avisos = eraAdmin ? _accesoEditor(ss, correo, false) : [];
    if (eraAdmin) { try { _sincronizarAccesoDirectorio(ss); } catch (e) {} }
    return Object.assign(equipoListar(), { avisos });
  } finally { lock.releaseLock(); }
}

function uiEquipo() {
  const roles = ROLES_EQUIPO.map(r => `<option value="${r}">${txh('rol.' + r)}</option>`).join('');
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:12px}
    table{width:100%;border-collapse:collapse;font-size:13px}th,td{border-bottom:1px solid #e2e8f0;padding:6px;text-align:left;vertical-align:middle}
    th{background:#f8fafc}input,select{padding:7px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px}
    button{padding:8px 11px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#e2e8f0}
    .prim{background:#0f766e;color:#fff}.peq{padding:4px 8px;font-size:12px}
    .fila{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin:8px 0}
    .error{background:#fee2e2;color:#991b1b;padding:8px;border-radius:6px;margin:8px 0}
    .aviso{background:#fef3c7;color:#92400e;padding:8px;border-radius:6px;margin:8px 0}
    small{color:#64748b}code{background:#f1f5f9;padding:2px 4px;border-radius:4px;font-size:11px;display:block;white-space:pre-wrap;word-break:break-all;margin:4px 0}
    .yo{color:#0f766e;font-weight:bold}
  </style>
  ${_tScript()}
  <div class="caja">
    <p style="margin-top:0">${txh('equipo.intro', { limite: LIMITE_APPSHEET_GRATIS })}</p>
    <div id="lista">${txh('comun.cargando')}</div>
    <h3 style="margin-bottom:6px">${txh('equipo.agregar')}</h3>
    <div class="fila">
      <input id="correo" type="email" placeholder="correo@gmail.com" style="flex:2;min-width:180px" aria-label="${txh('equipo.correo')}">
      <input id="nombre" placeholder="${txh('equipo.nombre')}" style="flex:1;min-width:120px" aria-label="${txh('equipo.nombre')}">
      <select id="rol" aria-label="${txh('equipo.rol')}">${roles}</select>
      <button class="prim" onclick="guardar()">${txh('equipo.guardar')}</button>
    </div>
    <div id="msg" role="status" aria-live="polite"></div>
    <div class="aviso">${txh('equipo.share')}</div>
    <details><summary><small>${txh('equipo.formulas')}</small></summary>
      <small>${txh('equipo.formula.permisos')}</small><code id="f1"></code>
      <small>${txh('equipo.formula.estado')}</small><code id="f2"></code>
      <small>${txh('equipo.formula.inicial')}</small><code>"En revisión"</code>
      <small>${txh('equipo.formula.tabla')}</small>
    </details>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    const ROLES = ${JSON.stringify(ROLES_EQUIPO)};
    const rolSel = (i, actual) => '<select data-i="' + i + '" class="cambiaRol">' + ROLES.map(r => '<option value="' + r + '"' + (r === actual ? ' selected' : '') + '>' + th('rol.' + r) + '</option>').join('') + '</select>';
    function pintar(r) {
      $('f1').textContent = r.formulas.permisos; $('f2').textContent = r.formulas.esAdmin;
      if (!r.lista.length) { $('lista').innerHTML = '<p>' + th('equipo.nadie') + '</p>'; return; }
      $('lista').innerHTML = '<table><tr><th>' + th('equipo.col.persona') + '</th><th>' + th('equipo.col.rol') + '</th><th>' + th('equipo.col.activo') + '</th><th></th></tr>' + r.lista.map((p, i) =>
        '<tr><td>' + (p.nombre ? '<b>' + esc(p.nombre) + '</b><br>' : '') + '<span class="' + (p.correo === r.yo ? 'yo' : '') + '">' + esc(p.correo) + (p.correo === r.yo ? ' ' + th('equipo.tu') : '') + (p.correo === r.duenio ? ' ' + th('equipo.duenio') : '') + '</span></td>' +
        '<td>' + rolSel(i, p.rol) + '</td>' +
        '<td><input type="checkbox" data-i="' + i + '" class="cambiaActivo"' + (p.activo ? ' checked' : '') + ' aria-label="' + th('equipo.col.activo') + '"></td>' +
        '<td><button class="peq" data-i="' + i + '" onclick="quitar(this)">' + th('equipo.quitar') + '</button></td></tr>').join('') + '</table>';
      window._lista = r.lista;
      document.querySelectorAll('.cambiaRol').forEach(s => s.onchange = () => { const p = window._lista[+s.dataset.i]; enviar({ correo: p.correo, nombre: p.nombre, rol: s.value, activo: p.activo }); });
      document.querySelectorAll('.cambiaActivo').forEach(c => c.onchange = () => { const p = window._lista[+c.dataset.i]; enviar({ correo: p.correo, nombre: p.nombre, rol: p.rol, activo: c.checked }); });
    }
    function resultado(r) {
      pintar(r);
      const av = (r.avisos || []).map(a => '<br>• ' + esc(a)).join('');
      $('msg').innerHTML = r.guardado || r.avisos ? '<div class="ok">' + th('equipo.guardado') + av + (r.nuevo ? '<br>' + th('equipo.agregaShare') : '') + '</div>' : '';
    }
    function fallo(e) { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; google.script.run.withSuccessHandler(pintar).equipoListar(); }
    function enviar(p) { $('msg').textContent = tt('comun.guardando'); google.script.run.withSuccessHandler(resultado).withFailureHandler(fallo).equipoGuardar(p); }
    function guardar() { enviar({ correo: $('correo').value, nombre: $('nombre').value, rol: $('rol').value, activo: true }); $('correo').value = ''; $('nombre').value = ''; }
    function quitar(b) {
      const p = window._lista[+b.dataset.i];
      if (!confirm(tt('equipo.confirmarQuitar', {correo: p.correo}))) return;
      $('msg').textContent = tt('equipo.quitando');
      google.script.run.withSuccessHandler(r => { resultado(r); $('msg').innerHTML += '<div class="aviso">' + th('equipo.quitaShare') + '</div>'; }).withFailureHandler(fallo).equipoQuitar(p.correo);
    }
    google.script.run.withSuccessHandler(pintar).withFailureHandler(fallo).equipoListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(720).setHeight(600), tx('equipo.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 16. IDIOMA DE LA INTERFAZ (menús, ventanas y avisos)
// ─────────────────────────────────────────────────────────────────────────────
// Todos los textos que ve quien administra tienen un código estable (ej. 'panel.publicar').
// El español vive en T_ES (abajo); las traducciones viven en la hoja TEXTOS, una columna por
// idioma de interfaz, y se pueden instalar desde el archivo que comparta otra comunidad.
//  - Cada persona elige su idioma (se guarda solo para ella); el catálogo tiene uno predeterminado.
//  - Si a un texto le falta traducción, se muestra en el predeterminado y, si tampoco, en español.
//  - Las traducciones son texto plano: *así* se marca una negrita y {n} es un dato que se rellena.
//    Nunca se interpreta HTML de una traducción (pueden venir de otras comunidades).
// Reglas para quien programe: frases completas con {huecos}, nunca pedazos pegados; no cambiar el
// sentido de un código existente (si cambia el sentido, código nuevo).

const FORMATO_INTERFAZ = 'catalogo-lenguas-interfaz';
const PROP_IDIOMA_INTERFAZ = 'IDIOMA_INTERFAZ';
let _dicMemo = null;

/** Idioma de interfaz de esta persona → el predeterminado del catálogo → español. */
function _idiomaInterfaz() {
  let c = '';
  try { c = PropertiesService.getUserProperties().getProperty(PROP_IDIOMA_INTERFAZ) || ''; } catch (e) { /* activador simple sin acceso */ }
  return _k(c) || _idiomaInterfazCatalogo() || 'es';
}
function _idiomaInterfazCatalogo() {
  try { return _k(getAppConfig(SpreadsheetApp.getActiveSpreadsheet()).idiomaInterfaz); } catch (e) { return ''; }
}

/** Lee la hoja TEXTOS: { codigos, nombres: {codigo: nombre}, cols: {codigo: {clave: texto}}, filas: {clave: índice} }. */
function _leerTextos(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.TEXTOS);
  if (!sh || sh.getLastRow() < 1) return null;
  const d = sh.getDataRange().getValues();
  const h = d[0].map(x => _k(x));
  const iClave = h.indexOf('clave');
  if (iClave < 0) return null;
  const codigos = h.map((c, i) => i !== iClave && c ? c : '').filter(Boolean);
  const r = { sh, h, iClave, codigos, nombres: {}, cols: {}, filas: {} };
  codigos.forEach(c => { r.cols[c] = {}; });
  d.slice(1).forEach((fila, i) => {
    const clave = _s(fila[iClave]);
    if (!clave) return;
    r.filas[clave] = i + 2;
    h.forEach((c, j) => {
      if (!r.cols[c]) return;
      const v = _s(fila[j]);
      if (clave === '_nombre') r.nombres[c] = v;
      else if (v) r.cols[c][clave] = v;
    });
  });
  return r;
}

/** Diccionario ya resuelto para esta persona (con las caídas a otro idioma). */
function _diccionario() {
  if (_dicMemo) return _dicMemo;
  const L = _idiomaInterfaz(), pred = _idiomaInterfazCatalogo();
  const d = Object.assign({}, T_ES);
  let tabla = null;
  try { tabla = _leerTextos(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { tabla = null; }
  // Quien eligió español lo ve todo en español; si eligió otra lengua, lo que falte sale en la del catálogo y luego en español
  if (tabla) (L === 'es' ? [] : [pred, L]).forEach(c => {
    if (!c || c === 'es' || !tabla.cols[c]) return;
    Object.keys(T_ES).forEach(k => { if (tabla.cols[c][k]) d[k] = tabla.cols[c][k]; });
  });
  _dicMemo = { L, d };
  return _dicMemo;
}

function _rellenar(s, v, escapar) {
  if (!v) return s;
  return s.replace(/\{(\w+)\}/g, (m, x) => (x in v) ? (escapar ? _esc(v[x]) : String(v[x] == null ? '' : v[x])) : m);
}
/** Texto plano (avisos, menús, hojas): sin marcas. */
function tx(clave, v) {
  const s = _diccionario().d[clave];
  return _rellenar(s == null ? clave : String(s), v, false).replace(/\*([^*\n]+)\*/g, '$1');
}
/** Texto para HTML: escapado, *negritas* y saltos de línea. */
function txh(clave, v) {
  const s = _diccionario().d[clave];
  return _rellenar(_esc(s == null ? clave : String(s)), v, true).replace(/\*([^*\n]+)\*/g, '<b>$1</b>').replace(/\n/g, '<br>');
}
/** Todas las formas conocidas de un texto (en cualquier idioma instalado): para leer hojas que hizo otra persona. */
function _todasLasFormas(clave) {
  const formas = new Set([_k(T_ES[clave] || '')]);
  try {
    const tabla = _leerTextos(SpreadsheetApp.getActiveSpreadsheet());
    if (tabla) tabla.codigos.forEach(c => { if (tabla.cols[c][clave]) formas.add(_k(tabla.cols[c][clave])); });
  } catch (e) { /* sin hoja de textos */ }
  formas.delete('');
  return [...formas];
}

/** Bloque para las ventanas: el diccionario resuelto y las funciones tt (texto) y th (HTML), iguales a tx / txh del servidor. */
function _tScript() {
  const json = JSON.stringify(_diccionario().d).replace(/</g, '\\u003c');
  return `<script>
    const T = ${json};
    const _escT = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const _rell = (s, v, e) => !v ? s : s.replace(/\\{(\\w+)\\}/g, (m, x) => (x in v) ? (e ? _escT(v[x]) : String(v[x] == null ? '' : v[x])) : m);
    function tt(k, v) { return _rell(T[k] == null ? k : String(T[k]), v, false).replace(/\\*([^*\\n]+)\\*/g, '$1'); }
    function th(k, v) { return _rell(_escT(T[k] == null ? k : String(T[k])), v, true).replace(/\\*([^*\\n]+)\\*/g, '<b>$1</b>').replace(/\\n/g, '<br>'); }
  </script>`;
}

/** Deja la hoja TEXTOS al día con los textos de esta versión, sin tocar ninguna traducción. */
function _sincronizarTextos(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEETS.TEXTOS);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEETS.TEXTOS);
  const previa = _leerTextos(ss);
  const codigos = ['es'].concat(previa ? previa.codigos.filter(c => c !== 'es') : []);
  const claves = Object.keys(T_ES);
  const extra = previa ? Object.keys(previa.filas).filter(k => k !== '_nombre' && !(k in T_ES)) : []; // de otras versiones: se conservan
  const fila = (clave, base) => [clave, base].concat(codigos.slice(1).map(c => previa && previa.cols[c] ? (previa.cols[c][clave] || '') : ''));
  const datos = [['clave'].concat(codigos), ['_nombre', 'Español'].concat(codigos.slice(1).map(c => (previa && previa.nombres[c]) || c))]
    .concat(claves.map(k => fila(k, T_ES[k])))
    .concat(extra.map(k => fila(k, '')));
  sh.clearContents();
  sh.getRange(1, 1, datos.length, datos[0].length).setNumberFormat('@').setValues(datos);
  sh.getRange(1, 1, 2, datos[0].length).setFontWeight('bold').setBackground('#475569').setFontColor('white');
  sh.setFrozenRows(2); sh.setFrozenColumns(2);
  sh.setTabColor('#475569');
  _dicMemo = null;
  return claves.length;
}

// ── Ventana 🗣️ Idioma de los menús ───────────────────────────────────────────

function interfazListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _dicMemo = null;
  let tabla = _leerTextos(ss);
  if (!tabla) { _sincronizarTextos(ss); tabla = _leerTextos(ss); }
  const total = Object.keys(T_ES).length;
  let mio = '';
  try { mio = _k(PropertiesService.getUserProperties().getProperty(PROP_IDIOMA_INTERFAZ)); } catch (e) {}
  const idiomas = ['es'].concat(tabla.codigos.filter(c => c !== 'es')).map(c => ({
    codigo: c, nombre: c === 'es' ? 'Español' : (tabla.nombres[c] || c),
    traducidos: c === 'es' ? total : Object.keys(T_ES).filter(k => tabla.cols[c] && tabla.cols[c][k]).length, total
  }));
  return { actual: _idiomaInterfaz(), mio, predeterminado: _idiomaInterfazCatalogo() || 'es', idiomas, total };
}

/** paraTodos = también como idioma predeterminado del catálogo. */
function interfazElegir(codigo, paraTodos) {
  codigo = _k(codigo) || 'es';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabla = _leerTextos(ss);
  if (codigo !== 'es' && (!tabla || !tabla.cols[codigo])) throw new Error(tx('interfaz.error.noInstalado'));
  PropertiesService.getUserProperties().setProperty(PROP_IDIOMA_INTERFAZ, codigo);
  if (paraTodos) _escribirConfig(ss, { IDIOMA_INTERFAZ: codigo });
  _dicMemo = null; _cfgMemo = null;
  try { _construirMenu(); } catch (e) { /* se verá al volver a abrir la hoja */ }
  return interfazListar();
}

function _codigoInterfaz(codigo) {
  const c = _k(codigo).replace(/[^a-z0-9-]/g, '');
  if (!/^[a-z]{2,8}(-[a-z0-9]{1,8})?$/.test(c)) throw new Error(tx('interfaz.error.codigo'));
  return c;
}

function interfazCrear(codigo, nombre) {
  codigo = _codigoInterfaz(codigo);
  if (!_s(nombre)) throw new Error(tx('interfaz.error.nombre'));
  const lock = _lock();
  try { _crearColumnaInterfaz(SpreadsheetApp.getActiveSpreadsheet(), codigo, nombre); return interfazListar(); }
  finally { lock.releaseLock(); }
}

function _crearColumnaInterfaz(ss, codigo, nombre) {
  let tabla = _leerTextos(ss);
  if (!tabla) { _sincronizarTextos(ss); tabla = _leerTextos(ss); }
  if (codigo === 'es' || tabla.cols[codigo]) throw new Error(tx('interfaz.error.yaExiste'));
  const col = tabla.h.length + 1;
  tabla.sh.getRange(1, col).setValue(codigo);
  tabla.sh.getRange(2, col).setValue(_s(nombre).slice(0, 60));
  tabla.sh.getRange(1, col, 2, 1).setFontWeight('bold').setBackground('#475569').setFontColor('white');
  _dicMemo = null;
}

/** Para el editor: cada texto en español y su traducción actual. */
function interfazTextos(codigo) {
  codigo = _k(codigo);
  const tabla = _leerTextos(SpreadsheetApp.getActiveSpreadsheet());
  if (!tabla || !tabla.cols[codigo] || codigo === 'es') throw new Error(tx('interfaz.error.noInstalado'));
  return Object.keys(T_ES).map(k => ({ clave: k, es: T_ES[k], tr: tabla.cols[codigo][k] || '' }));
}

function _limpiarTraduccion(v) { return _s(v).replace(/\r\n?/g, '\n').slice(0, 1500); }

/** cambios = { clave: texto }. Un texto vacío borra esa traducción (se volverá a ver en el idioma de respaldo). */
function interfazGuardar(codigo, cambios) {
  codigo = _k(codigo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const tabla = _leerTextos(ss);
    if (!tabla || !tabla.cols[codigo] || codigo === 'es') throw new Error(tx('interfaz.error.noInstalado'));
    const col = tabla.h.indexOf(codigo) + 1;
    let n = 0;
    Object.keys(cambios || {}).forEach(k => {
      if (!(k in T_ES) || !tabla.filas[k]) return;
      tabla.sh.getRange(tabla.filas[k], col).setNumberFormat('@').setValue(_limpiarTraduccion(cambios[k]));
      n++;
    });
    _dicMemo = null;
    if (codigo === _idiomaInterfaz()) { try { _construirMenu(); } catch (e) {} }
    return { n, lista: interfazListar() };
  } finally { lock.releaseLock(); }
}

function interfazExportar(codigo) {
  codigo = _k(codigo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabla = _leerTextos(ss);
  if (!tabla || !tabla.cols[codigo] || codigo === 'es') throw new Error(tx('interfaz.error.noInstalado'));
  const cfg = getAppConfig(ss);
  const textos = {};
  Object.keys(T_ES).forEach(k => { if (tabla.cols[codigo][k]) textos[k] = tabla.cols[codigo][k]; });
  const paquete = {
    formato: FORMATO_INTERFAZ, version: CONFIG.VERSION, creado: new Date().toISOString(),
    codigo, nombre: tabla.nombres[codigo] || codigo, origen: { titulo: cfg.titulo, idCatalogo: cfg.idCatalogo },
    total: Object.keys(T_ES).length, textos
  };
  const nombre = `Menus_${(tabla.nombres[codigo] || codigo).replace(/[^\p{L}\p{N}]+/gu, '_')}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.json`;
  const archivo = _carpetaExportaciones(ss).createFile(nombre, JSON.stringify(paquete, null, 1), 'application/json');
  return { url: archivo.getUrl(), nombre, n: Object.keys(textos).length, total: paquete.total };
}

/**
 * Instala (o solo muestra qué haría, si aplicar = false) una traducción compartida.
 * Por omisión solo llena lo que falta: lo que esta comunidad ya tradujo no se toca, salvo reemplazar = true.
 * Los textos de otras versiones que aquí no existen se ignoran; los que faltan se verán en el idioma de respaldo.
 */
function interfazInstalar(texto, reemplazar, aplicar) {
  let p;
  try { p = JSON.parse(texto); } catch (e) { throw new Error(tx('interfaz.error.archivo')); }
  if (!p || p.formato !== FORMATO_INTERFAZ || !p.textos || typeof p.textos !== 'object') throw new Error(tx('interfaz.error.archivo'));
  const codigo = _codigoInterfaz(p.codigo);
  if (codigo === 'es') throw new Error(tx('interfaz.error.espanol'));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = aplicar ? _lock() : null;
  try {
    let tabla = _leerTextos(ss);
    if (!tabla) { _sincronizarTextos(ss); tabla = _leerTextos(ss); }
    const existe = !!tabla.cols[codigo];
    const actual = existe ? tabla.cols[codigo] : {};
    const rep = { codigo, nombre: _s(p.nombre).slice(0, 60) || codigo, existe, nuevos: 0, iguales: 0, distintos: 0, reemplazados: 0, desconocidos: 0, faltan: 0 };
    const escribir = {};
    Object.keys(p.textos).forEach(k => {
      const v = _limpiarTraduccion(p.textos[k]);
      if (!(k in T_ES)) { rep.desconocidos++; return; }
      if (!v) return;
      if (!actual[k]) { rep.nuevos++; escribir[k] = v; }
      else if (actual[k] === v) rep.iguales++;
      else { rep.distintos++; if (reemplazar) { rep.reemplazados++; escribir[k] = v; } }
    });
    rep.faltan = Object.keys(T_ES).filter(k => !actual[k] && !escribir[k]).length;
    if (aplicar) {
      if (!existe) { _crearColumnaInterfaz(ss, codigo, rep.nombre); tabla = _leerTextos(ss); }
      const col = tabla.h.indexOf(codigo) + 1;
      const filas = tabla.sh.getRange(1, col, tabla.sh.getLastRow(), 1).getValues();
      Object.keys(escribir).forEach(k => { if (tabla.filas[k]) filas[tabla.filas[k] - 1][0] = escribir[k]; });
      tabla.sh.getRange(1, col, filas.length, 1).setNumberFormat('@').setValues(filas);
      _dicMemo = null;
      if (codigo === _idiomaInterfaz()) { try { _construirMenu(); } catch (e) {} }
    }
    rep.aplicado = !!aplicar;
    return rep;
  } finally { if (lock) lock.releaseLock(); }
}

function uiIdiomaMenus() {
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:14px}
    .sec{border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:12px;background:#f8fafc}
    h3{margin:0 0 6px;font-size:15px}
    label{display:block;margin:4px 0}
    select,input[type=text],textarea{padding:7px;border:1px solid #cbd5e1;border-radius:6px;font-size:14px;box-sizing:border-box}
    button{padding:8px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#e2e8f0;margin:6px 4px 0 0}
    .prim{background:#475569;color:#fff}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:8px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
    small{color:#64748b}
    table{width:100%;border-collapse:collapse;font-size:13px}td,th{border-bottom:1px solid #e2e8f0;padding:6px;text-align:left;vertical-align:top}
    td.es{width:45%;color:#334155}td textarea{width:100%;min-height:38px;font-family:inherit}
    .clave{color:#94a3b8;font-size:11px;display:block}
    .scroll{max-height:320px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:6px;background:#fff}
    .fila{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
  </style>
  ${_tScript()}
  <div class="caja">
    <p style="margin-top:0">${txh('interfaz.intro')}</p>
    <div class="sec">
      <h3>${txh('interfaz.elegir')}</h3>
      <div id="actual"></div>
      <div class="fila"><select id="usar"></select><button class="prim" onclick="elegir()">${txh('interfaz.paraMi')}</button></div>
      <label><input type="checkbox" id="todos"> ${txh('interfaz.paraTodos')} <small>${txh('interfaz.paraTodos.d')}</small></label>
      <div id="m1" role="status" aria-live="polite"></div>
    </div>
    <div class="sec">
      <h3>${txh('interfaz.traducir')}</h3>
      <p style="margin:0 0 6px"><small>${txh('interfaz.traducir.d')}</small></p>
      <div class="fila"><label for="cual" style="display:inline">${txh('interfaz.cual')}</label> <select id="cual" onchange="cargar()"></select></div>
      <details><summary>${txh('interfaz.nueva')}</summary>
        <div class="fila"><input type="text" id="nCodigo" placeholder="${txh('interfaz.codigo')}" style="width:180px"><input type="text" id="nNombre" placeholder="${txh('interfaz.nombre')}" style="flex:1"><button onclick="crear()">${txh('interfaz.crear')}</button></div>
      </details>
      <div class="fila" style="margin-top:6px"><label for="parte" style="display:inline">${txh('interfaz.parte')}</label>
        <select id="parte" onchange="limite=150;pintarTabla()"><option value="">${txh('interfaz.parte.todo')}</option><option value="hoja">${txh('interfaz.parte.hoja')}</option><option value="visor">${txh('interfaz.parte.visor')}</option></select></div>
      <p style="margin:4px 0"><small>${txh('interfaz.visor.d')}</small></p>
      <div class="fila" style="margin-top:6px"><label style="display:inline"><input type="checkbox" id="faltan" checked onchange="pintarTabla()"> ${txh('interfaz.soloFaltan')}</label>
        <input type="text" id="buscar" placeholder="${txh('interfaz.buscar')}" oninput="pintarTabla()" style="flex:1;min-width:140px"></div>
      <div class="scroll" style="margin-top:6px"><table><thead><tr><th>${txh('interfaz.col.espanol')}</th><th>${txh('interfaz.col.traduccion')}</th></tr></thead><tbody id="filas"></tbody></table></div>
      <div class="fila"><small id="cuenta"></small><button id="mas" style="display:none" onclick="limite+=150;pintarTabla()">${txh('interfaz.mas')}</button></div>
      <button class="prim" onclick="guardar()">${txh('interfaz.guardar')}</button>
      <div id="m2" role="status" aria-live="polite"></div>
    </div>
    <div class="sec">
      <h3>${txh('interfaz.compartir')}</h3>
      <p style="margin:0 0 6px"><small>${txh('interfaz.compartir.d')}</small></p>
      <button onclick="exportar()">${txh('interfaz.exportar')}</button>
      <div style="margin-top:8px"><b>${txh('interfaz.instalar')}</b></div>
      <input type="file" id="archivo" accept=".json,application/json">
      <label><input type="checkbox" id="reemplazar"> ${txh('interfaz.reemplazar')} <small>${txh('interfaz.reemplazar.d')}</small></label>
      <button onclick="instalar(false)">${txh('interfaz.ver')}</button><button class="prim" id="bInst" onclick="instalar(true)" disabled>${txh('interfaz.instalarBoton')}</button>
      <div id="m3" role="status" aria-live="polite"></div>
    </div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    let info = null, textos = [], cambios = {}, limite = 150, archivo = '';
    const ok = (id, h) => { $(id).innerHTML = '<div class="ok">' + h + '</div>'; };
    const err = id => e => { $(id).innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    function pintarInfo(r) {
      info = r;
      const nombre = c => (r.idiomas.find(i => i.codigo === c) || {nombre: c}).nombre;
      $('actual').innerHTML = '<p>' + th('interfaz.actual', {nombre: nombre(r.actual), pred: nombre(r.predeterminado)}) + '</p>';
      $('usar').innerHTML = r.idiomas.map(i => '<option value="' + esc(i.codigo) + '"' + (i.codigo === r.actual ? ' selected' : '') + '>' + esc(tt('interfaz.avance', {nombre: i.nombre, n: i.traducidos, total: i.total})) + '</option>').join('');
      const trad = r.idiomas.filter(i => i.codigo !== 'es');
      const previo = $('cual').value;
      $('cual').innerHTML = trad.map(i => '<option value="' + esc(i.codigo) + '">' + esc(tt('interfaz.avance', {nombre: i.nombre, n: i.traducidos, total: i.total})) + '</option>').join('');
      if (previo && trad.some(i => i.codigo === previo)) $('cual').value = previo;
      if (!trad.length) { textos = []; $('filas').innerHTML = '<tr><td colspan="2"><small>' + th('interfaz.nadaQueTraducir') + '</small></td></tr>'; $('cuenta').textContent = ''; }
    }
    function elegir() {
      google.script.run.withFailureHandler(err('m1')).withSuccessHandler(r => { pintarInfo(r); ok('m1', th('interfaz.cambiado')); })
        .interfazElegir($('usar').value, $('todos').checked);
    }
    function crear() {
      google.script.run.withFailureHandler(err('m2')).withSuccessHandler(r => { pintarInfo(r); $('cual').value = $('nCodigo').value.trim().toLowerCase(); $('nCodigo').value = ''; $('nNombre').value = ''; cargar(); })
        .interfazCrear($('nCodigo').value, $('nNombre').value);
    }
    function cargar() {
      cambios = {}; limite = 150;
      const c = $('cual').value; if (!c) return;
      $('filas').innerHTML = '<tr><td colspan="2">' + th('comun.cargando') + '</td></tr>';
      google.script.run.withFailureHandler(err('m2')).withSuccessHandler(l => { textos = l; pintarTabla(); }).interfazTextos(c);
    }
    function pintarTabla() {
      const q = $('buscar').value.toLowerCase().trim(), solo = $('faltan').checked, parte = $('parte').value;
      const v = textos.filter(x => (!parte || (parte === 'visor') === x.clave.startsWith('visor.')) && (!solo || !(x.clave in cambios ? cambios[x.clave] : x.tr)) && (!q || (x.es + ' ' + x.tr).toLowerCase().includes(q)));
      $('filas').innerHTML = v.slice(0, limite).map(x => '<tr><td class="es">' + esc(x.es) + '<span class="clave">' + esc(x.clave) + '</span></td><td><textarea data-k="' + esc(x.clave) + '">' +
        esc(x.clave in cambios ? cambios[x.clave] : x.tr) + '</textarea></td></tr>').join('');
      document.querySelectorAll('#filas textarea').forEach(a => a.oninput = () => { cambios[a.dataset.k] = a.value; });
      $('cuenta').textContent = tt('interfaz.mostrando', {n: Math.min(v.length, limite), total: v.length});
      $('mas').style.display = v.length > limite ? 'inline-block' : 'none';
    }
    function guardar() {
      const c = $('cual').value, n = Object.keys(cambios).length;
      if (!c || !n) { ok('m2', th('interfaz.sinCambios')); return; }
      $('m2').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(err('m2')).withSuccessHandler(r => {
        textos.forEach(x => { if (x.clave in cambios) x.tr = cambios[x.clave].trim(); }); cambios = {};
        pintarInfo(r.lista); $('cual').value = c; pintarTabla(); ok('m2', th('interfaz.guardadas', {n: r.n}));
      }).interfazGuardar(c, cambios);
    }
    function exportar() {
      const c = $('cual').value; if (!c) { err('m3')({message: tt('interfaz.nadaQueTraducir')}); return; }
      google.script.run.withFailureHandler(err('m3')).withSuccessHandler(r => ok('m3', '<a href="' + esc(r.url) + '" target="_blank">' + esc(r.nombre) + '</a><br>' + th('interfaz.exportado', {n: r.n, total: r.total}))).interfazExportar(c);
    }
    $('archivo').addEventListener('change', e => {
      const f = e.target.files[0]; $('bInst').disabled = true; archivo = '';
      if (!f) return;
      const lector = new FileReader(); lector.onload = () => { archivo = lector.result; }; lector.readAsText(f);
    });
    function instalar(aplicar) {
      if (!archivo) { err('m3')({message: tt('interfaz.eligeArchivo')}); return; }
      google.script.run.withFailureHandler(err('m3')).withSuccessHandler(r => {
        if (r.aplicado) { ok('m3', th('interfaz.instalada', {nombre: r.nombre})); $('bInst').disabled = true;
          google.script.run.withSuccessHandler(i => { pintarInfo(i); cargar(); }).interfazListar(); return; }
        ok('m3', th('interfaz.previa', r) + '<br><small>' + th('interfaz.previaExtra', r) + '</small>');
        $('bInst').disabled = false;
      }).interfazInstalar(archivo, $('reemplazar').checked, aplicar);
    }
    google.script.run.withFailureHandler(err('m1')).withSuccessHandler(r => { pintarInfo(r); cargar(); }).interfazListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(760).setHeight(720), tx('interfaz.titulo'));
}
// ─────────────────────────────────────────────────────────────────────────────
// 21. IMÁGENES Y SONIDOS: ASIGNAR, BANCO SIN ASIGNAR, REVISIÓN Y HOJAS PARA QUIEN ILUSTRA
// ─────────────────────────────────────────────────────────────────────────────
// Nadie tiene que buscar códigos: todo se elige por la palabra escrita en cualquier lengua.
//  - Una imagen es de la palabra (vale para todas las lenguas); un sonido es de una traducción (palabra + lengua).
//  - Sonidos en dos calidades: «consulta» (ligera, la que se publica) y «archivo» (alta calidad, privada).
//  - Lo que propone quien captura, o lo que llega de una hoja de quien ilustra, entra a MEDIOS_PROPUESTOS
//    como «Pendiente»; quien administra lo acepta o lo rechaza. Nada cambia en PALABRAS/TRADUCCIONES antes.
//  - Los archivos de Drive no se mueven ni se borran.

const COLS_MEDIOS = ['id_medio', 'tipo', 'id_palabra', 'id_idioma', 'enlace', 'calidad', 'credito', 'id_voz', 'origen',
  'propuesto_por', 'fecha', 'estado', 'revisado_por', 'fecha_revision', 'nota'];
const MEDIO_LIMITES = { imagenIdeal: 300 * 1024, imagenMax: 1024 * 1024, audioIdeal: 1024 * 1024, audioMax: 5 * 1024 * 1024 };
const PROP_HOJAS_ILUSTRACION = 'HOJAS_ILUSTRACION';
const ENC_ILUSTRACION = ['ID_PALABRA', 'CATEGORIA', 'PALABRA', 'ENLACE_IMAGEN', 'CREDITO', 'COMENTARIOS'];
const NOTAS_RAPIDAS = ['temporal', 'mejorar', 'comunidad', 'permiso', 'regrabar'];

// ── Créditos con lenguaje controlado: se eligen de 📚 Personas e instituciones (ACTORES) ──
function _actores(ss) {
  return readSheet(ss, CONFIG.SHEETS.ACTORES).filter(a => _s(a.id_actor) && _estadoCanonico(a.estado) !== 'retirado')
    .map(a => ({ id: _s(a.id_actor), nombre: _s(a.nombre), credito: _s(a.credito) }));
}
/** Devuelve el id de la persona o institución con ese nombre (o id). Si no existe y crear = true, la agrega. */
function _actorPorNombre(ss, texto, crear) {
  const v = _s(texto); if (!v) return { id: '', nuevo: false };
  const k = _plegar(v), l = _actores(ss);
  const a = l.find(x => _k(x.id) === _k(v) || _plegar(x.nombre) === k || (x.credito && _plegar(x.credito) === k));
  if (a) return { id: a.id, nuevo: false };
  if (!crear) return { id: '', nuevo: false };
  let t = _tabla(ss, CONFIG.SHEETS.ACTORES);
  if (!t) { asegurarEstructura(ss); t = _tabla(ss, CONFIG.SHEETS.ACTORES); }
  const id = Utilities.getUuid();
  _agregar(t, [{ id_actor: id, nombre: v.slice(0, 200), tipo: 'Persona', credito: v.slice(0, 200), notas: tx('medios.actorCreado'), estado: _etiquetaEstado('publicado') }]);
  return { id, nuevo: true };
}
/** Texto público del crédito: el «crédito» o el nombre de la persona o institución; si no es un id, el texto tal cual. */
function _creditoVisible(actores, v) {
  const a = actores.find(x => _k(x.id) === _k(v));
  return a ? (a.credito || a.nombre) : _s(v);
}

function _hojaMedios(ss) {
  let t = _tabla(ss, CONFIG.SHEETS.MEDIOS);
  if (!t) { asegurarEstructura(ss); t = _tabla(ss, CONFIG.SHEETS.MEDIOS); }
  return t;
}

/** Texto legible de cada palabra en cada lengua: { id_palabra: { id_idioma: {texto, fila} } } */
function _textosPorPalabra(tTr) {
  const r = {};
  if (tTr) tTr.filas.forEach((f, i) => {
    const p = _k(_get(tTr, f, 'id_palabra')), L = _k(_get(tTr, f, 'id_idioma')), x = _s(_get(tTr, f, 'texto'));
    if (!p || !L) return;
    (r[p] = r[p] || {})[L] = { texto: x, fila: i };
  });
  return r;
}

/** Archivo de Drive a partir de enlace, ID o nombre (o null si es una dirección web externa o no se encuentra). */
function _archivoMedio(enlace, tipo) {
  const v = _s(enlace);
  if (!v) return null;
  if (tipo === 'audio') return _buscarArchivoAudio(v);
  return _resolverArchivo(v);
}

/** Revisa el archivo contra las recomendaciones. Devuelve { avisos:[], error:'' }. */
function _revisarMedio(enlace, tipo, calidad) {
  const v = _s(enlace), r = { avisos: [], error: '' };
  if (!v) { r.error = tx('medios.error.enlace'); return r; }
  if (/^https?:\/\//i.test(v) && !/drive\.google|docs\.google|googleusercontent/i.test(v)) {
    if (tipo === 'audio') r.error = tx('medios.error.audioWeb');
    else r.avisos.push(tx('medios.aviso.web'));
    return r;
  }
  const f = _archivoMedio(v, tipo);
  if (!f) { r.error = tx('medios.error.noEncontrado'); return r; }
  const mime = _s(f.getMimeType()), kb = Math.round(f.getSize() / 1024);
  if (tipo === 'imagen') {
    if (!/^image\//.test(mime)) r.error = tx('medios.error.noImagen');
    else if (f.getSize() > MEDIO_LIMITES.imagenMax) r.avisos.push(tx('medios.aviso.imagenPesada', { kb }));
    else if (f.getSize() > MEDIO_LIMITES.imagenIdeal) r.avisos.push(tx('medios.aviso.imagenMejorable', { kb }));
  } else {
    if (!/^audio\/|^video\/(mp4|webm|ogg)/.test(mime)) r.error = tx('medios.error.noAudio');
    else if (calidad !== 'archivo' && f.getSize() > MEDIO_LIMITES.audioMax) r.error = tx('medios.error.audioMax', { kb });
    else if (calidad !== 'archivo' && f.getSize() > MEDIO_LIMITES.audioIdeal) r.avisos.push(tx('medios.aviso.audioPesado', { kb }));
    else if (calidad === 'archivo' && /mpeg|mp3|aac|mp4|m4a|ogg|webm/.test(mime)) r.avisos.push(tx('medios.aviso.archivoComprimido'));
  }
  return r;
}

function _enlaceNormal(v) {
  const id = _idDrive(v);
  return id ? 'https://drive.google.com/file/d/' + id + '/view' : _s(v);
}

/** Datos para la ventana: palabras legibles, lenguas, voces y cuántas propuestas esperan. */
function mediosDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getAppConfig(ss);
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  _requerir(tPal, ['id_palabra'], CONFIG.SHEETS.PALABRAS);
  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma))
    .map(i => ({ id: _s(i.id_idioma), nombre: _s(i.nombre_propio || i.nombre_completo) || _s(i.codigo_iso) }));
  const cats = {}, categorias = [];
  readSheet(ss, CONFIG.SHEETS.CATEGORIAS).forEach(c => { if (!_s(c.id_categoria)) return; cats[_k(c.id_categoria)] = _s(c.nombre_categoria); categorias.push({ id: _s(c.id_categoria), nombre: _s(c.nombre_categoria) }); });
  categorias.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  const actores = _actores(ss);
  const textos = _textosPorPalabra(tTr);
  const pend = {};
  const tM = _tabla(ss, CONFIG.SHEETS.MEDIOS);
  if (tM) tM.filas.forEach(f => {
    if (_k(_get(tM, f, 'estado')) !== 'pendiente') return;
    const k = _k(_get(tM, f, 'id_palabra')) + '|' + (_k(_get(tM, f, 'tipo')) === 'audio' ? _k(_get(tM, f, 'id_idioma')) : 'img');
    pend[k] = (pend[k] || 0) + 1;
  });
  const palabras = tPal.filas.filter(f => _s(_get(tPal, f, 'id_palabra'))).map(f => {
    const id = _s(_get(tPal, f, 'id_palabra')), t = textos[_k(id)] || {}, tx2 = {}, au = {};
    Object.keys(t).forEach(L => {
      tx2[L] = t[L].texto;
      const fila = tTr.filas[t[L].fila];
      au[L] = { c: !!_s(_get(tTr, fila, 'audio')), a: !!_s(_get(tTr, fila, 'audio_archivo')), p: pend[_k(id) + '|' + L] || 0, n: _s(_get(tTr, fila, 'audio_nota')) };
    });
    return { id, cat: cats[_k(_get(tPal, f, 'id_categoria'))] || '', catId: _s(_get(tPal, f, 'id_categoria')), estado: _s(_get(tPal, f, 'estado')), t: tx2, au,
      img: _s(_get(tPal, f, 'imagen_referencia')), credito: _creditoVisible(actores, _get(tPal, f, 'imagen_credito')), nota: _s(_get(tPal, f, 'imagen_nota')),
      pImg: pend[_k(id) + '|img'] || 0 };
  });
  const voces = readSheet(ss, CONFIG.SHEETS.VOCES).filter(v => _s(v.id_voz) && _k(v.activo) !== 'no')
    .map(v => ({ id: _s(v.id_voz), nombre: _s(v.etiqueta) || _s(v.nombre_publico) || _s(v.id_voz), L: _s(v.id_idioma) }));
  const carpeta = id => { const c = _carpetaValida(id); return c ? c.getUrl() : ''; };
  return { soyAdmin: _soyAdmin(ss), idiomas, palabras, voces, categorias, actores: actores.map(a => a.nombre).filter(Boolean).sort((a, b) => a.localeCompare(b, 'es')),
    notas: NOTAS_RAPIDAS.map(k => tx('medios.nota.' + k)), principal: _s(cfg.idiomaPrincipal),
    pendientes: Object.values(pend).reduce((a, n) => a + n, 0), carpetas: { imagen: carpeta(cfg.imagenes), audio: carpeta(cfg.audios) } };
}

/**
 * Asigna (o propone) una imagen a una palabra o un sonido a una traducción.
 * o = { tipo, id_palabra, id_idioma, enlace, calidad, credito, id_voz, directo, origen }
 */
function mediosAsignar(o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  o = o || {};
  const tipo = o.tipo === 'audio' ? 'audio' : 'imagen', calidad = o.calidad === 'archivo' ? 'archivo' : 'consulta';
  const rev = _revisarMedio(o.enlace, tipo, calidad);
  if (rev.error) throw new Error(rev.error);
  const lock = _lock();
  try {
    const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS);
    const iP = tPal.filas.findIndex(f => _k(_get(tPal, f, 'id_palabra')) === _k(o.id_palabra));
    if (iP < 0) throw new Error(tx('medios.error.palabra'));
    if (tipo === 'audio') {
      const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
      const iT = tTr.filas.findIndex(f => _k(_get(tTr, f, 'id_palabra')) === _k(o.id_palabra) && _k(_get(tTr, f, 'id_idioma')) === _k(o.id_idioma));
      if (iT < 0) throw new Error(tx('medios.error.sinTraduccion'));
    }
    const directo = !!o.directo && _soyAdmin(ss);
    if (directo) rev.avisos.push(..._aplicarMedio(ss, { tipo, id_palabra: o.id_palabra, id_idioma: o.id_idioma, enlace: o.enlace, calidad, credito: o.credito, id_voz: o.id_voz, nota: o.nota }));
    else {
      _agregar(_hojaMedios(ss), [{ id_medio: Utilities.getUuid(), tipo, id_palabra: _s(o.id_palabra), id_idioma: tipo === 'audio' ? _s(o.id_idioma) : '',
        enlace: _enlaceNormal(o.enlace), calidad: tipo === 'audio' ? calidad : '', credito: _s(o.credito).slice(0, 300), id_voz: _s(o.id_voz),
        origen: _s(o.origen) || tx('medios.origen.ventana'), propuesto_por: _correoActual(), fecha: _ahora(), estado: 'Pendiente', nota: _s(o.nota).slice(0, 300) }]);
    }
    return { aplicado: directo, avisos: rev.avisos };
  } finally { lock.releaseLock(); }
}

const _ahora = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

/** Escribe de verdad en PALABRAS o TRADUCCIONES (solo quien administra, directo o al aceptar una propuesta). */
function _aplicarMedio(ss, m) {
  const avisos = [];
  if (m.tipo === 'audio') {
    const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
    if (_col(t, 'audio_nota') < 0) { asegurarEstructura(ss); return _aplicarMedio(ss, m); }
    const i = t.filas.findIndex(f => _k(_get(t, f, 'id_palabra')) === _k(m.id_palabra) && _k(_get(t, f, 'id_idioma')) === _k(m.id_idioma));
    if (i < 0) throw new Error(tx('medios.error.sinTraduccion'));
    const antes = { audio: _get(t, t.filas[i], 'audio'), audio_archivo: _get(t, t.filas[i], 'audio_archivo') };
    const c = m.calidad === 'archivo' ? { audio_archivo: _enlaceNormal(m.enlace) } : { audio: _enlaceNormal(m.enlace) };
    if (_s(m.id_voz)) { c.id_voz = _s(m.id_voz); c.consentimiento = ''; }   // el consentimiento sale de la ficha de la voz
    else if (m.calidad !== 'archivo' && !_s(_get(t, t.filas[i], 'id_voz')) && !_s(_get(t, t.filas[i], 'consentimiento'))) c.consentimiento = 'Pendiente'; // sin voz no se publica
    if (_s(m.nota)) c.audio_nota = _s(m.nota).slice(0, 300);
    _actualizar(t, i, c);
    _historial(ss, CONFIG.SHEETS.TRADUCCIONES, _s(_get(t, t.filas[i], 'id_traduccion')), antes, c);
  } else {
    const t = _tabla(ss, CONFIG.SHEETS.PALABRAS);
    if (_col(t, 'imagen_nota') < 0) { asegurarEstructura(ss); return _aplicarMedio(ss, m); }
    const i = t.filas.findIndex(f => _k(_get(t, f, 'id_palabra')) === _k(m.id_palabra));
    if (i < 0) throw new Error(tx('medios.error.palabra'));
    const antes = { imagen_referencia: _get(t, t.filas[i], 'imagen_referencia') };
    const c = { imagen_referencia: _enlaceNormal(m.enlace), imagen_id: '', imagen_ref: '' };   // se resuelve y comparte al publicar
    if (_s(m.credito)) {   // lenguaje controlado: siempre una persona o institución de ACTORES
      const a = _actorPorNombre(ss, m.credito, true);
      c.imagen_credito = a.id;
      if (a.nuevo) avisos.push(tx('medios.aviso.actorNuevo', { nombre: _s(m.credito) }));
    }
    if (_s(m.nota)) c.imagen_nota = _s(m.nota).slice(0, 300);
    _actualizar(t, i, c);
    _historial(ss, CONFIG.SHEETS.PALABRAS, _s(m.id_palabra), antes, { imagen_referencia: c.imagen_referencia });
  }
  return avisos;
}

/** Nota interna de la imagen (de la palabra) o del sonido (de la traducción). Vacía = se borra. No se publica. */
function mediosNota(tipo, idPalabra, idIdioma, nota) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const audio = tipo === 'audio', t = _tabla(ss, audio ? CONFIG.SHEETS.TRADUCCIONES : CONFIG.SHEETS.PALABRAS), col = audio ? 'audio_nota' : 'imagen_nota';
    if (_col(t, col) < 0) { asegurarEstructura(ss); return mediosNota(tipo, idPalabra, idIdioma, nota); }
    const i = t.filas.findIndex(f => _k(_get(t, f, 'id_palabra')) === _k(idPalabra) && (!audio || _k(_get(t, f, 'id_idioma')) === _k(idIdioma)));
    if (i < 0) throw new Error(tx(audio ? 'medios.error.sinTraduccion' : 'medios.error.palabra'));
    _actualizar(t, i, { [col]: _s(nota).slice(0, 300) });
    return true;
  } finally { lock.releaseLock(); }
}

/** Quita la imagen de una palabra o un sonido de una traducción (solo quien administra). */
function mediosQuitar(tipo, idPalabra, idIdioma, calidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lock = _lock();
  try {
    if (tipo === 'audio') {
      const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
      const i = t.filas.findIndex(f => _k(_get(t, f, 'id_palabra')) === _k(idPalabra) && _k(_get(t, f, 'id_idioma')) === _k(idIdioma));
      if (i > -1) _actualizar(t, i, calidad === 'archivo' ? { audio_archivo: '' } : { audio: '' });
    } else {
      const t = _tabla(ss, CONFIG.SHEETS.PALABRAS);
      const i = t.filas.findIndex(f => _k(_get(t, f, 'id_palabra')) === _k(idPalabra));
      if (i > -1) _actualizar(t, i, { imagen_referencia: '', imagen_id: '', imagen_ref: '', imagen_credito: '', imagen_nota: '' });
    }
    return true;
  } finally { lock.releaseLock(); }
}

/** Vista previa pequeña de una imagen de Drive (miniatura) o de un audio (hasta 5 MB), sin compartir el archivo. */
function mediosVistaPrevia(enlace, tipo) {
  const v = _s(enlace);
  if (/^https?:\/\//i.test(v) && !/drive\.google|docs\.google|googleusercontent/i.test(v)) return { url: v };
  const f = _archivoMedio(v, tipo);
  if (!f) throw new Error(tx('medios.error.noEncontrado'));
  if (tipo === 'audio') {
    if (f.getSize() > MEDIO_LIMITES.audioMax) throw new Error(tx('medios.error.audioMax', { kb: Math.round(f.getSize() / 1024) }));
    const b = f.getBlob();
    return { url: 'data:' + (b.getContentType() || 'audio/mpeg') + ';base64,' + Utilities.base64Encode(b.getBytes()), nombre: f.getName(), kb: Math.round(f.getSize() / 1024) };
  }
  let th = null; try { th = f.getThumbnail(); } catch (e) { th = null; }
  if (!th && f.getSize() <= MEDIO_LIMITES.imagenMax) th = f.getBlob();
  return { url: th ? 'data:' + (th.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(th.getBytes()) : '', nombre: f.getName(), kb: Math.round(f.getSize() / 1024) };
}

// ── Banco sin asignar: archivos de una carpeta que todavía no están en ninguna palabra ──

function _medioReferenciados(ss) {
  const usados = new Set();
  const add = v => { v = _s(v); if (!v) return; const id = _idDrive(v) || (v.match(/[-\w]{25,}/) || [])[0]; if (id) usados.add(id); usados.add(_k(v.split('/').pop())); };
  const tP = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  if (tP) tP.filas.forEach(f => { add(_get(tP, f, 'imagen_referencia')); add(_get(tP, f, 'imagen_id')); });
  const tT = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  if (tT) tT.filas.forEach(f => { add(_get(tT, f, 'audio')); add(_get(tT, f, 'audio_archivo')); });
  const tM = _tabla(ss, CONFIG.SHEETS.MEDIOS);
  if (tM) tM.filas.forEach(f => { if (_k(_get(tM, f, 'estado')) === 'pendiente') add(_get(tM, f, 'enlace')); });
  return usados;
}

const _plegar = s => _s(s).normalize('NFD').replace(/\p{M}/gu, '').toLowerCase().replace(/[’'ʼꞌ´`]/g, "'");
const _nombreBase = n => _plegar(_s(n).replace(/\.[a-z0-9]{2,5}$/i, '')).replace(/[_\-.]+/g, ' ').replace(/\b\d+\b/g, ' ').replace(/\s+/g, ' ').trim();

/**
 * Lista los archivos de una carpeta que aún no están asignados. tipo: 'imagen' | 'audio'. desde: para paginar.
 * Cada archivo trae una sugerencia de palabra si su nombre coincide con una forma escrita.
 */
function mediosBanco(carpetaUrl, tipo, desde) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const id = _idDrive(carpetaUrl) || (_s(carpetaUrl).match(/folders\/([-\w]{20,})/) || [])[1] || _s(carpetaUrl);
  let carpeta;
  try { carpeta = DriveApp.getFolderById(id); } catch (e) { throw new Error(tx('medios.error.carpeta')); }
  const usados = _medioReferenciados(ss);
  const todos = [];
  const it = carpeta.getFiles();
  while (it.hasNext() && todos.length < 2000) {
    const f = it.next();
    if (f.isTrashed()) continue;
    const mime = _s(f.getMimeType());
    if (tipo === 'audio' ? !/^audio\/|^video\/(mp4|webm|ogg)/.test(mime) : !/^image\//.test(mime)) continue;
    if (usados.has(f.getId()) || usados.has(_k(f.getName()))) continue;
    todos.push(f);
  }
  todos.sort((a, b) => a.getName().localeCompare(b.getName()));
  // Índice de formas escritas → palabra (para sugerir por el nombre del archivo)
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES), formas = new Map();
  if (tTr) tTr.filas.forEach(f => {
    _formas(_s(_get(tTr, f, 'texto'))).forEach(x => {
      const k = _plegar(x); if (!k) return;
      const l = formas.get(k) || []; l.push({ p: _s(_get(tTr, f, 'id_palabra')), L: _s(_get(tTr, f, 'id_idioma')) }); formas.set(k, l);
    });
  });
  const ini = Math.max(0, Number(desde) || 0), pagina = todos.slice(ini, ini + 24);
  const archivos = pagina.map(f => {
    const kb = Math.round(f.getSize() / 1024);
    let mini = '';
    if (tipo === 'imagen') { try { const t = f.getThumbnail(); if (t) mini = 'data:' + (t.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(t.getBytes()); } catch (e) {} }
    const sug = formas.get(_nombreBase(f.getName())) || [];
    const avisos = [];
    if (tipo === 'imagen' && f.getSize() > MEDIO_LIMITES.imagenMax) avisos.push(tx('medios.aviso.imagenPesada', { kb }));
    if (tipo === 'audio' && f.getSize() > MEDIO_LIMITES.audioMax) avisos.push(tx('medios.aviso.soloArchivo', { kb }));
    return { id: f.getId(), nombre: f.getName(), kb, mini, sugerencias: sug.slice(0, 3), avisos };
  });
  return { total: todos.length, desde: ini, archivos, carpeta: carpeta.getName() };
}

// ── Revisión de propuestas ──

function mediosPendientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tM = _hojaMedios(ss), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  const textos = _textosPorPalabra(tTr);
  const idiomas = {}; readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { idiomas[_k(i.id_idioma)] = _s(i.nombre_propio || i.nombre_completo); });
  const voces = {}; readSheet(ss, CONFIG.SHEETS.VOCES).forEach(v => { voces[_k(v.id_voz)] = _s(v.etiqueta) || _s(v.nombre_publico); });
  return { soyAdmin: _soyAdmin(ss), lista: tM.filas.filter(f => _k(_get(tM, f, 'estado')) === 'pendiente').map(f => {
    const p = _k(_get(tM, f, 'id_palabra')), t = textos[p] || {};
    return { id: _s(_get(tM, f, 'id_medio')), tipo: _s(_get(tM, f, 'tipo')), id_palabra: _s(_get(tM, f, 'id_palabra')),
      palabra: Object.keys(t).map(L => (idiomas[L] || L) + ': ' + t[L].texto).join(' · '),
      lengua: idiomas[_k(_get(tM, f, 'id_idioma'))] || '', enlace: _s(_get(tM, f, 'enlace')), calidad: _s(_get(tM, f, 'calidad')),
      credito: _s(_get(tM, f, 'credito')), voz: voces[_k(_get(tM, f, 'id_voz'))] || '', origen: _s(_get(tM, f, 'origen')), nota: _s(_get(tM, f, 'nota')),
      quien: _s(_get(tM, f, 'propuesto_por')), fecha: _s(_get(tM, f, 'fecha')) };
  }) };
}

/** accion: 'aceptar' | 'rechazar'. ids: lista de id_medio. notas: { id_medio: nota } (opcional; reemplaza la de la propuesta). */
function mediosRevisar(ids, accion, notas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lock = _lock();
  try {
    const tM = _hojaMedios(ss), quiero = new Set((ids || []).map(_k));
    let hechas = 0; const errores = [], avisos = [];
    tM.filas.forEach((f, i) => {
      if (!quiero.has(_k(_get(tM, f, 'id_medio'))) || _k(_get(tM, f, 'estado')) !== 'pendiente') return;
      const idM = _k(_get(tM, f, 'id_medio')), nota = notas && notas[idM] != null ? _s(notas[idM]) : _s(_get(tM, f, 'nota'));
      if (accion === 'aceptar') {
        try {
          avisos.push(..._aplicarMedio(ss, { tipo: _s(_get(tM, f, 'tipo')), id_palabra: _get(tM, f, 'id_palabra'), id_idioma: _get(tM, f, 'id_idioma'),
            enlace: _get(tM, f, 'enlace'), calidad: _s(_get(tM, f, 'calidad')), credito: _get(tM, f, 'credito'), id_voz: _get(tM, f, 'id_voz'), nota }));
        } catch (e) { errores.push(e.message); return; }
      }
      _actualizar(tM, i, { estado: accion === 'aceptar' ? 'Aceptado' : 'Rechazado', revisado_por: _correoActual(), fecha_revision: _ahora(), nota: nota.slice(0, 300) });
      hechas++;
    });
    if (accion === 'aceptar' && hechas) { try { publicarCatalogo(ss, true); } catch (e) { /* se publica en la próxima hora */ } }
    return { hechas, errores: [...new Set(errores)], avisos: [...new Set(avisos)] };
  } finally { lock.releaseLock(); }
}

// ── Hojas para quien ilustra (se comparten como las de traducción y se traen después) ──

function _hojasIlustracion() {
  try { return JSON.parse(PropertiesService.getDocumentProperties().getProperty(PROP_HOJAS_ILUSTRACION) || '[]'); } catch (e) { return []; }
}

/** op = { nombre, todas, categoria, idiomas:[id] } */
function ilustracionCrear(op) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  op = op || {};
  const cfg = getAppConfig(ss);
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  const textos = _textosPorPalabra(tTr);
  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma));
  const usar = (op.idiomas && op.idiomas.length ? idiomas.filter(i => op.idiomas.map(_k).includes(_k(i.id_idioma))) : idiomas);
  const cats = {}; readSheet(ss, CONFIG.SHEETS.CATEGORIAS).forEach(c => { cats[_k(c.id_categoria)] = _s(c.nombre_categoria); });
  const filas = [];
  tPal.filas.forEach(f => {
    const id = _s(_get(tPal, f, 'id_palabra')); if (!id) return;
    if (_estadoCanonico(_get(tPal, f, 'estado')) === 'retirado') return;
    if (!op.todas && _s(_get(tPal, f, 'imagen_referencia'))) return;
    if (op.categorias && op.categorias.length && !op.categorias.map(_k).includes(_k(_get(tPal, f, 'id_categoria')))) return;
    const t = textos[_k(id)] || {};
    const pal = usar.filter(i => t[_k(i.id_idioma)] && t[_k(i.id_idioma)].texto).map(i => _s(i.nombre_propio || i.nombre_completo) + ': ' + t[_k(i.id_idioma)].texto).join('\n');
    if (!pal) return;
    filas.push([id, cats[_k(_get(tPal, f, 'id_categoria'))] || '', pal, '', '', '']);
  });
  if (!filas.length) throw new Error(tx(op.categorias && op.categorias.length ? 'ilustrar.nadaCategorias' : 'ilustrar.todas'));
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nomCats = (op.categorias || []).map(c => cats[_k(c)]).filter(Boolean);
  const nombre = _s(op.nombre).slice(0, 100) || (nomCats.length ? tx('ilustrar.archivoCat', { categorias: nomCats.slice(0, 3).join(', ') + (nomCats.length > 3 ? '…' : ''), fecha }) : tx('ilustrar.archivo', { fecha }));
  const libro = _aColaboraciones(ss, SpreadsheetApp.create(nombre));
  const h = libro.getActiveSheet().setName('DATOS_ILUSTRACION');
  const inst = libro.insertSheet(tx('plantilla.hojaInstrucciones'), 0);
  const lineas = [1, 2, 3, 4, 5, 6, 7].map(i => [tx('ilustrar.inst.' + i, { catalogo: cfg.titulo || ss.getName() })]);
  inst.getRange(1, 1, lineas.length, 1).setValues(lineas).setWrap(true);
  inst.setColumnWidth(1, 760).getRange(1, 1).setFontWeight('bold').setFontSize(14);
  const vista = [ENC_ILUSTRACION.map(k => tx('ilustrar.enc.' + k))];
  h.getRange(1, 1, 1, ENC_ILUSTRACION.length).setValues([ENC_ILUSTRACION]).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('white');
  h.getRange(2, 1, 1, ENC_ILUSTRACION.length).setValues(vista).setFontStyle('italic').setBackground('#e2e8f0');
  h.getRange(3, 1, filas.length, ENC_ILUSTRACION.length).setNumberFormat('@').setValues(filas).setVerticalAlignment('top').setWrap(true);
  h.setFrozenRows(2); h.hideColumns(1);
  h.setColumnWidth(2, 140); h.setColumnWidth(3, 280); h.setColumnWidth(4, 320); h.setColumnWidth(5, 180); h.setColumnWidth(6, 200);
  h.getRange(3, 2, filas.length, 2).setBackground('#f8fafc');
  // Crédito con lenguaje controlado: lista de 📚 Personas e instituciones (se puede escribir otro; se agrega al aceptar)
  const nombres = _actores(ss).map(a => a.nombre).filter(Boolean).slice(0, 450);
  if (nombres.length) { try { h.getRange(3, 5, filas.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(nombres, true).setAllowInvalid(true).build()); } catch (e) {} }
  try { h.protect().setUnprotectedRanges([h.getRange(3, 4, filas.length, 3)]).setWarningOnly(true); } catch (e) { /* sin protección: no pasa nada */ }
  const lista = _hojasIlustracion().filter(x => x.id !== libro.getId());
  lista.unshift({ id: libro.getId(), nombre, fecha, filas: filas.length });
  PropertiesService.getDocumentProperties().setProperty(PROP_HOJAS_ILUSTRACION, JSON.stringify(lista.slice(0, 40)));
  return { id: libro.getId(), url: libro.getUrl(), nombre, filas: filas.length };
}

function ilustracionLista() {
  return _hojasIlustracion().map(x => { let url = ''; try { url = DriveApp.getFileById(x.id).getUrl(); } catch (e) { url = ''; } return Object.assign({}, x, { url, existe: !!url }); });
}

/** Trae los enlaces de una hoja de ilustración: entran como propuestas «Pendiente» (o directo si quien administra lo pide). */
function ilustracionTraer(id, directo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!_hojasIlustracion().some(x => x.id === id)) throw new Error(tx('plantilla.error.noEsta'));
  const libro = SpreadsheetApp.openById(id);
  const h = libro.getSheetByName('DATOS_ILUSTRACION');
  if (!h) throw new Error(tx('ilustrar.error.hoja'));
  const d = h.getDataRange().getValues(), enc = d[0].map(x => _s(x).toUpperCase());
  const c = k => enc.indexOf(k);
  const yaPropuestas = new Set();
  const tM = _hojaMedios(ss);
  tM.filas.forEach(f => { if (_k(_get(tM, f, 'estado')) !== 'rechazado') yaPropuestas.add(_k(_get(tM, f, 'id_palabra')) + '|' + _k(_enlaceNormal(_get(tM, f, 'enlace')))); });
  const r = { nuevas: 0, repetidas: 0, errores: [], avisos: 0 };
  d.slice(2).forEach((fila, i) => {
    const idP = _s(fila[c('ID_PALABRA')]), enlace = _s(fila[c('ENLACE_IMAGEN')]);
    if (!idP || !enlace) return;
    if (yaPropuestas.has(_k(idP) + '|' + _k(_enlaceNormal(enlace)))) { r.repetidas++; return; }
    try {
      const com = _s(fila[c('COMENTARIOS')]);
      const x = mediosAsignar({ tipo: 'imagen', id_palabra: idP, enlace, credito: _s(fila[c('CREDITO')]), directo: !!directo, nota: com ? tx('medios.nota.deIlustracion', { c: com }) : '',
        origen: tx('medios.origen.ilustracion', { nombre: libro.getName() }) });
      if (x.avisos.length) r.avisos++;
      r.nuevas++; yaPropuestas.add(_k(idP) + '|' + _k(_enlaceNormal(enlace)));
    } catch (e) { r.errores.push(tx('ilustrar.errorFila', { fila: i + 3, error: e.message })); }
  });
  return r;
}

function uiMediosIlustrar() { uiMedios('ilustrar'); }

function uiMedios(tabInicial) {
  const tab = ['palabra', 'banco', 'revisar', 'ilustrar', 'rec'].includes(tabInicial) ? tabInicial : 'palabra';
  const html = `${_CSS_FICHAS}${_tScript()}
  <style>
    .tabs{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:8px}.tabs button{background:#e2e8f0}.tabs button.on{background:#475569;color:#fff}
    .pal{border:1px solid #e2e8f0;border-radius:8px;padding:8px;margin:6px 0;background:#fff}
    .pal .t{font-weight:bold}.pal small{color:#64748b}.fila{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
    .fila input[type=text]{flex:1;min-width:200px;width:auto}.fila select{width:auto;flex:0 1 auto;max-width:280px}.chip{display:inline-block;border-radius:10px;padding:1px 8px;font-size:12px;background:#f1f5f9;margin:2px}
    .si{background:#dcfce7;color:#166534}.no{background:#fee2e2;color:#b91c1c}.pe{background:#fef9c3;color:#854d0e}
    .mini{width:84px;height:84px;object-fit:contain;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px}
    .banco{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:8px}
    .rec{font-size:12px;line-height:1.5}.rec li{margin:2px 0}
    .res{max-height:150px;overflow:auto;border:1px solid #e2e8f0;border-radius:6px;background:#fff}
    .res div{padding:4px 6px;cursor:pointer;border-bottom:1px solid #f1f5f9}.res div:hover{background:#eef2ff}
  </style>
  <div class="caja">
    <p style="margin-top:0">${txh('medios.intro')}</p>
    <div class="tabs" role="tablist">
      <button data-tab="palabra">${txh('medios.tab.palabra')}</button>
      <button data-tab="banco">${txh('medios.tab.banco')}</button>
      <button data-tab="revisar" id="tRevisar">${txh('medios.tab.revisar')}</button>
      <button data-tab="ilustrar">${txh('medios.tab.ilustrar')}</button>
      <button data-tab="rec">${txh('medios.tab.rec')}</button>
    </div>
    <div id="vista"></div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null, TAB = '${tab}', modo = 'imagen', L = '', filtro = 'falta', q = '', cat = '', limite = 40, banco = null, previas = {};
    const ok = h => { $('msg').innerHTML = '<div class="ok">' + h + '</div>'; };
    const err = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    const plegar = s => String(s || '').normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase();
    const nombreL = id => (D.idiomas.find(i => i.id === id) || {nombre: id}).nombre;
    const etiqueta = p => { const pri = p.t[D.principal] || p.t[L] || Object.values(p.t)[0] || tt('medios.sinTexto'); return pri; };
    const otras = p => Object.keys(p.t).filter(x => p.t[x] !== etiqueta(p)).map(x => nombreL(x) + ': ' + p.t[x]).join(' · ');
    function listas() {
      let dl = $('dlActores'); if (!dl) { document.body.insertAdjacentHTML('beforeend', '<datalist id="dlActores"></datalist><datalist id="dlNotas"></datalist>'); }
      $('dlActores').innerHTML = D.actores.map(a => '<option value="' + esc(a) + '">').join('');
      $('dlNotas').innerHTML = D.notas.map(a => '<option value="' + esc(a) + '">').join('');
    }
    function cargar(f) { google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; listas(); if (!L) L = (D.idiomas[0] || {}).id || ''; $('tRevisar').textContent = tt('medios.tab.revisarN', {n: D.pendientes}); (f || pintar)(); }).mediosDatos(); }
    document.querySelector('.tabs').addEventListener('click', e => { const b = e.target.closest('[data-tab]'); if (!b) return; TAB = b.dataset.tab; $('msg').innerHTML = ''; pintar(); });
    function pintar() {
      document.querySelectorAll('.tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === TAB));
      if (!D) { $('vista').innerHTML = '<p>' + th('comun.cargando') + '</p>'; return; }
      ({ palabra: vPalabra, banco: vBanco, revisar: vRevisar, ilustrar: vIlustrar, rec: vRec })[TAB]();
    }
    const recCorta = () => '<details class="sec"><summary><b>' + th(modo === 'audio' ? 'medios.rec.audio.titulo' : 'medios.rec.imagen.titulo') + '</b></summary><ul class="rec">' +
      (modo === 'audio' ? [1,2,3,4,5] : [1,2,3,4,5]).map(i => '<li>' + th('medios.rec.' + modo + '.' + i) + '</li>').join('') + '</ul></details>';
    const selModo = () => '<div class="fila"><label style="font-weight:normal"><input type="radio" name="modo" value="imagen"' + (modo === 'imagen' ? ' checked' : '') + '> 🖼️ ' + th('medios.imagenes') + '</label>' +
      '<label style="font-weight:normal"><input type="radio" name="modo" value="audio"' + (modo === 'audio' ? ' checked' : '') + '> 🔊 ' + th('medios.sonidos') + '</label>' +
      (modo === 'audio' ? '<label for="selL">' + th('medios.lengua') + '</label><select id="selL">' + D.idiomas.map(i => '<option value="' + esc(i.id) + '"' + (i.id === L ? ' selected' : '') + '>' + esc(i.nombre) + '</option>').join('') + '</select>' : '') +
      '<label for="selCat">📁</label><select id="selCat" title="' + esc(tt('medios.categoria')) + '"><option value="">' + th('medios.todasCategorias') + '</option>' + D.categorias.map(c => '<option value="' + esc(c.id) + '"' + (c.id === cat ? ' selected' : '') + '>' + esc(c.nombre) + '</option>').join('') + '</select></div>';
    function enlazarModo(f) {
      document.querySelectorAll('input[name="modo"]').forEach(r => r.onchange = () => { modo = r.value; limite = 40; banco = null; f(); });
      if ($('selL')) $('selL').onchange = () => { L = $('selL').value; limite = 40; f(); };
      if ($('selCat')) $('selCat').onchange = () => { cat = $('selCat').value; limite = 40; f(); };
    }
    const calidadSel = id => '<select id="' + id + '" title="' + esc(tt('medios.calidad')) + '"><option value="consulta">' + th('medios.calidad.consulta') + '</option><option value="archivo">' + th('medios.calidad.archivo') + '</option></select>';
    const vozSel = id => { const v = D.voces.filter(x => !x.L || x.L === L); return '<select id="' + id + '"><option value="">' + th('medios.vozSin') + '</option>' + v.map(x => '<option value="' + esc(x.id) + '">' + esc(x.nombre) + '</option>').join('') + '</select>'; };
    const directoChk = (id, si = true) => D.soyAdmin ? '<label style="font-weight:normal"><input type="checkbox" id="' + id + '"' + (si ? ' checked' : '') + '> ' + th('medios.directo') + '</label>' : '<small>' + th('medios.propuesta') + '</small>';
    // ── Por palabra
    function vPalabra() {
      const lista = D.palabras.filter(p => {
        if (modo === 'audio' && !p.t[L]) return false;
        if (cat && p.catId !== cat) return false;
        const tiene = modo === 'audio' ? (p.au[L] || {}).c : !!p.img, nota = modo === 'audio' ? (p.au[L] || {}).n : p.nota;
        if (filtro === 'falta' && tiene) return false;
        if (filtro === 'tiene' && !tiene) return false;
        if (filtro === 'nota' && !nota) return false;
        return !q || plegar(Object.values(p.t).join(' ') + ' ' + p.cat).includes(plegar(q));
      });
      let h = selModo() + '<div class="fila" style="margin-top:6px"><input type="text" id="q" placeholder="' + esc(tt('medios.buscar')) + '" value="' + esc(q) + '">' +
        '<select id="filtro"><option value="falta">' + th(modo === 'audio' ? 'medios.f.sinAudio' : 'medios.f.sinImagen') + '</option><option value="tiene">' + th(modo === 'audio' ? 'medios.f.conAudio' : 'medios.f.conImagen') + '</option><option value="nota">' + th('medios.f.conNota') + '</option><option value="todas">' + th('medios.f.todas') + '</option></select></div>' +
        recCorta() + '<p><small>' + th('medios.mostrando', {n: Math.min(limite, lista.length), total: lista.length}) + '</small></p>';
      h += lista.slice(0, limite).map(p => {
        const i = D.palabras.indexOf(p), a = p.au[L] || {}, nota = modo === 'audio' ? a.n : p.nota;
        const estado = modo === 'audio'
          ? '<span class="chip ' + (a.c ? 'si' : 'no') + '">' + th(a.c ? 'medios.e.consulta' : 'medios.e.sinConsulta') + '</span><span class="chip ' + (a.a ? 'si' : '') + '">' + th(a.a ? 'medios.e.archivo' : 'medios.e.sinArchivo') + '</span>' + (a.p ? '<span class="chip pe">' + th('medios.e.pendientes', {n: a.p}) + '</span>' : '')
          : '<span class="chip ' + (p.img ? 'si' : 'no') + '">' + th(p.img ? 'medios.e.conImagen' : 'medios.e.sinImagen') + '</span>' + (p.pImg ? '<span class="chip pe">' + th('medios.e.pendientes', {n: p.pImg}) + '</span>' : '') + (p.credito ? '<small> · © ' + esc(p.credito) + '</small>' : '');
        return '<div class="pal" data-i="' + i + '"><div class="fila"><span class="t">' + esc(modo === 'audio' ? p.t[L] : etiqueta(p)) + '</span>' + estado +
          (modo === 'imagen' && p.img ? ' <button data-ver="' + i + '">👁 ' + th('medios.ver') + '</button>' : '') + (modo === 'audio' && a.c ? ' <button data-oir="' + i + '">▶ ' + th('medios.oir') + '</button>' : '') + '</div>' +
          '<small>' + esc((modo === 'audio' ? Object.keys(p.t).filter(x => x !== L).map(x => nombreL(x) + ': ' + p.t[x]).join(' · ') : otras(p)) + (p.cat ? ' · 📁 ' + p.cat : '')) + '</small>' +
          (nota ? '<div><span class="chip pe">📝 ' + esc(nota) + '</span></div>' : '') + '<div id="prev' + i + '"></div>' +
          '<div class="fila" style="margin-top:4px"><input type="text" id="e' + i + '" placeholder="' + esc(tt(modo === 'audio' ? 'medios.enlace.audio' : 'medios.enlace.imagen')) + '">' +
          (modo === 'audio' ? calidadSel('c' + i) + vozSel('v' + i) : '<input type="text" id="cr' + i + '" list="dlActores" placeholder="' + esc(tt('medios.credito')) + '" style="max-width:200px;min-width:140px">') +
          '<button class="prim" data-asignar="' + i + '">' + th('medios.asignar') + '</button></div>' +
          '<div class="fila"><input type="text" id="n' + i + '" list="dlNotas" placeholder="' + esc(tt('medios.notaInterna')) + '" value="' + esc(nota || '') + '"><button data-nota="' + i + '">📝 ' + th('medios.guardarNota') + '</button></div></div>';
      }).join('');
      if (lista.length > limite) h += '<button id="mas">' + th('estados.mas') + '</button>';
      h += '<div style="margin-top:6px">' + directoChk('directo') + '</div>';
      $('vista').innerHTML = h;
      $('filtro').value = filtro;
      $('filtro').onchange = () => { filtro = $('filtro').value; limite = 40; vPalabra(); };
      $('q').oninput = () => { q = $('q').value; clearTimeout(window._t); window._t = setTimeout(() => { limite = 40; vPalabra(); $('q').focus(); $('q').setSelectionRange(q.length, q.length); }, 350); };
      if ($('mas')) $('mas').onclick = () => { limite += 60; vPalabra(); };
      enlazarModo(vPalabra);
    }
    $('vista').addEventListener('click', e => {
      const b = e.target.closest('button'); if (!b) return;
      if (b.dataset.asignar != null) asignar(+b.dataset.asignar);
      if (b.dataset.nota != null) guardarNota(+b.dataset.nota);
      if (b.dataset.ver != null) previa(D.palabras[+b.dataset.ver].img, 'imagen', 'prev' + b.dataset.ver);
      if (b.dataset.oir != null) oirTraduccion(+b.dataset.oir);
      if (b.dataset.bver != null) previa(banco.archivos[+b.dataset.bver].id, modo, 'bprev' + b.dataset.bver);
      if (b.dataset.basignar != null) asignarBanco(+b.dataset.basignar);
      if (b.dataset.bmas != null) cargarBanco(banco.desde + 24);
      if (b.dataset.rev) revisar([b.dataset.rev], b.dataset.acc);
      if (b.dataset.revtodo) revisar(PEND.lista.map(x => x.id), 'aceptar');
      if (b.dataset.traer) traer(b.dataset.traer);
      if (b.dataset.sug != null) { const [k, j] = b.dataset.sug.split('|'); elegirPalabra(+k, D.palabras.findIndex(p => p.id === banco.archivos[+k].sugerencias[+j].p)); }
    });
    function previa(enlace, tipo, donde) {
      $(donde).innerHTML = '<small>' + th('comun.cargando') + '</small>';
      google.script.run.withFailureHandler(e => { $(donde).innerHTML = '<small class="no">' + esc(e.message) + '</small>'; }).withSuccessHandler(r => {
        $(donde).innerHTML = tipo === 'audio' ? '<audio controls autoplay src="' + esc(r.url) + '"></audio> <small>' + esc((r.nombre || '') + (r.kb ? ' · ' + r.kb + ' KB' : '')) + '</small>'
          : (r.url ? '<img class="mini" src="' + esc(r.url) + '" alt=""> ' : '') + '<small>' + esc((r.nombre || '') + (r.kb ? ' · ' + r.kb + ' KB' : '')) + '</small>';
      }).mediosVistaPrevia(enlace, tipo);
    }
    function oirTraduccion(i) { previaAudioDe(i); }
    function previaAudioDe(i) { google.script.run.withFailureHandler(err).withSuccessHandler(r => { $('prev' + i).innerHTML = r ? '<audio controls autoplay src="' + esc(r.url) + '"></audio>' : ''; }).mediosAudioDe(D.palabras[i].id, L); }
    function resultado(r) {
      ok(th(r.aplicado ? 'medios.listo' : 'medios.propuesto') + (r.avisos.length ? '<br>⚠️ ' + r.avisos.map(esc).join('<br>⚠️ ') : ''));
      cargar();
    }
    function asignar(i) {
      const p = D.palabras[i], v = $('e' + i).value.trim();
      if (!v) { err({message: tt('medios.error.enlace')}); return; }
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(err).withSuccessHandler(resultado).mediosAsignar({ tipo: modo, id_palabra: p.id, id_idioma: L, enlace: v,
        calidad: $('c' + i) ? $('c' + i).value : '', id_voz: $('v' + i) ? $('v' + i).value : '', credito: $('cr' + i) ? $('cr' + i).value : '', nota: $('n' + i) ? $('n' + i).value : '',
        directo: $('directo') ? $('directo').checked : false });
    }
    function guardarNota(i) {
      const p = D.palabras[i], v = $('n' + i).value;
      google.script.run.withFailureHandler(err).withSuccessHandler(() => { ok(th(v.trim() ? 'medios.notaGuardada' : 'medios.notaBorrada')); cargar(); }).mediosNota(modo, p.id, L, v);
    }
    // ── Banco sin asignar
    function vBanco() {
      let h = selModo() + '<p>' + th('medios.banco.d') + '</p><div class="fila"><input type="text" id="carpeta" placeholder="' + esc(tt('medios.banco.carpeta')) + '" value="' + esc((banco && banco.url) || D.carpetas[modo] || '') + '">' +
        '<button class="prim" id="bCargar">' + th('medios.banco.ver') + '</button></div>' + recCorta();
      if (banco && banco.archivos) {
        h += '<p><b>' + th('medios.banco.cuenta', {n: banco.total, carpeta: banco.carpeta}) + '</b></p><div class="banco">' + banco.archivos.map((a, k) =>
          '<div class="pal"><div class="fila">' + (a.mini ? '<img class="mini" src="' + esc(a.mini) + '" alt="">' : '<button data-bver="' + k + '">' + (modo === 'audio' ? '▶ ' + th('medios.oir') : '👁 ' + th('medios.ver')) + '</button>') +
          '<div style="flex:1;min-width:120px"><small>' + esc(a.nombre) + ' · ' + a.kb + ' KB</small>' + a.avisos.map(x => '<div class="chip no">' + esc(x) + '</div>').join('') + '</div></div><div id="bprev' + k + '"></div>' +
          (a.sugerencias.length ? '<div><small>' + th('medios.banco.sugerencia') + '</small> ' + a.sugerencias.map((s, j) => { const p = D.palabras.find(x => x.id === s.p); return p && (!cat || p.catId === cat) ? '<button data-sug="' + k + '|' + j + '">' + esc(etiqueta(p)) + '</button>' : ''; }).join('') + '</div>' : '') +
          '<input type="text" id="bq' + k + '" placeholder="' + esc(tt('medios.banco.buscar')) + '" style="width:100%;margin-top:4px"><div class="res" id="bres' + k + '" style="display:none"></div>' +
          '<div id="bsel' + k + '"></div>' + (modo === 'audio' ? '<div class="fila">' + calidadSel('bc' + k) + vozSel('bv' + k) + '</div>' : '<input type="text" id="bcr' + k + '" list="dlActores" placeholder="' + esc(tt('medios.credito')) + '" style="width:100%;margin-top:4px">') +
          '<input type="text" id="bn' + k + '" list="dlNotas" placeholder="' + esc(tt('medios.notaInterna')) + '" style="width:100%;margin-top:4px">' +
          '<button class="prim" data-basignar="' + k + '" style="margin-top:4px">' + th('medios.asignar') + '</button></div>').join('') + '</div>';
        if (banco.desde + 24 < banco.total) h += '<button data-bmas="1">' + th('medios.banco.mas') + '</button>';
        h += '<div style="margin-top:6px">' + directoChk('directo') + '</div>';
      }
      $('vista').innerHTML = h;
      $('bCargar').onclick = () => cargarBanco(0);
      enlazarModo(vBanco);
      if (banco && banco.archivos) banco.archivos.forEach((a, k) => {
        $('bq' + k).oninput = () => {
          const t = plegar($('bq' + k).value.trim());
          if (t.length < 2) { $('bres' + k).style.display = 'none'; return; }
          const r = D.palabras.map((p, i) => ({p, i})).filter(x => (modo !== 'audio' || x.p.t[L]) && (!cat || x.p.catId === cat) && plegar(Object.values(x.p.t).join(' ')).includes(t)).slice(0, 12);
          $('bres' + k).style.display = 'block';
          $('bres' + k).innerHTML = r.length ? r.map(x => '<div data-k="' + k + '" data-pi="' + x.i + '">' + esc(modo === 'audio' ? x.p.t[L] : etiqueta(x.p)) + ' <small>' + esc(otras(x.p)) + '</small></div>').join('') : '<div><small>' + th('estados.vacia') + '</small></div>';
        };
      });
    }
    $('vista').addEventListener('click', e => { const d = e.target.closest('[data-pi]'); if (d) elegirPalabra(+d.dataset.k, +d.dataset.pi); });
    function elegirPalabra(k, pi) {
      if (pi < 0) return;
      const p = D.palabras[pi]; banco.archivos[k].elegida = pi;
      $('bres' + k).style.display = 'none';
      $('bsel' + k).innerHTML = '<div class="chip si">✓ ' + esc(modo === 'audio' ? (p.t[L] || etiqueta(p)) + ' (' + nombreL(L) + ')' : etiqueta(p)) + '</div>' + (modo === 'imagen' && p.img ? '<div class="chip pe">' + th('medios.banco.reemplaza') + '</div>' : '');
    }
    function cargarBanco(desde) {
      const url = $('carpeta') ? $('carpeta').value.trim() : banco.url;
      if (!url) { err({message: tt('medios.error.carpeta')}); return; }
      $('msg').textContent = tt('comun.cargando');
      google.script.run.withFailureHandler(err).withSuccessHandler(r => { banco = Object.assign(r, {url}); $('msg').innerHTML = ''; vBanco(); }).mediosBanco(url, modo, desde);
    }
    function asignarBanco(k) {
      const a = banco.archivos[k];
      if (a.elegida == null) { err({message: tt('medios.banco.elige')}); return; }
      const p = D.palabras[a.elegida];
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(err).withSuccessHandler(r => { resultado(r); banco.archivos.splice(k, 1); banco.total--; vBanco(); })
        .mediosAsignar({ tipo: modo, id_palabra: p.id, id_idioma: L, enlace: a.id, calidad: $('bc' + k) ? $('bc' + k).value : '', id_voz: $('bv' + k) ? $('bv' + k).value : '',
          credito: $('bcr' + k) ? $('bcr' + k).value : '', nota: $('bn' + k) ? $('bn' + k).value : '', directo: $('directo') ? $('directo').checked : false, origen: tt('medios.origen.banco') });
    }
    // ── Por revisar
    let PEND = null;
    function vRevisar() {
      $('vista').innerHTML = '<p>' + th('comun.cargando') + '</p>';
      google.script.run.withFailureHandler(err).withSuccessHandler(r => {
        PEND = r;
        if (!r.lista.length) { $('vista').innerHTML = '<div class="ok">' + th('medios.rev.nada') + '</div>'; return; }
        $('vista').innerHTML = '<p>' + th(r.soyAdmin ? 'medios.rev.d' : 'medios.rev.dCaptura') + '</p>' + (r.soyAdmin ? '<button class="prim" data-revtodo="1">' + th('medios.rev.todo', {n: r.lista.length}) + '</button>' : '') +
          r.lista.map((x, k) => '<div class="pal"><div class="fila"><span class="chip">' + (x.tipo === 'audio' ? '🔊 ' + esc(x.lengua) + ' · ' + esc(tt('medios.calidad.' + (x.calidad || 'consulta'))) : '🖼️') + '</span><span class="t">' + esc(x.palabra) + '</span></div>' +
            '<small>' + esc([x.credito && tt('medios.rev.credito', {c: x.credito}), x.voz && tt('medios.rev.voz', {v: x.voz}), x.origen, x.quien, x.fecha].filter(Boolean).join(' · ')) + '</small>' +
            '<div id="rprev' + k + '"></div>' + (r.soyAdmin ? '<input type="text" id="rn' + k + '" list="dlNotas" placeholder="' + esc(tt('medios.notaInterna')) + '" value="' + esc(x.nota) + '" style="width:100%;margin-top:4px">' : (x.nota ? '<span class="chip pe">📝 ' + esc(x.nota) + '</span>' : '')) +
            '<div class="fila"><button data-rver="' + k + '">' + (x.tipo === 'audio' ? '▶ ' + th('medios.oir') : '👁 ' + th('medios.ver')) + '</button>' +
            (r.soyAdmin ? '<button class="prim" data-rev="' + esc(x.id) + '" data-acc="aceptar">✓ ' + th('medios.rev.aceptar') + '</button><button data-rev="' + esc(x.id) + '" data-acc="rechazar">✕ ' + th('medios.rev.rechazar') + '</button>' : '') + '</div></div>').join('');
        document.querySelectorAll('[data-rver]').forEach(b => b.onclick = () => { const x = PEND.lista[+b.dataset.rver]; previa(x.enlace, x.tipo === 'audio' ? 'audio' : 'imagen', 'rprev' + b.dataset.rver); });
      }).mediosPendientes();
    }
    function revisar(ids, acc) {
      $('msg').textContent = tt('comun.guardando');
      const notas = {}; PEND.lista.forEach((x, k) => { if ($('rn' + k)) notas[x.id.toLowerCase()] = $('rn' + k).value; });
      google.script.run.withFailureHandler(err).withSuccessHandler(r => { ok(th(acc === 'aceptar' ? 'medios.rev.aceptadas' : 'medios.rev.rechazadas', {n: r.hechas}) + (r.avisos.length ? '<br>' + r.avisos.map(esc).join('<br>') : '') + (r.errores.length ? '<br>⚠️ ' + r.errores.map(esc).join('<br>⚠️ ') : '')); cargar(vRevisar); }).mediosRevisar(ids, acc, notas);
    }
    // ── Hojas para quien ilustra
    function vIlustrar() {
      const cats = [...new Set(D.palabras.map(p => p.cat).filter(Boolean))].sort();
      $('vista').innerHTML = '<div class="sec"><h3>' + th('ilustrar.crear') + '</h3><p><small>' + th('ilustrar.crear.d') + '</small></p>' +
        '<label style="font-weight:normal"><input type="checkbox" id="iTodas"> ' + th('ilustrar.todasPalabras') + '</label>' +
        '<label for="iNombre">' + th('plantilla.nombre') + '</label><input type="text" id="iNombre" style="width:100%">' +
        '<p><b>' + th('ilustrar.categorias') + '</b> <small>' + th('ilustrar.categorias.d') + '</small><br>' + D.categorias.map(c => '<label style="font-weight:normal;display:inline-block;margin-right:10px"><input type="checkbox" name="iC" value="' + esc(c.id) + '"' + (c.id === cat ? ' checked' : '') + '> ' + esc(c.nombre) + '</label>').join('') + '</p>' +
        '<p><b>' + th('plantilla.lenguas') + '</b><br>' + D.idiomas.map(i => '<label style="font-weight:normal;display:inline-block;margin-right:10px"><input type="checkbox" name="iL" value="' + esc(i.id) + '" checked> ' + esc(i.nombre) + '</label>').join('') + '</p>' +
        '<button class="prim" id="iCrear">' + th('ilustrar.crearBoton') + '</button><div id="iRes"></div></div>' +
        '<div class="sec"><h3>' + th('ilustrar.traer') + '</h3><p><small>' + th('ilustrar.traer.d') + '</small></p><div id="iLista">' + th('comun.cargando') + '</div>' + directoChk('directoI', false) + '</div>';
      $('iCrear').onclick = () => {
        $('iRes').textContent = tt('plantilla.preparando');
        google.script.run.withFailureHandler(e => { $('iRes').innerHTML = '<div class="error">' + esc(e.message) + '</div>'; }).withSuccessHandler(r => {
          const msj = tt('ilustrar.mensaje', {nombre: r.nombre, url: r.url});
          $('iRes').innerHTML = '<div class="ok">' + th('ilustrar.creada', {n: r.filas}) + '<br><a href="' + esc(r.url) + '" target="_blank">' + esc(r.nombre) + '</a></div>' +
            '<p><b>' + th('plantilla.quien') + '</b></p><div class="fila"><input type="text" id="iCorreos" placeholder="persona@ejemplo.com"><button id="iPermiso">' + th('plantilla.darPermiso') + '</button></div><div id="iMP"></div>' +
            '<div class="fila"><a class="btn" style="background:#16a34a;color:#fff;padding:8px 12px;border-radius:6px;text-decoration:none" target="_blank" href="https://wa.me/?text=' + encodeURIComponent(msj) + '">' + th('plantilla.whatsapp') + '</a>' +
            '<a style="background:#e2e8f0;padding:8px 12px;border-radius:6px;text-decoration:none;color:#1f2937" target="_blank" href="mailto:?subject=' + encodeURIComponent(r.nombre) + '&body=' + encodeURIComponent(msj) + '">' + th('plantilla.correo') + '</a></div>';
          $('iPermiso').onclick = () => google.script.run.withFailureHandler(e => { $('iMP').innerHTML = '<div class="error">' + esc(e.message) + '</div>'; })
            .withSuccessHandler(x => { $('iMP').innerHTML = '<div class="ok">' + th('plantilla.permisoCorreos', {n: x.n}) + '</div>'; }).plantillaCompartir(r.id, $('iCorreos').value, false);
          listaIlustracion();
        }).ilustracionCrear({ nombre: $('iNombre').value, todas: $('iTodas').checked, idiomas: [...document.querySelectorAll('input[name="iL"]:checked')].map(x => x.value),
          categorias: [...document.querySelectorAll('input[name="iC"]:checked')].map(x => x.value) });
      };
      listaIlustracion();
    }
    function listaIlustracion() {
      google.script.run.withFailureHandler(err).withSuccessHandler(l => {
        $('iLista').innerHTML = l.length ? l.map(x => '<div class="fila" style="margin:4px 0"><span style="flex:1">' + (x.existe ? '<a href="' + esc(x.url) + '" target="_blank">' + esc(x.nombre) + '</a>' : esc(x.nombre) + ' <small>(' + th('ilustrar.noExiste') + ')</small>') + ' <small>' + esc(x.fecha) + ' · ' + x.filas + '</small></span>' +
          (x.existe ? '<button class="prim" data-traer="' + esc(x.id) + '">📥 ' + th('ilustrar.traerBoton') + '</button>' : '') + '</div>').join('') : '<small>' + th('ilustrar.ninguna') + '</small>';
      }).ilustracionLista();
    }
    function traer(id) {
      $('msg').textContent = tt('comun.cargando');
      google.script.run.withFailureHandler(err).withSuccessHandler(r => {
        ok(th('ilustrar.traidas', {n: r.nuevas, rep: r.repetidas}) + (r.avisos ? '<br>' + th('ilustrar.conAvisos', {n: r.avisos}) : '') + (r.errores.length ? '<br>⚠️ ' + r.errores.slice(0, 8).map(esc).join('<br>⚠️ ') : ''));
        cargar(vIlustrar);
      }).ilustracionTraer(id, $('directoI') ? $('directoI').checked : false);
    }
    // ── Recomendaciones
    function vRec() {
      $('vista').innerHTML = ['imagen', 'audio'].map(m => '<div class="sec"><h3>' + th('medios.rec.' + m + '.titulo') + '</h3><ul class="rec">' + [1,2,3,4,5].map(i => '<li>' + th('medios.rec.' + m + '.' + i) + '</li>').join('') + '</ul></div>').join('') +
        '<div class="sec"><h3>' + th('medios.rec.calidades.titulo') + '</h3><p class="rec">' + th('medios.rec.calidades') + '</p></div>';
    }
    pintar(); cargar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(900).setHeight(740), tx('medios.titulo'));
}

/** Audio de consulta de una traducción, para escucharlo en la ventana. */
function mediosAudioDe(idPalabra, idIdioma) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  const f = t.filas.find(x => _k(_get(t, x, 'id_palabra')) === _k(idPalabra) && _k(_get(t, x, 'id_idioma')) === _k(idIdioma));
  const a = f ? _s(_get(t, f, 'audio')) : '';
  return a ? mediosVistaPrevia(a, 'audio') : null;
}

// ─────────────────────────────────────────────────────────────────────────────
// 22. EDITAR PALABRAS DESDE LA HOJA (sin tocar tablas ni códigos)
// ─────────────────────────────────────────────────────────────────────────────
// Para quien administra o da soporte: buscar en cualquier lengua, ver la ficha completa, corregir,
// agregar lenguas, crear palabras, cambiar categoría o estado, y cambios en bloque con vista previa.
//  - Quien administra: sus cambios se aplican tal cual.
//  - Quien captura: no publica ni retira; si corrige algo ya publicado, eso vuelve a «En revisión».
//  - Los atajos de escritura se aplican al guardar (salvo fuentes que piden no alterar su contenido).
//  - Antes de guardar se avisa si una forma ya existe en otra palabra de la misma lengua.
//  - Cada cambio queda en HISTORIAL. Nada se borra: se retira.

function _catalogoActual(ss) {
  const cfg = getAppConfig(ss);
  return { titulo: cfg.titulo || ss.getName(), archivo: ss.getName(), rama: _esRama(ss), copia: _esCopia(ss) };
}

/** Retira las traducciones guardadas sin lengua válida (la primera versión de esta ventana las guardaba como «undefined»). */
function _repararTraduccionesSinLengua(ss) {
  const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  if (!t) return 0;
  const lenguas = new Set(readSheet(ss, CONFIG.SHEETS.IDIOMAS).map(i => _k(i.id_idioma)).filter(Boolean));
  let n = 0;
  t.filas.forEach((f, i) => {
    const L = _k(_get(t, f, 'id_idioma'));
    if ((L === 'undefined' || L === 'null') && !lenguas.has(L) && _estadoCanonico(_get(t, f, 'estado')) !== 'retirado') {
      _actualizar(t, i, { estado: _etiquetaEstado('retirado') });
      _historial(ss, CONFIG.SHEETS.TRADUCCIONES, _s(_get(t, f, 'id_traduccion')), { id_idioma: L }, { estado: tx('editar.reparadaSinLengua') });
      n++;
    }
  });
  return n;
}

function editarDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  let reparadas = 0; try { reparadas = _repararTraduccionesSinLengua(ss); } catch (e) { reparadas = 0; }
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  _requerir(tPal, ['id_palabra'], CONFIG.SHEETS.PALABRAS);
  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma) && _estadoCanonico(i.estado) !== 'retirado')
    .map(i => ({ id: _s(i.id_idioma), nombre: _s(i.nombre_propio || i.nombre_completo) || _s(i.codigo_iso) }));
  const categorias = readSheet(ss, CONFIG.SHEETS.CATEGORIAS).filter(c => _s(c.id_categoria) && _estadoCanonico(c.estado) !== 'retirado')
    .map(c => ({ id: _s(c.id_categoria), nombre: _s(c.nombre_categoria) })).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  const fuentes = readSheet(ss, CONFIG.SHEETS.FUENTES).filter(f => _s(f.id_fuente) && _estadoCanonico(f.estado) !== 'retirado')
    .map(f => ({ id: _s(f.id_fuente), titulo: _s(f.titulo) || _s(f.id_fuente) }));
  const por = {};
  if (tTr) tTr.filas.forEach(f => {
    const p = _k(_get(tTr, f, 'id_palabra')), L = _s(_get(tTr, f, 'id_idioma'));
    if (!p || !L || _estadoCanonico(_get(tTr, f, 'estado')) === 'retirado') return;   // lo retirado no aparece en la lista
    const o = por[p] = por[p] || { t: {}, e: {}, d: {}, sf: false };
    o.t[L] = _s(_get(tTr, f, 'texto'));
    if (_s(_get(tTr, f, 'descripcion'))) o.d[L] = 1;
    o.e[L] = _estadoCanonico(_get(tTr, f, 'estado')) || 'publicado';
    if (!_s(_get(tTr, f, 'id_fuente'))) o.sf = true;
  });
  const palabras = tPal.filas.filter(f => _s(_get(tPal, f, 'id_palabra'))).map(f => {
    const id = _s(_get(tPal, f, 'id_palabra')), o = por[_k(id)] || { t: {}, e: {}, d: {}, sf: false };
    return { id, cat: _s(_get(tPal, f, 'id_categoria')), estado: _estadoCanonico(_get(tPal, f, 'estado')) || 'publicado',
      t: o.t, e: o.e, d: o.d, sinFuente: o.sf, img: !!_s(_get(tPal, f, 'imagen_referencia')) };
  });
  let relOpciones = []; try { relOpciones = _opcionesRelacion(ss); } catch (e) { relOpciones = []; }
  return { soyAdmin: _soyAdmin(ss), catalogo: _catalogoActual(ss), idiomas, categorias, fuentes, palabras, reparadas, relOpciones,
    principal: _s(getAppConfig(ss).idiomaPrincipal), estados: ESTADOS.map(e => ({ canon: e.canon, etiqueta: tx('estado.' + e.canon) })) };
}

function editarFicha(idPalabra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  const f = tPal.filas.find(x => _k(_get(tPal, x, 'id_palabra')) === _k(idPalabra));
  if (!f) throw new Error(tx('medios.error.palabra'));
  const trads = {};
  tTr.filas.forEach(x => {
    if (_k(_get(tTr, x, 'id_palabra')) !== _k(idPalabra)) return;
    const L = _s(_get(tTr, x, 'id_idioma'));
    trads[L] = { id: _s(_get(tTr, x, 'id_traduccion')), texto: _s(_get(tTr, x, 'texto_captura')) || _s(_get(tTr, x, 'texto')), publicado: _s(_get(tTr, x, 'texto')),
      nota_variante: _s(_get(tTr, x, 'nota_variante')), silabas: _s(_get(tTr, x, 'silabas')), id_fuente: _s(_get(tTr, x, 'id_fuente')),
      ubicacion_fuente: _s(_get(tTr, x, 'ubicacion_fuente')), estado: _estadoCanonico(_get(tTr, x, 'estado')) || 'publicado',
      audio: !!_s(_get(tTr, x, 'audio')), audio_nota: _s(_get(tTr, x, 'audio_nota')), creado_por: _s(_get(tTr, x, 'creado_por')),
      descripcion: _s(_get(tTr, x, 'descripcion')) };
  });
  return { id: _s(idPalabra), id_categoria: _s(_get(tPal, f, 'id_categoria')), estado: _estadoCanonico(_get(tPal, f, 'estado')) || 'publicado',
    juegos: _k(_get(tPal, f, 'usar_en_juegos')) !== 'no', imagen: _s(_get(tPal, f, 'imagen_referencia')), imagen_nota: _s(_get(tPal, f, 'imagen_nota')),
    creado_por: _s(_get(tPal, f, 'creado_por')), trads, relaciones: _relacionesDe(ss, idPalabra), galeria: (() => { try { return galeriaDatos(idPalabra); } catch (e) { return null; } })() };
}

/** Texto que se guarda: con los atajos de escritura aplicados, salvo que la fuente pida no alterar su contenido. */
function _textoEditado(ctx, reglas, L, idFuente, escrito) {
  const v = _s(escrito).normalize('NFC');
  if (!v) return v;
  if (ctx && _s(idFuente) && _fuenteFiel(ctx, idFuente)) return v;
  return _normalizar(v, L, reglas, _correoActual(), false);
}

/** Descripción en la lengua: con los mismos atajos que el texto (salvo fuentes fieles), sin saltos de más y con tope de largo. */
function _descripcionEditada(ctx, reglas, L, idFuente, escrito) {
  const v = _textoEditado(ctx, reglas, L, idFuente, _s(escrito).replace(/\n{3,}/g, '\n\n'));
  return _s(v).slice(0, LARGO_DESCRIPCION);
}
const LARGO_DESCRIPCION = 600;
const AUTOR = 'Alejandro Estrada';   // crédito de autoría (AGPL-3.0, sección 7 b): no se traduce

/** Formas escritas en cada lengua → palabras que las usan (para avisar coincidencias antes de guardar). */
function _indiceFormas(tTr) {
  const m = new Map();
  tTr.filas.forEach(f => {
    if (_estadoCanonico(_get(tTr, f, 'estado')) === 'retirado') return;
    const L = _k(_get(tTr, f, 'id_idioma')), p = _s(_get(tTr, f, 'id_palabra'));
    _formas(_s(_get(tTr, f, 'texto'))).forEach(x => { const k = L + '|' + _claveForma(x); const l = m.get(k) || []; if (!l.includes(p)) l.push(p); m.set(k, l); });
  });
  return m;
}

/**
 * Revisa una edición sin guardar: cómo quedará cada texto (con atajos) y qué formas ya existen en otras palabras.
 * c = { id_palabra, trads: { id_idioma: { texto, id_fuente } } }
 */
function editarComprobar(c) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  let ctx = null; try { ctx = _contextoPermisos(ss); } catch (e) { ctx = null; }
  const reglas = _cargarReglas(ss), indice = _indiceFormas(tTr), textos = {};
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).forEach(t => { const p = _k(t.id_palabra); if (!textos[p]) textos[p] = _s(t.texto); });
  const idiomas = {}; readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { idiomas[_k(i.id_idioma)] = _s(i.nombre_propio || i.nombre_completo); });
  const r = { textos: {}, coincidencias: [] };
  Object.keys((c && c.trads) || {}).forEach(L => {
    const x = c.trads[L] || {};
    if (x.texto == null) return;
    const final = _textoEditado(ctx, reglas, L, x.id_fuente, x.texto);
    r.textos[L] = final;
    _formas(final).forEach(forma => {
      const otras = (indice.get(_k(L) + '|' + _claveForma(forma)) || []).filter(p => _k(p) !== _k(c.id_palabra));
      if (otras.length) r.coincidencias.push({ lengua: idiomas[_k(L)] || L, forma, palabras: otras.slice(0, 5).map(p => textos[_k(p)] || p) });
    });
  });
  return r;
}

/**
 * Guarda la ficha. c = { id_palabra ('' = nueva), id_categoria, categoriaNueva, estado, juegos,
 *   trads: { id_idioma: { texto, nota_variante, silabas, id_fuente, ubicacion_fuente, estado } } }
 * forzar = guardar aunque haya coincidencias.
 */
function editarGuardar(c, forzar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  c = c || {};
  const admin = _soyAdmin(ss), yo = _correoActual();
  const lenguas = new Set(readSheet(ss, CONFIG.SHEETS.IDIOMAS).map(i => _k(i.id_idioma)).filter(Boolean));
  const malas = Object.keys(c.trads || {}).filter(L => !lenguas.has(_k(L)));
  if (malas.length) throw new Error(tx('editar.error.lengua'));
  const comp = editarComprobar(c);
  let ctx = null; try { ctx = _contextoPermisos(ss); } catch (e) { ctx = null; }
  const reglas = _cargarReglas(ss);
  if (comp.coincidencias.length && !forzar) return Object.assign({ confirmar: true }, comp);
  const lock = _lock();
  try {
    const avisos = [], nombresL = {};
    readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { nombresL[_k(i.id_idioma)] = _s(i.nombre_propio || i.nombre_completo); });
    // La columna de descripción aparece la primera vez que alguien escribe una (catálogos anteriores no la tienen)
    if (Object.values(c.trads || {}).some(x => x && x.descripcion != null)) { const shT = ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES); if (shT) _asegurarColumnas(shT, ['descripcion']); }
    if (c.imagenes_max != null) { const shP = ss.getSheetByName(CONFIG.SHEETS.PALABRAS); if (shP) _asegurarColumnas(shP, ['imagenes_max']); }
    const tPal = _tablaConEstado(ss, CONFIG.SHEETS.PALABRAS), tTr = _tablaConEstado(ss, CONFIG.SHEETS.TRADUCCIONES);
    const revision = _etiquetaEstado('revisión');
    // Estado que se pide, limitado por el rol
    const pedir = canon => {
      if (!canon) return '';
      if (!ESTADOS.some(e => e.canon === canon)) throw new Error(tx('estados.error.estado'));
      if (!admin && (canon === 'publicado' || canon === 'retirado')) { avisos.push(tx('editar.aviso.soloAdmin')); return 'revisión'; }
      return canon;
    };
    // Categoría nueva escrita en la ficha
    let idCat = _s(c.id_categoria);
    if (_s(c.categoriaNueva)) {
      const tCat = _tabla(ss, CONFIG.SHEETS.CATEGORIAS);
      const ya = tCat.filas.find(f => _plegar(_get(tCat, f, 'nombre_categoria')) === _plegar(c.categoriaNueva));
      if (ya) idCat = _s(_get(tCat, ya, 'id_categoria'));
      else {
        idCat = Utilities.getUuid();
        _agregar(tCat, [{ id_categoria: idCat, nombre_categoria: _s(c.categoriaNueva).slice(0, 120), estado: _etiquetaEstado(admin ? 'publicado' : 'revisión') }]);
        avisos.push(tx('editar.aviso.categoriaNueva', { nombre: _s(c.categoriaNueva) }));
      }
    }
    // Palabra
    let idP = _s(c.id_palabra), iP = idP ? tPal.filas.findIndex(f => _k(_get(tPal, f, 'id_palabra')) === _k(idP)) : -1;
    let nuevos = 0, cambios = 0;
    if (idP && iP < 0) throw new Error(tx('medios.error.palabra'));
    const estadoPal = pedir(_estadoCanonico(c.estado));
    if (!idP) {
      if (!Object.values(c.trads || {}).some(x => _s(x && x.texto))) throw new Error(tx('editar.error.vacia'));
      idP = Utilities.getUuid();
      _agregar(tPal, [{ id_palabra: idP, id_categoria: idCat, estado: _etiquetaEstado(estadoPal || (admin ? 'publicado' : 'revisión')),
        usar_en_juegos: c.juegos === false ? 'No' : 'Sí', creado_por: yo, origen: tx('editar.origen') }]);
      _historial(ss, CONFIG.SHEETS.PALABRAS, idP, {}, { nueva: tx('editar.origen') });
      nuevos++;
    } else {
      const f = tPal.filas[iP], antes = {}, despues = {};
      const pon = (col, v) => { if (_s(_get(tPal, f, col)) !== _s(v)) { antes[col] = _get(tPal, f, col); despues[col] = v; } };
      if (c.id_categoria != null || c.categoriaNueva) pon('id_categoria', idCat);
      if (estadoPal) pon('estado', _etiquetaEstado(estadoPal));
      if (c.juegos != null) pon('usar_en_juegos', c.juegos ? 'Sí' : 'No');
      if (c.imagenes_max != null && admin) pon('imagenes_max', _s(c.imagenes_max) === '' ? '' : String(Math.max(0, Math.min(50, Math.round(Number(c.imagenes_max) || 0)))));
      if (Object.keys(despues).length) { _actualizar(tPal, iP, despues); _historial(ss, CONFIG.SHEETS.PALABRAS, idP, antes, despues); cambios++; }
    }
    const palabraPublicada = _estadoCanonico(_get(tPal, tPal.filas[tPal.filas.findIndex(f => _k(_get(tPal, f, 'id_palabra')) === _k(idP))], 'estado')) === 'publicado';
    // Traducciones
    const modoCaptura = _col(tTr, 'texto_captura') > -1;
    Object.keys(c.trads || {}).forEach(L => {
      const x = c.trads[L] || {};
      const i = tTr.filas.findIndex(f => _k(_get(tTr, f, 'id_palabra')) === _k(idP) && _k(_get(tTr, f, 'id_idioma')) === _k(L));
      const final = x.texto == null ? null : comp.textos[L];
      const est = pedir(_estadoCanonico(x.estado));
      if (i < 0) {
        if (!_s(final)) return;
        const fila = { id_traduccion: Utilities.getUuid(), id_palabra: idP, id_idioma: L, texto: final, nota_variante: _s(x.nota_variante), silabas: _s(x.silabas),
          id_fuente: _s(x.id_fuente), ubicacion_fuente: _s(x.ubicacion_fuente), creado_por: yo, origen: tx('editar.origen'),
          descripcion: _descripcionEditada(ctx, reglas, L, x.id_fuente, x.descripcion),
          estado: _etiquetaEstado(est || (admin ? (palabraPublicada ? 'publicado' : 'revisión') : 'revisión')) };
        if (modoCaptura) fila.texto_captura = _s(x.texto);
        _agregar(tTr, [fila]);
        _historial(ss, CONFIG.SHEETS.TRADUCCIONES, fila.id_traduccion, {}, { texto: final });
        nuevos++;
        return;
      }
      const f = tTr.filas[i], antes = {}, despues = {};
      const pon = (col, v) => { if (v != null && _s(_get(tTr, f, col)) !== _s(v)) { antes[col] = _get(tTr, f, col); despues[col] = v; } };
      if (final != null) {
        if (!_s(final)) avisos.push(tx('editar.aviso.vacio', { lengua: nombresL[_k(L)] || L }));
        else { pon('texto', final); if (modoCaptura) pon('texto_captura', _s(x.texto)); }
      }
      ['nota_variante', 'silabas', 'id_fuente', 'ubicacion_fuente'].forEach(k => pon(k, x[k] == null ? null : _s(x[k])));
      if (x.descripcion != null) pon('descripcion', _descripcionEditada(ctx, reglas, L, x.id_fuente != null ? x.id_fuente : _get(tTr, f, 'id_fuente'), x.descripcion));
      if (est) pon('estado', _etiquetaEstado(est));
      // Quien captura corrige algo ya publicado → vuelve a revisión
      if (!admin && Object.keys(despues).some(k => k !== 'estado') && _estadoCanonico(_get(tTr, f, 'estado')) === 'publicado') {
        despues.estado = revision; antes.estado = _get(tTr, f, 'estado');
        avisos.push(tx('editar.aviso.aRevision'));
      }
      if (Object.keys(despues).length) { _actualizar(tTr, i, despues); _historial(ss, CONFIG.SHEETS.TRADUCCIONES, _s(_get(tTr, f, 'id_traduccion')), antes, despues); cambios++; }
    });
    return { id_palabra: idP, nuevos, cambios, avisos: [...new Set(avisos)], ficha: editarFicha(idP) };
  } finally { lock.releaseLock(); }
}

/**
 * Cambios en bloque sobre palabras elegidas (solo quien administra). Primero vista previa (aplicar = false).
 * accion: 'categoria' | 'estado' | 'juegos' | 'fuente'. valor: id de categoría, estado canónico, 'si'/'no', id de fuente.
 * idIdioma: para 'fuente', solo esa lengua ('' = todas). ubicacion: para 'fuente'.
 */
function editarBloque(ids, accion, valor, idIdioma, ubicacion, aplicar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const sel = new Set((ids || []).map(_k));
  if (!sel.size) throw new Error(tx('estados.error.ninguna'));
  const tPal = _tablaConEstado(ss, CONFIG.SHEETS.PALABRAS), tTr = _tablaConEstado(ss, CONFIG.SHEETS.TRADUCCIONES);
  const filasP = tPal.filas.map((f, i) => sel.has(_k(_get(tPal, f, 'id_palabra'))) ? i : -1).filter(i => i > -1);
  const filasT = tTr.filas.map((f, i) => sel.has(_k(_get(tTr, f, 'id_palabra'))) && (!idIdioma || _k(_get(tTr, f, 'id_idioma')) === _k(idIdioma)) ? i : -1).filter(i => i > -1);
  let col, nuevo, tabla, filas;
  if (accion === 'categoria') {
    if (!readSheet(ss, CONFIG.SHEETS.CATEGORIAS).some(c => _k(c.id_categoria) === _k(valor))) throw new Error(tx('editar.error.categoria'));
    col = 'id_categoria'; nuevo = _s(valor); tabla = tPal; filas = filasP;
  } else if (accion === 'estado') {
    const canon = _estadoCanonico(valor);
    if (!ESTADOS.some(e => e.canon === canon)) throw new Error(tx('estados.error.estado'));
    col = 'estado'; nuevo = _etiquetaEstado(canon); tabla = null;
  } else if (accion === 'juegos') {
    col = 'usar_en_juegos'; nuevo = _k(valor) === 'no' ? 'No' : 'Sí'; tabla = tPal; filas = filasP;
  } else if (accion === 'fuente') {
    if (!readSheet(ss, CONFIG.SHEETS.FUENTES).some(f => _k(f.id_fuente) === _k(valor))) throw new Error(tx('editar.error.fuente'));
    col = 'id_fuente'; nuevo = _s(valor); tabla = tTr; filas = filasT;
  } else throw new Error(tx('editar.error.accion'));
  if (accion === 'estado') {
    const cambiaP = filasP.filter(i => _s(_get(tPal, tPal.filas[i], 'estado')) !== nuevo).length;
    const cambiaT = filasT.filter(i => _s(_get(tTr, tTr.filas[i], 'estado')) !== nuevo).length;
    if (!aplicar) return { palabras: cambiaP, traducciones: cambiaT };
  } else {
    const n = filas.filter(i => _s(_get(tabla, tabla.filas[i], col)) !== nuevo).length;
    if (!aplicar) return tabla === tTr ? { palabras: new Set(filas.map(i => _k(_get(tTr, tTr.filas[i], 'id_palabra')))).size, traducciones: n } : { palabras: n, traducciones: 0 };
  }
  const lock = _lock();
  try {
    let n = 0, nt = 0;
    if (accion === 'estado') {
      n = _cambiarEstadoFilas(tPal, f => sel.has(_k(_get(tPal, f, 'id_palabra'))), nuevo);
      nt = _cambiarEstadoFilas(tTr, f => sel.has(_k(_get(tTr, f, 'id_palabra'))) && (!idIdioma || _k(_get(tTr, f, 'id_idioma')) === _k(idIdioma)), nuevo);
    } else {
      if (_col(tabla, col) < 0) throw new Error(tx('comun.faltanColumnas', { hoja: tabla.nombre, columnas: col }));
      filas.forEach(i => {
        const f = tabla.filas[i], c = { [col]: nuevo };
        if (accion === 'fuente' && _s(ubicacion)) c.ubicacion_fuente = _s(ubicacion);
        if (_s(_get(tabla, f, col)) === nuevo && !c.ubicacion_fuente) return;
        _actualizar(tabla, i, c);
        tabla === tTr ? nt++ : n++;
      });
    }
    _historial(ss, tx('editar.historialBloque'), accion, { palabras: sel.size }, { [col]: nuevo, palabras: n, traducciones: nt });
    return { palabras: n, traducciones: nt };
  } finally { lock.releaseLock(); }
}

/** Publicar desde la ventana (solo quien administra; lo demás se publica solo cada hora). */
function editarPublicar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const r = publicarCatalogo(ss, true);
  return { conteos: r.conteos || {}, diagnostico: _porQueNoSeVen(ss) };
}

function uiEditar() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <style>
    body{margin:0}.caja{padding:10px}
    .cab{display:flex;gap:8px;align-items:center;flex-wrap:wrap;background:#f1f5f9;border-radius:8px;padding:6px 10px;margin-bottom:8px;font-size:13px}
    .cab .rama{background:#fef9c3;color:#854d0e;border-radius:10px;padding:1px 8px}
    .dos{display:grid;grid-template-columns:330px 1fr;gap:10px;align-items:start}
    .lista{max-height:560px;overflow:auto;border:1px solid #e2e8f0;border-radius:8px;background:#fff}
    .it{display:flex;gap:6px;align-items:flex-start;padding:6px 8px;border-bottom:1px solid #f1f5f9;cursor:pointer}
    .it:hover{background:#f8fafc}.it.on{background:#eef2ff}.it small{color:#64748b;display:block}
    .fila{display:flex;gap:6px;flex-wrap:wrap;align-items:center}.fila>input[type=text],.fila>select{flex:1;min-width:120px;width:auto}
    .chip{display:inline-block;border-radius:10px;padding:0 7px;font-size:11px;background:#f1f5f9;margin-left:4px}
    .ch-publicado{background:#dcfce7;color:#166534}.ch-revisión{background:#fef9c3;color:#854d0e}.ch-borrador{background:#e0e7ff;color:#3730a3}.ch-retirado{background:#fee2e2;color:#991b1b}
    .tr{border:1px solid #e2e8f0;border-radius:8px;padding:8px;margin:6px 0;background:#fff}.tr h4{margin:0 0 4px;font-size:14px}
    .tr textarea{width:100%;box-sizing:border-box;font-size:15px;min-height:38px}
    .final{font-size:12px;color:#166534}.bloque{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:8px;margin-bottom:8px}
    details summary{cursor:pointer}
    .sug{display:flex;flex-direction:column;gap:2px;margin:2px 0}.sug button{text-align:left;margin:0;font-weight:normal}
    .rels{margin:4px 0;padding-left:18px}.rels li{margin:4px 0}.rels button{padding:2px 8px;margin:0 0 0 4px;font-size:12px}
  </style>
  <div class="caja">
    <div class="cab" id="cab"></div>
    <div class="dos">
      <div>
        <div class="fila"><input type="text" id="q" placeholder="${txh('editar.buscar')}" aria-label="${txh('editar.buscar')}"></div>
        <div class="fila" style="margin-top:4px">
          <select id="fCat" aria-label="${txh('medios.categoria')}"></select>
          <select id="fEst" aria-label="${txh('estados.col.estado')}"></select>
        </div>
        <div class="fila" style="margin-top:4px"><select id="fFalta" aria-label="${txh('editar.f.falta')}"></select></div>
        <div class="fila" style="margin:6px 0"><label style="font-weight:normal"><input type="checkbox" id="todas"> ${txh('estados.selTodas')}</label><small id="cuenta"></small></div>
        <div class="lista" id="lista" role="listbox" aria-label="${txh('editar.lista')}"></div>
        <button class="prim" id="bNueva" style="margin-top:6px;width:100%">＋ ${txh('editar.nueva')}</button>
      </div>
      <div>
        <div id="bloque"></div>
        <div id="ficha"><p>${txh('editar.elige')}</p></div>
        <div id="msg" role="status" aria-live="polite"></div>
      </div>
    </div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null, F = null, sel = new Set(), actual = null, q = '', limite = 200;
    const ok = h => { $('msg').innerHTML = '<div class="ok">' + h + '</div>'; };
    const err = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    const plegar = s => String(s || '').normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').toLowerCase();
    const nomL = id => (D.idiomas.find(i => i.id === id) || {nombre: id}).nombre;
    const nomCat = id => (D.categorias.find(c => c.id === id) || {nombre: ''}).nombre;
    const chip = c => '<span class="chip ch-' + esc(c) + '">' + esc(tt('estado.' + c)) + '</span>';
    const ref = p => p.t[D.principal] || Object.values(p.t)[0] || tt('medios.sinTexto');
    function cargar(despues) {
      google.script.run.withFailureHandler(err).withSuccessHandler(d => {
        D = d;
        const c = D.catalogo;
        $('cab').innerHTML = '📚 <b>' + esc(c.titulo) + '</b> <small>(' + esc(c.archivo) + ')</small>' + (c.rama ? ' <span class="rama">🌿 ' + th('editar.esRama') + '</span>' : '') +
          ' · ' + th(D.soyAdmin ? 'editar.rolAdmin' : 'editar.rolCaptura') + (D.soyAdmin ? ' <button id="bPub" style="margin-left:auto">' + th('menu.publicar') + '</button>' : '');
        if ($('bPub')) $('bPub').onclick = publicar;
        const vc = $('fCat').value, ve = $('fEst').value, vf = $('fFalta').value;
        $('fCat').innerHTML = '<option value="">' + th('medios.todasCategorias') + '</option><option value="__sin">' + th('editar.sinCategoria') + '</option>' + D.categorias.map(x => '<option value="' + esc(x.id) + '">' + esc(x.nombre) + '</option>').join('');
        $('fEst').innerHTML = '<option value="">' + th('editar.todosEstados') + '</option>' + D.estados.map(x => '<option value="' + x.canon + '">' + esc(x.etiqueta) + '</option>').join('');
        $('fFalta').innerHTML = '<option value="">' + th('editar.f.nada') + '</option>' + D.idiomas.map(i => '<option value="L|' + esc(i.id) + '">' + esc(tt('editar.f.sinLengua', {lengua: i.nombre})) + '</option>').join('') +
          D.idiomas.map(i => '<option value="D|' + esc(i.id) + '">' + esc(tt('editar.f.sinDescripcion', {lengua: i.nombre})) + '</option>').join('') +
          '<option value="sinFuente">' + th('editar.f.sinFuente') + '</option><option value="sinImagen">' + th('medios.f.sinImagen') + '</option>';
        $('fCat').value = vc; $('fEst').value = ve; $('fFalta').value = vf;
        pintarLista(); pintarBloque();
        if (D.reparadas) ok(th('editar.reparadas', {n: D.reparadas}));
        if (despues) despues();
      }).editarDatos();
    }
    function visibles() {
      const c = $('fCat').value, e = $('fEst').value, f = $('fFalta').value, k = plegar(q);
      return D.palabras.filter(p => (!c || (c === '__sin' ? !p.cat : p.cat === c)) && (!e || p.estado === e) &&
        (!f || (f.startsWith('L|') ? !p.t[f.slice(2)] : f.startsWith('D|') ? (p.t[f.slice(2)] && !(p.d || {})[f.slice(2)]) : f === 'sinFuente' ? p.sinFuente : !p.img)) &&
        (!k || plegar(Object.values(p.t).join(' ') + ' ' + nomCat(p.cat)).includes(k)));
    }
    function pintarLista() {
      const v = visibles();
      $('cuenta').textContent = ' ' + tt('estados.cuenta', {n: Math.min(v.length, limite), total: v.length, sel: sel.size});
      $('lista').innerHTML = v.slice(0, limite).map(p => '<div class="it' + (actual === p.id ? ' on' : '') + '" data-id="' + esc(p.id) + '" role="option"><input type="checkbox" data-sel="' + esc(p.id) + '"' + (sel.has(p.id) ? ' checked' : '') + ' aria-label="' + esc(tt('estados.sel')) + '">' +
        '<div><b>' + esc(ref(p)) + '</b>' + chip(p.estado) + '<small>' + esc(Object.keys(p.t).filter(L => p.t[L] !== ref(p)).map(L => nomL(L) + ': ' + p.t[L]).join(' · ')) + (p.cat ? ' · 📁 ' + esc(nomCat(p.cat)) : '') + '</small></div></div>').join('') +
        (v.length > limite ? '<button id="mas" style="width:100%">' + th('estados.mas') + '</button>' : '') || '<p style="padding:8px"><small>' + th('estados.vacia') + '</small></p>';
      if ($('mas')) $('mas').onclick = () => { limite += 300; pintarLista(); };
      $('todas').checked = v.length > 0 && v.every(p => sel.has(p.id));
    }
    $('q').oninput = () => { q = $('q').value; limite = 200; clearTimeout(window._t); window._t = setTimeout(pintarLista, 250); };
    ['fCat', 'fEst', 'fFalta'].forEach(i => $(i).onchange = () => { limite = 200; pintarLista(); });
    $('todas').onchange = () => { visibles().forEach(p => $('todas').checked ? sel.add(p.id) : sel.delete(p.id)); pintarLista(); pintarBloque(); };
    $('lista').addEventListener('click', e => {
      const c = e.target.closest('[data-sel]');
      if (c) { c.checked ? sel.add(c.dataset.sel) : sel.delete(c.dataset.sel); pintarBloque(); $('cuenta').textContent = ' ' + tt('estados.cuenta', {n: Math.min(visibles().length, limite), total: visibles().length, sel: sel.size}); return; }
      const it = e.target.closest('[data-id]'); if (it) abrirFicha(it.dataset.id);
    });
    $('bNueva').onclick = () => { actual = null; F = { id: '', id_categoria: $('fCat').value && $('fCat').value !== '__sin' ? $('fCat').value : '', estado: D.soyAdmin ? 'publicado' : 'revisión', juegos: true, trads: {} }; pintarFicha(); pintarLista(); };
    function abrirFicha(id) {
      actual = id; pintarLista(); $('msg').innerHTML = '';
      $('ficha').innerHTML = '<p>' + th('comun.cargando') + '</p>';
      google.script.run.withFailureHandler(err).withSuccessHandler(f => { F = f; pintarFicha(); }).editarFicha(id);
    }
    const opts = (l, v) => l.map(x => '<option value="' + esc(x.v) + '"' + (x.v === v ? ' selected' : '') + '>' + esc(x.t) + '</option>').join('');
    function pintarFicha() {
      const nueva = !F.id, estados = D.estados.filter(e => D.soyAdmin || ['borrador', 'revisión'].includes(e.canon) || e.canon === F.estado);
      let h = '<h3 style="margin:0 0 6px">' + (nueva ? '＋ ' + th('editar.nueva') : '✏️ ' + esc(F.trads[D.principal] ? F.trads[D.principal].texto : (Object.values(F.trads)[0] || {}).texto || '')) + '</h3>' +
        '<div class="fila"><label for="pCat">📁</label><select id="pCat">' + '<option value="">' + th('editar.sinCategoria') + '</option>' + opts(D.categorias.map(c => ({v: c.id, t: c.nombre})), F.id_categoria) + '<option value="__nueva">＋ ' + th('editar.categoriaNueva') + '</option></select>' +
        '<input type="text" id="pCatNueva" placeholder="' + esc(tt('editar.categoriaNueva')) + '" style="display:none">' +
        '<select id="pEst" aria-label="' + esc(tt('estados.col.estado')) + '">' + opts(estados.map(e => ({v: e.canon, t: e.etiqueta})), F.estado) + '</select>' +
        '<label style="font-weight:normal"><input type="checkbox" id="pJuegos"' + (F.juegos ? ' checked' : '') + '> ' + th('estados.col.juegos') + '</label></div>' +
        (F.imagen ? '<small>🖼️ ' + th('editar.tieneImagen') + (F.imagen_nota ? ' · 📝 ' + esc(F.imagen_nota) : '') + '</small>' : '<small>' + th('editar.sinImagen') + '</small>') +
        (!nueva && F.galeria ? '<div id="gal"></div>' : '');
      h += D.idiomas.map(i => {
        const t = F.trads[i.id] || {}, existe = !!t.id;
        return '<div class="tr"><h4>' + esc(i.nombre) + ' ' + (existe ? chip(t.estado) : '<span class="chip">' + th('editar.sinTraduccion') + '</span>') + (t.audio ? ' <span class="chip">🔊</span>' : '') + (t.descripcion ? ' <span class="chip" title="' + esc(tt('editar.descripcion.tiene')) + '">📖</span>' : '') + '</h4>' +
          '<textarea data-lengua="' + esc(i.id) + '" data-k="texto" rows="1" placeholder="' + esc(tt('editar.escribe', {lengua: i.nombre})) + '">' + esc(t.texto || '') + '</textarea>' +
          '<div class="final" id="fin_' + esc(i.id) + '">' + (t.publicado && t.publicado !== t.texto ? esc(tt('editar.seGuardaComo', {texto: t.publicado})) : '') + '</div>' +
          '<details><summary><small>' + th('editar.mas') + '</small></summary><div class="fila">' +
            '<input type="text" data-lengua="' + esc(i.id) + '" data-k="nota_variante" placeholder="' + esc(tt('editar.notaVariante')) + '" value="' + esc(t.nota_variante || '') + '">' +
            '<input type="text" data-lengua="' + esc(i.id) + '" data-k="silabas" placeholder="' + esc(tt('editar.silabas')) + '" value="' + esc(t.silabas || '') + '"></div>' +
            '<textarea data-lengua="' + esc(i.id) + '" data-k="descripcion" rows="2" maxlength="600" style="font-size:13px" placeholder="' + esc(tt('editar.descripcion', {lengua: i.nombre})) + '" aria-label="' + esc(tt('editar.descripcion', {lengua: i.nombre})) + '">' + esc(t.descripcion || '') + '</textarea>' +
            '<small>' + th('editar.descripcion.d') + '</small>' +
            '<div class="fila"><select data-lengua="' + esc(i.id) + '" data-k="id_fuente" aria-label="' + esc(tt('fuentes.titulo')) + '"><option value="">' + th('editar.sinFuente') + '</option>' + opts(D.fuentes.map(f => ({v: f.id, t: f.titulo})), t.id_fuente) + '</select>' +
            '<input type="text" data-lengua="' + esc(i.id) + '" data-k="ubicacion_fuente" placeholder="' + esc(tt('bloque.ubicacion')) + '" value="' + esc(t.ubicacion_fuente || '') + '"></div>' +
            (existe ? '<div class="fila"><label>' + th('estados.col.estado') + '</label><select data-lengua="' + esc(i.id) + '" data-k="estado">' + opts(D.estados.filter(e => D.soyAdmin || ['borrador', 'revisión'].includes(e.canon) || e.canon === t.estado).map(e => ({v: e.canon, t: e.etiqueta})), t.estado) + '</select></div>' : '') +
            (t.audio_nota ? '<small>🔊 📝 ' + esc(t.audio_nota) + '</small>' : '') +
          '</details></div>';
      }).join('');
      h += '<div class="fila" style="margin-top:6px"><button class="prim" id="bGuardar">💾 ' + th('editar.guardar') + '</button>' +
        (!nueva && D.soyAdmin && F.estado !== 'retirado' ? '<button id="bRetirar">🗄️ ' + th('editar.retirar') + '</button>' : '') + '</div><div id="aviso"></div>' +
        '<div class="tr" id="relCaja"><h4>🔗 ' + th('rel.titulo') + '</h4><div id="rel"></div></div>';
      $('ficha').innerHTML = h;
      pintarRel(); pintarGal();
      $('pCat').onchange = () => { $('pCatNueva').style.display = $('pCat').value === '__nueva' ? '' : 'none'; };
      $('bGuardar').onclick = () => guardar(false);
      if ($('bRetirar')) $('bRetirar').onclick = () => { $('pEst').value = 'retirado'; guardar(false); };
      document.querySelectorAll('#ficha textarea').forEach(a => { a.oninput = () => { a.style.height = 'auto'; a.style.height = a.scrollHeight + 'px'; }; a.oninput(); });
    }
    function recoger() {
      const c = { id_palabra: F.id, estado: $('pEst').value, juegos: $('pJuegos').checked, trads: {} };
      if ($('pMaxImg') && $('pMaxImg').value.trim() !== ((F.galeria || {}).maxPropio || '')) c.imagenes_max = $('pMaxImg').value.trim();
      if ($('pCat').value === '__nueva') c.categoriaNueva = $('pCatNueva').value.trim(); else c.id_categoria = $('pCat').value;
      document.querySelectorAll('#ficha [data-lengua]').forEach(e => {
        const L = e.dataset.lengua, k = e.dataset.k, t = F.trads[L] || {};
        const v = e.value;
        const antes = k === 'texto' ? (t.texto || '') : (t[k] || '');
        if (v === antes) return;
        (c.trads[L] = c.trads[L] || {})[k] = v;
      });
      // La fuente acompaña al texto para decidir si se aplican atajos
      Object.keys(c.trads).forEach(L => { if (c.trads[L].id_fuente == null) c.trads[L].id_fuente = (F.trads[L] || {}).id_fuente || ''; });
      return c;
    }
    function guardar(forzar) {
      const c = recoger();
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(err).withSuccessHandler(r => {
        if (r.confirmar) {
          $('msg').innerHTML = '';
          $('aviso').innerHTML = '<div class="error"><b>' + th('editar.coinciden') + '</b><ul>' + r.coincidencias.map(x => '<li>' + esc(tt('editar.coincide', {forma: x.forma, lengua: x.lengua, palabras: x.palabras.join(' · ')})) + '</li>').join('') + '</ul>' +
            '<small>' + th('editar.coinciden.d') + '</small><div class="fila"><button class="prim" id="bForzar">' + th('editar.guardarIgual') + '</button><button id="bCancelar">' + th('conexion.cancelar') + '</button></div></div>';
          $('bForzar').onclick = () => { $('aviso').innerHTML = ''; guardar(true); };
          $('bCancelar').onclick = () => { $('aviso').innerHTML = ''; };
          return;
        }
        F = r.ficha; actual = r.id_palabra;
        ok(th(r.nuevos || r.cambios ? 'editar.guardado' : 'editar.sinCambios') + (r.avisos.length ? '<br>' + r.avisos.map(esc).join('<br>') : ''));
        cargar(pintarFicha);
      }).editarGuardar(c, !!forzar);
    }
    // ── Galería de imágenes de la palabra (las extra; la principal es la de arriba)
    function pintarGal() {
      const g = F.galeria, c = $('gal'); if (!c || !g) return;
      const th0 = id => '<img src="https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w200" alt="" style="height:70px;border:1px solid #e2e8f0;border-radius:4px">';
      c.innerHTML = '<details' + (g.extra.length ? ' open' : '') + '><summary><small>🖼️ ' + th('editar.galeria', {n: (g.principal ? 1 : 0) + g.extra.length, max: g.max}) + '</small></summary>' +
        (g.extra.length ? '<div class="fila">' + g.extra.map(x => '<div style="text-align:center;font-size:11px">' + th0(x.archivo) + '<br>' + esc(x.credito || '') + (x.soloInterna ? ' 🔒' : '') +
          (D.soyAdmin ? '<br><button class="peq" data-gal="principal" data-img="' + esc(x.id) + '" title="' + esc(tt('editar.gal.principal')) + '">⭐</button><button class="peq" data-gal="' + (x.soloInterna ? 'publica' : 'interna') + '" data-img="' + esc(x.id) + '" title="' + esc(tt(x.soloInterna ? 'editar.gal.publica' : 'editar.gal.interna')) + '">' + (x.soloInterna ? '🌐' : '🔒') + '</button><button class="peq" data-gal="retirar" data-img="' + esc(x.id) + '" title="' + esc(tt('rel.retirar')) + '">✕</button>' : '') + '</div>').join('') + '</div>' : '<small>' + th('editar.gal.vacia') + '</small>') +
        (D.soyAdmin ? '<div class="fila" style="margin-top:4px"><label for="pMaxImg" style="font-weight:normal">' + th('editar.gal.max') + '</label><input type="text" id="pMaxImg" style="max-width:70px" value="' + esc(g.maxPropio) + '" placeholder="' + esc(String(g.max)) + '"><small>' + th('editar.gal.max.d') + '</small></div>' : '') + '</details>';
    }
    $('ficha').addEventListener('click', e => {
      const b = e.target.closest('[data-gal]'); if (!b) return;
      if (b.dataset.gal === 'retirar' && !confirm(tt('editar.gal.retirarC'))) return;
      google.script.run.withFailureHandler(err).withSuccessHandler(g => { F.galeria = g; pintarGal(); ok(th('rel.actualizada')); }).galeriaAccion(b.dataset.img, b.dataset.gal);
    });
    // ── Relaciones con otras palabras (se guardan al momento, aparte del botón Guardar)
    let rElegida = '';
    function pintarRel() {
      const c = $('rel'); if (!c) return;
      if (!F.id) { c.innerHTML = '<small>' + th('rel.guardaPrimero') + '</small>'; return; }
      const l = F.relaciones || [];
      c.innerHTML = (l.length ? '<ul class="rels">' + l.map(r => '<li><span class="chip">' + esc(r.etiqueta) + '</span> <a href="#" data-ir="' + esc(r.otra) + '">' + esc(r.texto) + '</a>' + chip(r.estado) +
          (r.publica ? '' : ' <span class="chip" title="' + esc(tt('rel.noPublica.d')) + '">🔒 ' + th('rel.noPublica') + '</span>') + (r.nota ? ' <small>📝 ' + esc(r.nota) + '</small>' : '') +
          (D.soyAdmin && r.estado !== 'publicado' ? '<button data-rpub="' + esc(r.id) + '">✅ ' + th('rel.publicar') + '</button>' : '') +
          (D.soyAdmin || r.estado !== 'publicado' ? '<button data-rret="' + esc(r.id) + '" title="' + esc(tt('rel.retirar.d')) + '">✕ ' + th('rel.retirar') + '</button>' : '') + '</li>').join('') + '</ul>'
        : '<small>' + th('rel.ninguna') + '</small>') +
        (D.relOpciones.length ? '<div class="fila" style="margin-top:6px"><select id="rTipo" aria-label="' + esc(tt('rel.tipo')) + '">' + opts(D.relOpciones.map(o => ({v: o.v, t: o.t + (o.publica ? '' : ' 🔒')}))) + '</select>' +
          '<input type="text" id="rBusca" placeholder="' + esc(tt('rel.buscaPalabra')) + '" aria-label="' + esc(tt('rel.buscaPalabra')) + '" autocomplete="off"></div><div id="rSug" class="sug" role="listbox"></div>' +
          '<div class="fila"><input type="text" id="rNota" maxlength="300" placeholder="' + esc(tt('rel.nota')) + '" aria-label="' + esc(tt('rel.nota')) + '"><button class="prim" id="rAdd" disabled>＋ ' + th('rel.agregar') + '</button></div>' +
          '<small>' + th(D.soyAdmin ? 'rel.ayudaAdmin' : 'rel.ayudaCaptura') + '</small>' : '');
      rElegida = '';
      if (!$('rBusca')) return;
      $('rBusca').oninput = () => {
        rElegida = ''; $('rAdd').disabled = true;
        const k = plegar($('rBusca').value.trim());
        if (k.length < 1) { $('rSug').innerHTML = ''; return; }
        const m = D.palabras.filter(p => p.id !== F.id && plegar(Object.values(p.t).join(' ')).includes(k)).slice(0, 8);
        $('rSug').innerHTML = m.length ? m.map(p => '<button data-pick="' + esc(p.id) + '" role="option"><b>' + esc(ref(p)) + '</b> <small>' + esc(Object.keys(p.t).filter(L => p.t[L] !== ref(p)).map(L => p.t[L]).join(' · ')) + '</small></button>').join('') : '<small>' + th('rel.sinResultados') + '</small>';
      };
      $('rAdd').onclick = () => {
        if (!rElegida) return;
        $('msg').textContent = tt('comun.guardando');
        google.script.run.withFailureHandler(err).withSuccessHandler(r => { F.relaciones = r.relaciones; pintarRel(); ok(th('rel.agregada') + (r.aviso ? '<br>' + esc(r.aviso) : '')); })
          .relacionAgregar({ desde: F.id, hacia: rElegida, opcion: $('rTipo').value, nota: $('rNota').value });
      };
    }
    $('ficha').addEventListener('click', e => {
      const pick = e.target.closest('[data-pick]'), ir = e.target.closest('[data-ir]'), ret = e.target.closest('[data-rret]'), pub = e.target.closest('[data-rpub]');
      if (pick) { rElegida = pick.dataset.pick; const p = D.palabras.find(x => x.id === rElegida); $('rBusca').value = p ? ref(p) : ''; $('rSug').innerHTML = ''; $('rAdd').disabled = false; $('rAdd').focus(); }
      if (ir) { e.preventDefault(); abrirFicha(ir.dataset.ir); }
      const listo = r => { F.relaciones = r.relaciones; pintarRel(); ok(th('rel.actualizada')); };
      if (ret) google.script.run.withFailureHandler(err).withSuccessHandler(listo).relacionRetirar(ret.dataset.rret, F.id);
      if (pub) google.script.run.withFailureHandler(err).withSuccessHandler(listo).relacionEstado(pub.dataset.rpub, 'publicado', F.id);
    });
    // ── En bloque
    function pintarBloque() {
      if (!D || !D.soyAdmin || !sel.size) { $('bloque').innerHTML = ''; return; }
      $('bloque').innerHTML = '<div class="bloque"><b>' + th('editar.bloque', {n: sel.size}) + '</b><div class="fila" style="margin-top:4px">' +
        '<select id="bAcc"><option value="categoria">📁 ' + th('editar.b.categoria') + '</option><option value="estado">' + th('editar.b.estado') + '</option><option value="fuente">📚 ' + th('editar.b.fuente') + '</option><option value="juegos">🎮 ' + th('editar.b.juegos') + '</option></select>' +
        '<select id="bVal"></select><select id="bL" style="display:none"></select><input type="text" id="bUbic" placeholder="' + esc(tt('bloque.ubicacion')) + '" style="display:none"></div>' +
        '<div class="fila"><button id="bVer">' + th('editar.b.ver') + '</button><button class="prim" id="bAplicar" disabled>' + th('estados.aplicar') + '</button><button id="bLimpiar">' + th('editar.b.quitarSel') + '</button></div><div id="bRes"></div></div>';
      const valores = () => {
        const a = $('bAcc').value;
        $('bVal').innerHTML = a === 'categoria' ? opts(D.categorias.map(c => ({v: c.id, t: c.nombre}))) : a === 'estado' ? opts(D.estados.map(e => ({v: e.canon, t: e.etiqueta}))) :
          a === 'fuente' ? opts(D.fuentes.map(f => ({v: f.id, t: f.titulo}))) : opts([{v: 'si', t: tt('estados.juegosSi')}, {v: 'no', t: tt('estados.juegosNo')}]);
        $('bL').style.display = ['fuente', 'estado'].includes(a) ? '' : 'none';
        $('bL').innerHTML = '<option value="">' + th('editar.b.todasLenguas') + '</option>' + opts(D.idiomas.map(i => ({v: i.id, t: i.nombre})));
        $('bUbic').style.display = a === 'fuente' ? '' : 'none';
        $('bAplicar').disabled = true; $('bRes').innerHTML = '';
      };
      $('bAcc').onchange = valores; valores();
      ['bVal', 'bL'].forEach(i => $(i).onchange = () => { $('bAplicar').disabled = true; $('bRes').innerHTML = ''; });
      $('bLimpiar').onclick = () => { sel.clear(); pintarLista(); pintarBloque(); };
      const correr = aplicar => google.script.run.withFailureHandler(err).withSuccessHandler(r => {
        $('bRes').innerHTML = '<p>' + th(aplicar ? 'editar.b.listo' : 'editar.b.vista', {p: r.palabras, t: r.traducciones}) + '</p>';
        $('bAplicar').disabled = aplicar || !(r.palabras || r.traducciones);
        if (aplicar) cargar();
      }).editarBloque([...sel], $('bAcc').value, $('bVal').value, $('bL').style.display === 'none' ? '' : $('bL').value, $('bUbic').value, aplicar);
      $('bVer').onclick = () => correr(false);
      $('bAplicar').onclick = () => correr(true);
    }
    function publicar() {
      $('msg').textContent = tt('comun.cargando');
      google.script.run.withFailureHandler(err).withSuccessHandler(r => ok(th('editar.publicado', {palabras: r.conteos.palabras || 0, traducciones: r.conteos.traducciones || 0}) + (r.diagnostico ? '<br><small>' + esc(r.diagnostico).replace(/\\n/g, '<br>') + '</small>' : ''))).editarPublicar();
    }
    cargar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1000).setHeight(760), tx('editar.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 23. VISOR: MANTENIMIENTO, CLAVE DEL EQUIPO, AVISOS, IDENTIDAD Y ACTUALIZACIÓN
// ─────────────────────────────────────────────────────────────────────────────
// Todo se ajusta desde ventanas; nada se escribe a mano en las tablas.
//  - Mantenimiento «congelar»: el público ve la última versión anterior; el equipo, lo nuevo.
//  - Mantenimiento «cerrar»: el público ve una pantalla de mantenimiento (y, si se elige, los visores
//    borran su copia guardada al conectarse). El equipo, con su clave, sigue viendo todo.
//  - Avisos: una franja arriba del visor, con fechas, tipo y botón opcional.
//  - El estado del visor (avisos, mantenimiento, identidad, cada cuánto revisar) viaja en la consulta
//    pequeña «manifiesto», sin leer las hojas: se guarda listo en las propiedades del script.

const COLS_AVISOS = ['id_aviso', 'tipo', 'texto', 'desde', 'hasta', 'boton', 'enlace', 'cerrable', 'orden', 'estado', 'creado_por', 'fecha'];
const TIPOS_AVISO = ['info', 'novedad', 'pausa', 'importante'];
const CAMPOS_CONTACTO = ['nombre', 'correo', 'whatsapp', 'telefono', 'horario'];
const PROP_ESTADO_VISOR = 'ESTADO_VISOR', PROP_TOKEN_EQUIPO = 'TOKEN_EQUIPO_PUBLICADO';
const PROP_SNAP_CONGELADO = 'SNAPSHOT_CONGELADO', PROP_MANIF_CONGELADO = 'MANIFIESTO_CONGELADO';
const CACHE_CONGELADO = 'CATALOGO_CONG';

const _AJUSTES_VISOR = () => ({
  revision: 10, avisosMax: 1,
  identidad: { imagen: '', alt: '', enTitulo: false },
  mant: { activo: false, modo: 'congelar', mensaje: '', hasta: '', imagen: '', alt: '', contacto: ['correo', 'whatsapp'], purgar: true, avisoPublico: true, desde: '', quien: '', programado: false },
  programa: null,
  purga: '',
  impresion: { leyenda: '' }   // leyenda que el catálogo pone al pie de todo lo impreso
});

function _visorAjustes(ss) {
  const base = _AJUSTES_VISOR();
  let a = {};
  try { a = JSON.parse(_s(getAppConfig(ss).visorAjustes) || '{}'); } catch (e) { a = {}; }
  return Object.assign(base, a, { identidad: Object.assign(base.identidad, a.identidad || {}), mant: Object.assign(base.mant, a.mant || {}),
    impresion: Object.assign(base.impresion, a.impresion || {}) });
}

function _guardarVisorAjustes(ss, a) {
  _escribirConfig(ss, { VISOR_AJUSTES: JSON.stringify(a) });
  _publicarEstadoVisor(ss);
}

const _hoyISO = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

/** Avisos vigentes (o futuros) para el visor: los vencidos no viajan. */
function _avisosPublicos(ss) {
  const hoy = _hoyISO();
  return readSheet(ss, CONFIG.SHEETS.AVISOS).filter(a => _s(a.id_aviso) && _estadoCanonico(a.estado) === 'publicado' && _s(a.texto) && (!_fechaISO(a.hasta) || _fechaISO(a.hasta) >= hoy))
    .sort((a, b) => (Number(a.orden) || 0) - (Number(b.orden) || 0)).slice(0, 12)
    .map(a => ({ id: _s(a.id_aviso), tipo: TIPOS_AVISO.includes(_k(a.tipo)) ? _k(a.tipo) : 'info', texto: _s(a.texto).slice(0, 500),
      desde: _fechaISO(a.desde), hasta: _fechaISO(a.hasta), boton: _s(a.boton).slice(0, 60), enlace: _enlaceAviso(a.enlace), cerrable: _k(a.cerrable) !== 'no' }));
}

/** Fechas de la hoja (Date o texto) → 'aaaa-mm-dd'. */
function _fechaISO(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const m = _s(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}

/** Solo enlaces seguros: https://…, buscar:texto o categoria:id. */
function _enlaceAviso(v) {
  const s = _s(v);
  if (!s) return '';
  if (/^https:\/\/[^\s<>"']+$/i.test(s) || /^buscar:.{1,80}$/i.test(s) || /^categoria:[\w-]{1,80}$/i.test(s)) return s;
  return '';
}

/** Imagen pública para el visor: de Drive (se comparte para verse) o una dirección https. */
function _imagenPublica(v) {
  const s = _s(v);
  if (!s) return '';
  const id = _idDrive(s);
  if (id) {
    try { const f = DriveApp.getFileById(id); const a = f.getSharingAccess(); if (a !== DriveApp.Access.ANYONE_WITH_LINK && a !== DriveApp.Access.ANYONE) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* sin permiso: se intenta igual */ }
    return 'https://lh3.googleusercontent.com/d/' + id;
  }
  return /^https:\/\//i.test(s) ? s : '';
}

/** Arma el estado público del visor y lo guarda listo para la consulta pequeña. */
function _publicarEstadoVisor(ss) {
  const a = _visorAjustes(ss), cfg = getAppConfig(ss), c = cfg.contacto || {};
  let mant = null;
  if (a.mant.activo) {
    const contacto = {};
    (a.mant.contacto || []).forEach(k => { if (CAMPOS_CONTACTO.includes(k) && _s(c[k])) contacto[k] = _s(c[k]); });
    mant = { modo: a.mant.modo === 'cerrar' ? 'cerrar' : 'congelar', mensaje: _s(a.mant.mensaje).slice(0, 800), hasta: _s(a.mant.hasta),
      imagen: _imagenPublica(a.mant.imagen) || _imagenPublica(a.identidad.imagen), alt: _s(a.mant.alt) || _s(a.identidad.alt), contacto, avisoPublico: a.mant.avisoPublico !== false };
  }
  const estado = {
    titulo: cfg.titulo || '', subtitulo: cfg.subtitulo || '', revision: [0, 5, 10, 30].includes(Number(a.revision)) ? Number(a.revision) : 10,
    avisosMax: Math.min(3, Math.max(1, Number(a.avisosMax) || 1)), purga: _s(a.purga),
    identidad: { imagen: _imagenPublica(a.identidad.imagen), alt: _s(a.identidad.alt), enTitulo: !!a.identidad.enTitulo },
    avisos: _avisosPublicos(ss), mant, leyenda: _s(a.impresion.leyenda), propuestas: _propuestasPublicas(ss)
  };
  let txt = JSON.stringify(estado);
  while (txt.length > 8500 && estado.avisos.length) { estado.avisos.pop(); txt = JSON.stringify(estado); }   // límite de una propiedad
  const p = PropertiesService.getScriptProperties();
  p.setProperty(PROP_ESTADO_VISOR, txt);
  p.setProperty(PROP_TOKEN_EQUIPO, _s(cfg.claveEquipo).length > 2 ? _s(cfg.claveEquipo) : '');
  return estado;
}

function _estadoVisor() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_ESTADO_VISOR) || 'null'); } catch (e) { return null; }
}

/** 'equipo' | 'publico' | null (clave incorrecta). */
function _rolToken(t) {
  const eq = PropertiesService.getScriptProperties().getProperty(PROP_TOKEN_EQUIPO) || '';
  if (eq && String(t || '') === eq) return 'equipo';
  return _tokenValido(t) ? 'publico' : null;
}

// ── Versión congelada (para el público durante el mantenimiento «congelar») ──

function _congelar(ss) {
  const texto = _leerSnapshot(), props = PropertiesService.getScriptProperties();
  let f = null; const id = props.getProperty(PROP_SNAP_CONGELADO);
  if (id) { try { f = DriveApp.getFileById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (f) f.setContent(texto); else { f = _carpetaDelLibro(ss).createFile('catalogo_congelado.json', texto, 'application/json'); props.setProperty(PROP_SNAP_CONGELADO, f.getId()); }
  props.setProperty(PROP_MANIF_CONGELADO, JSON.stringify(_leerManifiesto() || {}));
  _guardarEnCache(texto, CACHE_CONGELADO);
}

function _descongelar() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_SNAP_CONGELADO);
  if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  [PROP_SNAP_CONGELADO, PROP_MANIF_CONGELADO].forEach(k => { try { props.deleteProperty(k); } catch (e) {} });
  try { CacheService.getScriptCache().remove(`${CACHE_CONGELADO}_meta`); } catch (e) {}
}

function _leerSnapshotCongelado() {
  const enCache = _leerDeCache(CACHE_CONGELADO);
  if (enCache) return enCache;
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SNAP_CONGELADO);
  if (!id) return null;
  try { const t = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'); _guardarEnCache(t, CACHE_CONGELADO); return t; } catch (e) { return null; }
}

function _manifiestoCongelado() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_MANIF_CONGELADO) || 'null'); } catch (e) { return null; }
}

// ── Lo que ve cada quien según el mantenimiento (lo usa doGet / doPost) ──

/** null = sigue normal; si no, la respuesta de mantenimiento para quien no es del equipo. */
function _respuestaMantenimiento(rol) {
  const e = _estadoVisor();
  if (!e || !e.mant || rol === 'equipo' || e.mant.modo !== 'cerrar') return null;
  return { status: 503, mantenimiento: e.mant, estado: _estadoParaVisor(e, rol) };
}
function _estadoParaVisor(e, rol) {
  if (!e) return null;
  return { titulo: e.titulo, subtitulo: e.subtitulo, revision: e.revision, avisosMax: e.avisosMax, purga: e.purga, identidad: e.identidad, avisos: e.avisos, mant: e.mant, leyenda: e.leyenda || '', propuestas: e.propuestas || null, rol };
}
function _publicoCongelado(rol) {
  const e = _estadoVisor();
  return !!(e && e.mant && e.mant.modo === 'congelar' && rol !== 'equipo' && PropertiesService.getScriptProperties().getProperty(PROP_SNAP_CONGELADO));
}

// ── Acciones desde Sheets ──

/** o = { modo, mensaje, hasta, imagen, alt, contacto:[], purgar, avisoPublico } */
function mantenimientoEncender(o, programado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!programado) _requerirAdmin(ss);
  const lock = _lock();
  try {
    const a = _visorAjustes(ss), modo = o && o.modo === 'cerrar' ? 'cerrar' : 'congelar';
    if (!_leerManifiesto()) publicarCatalogo(ss, true);
    if (modo === 'congelar') _congelar(ss); else _descongelar();
    a.mant = Object.assign(a.mant, {
      activo: true, modo, mensaje: _s(o.mensaje).slice(0, 800), hasta: _s(o.hasta).slice(0, 40), imagen: _s(o.imagen), alt: _s(o.alt).slice(0, 200),
      contacto: (o.contacto || []).filter(k => CAMPOS_CONTACTO.includes(k)), purgar: modo === 'cerrar' && o.purgar !== false,
      avisoPublico: o.avisoPublico !== false, desde: _ahora(), quien: _correoActual() || 'automático', programado: !!programado
    });
    if (a.mant.purgar) a.purga = Utilities.getUuid().slice(0, 8);
    _guardarVisorAjustes(ss, a);
    _historial(ss, 'MANTENIMIENTO', modo, {}, { activo: 'Sí', modo, purgar: a.mant.purgar ? 'Sí' : 'No', programado: programado ? 'Sí' : 'No' });
    return conexionDatos();
  } finally { lock.releaseLock(); }
}

function mantenimientoTerminar(programado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!programado) _requerirAdmin(ss);
  const lock = _lock();
  try {
    const a = _visorAjustes(ss), antes = a.mant.modo;
    a.mant.activo = false; a.mant.programado = false;
    if (programado) a.programa = null;
    _descongelar();
    _guardarVisorAjustes(ss, a);
    try { publicarCatalogo(ss, true); } catch (e) { /* en la próxima hora */ }
    _historial(ss, 'MANTENIMIENTO', antes, { activo: 'Sí' }, { activo: 'No' });
    return conexionDatos();
  } finally { lock.releaseLock(); }
}

/** Programa un mantenimiento: se enciende y se apaga solo (lo revisa la tarea de cada hora). o = mismos campos + inicio, fin, aviso */
function mantenimientoProgramar(o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const ini = _fechaHora(o.inicio), fin = _fechaHora(o.fin);
  if (!ini) throw new Error(tx('mant.error.inicio'));
  if (fin && fin <= ini) throw new Error(tx('mant.error.fin'));
  const a = _visorAjustes(ss);
  a.programa = Object.assign({}, o, { inicio: _s(o.inicio), fin: _s(o.fin), creado: _ahora(), quien: _correoActual() });
  _guardarVisorAjustes(ss, a);
  if (o.crearAviso && _s(o.textoAviso)) {
    _agregar(_hojaAvisos(ss), [{ id_aviso: Utilities.getUuid(), tipo: 'pausa', texto: _s(o.textoAviso).slice(0, 500), desde: _hoyISO(),
      hasta: fin ? Utilities.formatDate(fin, Session.getScriptTimeZone(), 'yyyy-MM-dd') : Utilities.formatDate(ini, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      cerrable: 'Sí', orden: 0, estado: _etiquetaEstado('publicado'), creado_por: _correoActual(), fecha: _ahora() }]);
    _publicarEstadoVisor(ss);
  }
  _historial(ss, 'MANTENIMIENTO', 'programa', {}, { inicio: o.inicio, fin: o.fin, modo: o.modo });
  return conexionDatos();
}

function mantenimientoCancelarPrograma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const a = _visorAjustes(ss); a.programa = null; _guardarVisorAjustes(ss, a);
  return conexionDatos();
}

/** 'aaaa-mm-ddThh:mm' (hora de la hoja) → Date. */
function _fechaHora(v) {
  const s = _s(v);
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) return null;
  try { return Utilities.parseDate(s, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"); } catch (e) { return null; }
}

/** Lo llama la tarea de cada hora: enciende o apaga el mantenimiento programado. */
function _revisarMantenimientoProgramado(ss) {
  const a = _visorAjustes(ss), p = a.programa;
  if (!p) return '';
  const ahora = new Date(), ini = _fechaHora(p.inicio), fin = _fechaHora(p.fin);
  if (!a.mant.activo && ini && ahora >= ini && (!fin || ahora < fin)) { mantenimientoEncender(p, true); return 'encendido'; }
  if (a.mant.activo && a.mant.programado && fin && ahora >= fin) { mantenimientoTerminar(true); return 'terminado'; }
  if (!a.mant.activo && fin && ahora >= fin) { a.programa = null; _guardarVisorAjustes(ss, a); }
  return '';
}

/** Todos los visores borran su copia guardada al conectarse (y descargan lo actual). */
function visorBorrarCopias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const a = _visorAjustes(ss); a.purga = Utilities.getUuid().slice(0, 8);
  _guardarVisorAjustes(ss, a);
  _historial(ss, 'VISOR', 'borrar copias', {}, { purga: a.purga });
  return conexionDatos();
}

/** o = { revision, avisosMax, identidad: { imagen, alt, enTitulo } } */
function visorAjustesGuardar(o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const a = _visorAjustes(ss);
  if (o.revision != null) a.revision = [0, 5, 10, 30].includes(Number(o.revision)) ? Number(o.revision) : 10;
  if (o.avisosMax != null) a.avisosMax = Math.min(3, Math.max(1, Number(o.avisosMax) || 1));
  if (o.identidad) {
    const im = _s(o.identidad.imagen);
    if (im && !_idDrive(im) && !/^https:\/\//i.test(im)) throw new Error(tx('visorAj.error.imagen'));
    a.identidad = { imagen: im, alt: _s(o.identidad.alt).slice(0, 200), enTitulo: !!o.identidad.enTitulo };
  }
  if (o.impresion) a.impresion = { leyenda: _s(o.impresion.leyenda).replace(/\s+/g, ' ').slice(0, 160) };
  _guardarVisorAjustes(ss, a);
  return conexionDatos();
}

/** accion: 'nueva' | 'quitar'. La clave del equipo deja ver el catálogo durante el mantenimiento. */
function claveEquipoCambiar(accion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lock = _lock();
  try {
    const antes = _s(getAppConfig(ss).claveEquipo);
    const nueva = accion === 'nueva' ? 'eq' + _nuevaClave() : accion === 'quitar' ? '' : null;
    if (nueva === null) throw new Error(tx('conexion.error.accion'));
    if (nueva && nueva === _s(getAppConfig(ss).apiToken)) throw new Error(tx('conexion.error.accion'));
    _escribirConfig(ss, { CLAVE_EQUIPO: nueva });
    _publicarEstadoVisor(ss);
    _historial(ss, 'CONFIG', 'CLAVE_EQUIPO', { clave: antes ? '•••' : '' }, { clave: nueva ? '•••' : '' });
    return conexionDatos();
  } finally { lock.releaseLock(); }
}

// ── Avisos ──

function _hojaAvisos(ss) {
  let t = _tabla(ss, CONFIG.SHEETS.AVISOS);
  if (!t) { asegurarEstructura(ss); t = _tabla(ss, CONFIG.SHEETS.AVISOS); }
  return t;
}

function avisosListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), hoy = _hoyISO();
  const t = _hojaAvisos(ss);
  return { soyAdmin: _soyAdmin(ss), hoy, avisosMax: _visorAjustes(ss).avisosMax,
    categorias: readSheet(ss, CONFIG.SHEETS.CATEGORIAS).filter(c => _s(c.id_categoria)).map(c => ({ id: _s(c.id_categoria), nombre: _s(c.nombre_categoria) })),
    lista: t.filas.filter(f => _s(_get(t, f, 'id_aviso')) && _estadoCanonico(_get(t, f, 'estado')) !== 'retirado').map(f => {
      const desde = _fechaISO(_get(t, f, 'desde')), hasta = _fechaISO(_get(t, f, 'hasta'));
      return { id: _s(_get(t, f, 'id_aviso')), tipo: _k(_get(t, f, 'tipo')) || 'info', texto: _s(_get(t, f, 'texto')), desde, hasta,
        boton: _s(_get(t, f, 'boton')), enlace: _s(_get(t, f, 'enlace')), cerrable: _k(_get(t, f, 'cerrable')) !== 'no', orden: Number(_get(t, f, 'orden')) || 0,
        visible: _estadoCanonico(_get(t, f, 'estado')) === 'publicado', vigencia: hasta && hasta < hoy ? 'vencido' : desde && desde > hoy ? 'futuro' : 'activo' };
    }).sort((a, b) => a.orden - b.orden) };
}

/** o = { id ('' = nuevo), tipo, texto, desde, hasta, boton, enlace, cerrable, orden, visible } */
function avisoGuardar(o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  o = o || {};
  if (!_s(o.texto)) throw new Error(tx('avisos.error.texto'));
  if (_s(o.enlace) && !_enlaceAviso(o.enlace)) throw new Error(tx('avisos.error.enlace'));
  if (_fechaISO(o.desde) && _fechaISO(o.hasta) && _fechaISO(o.hasta) < _fechaISO(o.desde)) throw new Error(tx('avisos.error.fechas'));
  const lock = _lock();
  try {
    const t = _hojaAvisos(ss);
    const fila = { tipo: TIPOS_AVISO.includes(_k(o.tipo)) ? _k(o.tipo) : 'info', texto: _s(o.texto).slice(0, 500), desde: _fechaISO(o.desde), hasta: _fechaISO(o.hasta),
      boton: _s(o.boton).slice(0, 60), enlace: _enlaceAviso(o.enlace), cerrable: o.cerrable === false ? 'No' : 'Sí', orden: Number(o.orden) || 0,
      estado: _etiquetaEstado(o.visible === false ? 'borrador' : 'publicado') };
    const i = _s(o.id) ? t.filas.findIndex(f => _k(_get(t, f, 'id_aviso')) === _k(o.id)) : -1;
    if (i > -1) _actualizar(t, i, fila);
    else _agregar(t, [Object.assign({ id_aviso: Utilities.getUuid(), creado_por: _correoActual(), fecha: _ahora() }, fila)]);
    _publicarEstadoVisor(ss);
    return avisosListar();
  } finally { lock.releaseLock(); }
}

function avisoRetirar(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const t = _hojaAvisos(ss), i = t.filas.findIndex(f => _k(_get(t, f, 'id_aviso')) === _k(id));
  if (i > -1) _actualizar(t, i, { estado: _etiquetaEstado('retirado') });
  _publicarEstadoVisor(ss);
  return avisosListar();
}

function avisosAjustes(max) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const a = _visorAjustes(ss); a.avisosMax = Math.min(3, Math.max(1, Number(max) || 1)); _guardarVisorAjustes(ss, a);
  return avisosListar();
}

function uiAvisos() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <style>.av{border:1px solid #e2e8f0;border-left-width:6px;border-radius:8px;padding:8px;margin:6px 0;background:#fff}
    .av.info{border-left-color:#2563eb}.av.novedad{border-left-color:#16a34a}.av.pausa{border-left-color:#d97706}.av.importante{border-left-color:#dc2626}
    .av small{color:#64748b}.fila{display:flex;gap:6px;flex-wrap:wrap;align-items:center}.fila>*{flex:1;min-width:140px}
    .chip{display:inline-block;border-radius:10px;padding:0 8px;font-size:11px;background:#f1f5f9;margin-left:4px}</style>
  <div class="caja">
    <p style="margin-top:0">${txh('avisos.intro')}</p>
    <div class="sec"><h3 id="fTit">${txh('avisos.nuevo')}</h3>
      <div class="fila"><select id="tipo">${TIPOS_AVISO.map(x => `<option value="${x}">${txh('avisos.tipo.' + x)}</option>`).join('')}</select>
        <label style="font-weight:normal;flex:0 0 auto"><input type="checkbox" id="cerrable" checked> ${txh('avisos.cerrable')}</label></div>
      <label for="texto">${txh('avisos.texto')}</label><textarea id="texto" rows="3" maxlength="500" style="width:100%;box-sizing:border-box"></textarea>
      <small>${txh('avisos.texto.d')}</small>
      <div class="fila"><div><label for="desde">${txh('avisos.desde')}</label><input type="date" id="desde"></div><div><label for="hasta">${txh('avisos.hasta')}</label><input type="date" id="hasta"></div></div>
      <details><summary>${txh('avisos.boton')}</summary>
        <div class="fila"><input type="text" id="boton" maxlength="60" placeholder="${txh('avisos.boton.texto')}">
          <select id="destino"><option value="">${txh('avisos.boton.ninguno')}</option><option value="https">${txh('avisos.boton.web')}</option><option value="buscar">${txh('avisos.boton.buscar')}</option><option value="categoria">${txh('avisos.boton.categoria')}</option></select></div>
        <div class="fila"><input type="text" id="destVal" placeholder=""><select id="destCat" style="display:none"></select></div></details>
      <label style="font-weight:normal"><input type="checkbox" id="visible" checked> ${txh('avisos.visible')}</label>
      <div class="fila" style="margin-top:6px"><button class="prim" id="guardar">💾 ${txh('editar.guardar')}</button><button id="limpiar">${txh('conexion.cancelar')}</button></div>
    </div>
    <div class="sec"><h3>${txh('avisos.lista')}</h3>
      <div class="fila"><label for="max" style="flex:0 0 auto">${txh('avisos.max')}</label><select id="max" style="flex:0 0 auto"><option>1</option><option>2</option><option>3</option></select></div>
      <div id="lista"></div></div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null, editando = '';
    const ok = h => { $('msg').innerHTML = '<div class="ok">' + h + '</div>'; };
    const err = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    const iconos = {info: 'ℹ️', novedad: '✨', pausa: '⏸️', importante: '⚠️'};
    function destinoUI() {
      const d = $('destino').value;
      $('destCat').style.display = d === 'categoria' ? '' : 'none'; $('destVal').style.display = d === 'https' || d === 'buscar' ? '' : 'none';
      $('destVal').placeholder = d === 'https' ? 'https://…' : tt('avisos.boton.buscarQue');
    }
    $('destino').onchange = destinoUI;
    function limpiar() { editando = ''; $('fTit').textContent = tt('avisos.nuevo'); ['texto', 'desde', 'hasta', 'boton', 'destVal'].forEach(i => $(i).value = ''); $('tipo').value = 'info'; $('cerrable').checked = true; $('visible').checked = true; $('destino').value = ''; destinoUI(); }
    $('limpiar').onclick = limpiar;
    function enlace() {
      const d = $('destino').value;
      return d === 'https' ? $('destVal').value.trim() : d === 'buscar' ? 'buscar:' + $('destVal').value.trim() : d === 'categoria' ? 'categoria:' + $('destCat').value : '';
    }
    $('guardar').onclick = () => {
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; pintar(); limpiar(); ok(th('avisos.guardado')); })
        .avisoGuardar({ id: editando, tipo: $('tipo').value, texto: $('texto').value, desde: $('desde').value, hasta: $('hasta').value, boton: $('boton').value, enlace: enlace(), cerrable: $('cerrable').checked, visible: $('visible').checked });
    };
    $('max').onchange = () => google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; pintar(); ok(th('avisos.guardado')); }).avisosAjustes($('max').value);
    function pintar() {
      $('max').value = String(D.avisosMax);
      $('destCat').innerHTML = D.categorias.map(c => '<option value="' + esc(c.id) + '">' + esc(c.nombre) + '</option>').join('');
      $('lista').innerHTML = D.lista.length ? D.lista.map((a, i) => '<div class="av ' + esc(a.tipo) + '"><b>' + iconos[a.tipo] + ' ' + esc(tt('avisos.tipo.' + a.tipo)) + '</b>' +
        '<span class="chip">' + esc(tt('avisos.v.' + (a.visible ? a.vigencia : 'oculto'))) + '</span><div>' + esc(a.texto) + '</div><small>' +
        esc([a.desde && tt('avisos.desdeF', {f: a.desde}), a.hasta && tt('avisos.hastaF', {f: a.hasta}), a.boton && ('🔘 ' + a.boton), a.cerrable ? '' : tt('avisos.noCerrable')].filter(Boolean).join(' · ')) + '</small>' +
        '<div class="fila"><button data-ed="' + i + '">✏️ ' + th('avisos.editar') + '</button><button data-ret="' + esc(a.id) + '">🗄️ ' + th('avisos.quitar') + '</button></div></div>').join('') : '<small>' + th('avisos.ninguno') + '</small>';
    }
    $('lista').addEventListener('click', e => {
      const b = e.target.closest('button'); if (!b) return;
      if (b.dataset.ed != null) { const a = D.lista[+b.dataset.ed]; editando = a.id; $('fTit').textContent = tt('avisos.editar');
        $('tipo').value = a.tipo; $('texto').value = a.texto; $('desde').value = a.desde; $('hasta').value = a.hasta; $('boton').value = a.boton; $('cerrable').checked = a.cerrable; $('visible').checked = a.visible;
        const en = a.enlace || ''; $('destino').value = en.startsWith('https') ? 'https' : en.startsWith('buscar:') ? 'buscar' : en.startsWith('categoria:') ? 'categoria' : '';
        destinoUI(); $('destVal').value = en.startsWith('https') ? en : en.replace(/^buscar:/, ''); if (en.startsWith('categoria:')) $('destCat').value = en.slice(10); window.scrollTo(0, 0); }
      if (b.dataset.ret) google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; pintar(); ok(th('avisos.quitado')); }).avisoRetirar(b.dataset.ret);
    });
    destinoUI();
    google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; pintar(); }).avisosListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(720), tx('avisos.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 17. VARIAS FORMAS, APÓSTROFOS Y COINCIDENCIAS
// ─────────────────────────────────────────────────────────────────────────────
// Quien captura escribe como le sale natural; el programa se encarga al publicar:
//  - "El rancho, La casa" (con , ; o /) = varias formas de decir lo mismo. El texto se guarda tal cual;
//    al publicar se agrega `formas` (la primera es la principal: cartas y tarjetas; todas valen en el dictado y la búsqueda).
//  - Una frase sin comas ('Iteirite mɨtixutuwe) es una sola forma.
//  - ‘ ’ ´ ` ' ʼ ꞌ se unifican en el que use la lengua según su alfabeto (solo al publicar; la hoja no cambia).
//  - Si una misma forma aparece en dos palabras, quien revisa decide en 🔎 Coincidencias: ideas distintas
//    (se anota en RELACIONES y deja de aparecer) o la misma idea (se unen; nada se borra, lo sobrante queda Retirado).

const APOSTROFOS = ['ʼ', 'ꞌ', '’', '‘', "'", 'ʻ', '´', '`'];
const _reApostrofo = /[ʼꞌ’‘'ʻ´`]/g;
// Valores de datos (no se traducen): tipos de relación que escribe esta herramienta
const REL_OTRA_IDEA = 'otra idea, misma forma';
const REL_UNIDA = 'unida a';

/** Apóstrofo de una lengua según su alfabeto. '' = el alfabeto usa varios distintos: no se toca nada. */
function _apostrofoDe(alfabeto) {
  const a = _s(alfabeto).normalize('NFC');
  const hay = APOSTROFOS.filter(c => a.includes(c));
  return hay.length > 1 ? '' : (hay[0] || "'");
}
function _mapaApostrofos(ss) {
  const m = {};
  readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { if (_s(i.id_idioma)) m[_k(i.id_idioma)] = _apostrofoDe(i.alfabeto); });
  return m;
}
function _unificarApostrofos(v, canon) {
  const s = _s(v).normalize('NFC');
  return canon ? s.replace(_reApostrofo, canon) : s;
}
/** "El rancho, La casa" → ['El rancho', 'La casa']. Sin separadores → [texto]. */
function _formas(v) {
  const vistos = new Set();
  return _s(v).split(/\s*[,;\/]\s*/).map(x => x.replace(/\s+/g, ' ').trim())
    .filter(x => { const k = _claveForma(x); return k && !vistos.has(k) && vistos.add(k); });
}
/** Para comparar formas: sin mayúsculas, con un solo tipo de apóstrofo y sin signos alrededor. */
function _claveForma(v) {
  return _k(v).normalize('NFC').replace(_reApostrofo, "'").replace(/\s+/g, ' ').replace(/^[¡¿"«(\s]+|[!?.:»")\s]+$/g, '');
}

// ── Coincidencias: la misma forma en varias palabras ─────────────────────────

const _retirado = v => _estadoCanonico(v) === 'retirado';
const _par = (a, b) => [_k(a), _k(b)].sort().join('|');

/** Grupos de palabras (no retiradas) que comparten una forma en la misma lengua y que nadie ha revisado. */
function _coincidencias(ss) {
  const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES);
  if (!tPal || !tTr) return [];
  const vivas = new Set(tPal.filas.filter(f => _s(_get(tPal, f, 'id_palabra')) && !_retirado(_get(tPal, f, 'estado'))).map(f => _k(_get(tPal, f, 'id_palabra'))));
  const ap = _mapaApostrofos(ss);
  const revisados = new Set(readSheet(ss, CONFIG.SHEETS.RELACIONES).filter(r => _k(r.tipo) === REL_OTRA_IDEA).map(r => _par(r.id_origen, r.id_destino)));
  const grupos = {};
  tTr.filas.forEach(f => {
    const p = _k(_get(tTr, f, 'id_palabra')), L = _k(_get(tTr, f, 'id_idioma'));
    if (!vivas.has(p) || !L || _retirado(_get(tTr, f, 'estado'))) return;
    const texto = _unificarApostrofos(_get(tTr, f, 'texto') || _get(tTr, f, 'texto_captura'), ap[L]);
    _formas(texto).forEach(fm => {
      const g = grupos[L + '|' + _claveForma(fm)] = grupos[L + '|' + _claveForma(fm)] || { idioma: L, forma: fm, palabras: new Set() };
      g.palabras.add(p);
    });
  });
  return Object.values(grupos).filter(g => {
    const l = [...g.palabras];
    if (l.length < 2) return false;
    for (let i = 0; i < l.length; i++) for (let j = i + 1; j < l.length; j++) if (!revisados.has(_par(l[i], l[j]))) return true;
    return false;
  }).map(g => ({ idioma: g.idioma, forma: g.forma, palabras: [...g.palabras] }));
}

/** Para la ventana: cada grupo con lo necesario para decidir (imagen, traducciones, notas, categoría). */
function coincidenciasListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grupos = _coincidencias(ss);
  if (!grupos.length) return [];
  const idiomas = {}, cats = {};
  readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { idiomas[_k(i.id_idioma)] = _s(i.nombre_propio || i.nombre_completo || i.nombre) || _s(i.id_idioma); });
  readSheet(ss, CONFIG.SHEETS.CATEGORIAS).forEach(c => { cats[_k(c.id_categoria)] = _s(c.nombre_categoria); });
  const pal = {};
  readSheet(ss, CONFIG.SHEETS.PALABRAS).forEach(p => {
    const id = _s(p.id_palabra), img = _s(p.imagen_id) || _idDrive(p.imagen_referencia);
    if (id) pal[_k(id)] = { id, categoria: cats[_k(p.id_categoria)] || '', estado: tx('estado.' + (_estadoCanonico(p.estado) || 'publicado')),
      imagen: img ? `https://drive.google.com/thumbnail?id=${encodeURIComponent(img)}&sz=w160` : _imagenDirecta(p.imagen_referencia), textos: [] };
  });
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).forEach(t => {
    const p = pal[_k(t.id_palabra)];
    if (p && !_retirado(t.estado)) p.textos.push({ lengua: idiomas[_k(t.id_idioma)] || _s(t.id_idioma), texto: _s(t.texto || t.texto_captura), nota: _s(t.nota_variante) });
  });
  return grupos.map(g => ({ forma: g.forma, lengua: idiomas[g.idioma] || g.idioma, palabras: g.palabras.map(id => pal[id]).filter(Boolean) }));
}

/** "Son ideas distintas": se anota en RELACIONES y el grupo deja de aparecer. */
function coincidenciaDistintas(ids) {
  ids = (ids || []).map(_s).filter(Boolean);
  if (ids.length < 2) throw new Error(tx('coinc.error.pocas'));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const tRel = _tabla(ss, CONFIG.SHEETS.RELACIONES);
    if (!tRel) throw new Error(tx('comun.faltaHoja', { hoja: CONFIG.SHEETS.RELACIONES }));
    const ya = new Set(tRel.filas.filter(f => _k(_get(tRel, f, 'tipo')) === REL_OTRA_IDEA).map(f => _par(_get(tRel, f, 'id_origen'), _get(tRel, f, 'id_destino'))));
    const nuevas = [];
    for (let i = 0; i < ids.length; i++) for (let j = i + 1; j < ids.length; j++) {
      if (ya.has(_par(ids[i], ids[j]))) continue;
      nuevas.push({ id_relacion: Utilities.getUuid(), id_origen: ids[i], tipo: REL_OTRA_IDEA, id_destino: ids[j], nota: '',
        estado: _etiquetaEstado('publicado'), creado_por: _correoActual() });
    }
    if (nuevas.length) _agregar(tRel, nuevas);
    return { n: nuevas.length };
  } finally { lock.releaseLock(); }
}

/**
 * "Es la misma idea": todo pasa a la palabra `queda`. No se borra nada:
 *  - las traducciones de las otras se mueven; si `queda` ya tiene esa lengua, las formas nuevas se suman a su texto
 *    ("rancho" + "casa" → "rancho, casa"), con su nota y su audio si no tenía, y la otra fila queda Retirado;
 *  - si `queda` no tiene imagen o categoría, toma la de otra;
 *  - las relaciones se apuntan a `queda`; las otras palabras quedan Retirado, con una relación "unida a" que lo recuerda.
 */
function coincidenciaUnir(queda, otras) {
  queda = _s(queda);
  otras = (otras || []).map(_s).filter(x => x && _k(x) !== _k(queda));
  if (!queda || !otras.length) throw new Error(tx('coinc.error.pocas'));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const tPal = _tabla(ss, CONFIG.SHEETS.PALABRAS), tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES), tRel = _tabla(ss, CONFIG.SHEETS.RELACIONES);
    const iPal = id => tPal.filas.findIndex(f => _k(_get(tPal, f, 'id_palabra')) === _k(id));
    const iQ = iPal(queda);
    if (iQ < 0 || otras.some(o => iPal(o) < 0)) throw new Error(tx('coinc.error.noEsta'));
    const ap = _mapaApostrofos(ss);
    const RET = _etiquetaEstado('retirado');
    const formasDe = f => _formas(_unificarApostrofos(_get(tTr, f, 'texto') || _get(tTr, f, 'texto_captura'), ap[_k(_get(tTr, f, 'id_idioma'))])).map(_claveForma);
    // Cada palabra tiene una traducción por lengua: si `queda` ya tiene esa lengua, las formas nuevas se suman a la suya
    const dequeda = {};
    tTr.filas.forEach((f, i) => { if (_k(_get(tTr, f, 'id_palabra')) === _k(queda) && !_retirado(_get(tTr, f, 'estado'))) dequeda[_k(_get(tTr, f, 'id_idioma'))] = dequeda[_k(_get(tTr, f, 'id_idioma'))] ?? i; });
    const rep = { movidas: 0, sumadas: 0, repetidas: 0, palabras: 0 };
    const otrasK = new Set(otras.map(_k));
    const juntar = (a, b, sep) => [_s(a), _s(b)].filter(Boolean).join(sep);
    // Antes de tocar nada: si habría que sumar formas a una traducción de una fuente que pide no alterarla, no se junta
    const ctxP = _contextoPermisos(ss);
    tTr.filas.forEach(f => {
      if (!otrasK.has(_k(_get(tTr, f, 'id_palabra'))) || _retirado(_get(tTr, f, 'estado'))) return;
      const L = _k(_get(tTr, f, 'id_idioma')), j = dequeda[L];
      if (j == null) return;
      const fq = tTr.filas[j], ya = new Set(formasDe(fq));
      const nuevas = _formas(_get(tTr, f, 'texto') || _get(tTr, f, 'texto_captura')).filter(x => !ya.has(_claveForma(_unificarApostrofos(x, ap[L]))));
      if (nuevas.length && _fuenteFiel(ctxP, _get(tTr, fq, 'id_fuente'))) throw new Error(tx('coinc.error.fiel', { texto: _s(_get(tTr, fq, 'texto')) }));
    });
    tTr.filas.forEach((f, i) => {
      if (!otrasK.has(_k(_get(tTr, f, 'id_palabra'))) || _retirado(_get(tTr, f, 'estado'))) return;
      const L = _k(_get(tTr, f, 'id_idioma'));
      if (dequeda[L] == null) { _actualizar(tTr, i, { id_palabra: queda }); dequeda[L] = i; rep.movidas++; return; }
      const j = dequeda[L], fq = tTr.filas[j];
      const ya = new Set(formasDe(fq));
      const nuevas = _formas(_get(tTr, f, 'texto') || _get(tTr, f, 'texto_captura')).filter(x => !ya.has(_claveForma(_unificarApostrofos(x, ap[L]))));
      const cambios = {};
      if (nuevas.length) {
        cambios.texto = juntar(_get(tTr, fq, 'texto'), nuevas.join(', '), ', ');
        if (_s(_get(tTr, fq, 'texto_captura'))) cambios.texto_captura = juntar(_get(tTr, fq, 'texto_captura'), nuevas.join(', '), ', ');
      }
      const nota = _s(_get(tTr, f, 'nota_variante'));
      if (nota && !_s(_get(tTr, fq, 'nota_variante')).includes(nota)) cambios.nota_variante = juntar(_get(tTr, fq, 'nota_variante'), nota, ' · ');
      if (!_s(_get(tTr, fq, 'audio')) && _s(_get(tTr, f, 'audio')))
        ['audio', 'hablante', 'lugar_hablante', 'consentimiento'].forEach(c => { if (_col(tTr, c) > -1) cambios[c] = _get(tTr, f, c); });
      if (Object.keys(cambios).length) _actualizar(tTr, j, cambios);
      _actualizar(tTr, i, { estado: RET });
      nuevas.length ? rep.sumadas++ : rep.repetidas++;
    });
    const fQ = tPal.filas[iQ];
    otras.forEach(o => {
      const i = iPal(o), f = tPal.filas[i], cambiosQ = {};
      if (!_s(_get(tPal, fQ, 'imagen_referencia')) && _s(_get(tPal, f, 'imagen_referencia')))
        ['imagen_referencia', 'imagen_id', 'imagen_ref', 'imagen_publica'].forEach(c => { if (_col(tPal, c) > -1) cambiosQ[c] = _get(tPal, f, c); });
      if (!_s(_get(tPal, fQ, 'id_categoria')) && _s(_get(tPal, f, 'id_categoria'))) cambiosQ.id_categoria = _get(tPal, f, 'id_categoria');
      if (Object.keys(cambiosQ).length) _actualizar(tPal, iQ, cambiosQ);
      _actualizar(tPal, i, { estado: RET });
      rep.palabras++;
    });
    if (tRel) {
      tRel.filas.forEach((f, i) => {
        const a = _k(_get(tRel, f, 'id_origen')), b = _k(_get(tRel, f, 'id_destino'));
        if (!otrasK.has(a) && !otrasK.has(b)) return;
        if (_k(_get(tRel, f, 'tipo')) === REL_UNIDA) return;
        const na = otrasK.has(a) ? queda : _get(tRel, f, 'id_origen'), nb = otrasK.has(b) ? queda : _get(tRel, f, 'id_destino');
        _actualizar(tRel, i, _k(na) === _k(nb) ? { estado: RET } : { id_origen: na, id_destino: nb });
      });
      _agregar(tRel, otras.map(o => ({ id_relacion: Utilities.getUuid(), id_origen: o, tipo: REL_UNIDA, id_destino: queda, nota: '',
        estado: RET, creado_por: _correoActual() })));
    }
    return rep;
  } finally { lock.releaseLock(); }
}

function uiCoincidencias() {
  const html = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:14px}
    .grupo{border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:12px;background:#f8fafc}
    .grupo h3{margin:0 0 8px;font-size:15px}
    .pals{display:flex;flex-wrap:wrap;gap:8px}
    .pal{flex:1;min-width:200px;background:#fff;border:1px solid #cbd5e1;border-radius:8px;padding:8px}
    .pal img{width:100%;max-height:110px;object-fit:contain;background:#f1f5f9;border-radius:6px}
    .sinimg{height:60px;display:flex;align-items:center;justify-content:center;background:#f1f5f9;border-radius:6px;color:#94a3b8;font-size:12px}
    .pal ul{margin:6px 0;padding-left:18px}.nota{color:#92400e;font-size:12px}
    small{color:#64748b}
    button{padding:8px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#e2e8f0;margin:8px 4px 0 0}
    .prim{background:#475569;color:#fff}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:8px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
  </style>
  ${_tScript()}
  <div class="caja">
    <p style="margin-top:0">${txh('coinc.intro')}</p>
    <div id="lista">${txh('comun.cargando')}</div>
  </div>
  <script>
    const $ = id => document.getElementById(id);
    const esc = _escT;
    let grupos = [];
    function pintar() {
      if (!grupos.length) { $('lista').innerHTML = '<div class="ok">' + th('coinc.ninguna') + '</div>'; return; }
      $('lista').innerHTML = grupos.map((g, gi) => '<div class="grupo" id="g' + gi + '"><h3>' + th('coinc.grupo', {forma: g.forma, lengua: g.lengua, n: g.palabras.length}) + '</h3><div class="pals">' +
        g.palabras.map((p, pi) => '<label class="pal"><input type="radio" name="q' + gi + '" value="' + esc(p.id) + '"' + (pi ? '' : ' checked') + '> <small>' + th('coinc.queda') + '</small>' +
          (p.imagen ? '<img src="' + esc(p.imagen) + '" alt="">' : '<div class="sinimg">' + th('coinc.sinImagen') + '</div>') +
          '<ul>' + p.textos.map(x => '<li><b>' + esc(x.lengua) + ':</b> ' + esc(x.texto) + (x.nota ? '<br><span class="nota">📝 ' + esc(x.nota) + '</span>' : '') + '</li>').join('') + '</ul>' +
          '<small>' + esc([p.categoria, p.estado].filter(Boolean).join(' · ')) + '</small></label>').join('') +
        '</div><button onclick="distintas(' + gi + ')">' + th('coinc.distintas') + '</button><button class="prim" onclick="unir(' + gi + ')">' + th('coinc.unir') + '</button>' +
        '<div><small>' + th('coinc.unir.d') + '</small></div><div id="m' + gi + '" role="status" aria-live="polite"></div></div>').join('');
    }
    const err = gi => e => { $('m' + gi).innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    const hecho = (gi, h) => { const c = $('g' + gi); c.innerHTML = '<div class="ok">' + h + '</div>'; };
    function distintas(gi) {
      google.script.run.withFailureHandler(err(gi)).withSuccessHandler(() => hecho(gi, th('coinc.anotado', {forma: grupos[gi].forma})))
        .coincidenciaDistintas(grupos[gi].palabras.map(p => p.id));
    }
    function unir(gi) {
      const q = (document.querySelector('input[name=q' + gi + ']:checked') || {}).value;
      google.script.run.withFailureHandler(err(gi)).withSuccessHandler(r => hecho(gi, th('coinc.unidas', r)))
        .coincidenciaUnir(q, grupos[gi].palabras.map(p => p.id).filter(id => id !== q));
    }
    google.script.run.withFailureHandler(e => $('lista').textContent = tt('comun.error', {m: e.message})).withSuccessHandler(l => { grupos = l; pintar(); }).coincidenciasListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(820).setHeight(640), tx('coinc.titulo'));
}
// ─────────────────────────────────────────────────────────────────────────────
// 18. FUENTES, PERMISOS, COMUNIDADES Y VOCES (TERRENO PARA UN CORPUS)
// ─────────────────────────────────────────────────────────────────────────────
// Modelo en forma de grafo (así se exporta también):
//   ACTOR (persona, universidad, institución, comunidad…) —AUTORIZA→ FUENTE (obra o recurso) ←VIENE_DE— TRADUCCIÓN
//   ACTOR / AUTORIZACIÓN —ACUERDO→ COMUNIDAD (acuerdos colectivos)       VOZ —PERTENECE→ COMUNIDAD
// El permiso de cada traducción NO se marca fila por fila: se calcula siguiendo el camino
// traducción → fuente → autorizaciones. Cambiar una autorización actualiza miles de traducciones.
//
// Principios:
//  - Registrar es un derecho de cada hablante: capturar, revisar y publicar dependen solo de la base de
//    uso de la fuente y del consentimiento de cada voz. Ningún grupo funciona como llave.
//  - El permiso COLECTIVO (de la comunidad) solo se exige para usos colectivos: IA y terceros.
//    Una comunidad puede tener varias constancias; si se contradicen o alguien impugna, queda En disputa
//    y se congelan solo esos usos. El catálogo no decide quién representa a un pueblo: registra lo que existe.
//  - Ninguna variante es la correcta; cada fuente y cada voz dicen su comunidad, variante y escritura.
//  - Los datos de contacto viven en un DIRECTORIO PRIVADO fuera de la carpeta del catálogo.
//  - Nada de esto exporta datos para IA.

// Valores de datos (no se traducen)
const USOS = ['publicar', 'didactico', 'adaptar', 'compartir', 'ia'];
const VALORES_USO = ['Sí', 'No', 'Pendiente'];
const TIPOS_ACTOR = ['Persona', 'Universidad o centro académico', 'Institución pública', 'Organización', 'Comunidad o asamblea', 'Editorial', 'Otro'];
const ROLES_AUTORIZACION = ['Autoría', 'Compilación', 'Titular de derechos', 'Comunidad', 'Otro'];
const ROLES_CON_DERECHOS = ['Autoría', 'Compilación', 'Titular de derechos'];
const TIPOS_PERMISO = ['Escrito', 'Oral grabado', 'Acceso público', 'Acuerdo de asamblea', 'Otro'];
const BASES_USO = ['Autorización expresa', 'Acceso público con condiciones', 'Acuerdo comunitario', 'Consentimiento de las voces', 'Aportación con consentimiento'];
const TIPOS_FUENTE = ['Diccionario', 'Libro', 'Texto técnico', 'Alfabeto o norma', 'Gramática o reglas', 'Grabación', 'Captura del equipo', 'Otro'];
const PERMISOS_IA = [
  { canon: 'comunidad',     etiqueta: 'Sí, para herramientas de la comunidad' },
  { canon: 'investigacion', etiqueta: 'Sí, también investigación con convenio' },
  { canon: 'no',            etiqueta: 'No' },
  { canon: 'pendiente',     etiqueta: 'Pendiente' }
];
const TIPOS_REGISTRO = ['Palabra', 'Frase', 'Oración'];
const ORIGINALIDAD = ['Original', 'Traducción'];
const FORMAS_PERMISO = ['Verbal grabado', 'Escrito', 'Firmado en el teléfono', 'Verbal sin grabar', 'Acta de asamblea'];

const COLS_ACTORES = ['id_actor', 'nombre', 'tipo', 'id_comunidad', 'contacto', 'credito', 'enlace', 'notas', 'estado'];
const COLS_AUTORIZACIONES = ['id_autorizacion', 'id_actor', 'id_fuente', 'id_comunidad', 'rol', 'tipo_permiso',
  'uso_publicar', 'uso_didactico', 'uso_adaptar', 'uso_compartir', 'uso_ia', 'condiciones', 'constancia', 'fecha', 'vigente', 'impugnada', 'notas', 'registrado_por'];
const COLS_COMUNIDADES = ['id_comunidad', 'nombre_propio', 'nombre_otro', 'id_idioma', 'variante', 'region', 'sistema_escritura', 'consulta', 'en_disputa', 'notas', 'estado'];
const COLS_FUENTES = ['id_fuente', 'titulo', 'autoria', 'anio', 'tipo', 'base_uso', 'lenguas', 'id_comunidad', 'sistema_escritura', 'credito', 'publicar_credito',
  'fidelidad', 'enlace', 'fecha_consulta', 'condiciones', 'licencia', 'uso_publicar', 'uso_didactico', 'uso_adaptar', 'uso_compartir', 'uso_ia', 'archivo', 'notas', 'estado'];
const COLS_VOCES = ['id_voz', 'etiqueta', 'nombre_publico', 'id_comunidad', 'lugar', 'id_idioma', 'variante', 'sistema_escritura', 'consentimiento', 'permiso_ia',
  'forma_permiso', 'fecha_permiso', 'evidencia_permiso', 'version_terminos', 'activo', 'notas'];
const COLS_HISTORIAL = ['fecha', 'quien', 'hoja', 'id', 'campo', 'antes', 'despues'];
const COLS_DIRECTORIO = ['id_voz', 'nombre_completo', 'telefono', 'correo', 'comunidad', 'notas', 'alta'];
const FUENTE_EQUIPO = 'equipo';
const PROP_DIRECTORIO = 'DIRECTORIO_VOCES_ID';

// ── Valores ──────────────────────────────────────────────────────────────────

const _siNo = v => /^s[ií]/.test(_sinAcentos(v));
/** 'si' | 'no' | 'pendiente' */
function _valorUso(v) {
  const s = _sinAcentos(v);
  if (/^s[ií]/.test(s)) return 'si';
  if (/^no\b/.test(s)) return 'no';
  return 'pendiente';
}
const _etiquetaUso = c => ({ si: 'Sí', no: 'No' })[_valorUso(c)] || 'Pendiente';
function _permisoIA(v) {
  const s = _sinAcentos(v);
  if (!s) return '';
  const e = PERMISOS_IA.find(p => p.canon === s || _sinAcentos(p.etiqueta) === s);
  if (e) return e.canon;
  if (/^si\b.*invest/.test(s)) return 'investigacion';
  if (/^si\b/.test(s)) return 'comunidad';
  if (/^no\b/.test(s)) return 'no';
  return 'pendiente';
}
const _etiquetaPermisoIA = c => (PERMISOS_IA.find(p => p.canon === _permisoIA(c)) || PERMISOS_IA[3]).etiqueta;
const _iaSi = c => c === 'comunidad' || c === 'investigacion';
function _baseUso(f) {
  const s = _sinAcentos(f && f.base_uso);
  if (!f || _k(f.id_fuente) === FUENTE_EQUIPO || /voces|aportacion/.test(s) || (!s && _sinAcentos(f.tipo) === 'captura del equipo')) return 'voces';   // aportación: el consentimiento quedó en la propia fuente
  if (/publico/.test(s)) return 'publico';
  if (/comunitari/.test(s)) return 'comunitario';
  return 'expresa';
}
const _solicitarIA = () => { try { return _siNo(getAppConfig(SpreadsheetApp.getActiveSpreadsheet()).solicitarIA); } catch (e) { return false; } };

/** Palabra, frase u oración. Si nadie lo indicó, se deduce de la primera forma. */
function _tipoRegistro(t) {
  const dado = TIPOS_REGISTRO.find(x => _sinAcentos(x) === _sinAcentos(t && t.tipo_registro));
  if (dado) return dado;
  const f = (_formas(t && t.texto)[0] || '').trim();
  if (!f) return '';
  const n = f.split(/\s+/).length;
  if (/[.!?…]$/.test(f) || /^[¿¡]/.test(f) || n >= 6) return 'Oración';
  return n >= 2 ? 'Frase' : 'Palabra';
}

// ── Contexto de permisos (se lee una vez y se consulta muchas) ──────────────

function _mapaPor(ss, hoja, campo) {
  const m = {};
  readSheet(ss, hoja).forEach(o => { if (_s(o[campo])) m[_k(o[campo])] = o; });
  return m;
}
const _mapaVoces = ss => _mapaPor(ss, CONFIG.SHEETS.VOCES, 'id_voz');
const _mapaFuentes = ss => _mapaPor(ss, CONFIG.SHEETS.FUENTES, 'id_fuente');

function _contextoPermisos(ss) {
  const ctx = { fuentes: _mapaFuentes(ss), voces: _mapaVoces(ss), comunidades: _mapaPor(ss, CONFIG.SHEETS.COMUNIDADES, 'id_comunidad'),
    actores: _mapaPor(ss, CONFIG.SHEETS.ACTORES, 'id_actor'), porFuente: {}, porComunidad: {}, memo: {} };
  // Se guardan también las no vigentes: si una persona o institución solo tiene autorizaciones retiradas, cuenta como No
  readSheet(ss, CONFIG.SHEETS.AUTORIZACIONES).forEach(a => {
    if (!_s(a.id_autorizacion)) return;
    if (_s(a.id_fuente)) (ctx.porFuente[_k(a.id_fuente)] = ctx.porFuente[_k(a.id_fuente)] || []).push(a);
    if (_s(a.id_comunidad)) (ctx.porComunidad[_k(a.id_comunidad)] = ctx.porComunidad[_k(a.id_comunidad)] || []).push(a);
  });
  return ctx;
}

/** ¿La fuente permite este uso? 'si' | 'no' | 'pendiente' | 'disputa'. Sin fuente = captura del equipo. */
function _usoFuente(ctx, idFuente, uso) {
  const k = _k(idFuente) || FUENTE_EQUIPO, mk = k + '|' + uso;
  if (mk in ctx.memo) return ctx.memo[mk];
  const f = ctx.fuentes[k];
  let r;
  if (!f) r = k === FUENTE_EQUIPO ? (uso === 'ia' ? 'pendiente' : 'si') : 'pendiente';
  else {
    const base = _baseUso(f);
    if (base === 'voces' || base === 'publico') {
      const v = _s(f['uso_' + uso]);
      r = v ? _valorUso(v) : (base === 'voces' && uso !== 'ia' ? 'si' : 'pendiente');
    } else {
      const roles = base === 'comunitario' ? ['Comunidad'] : ROLES_CON_DERECHOS;
      const vals = _votosPorActor((ctx.porFuente[k] || []).filter(a => roles.includes(_s(a.rol))), uso);
      if (!vals.length) r = 'pendiente';
      else if (vals.includes('disputa')) r = 'disputa';
      else r = vals.includes('no') ? 'no' : vals.every(v => v === 'si') ? 'si' : 'pendiente';
    }
  }
  ctx.memo[mk] = r;
  return r;
}
/**
 * Un voto por persona o institución: lo que dicen sus autorizaciones vigentes.
 * Fuentes: si todas las suyas están retiradas (no vigentes), su voto es No (quien tiene derechos los retiró).
 * Comunidades: una constancia retirada simplemente deja de contar (la comunidad dice No con un acuerdo en No).
 * Una impugnada vale 'disputa'.
 */
function _votosPorActor(auts, uso, retirada) {
  const por = {};
  auts.forEach(a => { (por[_k(a.id_actor) || _k(a.id_autorizacion)] = por[_k(a.id_actor) || _k(a.id_autorizacion)] || []).push(a); });
  return Object.values(por).map(l => {
    const vig = l.filter(a => _sinAcentos(a.vigente) !== 'no');
    if (!vig.length) return retirada || 'no';
    if (vig.some(a => _siNo(a.impugnada))) return 'disputa';
    const v = vig.map(a => _valorUso(a['uso_' + uso]));
    return v.includes('no') ? 'no' : v.includes('si') ? 'si' : 'pendiente';
  });
}

/** ¿La fuente exige no alterar su contenido? */
function _fuenteFiel(ctx, idFuente) {
  const f = ctx.fuentes[_k(idFuente)];
  return !!f && (_siNo(f.fidelidad) || _usoFuente(ctx, idFuente, 'adaptar') === 'no');
}
/** Acuerdo colectivo para IA de una comunidad: 'otorgado' | 'no' | 'pendiente' | 'disputa'. */
function _estadoColectivo(ctx, idComunidad) {
  const c = ctx.comunidades[_k(idComunidad)];
  if (!c) return 'pendiente';
  if (_siNo(c.en_disputa)) return 'disputa';
  const vals = _votosPorActor(ctx.porComunidad[_k(idComunidad)] || [], 'ia', 'pendiente');
  if (vals.includes('disputa')) return 'disputa';
  const dichos = vals.filter(v => v !== 'pendiente');
  if (!dichos.length) return 'pendiente';
  if (dichos.includes('si') && dichos.includes('no')) return 'disputa';
  return dichos[0] === 'si' ? 'otorgado' : 'no';
}
/** Comunidad de una traducción: la de su voz o, si no, la de su fuente. */
function _comunidadDe(ctx, t) {
  const v = _s(t.id_voz) && ctx.voces[_k(t.id_voz)];
  if (v && _s(v.id_comunidad)) return _s(v.id_comunidad);
  const f = ctx.fuentes[_k(t.id_fuente) || FUENTE_EQUIPO];
  return f ? _s(f.id_comunidad) : '';
}
/** ¿Esta traducción podría entrar a un corpus para IA? Todos los permisos deben coincidir. */
function _aptaIA(ctx, t) {
  const motivos = [];
  if (!_esPublico(t)) motivos.push('revision');
  if (_usoFuente(ctx, t.id_fuente, 'ia') !== 'si') motivos.push('fuente');
  if (_s(t.audio)) { const v = _s(t.id_voz) && ctx.voces[_k(t.id_voz)]; if (!v || !_iaSi(_permisoIA(v.permiso_ia))) motivos.push('voz'); }
  const com = _comunidadDe(ctx, t);
  if (!com) motivos.push('sinComunidad');
  else if (_estadoColectivo(ctx, com) !== 'otorgado') motivos.push('comunidad');
  return { apta: !motivos.length, motivos };
}

/** Completa una traducción con lo que dice su voz: el consentimiento propio de la grabación manda; si no hay, el de la voz. */
function _conVoz(t, voces) {
  const v = t && _s(t.id_voz) && voces[_k(t.id_voz)];
  if (!v) return t;
  if (!_s(t.consentimiento)) t.consentimiento = _s(v.consentimiento) || 'Pendiente';
  t.hablante = _consentimiento(v.consentimiento) === 'publico_nombre' ? _s(v.nombre_publico) : '';
  t.lugar_hablante = _s(v.lugar);
  return t;
}

// ── Historial (no se edita desde las ventanas) ──────────────────────────────

function _historial(ss, hoja, id, antes, despues) {
  const t = _tabla(ss, CONFIG.SHEETS.HISTORIAL);
  if (!t) return;
  const quien = _correoActual() || '?', fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const filas = Object.keys(despues || {}).filter(c => _s((antes || {})[c]) !== _s(despues[c]))
    .map(c => ({ fecha, quien, hoja, id: _s(id), campo: c, antes: _s((antes || {})[c]).slice(0, 300), despues: _s(despues[c]).slice(0, 300) }));
  if (filas.length) _agregar(t, filas);
}

// ── Administración y directorio privado (fuera de la carpeta del catálogo) ──

function _soyAdmin(ss) {
  const t = _tabla(ss, CONFIG.SHEETS.EQUIPO), yo = _correoActual();
  if (!t || !yo || !t.filas.some(f => _s(_get(t, f, 'correo')))) return true; // instalación inicial
  return t.filas.some(f => _k(_get(t, f, 'correo')) === yo && _rolEquipo(_get(t, f, 'rol')) === 'Administrador' && _activo(_get(t, f, 'activo')));
}
function _requerirAdmin(ss) { if (!_soyAdmin(ss)) throw new Error(tx('voces.error.soloAdmin')); }

function _directorio(ss, crear) {
  if (_esRama(ss)) { if (crear) throw new Error(tx('rama.error.directorio')); return null; } // los contactos solo viven en el principal
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_DIRECTORIO);
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (e) { if (!crear) return null; throw new Error(tx('voces.error.sinDirectorio')); } // existe pero no hay acceso: nunca se crea otro
  }
  if (!crear) return null;
  const cfg = getAppConfig(ss);
  const libro = SpreadsheetApp.create(tx('voces.directorio.archivo', { catalogo: cfg.titulo || ss.getName() }));
  const sh = libro.getSheets()[0].setName('DIRECTORIO');
  sh.getRange(1, 1, 1, COLS_DIRECTORIO.length).setValues([COLS_DIRECTORIO]).setFontWeight('bold').setBackground('#7f1d1d').setFontColor('white');
  sh.setFrozenRows(1);
  try {
    const carpeta = DriveApp.getRootFolder().createFolder(tx('voces.directorio.carpeta', { catalogo: cfg.titulo || ss.getName() }));
    DriveApp.getFileById(libro.getId()).moveTo(carpeta);
    props.setProperty('CARPETA_PRIVADA_ID', carpeta.getId());
  } catch (e) { /* se queda en Mi unidad, igual de privado */ }
  props.setProperty(PROP_DIRECTORIO, libro.getId());
  _sincronizarAccesoDirectorio(ss, libro);
  return libro;
}
function _sincronizarAccesoDirectorio(ss, libro) {
  libro = libro || _directorio(ss, false);
  if (!libro) return;
  const t = _tabla(ss, CONFIG.SHEETS.EQUIPO);
  if (!t) return;
  const ids = [libro.getId(), PropertiesService.getScriptProperties().getProperty('CARPETA_PRIVADA_ID')].filter(Boolean);
  const admins = new Set(t.filas.filter(r => _rolEquipo(_get(t, r, 'rol')) === 'Administrador' && _activo(_get(t, r, 'activo'))).map(r => _k(_get(t, r, 'correo'))).filter(Boolean));
  ids.forEach((id, n) => {
    let f;
    try { f = n === 0 ? DriveApp.getFileById(id) : DriveApp.getFolderById(id); } catch (e) { return; }
    let duenio = '';
    try { duenio = _k(f.getOwner().getEmail()); } catch (e) {}
    let actuales = [];
    try { actuales = f.getEditors().map(u => _k(u.getEmail())); } catch (e) {}
    admins.forEach(c => { if (c !== duenio && !actuales.includes(c)) { try { f.addEditor(c); } catch (e) {} } });
    actuales.forEach(c => { if (c !== duenio && !admins.has(c)) { try { f.removeEditor(c); } catch (e) {} } });
  });
}

// ── Voces ────────────────────────────────────────────────────────────────────

function _opcionesComunes(ss) {
  const idiomas = {};
  readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { if (_s(i.id_idioma)) idiomas[_k(i.id_idioma)] = _s(i.nombre_propio || i.nombre_completo) || _s(i.id_idioma); });
  const comunidades = readSheet(ss, CONFIG.SHEETS.COMUNIDADES).filter(c => _s(c.id_comunidad) && !_retirado(c.estado))
    .map(c => ({ id: _s(c.id_comunidad), nombre: _s(c.nombre_propio) + (_s(c.nombre_otro) ? ' (' + _s(c.nombre_otro) + ')' : ''), variante: _s(c.variante), idioma: _k(c.id_idioma) }));
  return { idiomas: Object.keys(idiomas).map(k => ({ id: k, nombre: idiomas[k] })), comunidades };
}

function vocesListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  _cfgMemo = null;
  const t = _tabla(ss, CONFIG.SHEETS.VOCES);
  if (!t) throw new Error(tx('comun.faltaHoja', { hoja: CONFIG.SHEETS.VOCES }));
  const dir = _directorio(ss, false);
  const priv = {};
  if (dir) {
    const d = dir.getSheets()[0].getDataRange().getValues(), h = d[0].map(_k);
    d.slice(1).forEach(r => { const o = {}; h.forEach((c, i) => { o[c] = _s(r[i]); }); if (o.id_voz) priv[_k(o.id_voz)] = o; });
  }
  const usos = {};
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).forEach(x => { const k = _k(x.id_voz); if (k) usos[k] = (usos[k] || 0) + 1; });
  const op = _opcionesComunes(ss);
  const nomCom = {}; op.comunidades.forEach(c => { nomCom[_k(c.id)] = c.nombre; });
  return Object.assign(op, {
    directorio: dir ? dir.getUrl() : '', directorioPerdido: !dir && !!PropertiesService.getScriptProperties().getProperty(PROP_DIRECTORIO),
    solicitarIA: _solicitarIA(),
    consentimientos: CONSENTIMIENTOS.map(c => c.etiqueta), permisos: PERMISOS_IA.map(p => p.etiqueta), formas: FORMAS_PERMISO,
    voces: t.filas.filter(f => _s(_get(t, f, 'id_voz'))).map(f => {
      const o = {}; COLS_VOCES.forEach(c => { o[c] = _s(_get(t, f, c)); });
      const p = priv[_k(o.id_voz)] || {};
      return Object.assign(o, { nombre_completo: p.nombre_completo || '', telefono: p.telefono || '', correo: p.correo || '', comunidad_privada: p.comunidad || '',
        notas_privadas: p.notas || '', grabaciones: usos[_k(o.id_voz)] || 0, comunidad: nomCom[_k(o.id_comunidad)] || '' });
    })
  });
}

function _siguienteVoz(t) {
  let n = 0;
  t.filas.forEach(f => { const m = _s(_get(t, f, 'id_voz')).match(/^V-(\d+)$/i); if (m) n = Math.max(n, Number(m[1])); });
  return 'V-' + String(n + 1).padStart(3, '0');
}

/** Agrega o actualiza una voz. Lo lingüístico va a VOCES; el contacto, al directorio privado. */
function vozGuardar(p) {
  p = p || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (!_s(p.etiqueta) && !_s(p.nombre_completo)) throw new Error(tx('voces.error.nombre'));
  const lock = _lock();
  try {
    const t = _tabla(ss, CONFIG.SHEETS.VOCES);
    const id = _s(p.id_voz) || _siguienteVoz(t);
    const i = t.filas.findIndex(f => _k(_get(t, f, 'id_voz')) === _k(id));
    const antes = {}; if (i > -1) COLS_VOCES.forEach(c => { antes[c] = _get(t, t.filas[i], c); });
    const cons = CONSENTIMIENTOS.find(c => c.canon === _consentimiento(p.consentimiento)) || CONSENTIMIENTOS[3];
    const publico = {
      id_voz: id, etiqueta: _s(p.etiqueta) || id,
      nombre_publico: cons.canon === 'publico_nombre' ? (_s(p.nombre_publico) || _s(p.nombre_completo)) : '',
      id_comunidad: _s(p.id_comunidad), lugar: _s(p.lugar), id_idioma: _s(p.id_idioma), variante: _s(p.variante), sistema_escritura: _s(p.sistema_escritura),
      consentimiento: cons.etiqueta,
      // Con el interruptor apagado, el permiso para IA no se pregunta: se conserva lo que había (o Pendiente)
      permiso_ia: _solicitarIA() && p.permiso_ia != null && p.permiso_ia !== '' ? _etiquetaPermisoIA(p.permiso_ia) : (_s(antes.permiso_ia) || 'Pendiente'),
      forma_permiso: FORMAS_PERMISO.includes(_s(p.forma_permiso)) ? _s(p.forma_permiso) : '',
      fecha_permiso: _s(p.fecha_permiso), evidencia_permiso: _s(p.evidencia_permiso),
      activo: p.activo === false || _sinAcentos(p.activo) === 'no' ? 'No' : 'Sí', notas: _s(p.notas)
    };
    const privado = { nombre_completo: _s(p.nombre_completo), telefono: _s(p.telefono), correo: _s(p.correo), comunidad: _s(p.comunidad_privada), notas: _s(p.notas_privadas) };
    const dir = Object.values(privado).some(Boolean) ? _directorio(ss, true) : null; // primero el directorio: si no hay acceso, no se escribe nada
    if (i > -1) _actualizar(t, i, publico); else _agregar(t, [publico]);
    const vig = {}; ['consentimiento', 'permiso_ia', 'forma_permiso', 'fecha_permiso', 'evidencia_permiso', 'id_comunidad'].forEach(c => { vig[c] = publico[c]; });
    _historial(ss, CONFIG.SHEETS.VOCES, id, antes, vig);
    if (dir) {
      const sh = dir.getSheets()[0];
      const d = sh.getDataRange().getValues(), h = d[0].map(_k);
      const fila = d.findIndex((r, j) => j > 0 && _k(r[h.indexOf('id_voz')]) === _k(id));
      const valores = COLS_DIRECTORIO.map(c => c === 'id_voz' ? id : c === 'alta' ? (fila > 0 ? d[fila][h.indexOf('alta')] : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')) : privado[c] || '');
      sh.getRange(fila > 0 ? fila + 1 : d.length + 1, 1, 1, valores.length).setNumberFormat('@').setValues([valores]);
    }
    return vocesListar();
  } finally { lock.releaseLock(); }
}

// ── Fuentes, actores, autorizaciones y comunidades ──────────────────────────

function _asegurarFuenteEquipo(ss) {
  const t = _tabla(ss, CONFIG.SHEETS.FUENTES);
  if (!t || t.filas.some(f => _k(_get(t, f, 'id_fuente')) === FUENTE_EQUIPO)) return;
  _agregar(t, [{ id_fuente: FUENTE_EQUIPO, titulo: tx('fuentes.equipo.titulo'), tipo: 'Captura del equipo', base_uso: 'Consentimiento de las voces',
    publicar_credito: 'No', uso_publicar: 'Sí', uso_didactico: 'Sí', uso_adaptar: 'Sí', uso_compartir: 'Sí', uso_ia: 'Pendiente',
    notas: tx('fuentes.equipo.nota'), estado: _etiquetaEstado('publicado') }]);
}

const _filaA = (t, f, cols) => { const o = {}; cols.forEach(c => { o[c] = _s(_get(t, f, c)); }); return o; };

/** Todo lo que necesita la ventana 📚 en una sola llamada. */
function permisosListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  const ctx = _contextoPermisos(ss);
  const tF = _tabla(ss, CONFIG.SHEETS.FUENTES), tA = _tabla(ss, CONFIG.SHEETS.ACTORES), tAu = _tabla(ss, CONFIG.SHEETS.AUTORIZACIONES), tC = _tabla(ss, CONFIG.SHEETS.COMUNIDADES);
  if (!tF || !tA || !tAu || !tC) throw new Error(tx('comun.faltaHoja', { hoja: [CONFIG.SHEETS.FUENTES, CONFIG.SHEETS.ACTORES, CONFIG.SHEETS.AUTORIZACIONES, CONFIG.SHEETS.COMUNIDADES].join(', ') }));
  const lotes = {};
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).forEach(t => { const l = _s(t.lote); if (l) lotes[l] = (lotes[l] || 0) + 1; });
  return Object.assign(_opcionesComunes(ss), {
    soyAdmin: _soyAdmin(ss), solicitarIA: _solicitarIA(),
    listas: { usos: USOS, valores: VALORES_USO, tiposActor: TIPOS_ACTOR, roles: ROLES_AUTORIZACION, tiposPermiso: TIPOS_PERMISO, bases: BASES_USO, tiposFuente: TIPOS_FUENTE },
    fuentes: tF.filas.filter(f => _s(_get(tF, f, 'id_fuente'))).map(f => {
      const o = _filaA(tF, f, COLS_FUENTES);
      o.permisos = {}; USOS.forEach(u => { o.permisos[u] = _usoFuente(ctx, o.id_fuente, u); });
      o.fiel = _fuenteFiel(ctx, o.id_fuente); o.base = _baseUso(o);
      return o;
    }),
    actores: tA.filas.filter(f => _s(_get(tA, f, 'id_actor'))).map(f => _filaA(tA, f, COLS_ACTORES)),
    autorizaciones: tAu.filas.filter(f => _s(_get(tAu, f, 'id_autorizacion'))).map(f => _filaA(tAu, f, COLS_AUTORIZACIONES)),
    comunidadesDet: tC.filas.filter(f => _s(_get(tC, f, 'id_comunidad'))).map(f => Object.assign(_filaA(tC, f, COLS_COMUNIDADES), { colectivo: _estadoColectivo(ctx, _get(tC, f, 'id_comunidad')) })),
    lotes: Object.keys(lotes).map(l => ({ lote: l, n: lotes[l] })),
    categorias: readSheet(ss, CONFIG.SHEETS.CATEGORIAS).filter(c => _s(c.id_categoria)).map(c => ({ id: _s(c.id_categoria), nombre: _s(c.nombre_categoria) })),
    corpus: _resumenCorpus(ss, ctx)
  });
}

/** Guarda una fila genérica (fuente, actor o comunidad) con historial. */
function _guardarFila(ss, hoja, campoId, cols, p, limpiar) {
  const t = _tabla(ss, hoja);
  if (!t) throw new Error(tx('comun.faltaHoja', { hoja }));
  const id = _s(p[campoId]) || Utilities.getUuid();
  const i = t.filas.findIndex(f => _k(_get(t, f, campoId)) === _k(id));
  const antes = i > -1 ? _filaA(t, t.filas[i], cols) : {};
  const o = { [campoId]: id };
  cols.forEach(c => { if (c !== campoId && c in p) o[c] = _s(p[c]); });
  limpiar(o, antes);
  if (i > -1) _actualizar(t, i, o); else _agregar(t, [o]);
  _historial(ss, hoja, id, antes, o);
  return id;
}

function fuenteGuardar(p) {
  p = p || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (!_s(p.titulo)) throw new Error(tx('fuentes.error.titulo'));
  const lock = _lock();
  try {
    _guardarFila(ss, CONFIG.SHEETS.FUENTES, 'id_fuente', COLS_FUENTES, p, o => {
      o.tipo = TIPOS_FUENTE.includes(o.tipo) ? o.tipo : 'Otro';
      o.base_uso = BASES_USO.includes(o.base_uso) ? o.base_uso : 'Autorización expresa';
      ['publicar_credito', 'fidelidad'].forEach(c => { if (c in o) o[c] = _siNo(o[c]) ? 'Sí' : 'No'; });
      USOS.forEach(u => { if ('uso_' + u in o) o['uso_' + u] = _etiquetaUso(o['uso_' + u]); });
      // Si la fuente pide no alterar su contenido, el crédito siempre se muestra
      if (o.fidelidad === 'Sí') o.publicar_credito = 'Sí';
      o.estado = _etiquetaEstado(_estadoCanonico(o.estado) || 'publicado');
    });
    return permisosListar();
  } finally { lock.releaseLock(); }
}

function actorGuardar(p) {
  p = p || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (!_s(p.nombre)) throw new Error(tx('actores.error.nombre'));
  const lock = _lock();
  try {
    _guardarFila(ss, CONFIG.SHEETS.ACTORES, 'id_actor', COLS_ACTORES, p, o => {
      o.tipo = TIPOS_ACTOR.includes(o.tipo) ? o.tipo : 'Otro';
      o.estado = _etiquetaEstado(_estadoCanonico(o.estado) || 'publicado');
    });
    return permisosListar();
  } finally { lock.releaseLock(); }
}

function comunidadGuardar(p) {
  p = p || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (!_s(p.nombre_propio) && !_s(p.nombre_otro)) throw new Error(tx('comunidades.error.nombre'));
  const lock = _lock();
  try {
    _guardarFila(ss, CONFIG.SHEETS.COMUNIDADES, 'id_comunidad', COLS_COMUNIDADES, p, o => {
      if (!o.nombre_propio) o.nombre_propio = o.nombre_otro;
      o.en_disputa = _siNo(o.en_disputa) ? 'Sí' : 'No';
      o.estado = _etiquetaEstado(_estadoCanonico(o.estado) || 'publicado');
    });
    return permisosListar();
  } finally { lock.releaseLock(); }
}

/** Una autorización = una relación actor → fuente y/o actor → comunidad, con su alcance por uso. */
function autorizacionGuardar(p) {
  p = p || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (!_s(p.id_actor)) throw new Error(tx('autorizaciones.error.actor'));
  if (!_s(p.id_fuente) && !_s(p.id_comunidad)) throw new Error(tx('autorizaciones.error.objeto'));
  const lock = _lock();
  try {
    _guardarFila(ss, CONFIG.SHEETS.AUTORIZACIONES, 'id_autorizacion', COLS_AUTORIZACIONES, p, (o, antes) => {
      o.rol = ROLES_AUTORIZACION.includes(o.rol) ? o.rol : (o.id_fuente ? 'Autoría' : 'Comunidad');
      o.tipo_permiso = TIPOS_PERMISO.includes(o.tipo_permiso) ? o.tipo_permiso : 'Escrito';
      USOS.forEach(u => { o['uso_' + u] = _etiquetaUso(o['uso_' + u] != null ? o['uso_' + u] : antes['uso_' + u]); });
      o.vigente = o.vigente != null && _sinAcentos(o.vigente) === 'no' ? 'No' : 'Sí';
      o.impugnada = _siNo(o.impugnada) ? 'Sí' : 'No';
      o.registrado_por = _s(antes.registrado_por) || _correoActual();
    });
    return permisosListar();
  } finally { lock.releaseLock(); }
}

/** Interruptor "Solicitar autorización para IA". Al activarlo, lo que estaba en silencio aparece como pendiente. */
function solicitarIAGuardar(activo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const antes = { SOLICITAR_IA: getAppConfig(ss).solicitarIA };
  const v = { SOLICITAR_IA: activo ? 'Sí' : 'No' };
  _escribirConfig(ss, v);
  _historial(ss, CONFIG.SHEETS.CONFIG, 'SOLICITAR_IA', antes, v);
  _cfgMemo = null;
  return permisosListar();
}

// ── Lotes y asignación en bloque ─────────────────────────────────────────────

/**
 * Asigna una fuente (y opcionalmente su ubicación) a un grupo de traducciones a la vez.
 * filtro: { lote, categoria, idioma, creado_por, fuente ('__sin' = sin fuente) }. aplicar=false solo cuenta.
 */
function asignarFuenteBloque(filtro, idFuente, ubicacion, aplicar, soloSinFuente) {
  filtro = filtro || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lock = aplicar ? _lock() : null;
  try {
    const t = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES), tP = _tabla(ss, CONFIG.SHEETS.PALABRAS);
    if (aplicar && !_s(idFuente)) throw new Error(tx('bloque.error.fuente'));
    if (aplicar && !_mapaFuentes(ss)[_k(idFuente)]) throw new Error(tx('bloque.error.fuente'));
    const catDe = {}; tP.filas.forEach(f => { catDe[_k(_get(tP, f, 'id_palabra'))] = _k(_get(tP, f, 'id_categoria')); });
    const indices = [];
    t.filas.forEach((f, i) => {
      if (!_s(_get(t, f, 'id_traduccion'))) return;
      if (filtro.lote && _s(_get(t, f, 'lote')) !== _s(filtro.lote)) return;
      if (filtro.idioma && _k(_get(t, f, 'id_idioma')) !== _k(filtro.idioma)) return;
      if (filtro.creado_por && _k(_get(t, f, 'creado_por')) !== _k(filtro.creado_por)) return;
      if (filtro.categoria && catDe[_k(_get(t, f, 'id_palabra'))] !== _k(filtro.categoria)) return;
      const fu = _k(_get(t, f, 'id_fuente'));
      if (filtro.fuente === '__sin' ? fu : (filtro.fuente && fu !== _k(filtro.fuente))) return;
      if (soloSinFuente && fu) return;
      indices.push(i);
    });
    if (!aplicar) return { n: indices.length };
    const c = _col(t, 'id_fuente'), cu = _col(t, 'ubicacion_fuente');
    if (c < 0) throw new Error(tx('comun.faltanColumnas', { hoja: t.nombre, columnas: 'id_fuente' }));
    const colF = t.filas.map(f => [f[c]]), colU = cu > -1 ? t.filas.map(f => [f[cu]]) : null;
    indices.forEach(i => { colF[i][0] = idFuente; if (colU && _s(ubicacion)) colU[i][0] = _s(ubicacion); });
    if (t.filas.length) {
      t.sh.getRange(2, c + 1, colF.length, 1).setValues(colF);
      if (colU && _s(ubicacion)) t.sh.getRange(2, cu + 1, colU.length, 1).setNumberFormat('@').setValues(colU);
    }
    _historial(ss, CONFIG.SHEETS.TRADUCCIONES, tx('bloque.historialId', { n: indices.length }), { id_fuente: tx('bloque.historialAntes', { filtro: JSON.stringify(filtro) }) }, { id_fuente: idFuente });
    return Object.assign(permisosListar(), { asignadas: indices.length });
  } finally { if (lock) lock.releaseLock(); }
}

// ── Corpus (solo cuenta; no exporta nada) ────────────────────────────────────

function _resumenCorpus(ss, ctx) {
  ctx = ctx || _contextoPermisos(ss);
  const idiomas = {};
  readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { idiomas[_k(i.id_idioma)] = _s(i.nombre_propio || i.nombre_completo) || _s(i.id_idioma); });
  const vivas = new Set(readSheet(ss, CONFIG.SHEETS.PALABRAS).filter(p => !_retirado(p.estado)).map(p => _k(p.id_palabra)));
  const por = {}, lenguasPorConcepto = {}, motivos = {};
  let audios = 0, audiosAptos = 0;
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).forEach(t => {
    if (_retirado(t.estado) || !vivas.has(_k(t.id_palabra)) || !_s(t.texto)) return;
    const L = _k(t.id_idioma), tipo = _tipoRegistro(t) || 'Palabra';
    const r = por[L] = por[L] || { lengua: idiomas[L] || L, Palabra: 0, Frase: 0, 'Oración': 0, conFuente: 0, aptas: 0, revisadas: 0 };
    r[tipo]++;
    if (_s(t.id_fuente)) r.conFuente++;
    if (_esPublico(t)) r.revisadas++;
    const a = _aptaIA(ctx, t);
    if (a.apta) r.aptas++; else a.motivos.forEach(m => { motivos[m] = (motivos[m] || 0) + 1; });
    if (_s(t.audio)) { audios++; if (a.apta) audiosAptos++; }
    (lenguasPorConcepto[_k(t.id_palabra)] = lenguasPorConcepto[_k(t.id_palabra)] || new Set()).add(L);
  });
  const pares = {};
  Object.values(lenguasPorConcepto).forEach(s => {
    const l = [...s].sort();
    for (let i = 0; i < l.length; i++) for (let j = i + 1; j < l.length; j++) { const k = (idiomas[l[i]] || l[i]) + ' – ' + (idiomas[l[j]] || l[j]); pares[k] = (pares[k] || 0) + 1; }
  });
  const vocesPend = Object.values(ctx.voces).filter(v => _sinAcentos(v.activo) !== 'no' && (!_s(v.permiso_ia) || _permisoIA(v.permiso_ia) === 'pendiente')).map(v => _s(v.id_voz));
  const fuentesPend = Object.values(ctx.fuentes).filter(f => _usoFuente(ctx, f.id_fuente, 'ia') === 'pendiente').map(f => _s(f.titulo));
  const comPend = Object.values(ctx.comunidades).filter(c => _estadoColectivo(ctx, c.id_comunidad) !== 'otorgado').map(c => _s(c.nombre_propio) + ' · ' + _estadoColectivo(ctx, c.id_comunidad));
  return { lenguas: Object.values(por), pares: Object.keys(pares).map(k => ({ par: k, n: pares[k] })), audios, audiosAptos, motivos,
    pendientes: { voces: vocesPend, fuentes: fuentesPend, comunidades: comPend } };
}

// ── Términos de uso y página pública "Acerca de este catálogo" ──────────────

const TERMINOS_PUNTOS = ['para', 'guarda', 'publica', 'ia', 'nunca', 'retiro', 'equipo', 'responde'];
/** Términos en el idioma de la interfaz, con los datos del catálogo y (si se da) de una comunidad. */
function terminosTexto(idComunidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _cfgMemo = null;
  const cfg = getAppConfig(ss), ia = _siNo(cfg.solicitarIA);
  const com = _s(idComunidad) && _mapaPor(ss, CONFIG.SHEETS.COMUNIDADES, 'id_comunidad')[_k(idComunidad)];
  const c = cfg.contacto || {};
  const v = {
    catalogo: cfg.titulo || ss.getName(), responsable: c.nombre || '—',
    contacto: [c.correo, c.whatsapp || c.telefono].filter(Boolean).join(' · ') || '—',
    comunidad: com ? [_s(com.nombre_propio), _s(com.variante), _s(com.region)].filter(Boolean).join(' · ') : tx('terminos.sinComunidad'),
    consulta: com && _s(com.consulta) ? _s(com.consulta) : tx('terminos.sinConsulta')
  };
  const puntos = TERMINOS_PUNTOS.filter(p => ia || p !== 'ia').map(p => ({ titulo: tx('terminos.' + p + '.t'), texto: tx('terminos.' + p, v) }));
  // La versión cambia si cambia el texto base (español) o el interruptor: así se sabe qué aceptó cada persona
  const base = TERMINOS_PUNTOS.filter(p => ia || p !== 'ia').map(p => T_ES['terminos.' + p]).join('|');
  return { titulo: tx('terminos.titulo', v), puntos, acepto: tx('terminos.acepto'), version: 'T' + _hash(base).slice(0, 8), ia };
}

/** Lo que el visor muestra en "Acerca de este catálogo": sin datos personales. */
function _gobernanzaPublica(ss) {
  const ctx = _contextoPermisos(ss);
  return {
    aviso: tx('gob.aviso'),
    comunidades: Object.values(ctx.comunidades).filter(c => _esPublico(c)).map(c => ({
      id_comunidad: _s(c.id_comunidad), nombre_propio: _s(c.nombre_propio), nombre_otro: _s(c.nombre_otro), variante: _s(c.variante), region: _s(c.region),
      consulta: _s(c.consulta), acuerdo_ia: _estadoColectivo(ctx, c.id_comunidad)
    })),
    terminos: terminosTexto('')
  };
}

// ── Ventanas ─────────────────────────────────────────────────────────────────

const _CSS_FICHAS = `
  <style>
    body{font-family:sans-serif;color:#1f2937;font-size:14px;margin:0}.caja{padding:14px}
    .sec{border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:12px;background:#f8fafc}
    h3{margin:0 0 6px;font-size:15px}
    label{display:block;margin:6px 0 2px;font-weight:bold;font-size:13px}
    input[type=text],input[type=date],select,textarea{width:100%;padding:7px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px;box-sizing:border-box;font-family:inherit}
    .dos{display:flex;gap:8px;flex-wrap:wrap}.dos>div{flex:1;min-width:140px}
    button{padding:8px 12px;border:none;border-radius:6px;font-weight:bold;cursor:pointer;background:#e2e8f0;margin:6px 4px 0 0}
    .prim{background:#475569;color:#fff}.peq{padding:4px 8px;margin:0}
    table{width:100%;border-collapse:collapse;font-size:13px}td,th{border-bottom:1px solid #e2e8f0;padding:6px;text-align:left;vertical-align:top}
    small{color:#64748b;font-weight:normal}
    .ok{background:#dcfce7;color:#166534;padding:8px;border-radius:6px;margin-top:8px}
    .error{background:#fee2e2;color:#b91c1c;padding:8px;border-radius:6px;margin-top:8px}
    .priv{border-left:4px solid #7f1d1d;padding-left:8px}
    .chip{display:inline-block;padding:2px 8px;border-radius:10px;font-size:12px;background:#e2e8f0;margin:1px}
    .verde{background:#dcfce7;color:#166534}.ambar{background:#fef3c7;color:#92400e}.rojo{background:#fee2e2;color:#b91c1c}
    .tabs{display:flex;gap:4px;flex-wrap:wrap;border-bottom:2px solid #e2e8f0;margin-bottom:10px}
    .tabs button{margin:0;border-radius:6px 6px 0 0;background:#f1f5f9}.tabs button[aria-selected=true]{background:#475569;color:#fff}
    details{margin-top:8px}summary{cursor:pointer;font-weight:bold}
    .usos td{padding:3px}.usos select{padding:3px}
  </style>`;

function uiVoces() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <div class="caja">
    <p style="margin-top:0">${txh('voces.intro')}</p>
    <div id="dir"></div>
    <div class="sec"><table><thead><tr><th>${txh('voces.col.codigo')}</th><th>${txh('voces.col.quien')}</th><th>${txh('voces.col.comunidad')}</th><th>${txh('voces.col.publicar')}</th><th class="ia">${txh('voces.col.ia')}</th><th>${txh('voces.col.grabaciones')}</th><th></th></tr></thead><tbody id="filas"><tr><td colspan="7">${txh('comun.cargando')}</td></tr></tbody></table>
      <button class="prim" onclick="editar(null)">${txh('voces.nueva')}</button></div>
    <div class="sec" id="form" style="display:none">
      <h3 id="fTitulo"></h3>
      <div class="priv">
        <p style="margin:0"><small>${txh('voces.privado.d')}</small></p>
        <label for="nombre_completo">${txh('voces.f.nombre')}</label><input type="text" id="nombre_completo">
        <div class="dos"><div><label for="telefono">${txh('voces.f.telefono')}</label><input type="text" id="telefono"></div>
          <div><label for="correo">${txh('voces.f.correo')}</label><input type="text" id="correo"></div></div>
        <label for="notas_privadas">${txh('voces.f.notasPrivadas')}</label><textarea id="notas_privadas" rows="2"></textarea>
      </div>
      <label for="etiqueta">${txh('voces.f.etiqueta')} <small>${txh('voces.f.etiqueta.d')}</small></label><input type="text" id="etiqueta">
      <div class="dos"><div><label for="id_comunidad">${txh('voces.f.comunidad')} <small>${txh('voces.f.comunidad.d')}</small></label><select id="id_comunidad"></select></div>
        <div><label for="id_idioma">${txh('voces.f.lengua')}</label><select id="id_idioma"></select></div></div>
      <div class="dos"><div><label for="variante">${txh('voces.f.variante')}</label><input type="text" id="variante"></div>
        <div><label for="sistema_escritura">${txh('voces.f.escritura')}</label><input type="text" id="sistema_escritura"></div></div>
      <label for="lugar">${txh('voces.f.lugar')} <small>${txh('voces.f.lugar.d')}</small></label><input type="text" id="lugar">
      <label for="consentimiento">${txh('voces.f.publicar')}</label><select id="consentimiento"></select>
      <div class="ia"><label for="permiso_ia">${txh('voces.f.ia')} <small>${txh('voces.f.ia.d')}</small></label><select id="permiso_ia"></select></div>
      <div class="dos"><div><label for="forma_permiso">${txh('voces.f.forma')}</label><select id="forma_permiso"></select></div>
        <div><label for="fecha_permiso">${txh('voces.f.fecha')}</label><input type="date" id="fecha_permiso"></div></div>
      <label for="evidencia_permiso">${txh('voces.f.evidencia')} <small>${txh('voces.f.evidencia.d')}</small></label><input type="text" id="evidencia_permiso">
      <label for="notas">${txh('voces.f.notas')}</label><input type="text" id="notas">
      <label style="font-weight:normal"><input type="checkbox" id="activo" checked> ${txh('voces.f.activo')}</label>
      <button class="prim" onclick="guardar()">${txh('voces.guardar')}</button><button onclick="cerrar()">${txh('voces.cancelar')}</button>
    </div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null, actual = null;
    const CAMPOS = ['nombre_completo','telefono','correo','notas_privadas','etiqueta','id_comunidad','id_idioma','variante','sistema_escritura','lugar','consentimiento','permiso_ia','forma_permiso','fecha_permiso','evidencia_permiso','notas'];
    const SELECTS = ['id_comunidad','id_idioma','consentimiento','permiso_ia','forma_permiso'];
    const opciones = (l, v) => l.map(x => '<option value="' + esc(x.v) + '"' + (String(x.v).toLowerCase() === String(v || '').toLowerCase() ? ' selected' : '') + '>' + esc(x.t) + '</option>').join('');
    const chip = (t, c) => '<span class="chip ' + c + '">' + esc(t) + '</span>';
    const color = t => /^Sí|^Público/.test(t) ? 'verde' : /^Pend/.test(t) || !t ? 'ambar' : 'rojo';
    function pintar(d) {
      D = d;
      document.querySelectorAll('.ia').forEach(e => e.style.display = d.solicitarIA ? '' : 'none');
      $('dir').innerHTML = d.directorio ? '<p><a href="' + esc(d.directorio) + '" target="_blank">' + th('voces.abrirDirectorio') + '</a> <small>' + th('voces.abrirDirectorio.d') + '</small></p>'
        : (d.directorioPerdido ? '<div class="error">' + th('voces.error.sinDirectorio') + '</div>' : '');
      $('filas').innerHTML = d.voces.length ? d.voces.map((v, i) => '<tr><td><b>' + esc(v.id_voz) + '</b></td><td>' + esc(v.nombre_completo || v.etiqueta) +
        (v.nombre_completo && v.etiqueta !== v.nombre_completo ? '<br><small>' + esc(v.etiqueta) + '</small>' : '') + (v.activo === 'No' ? '<br><small>' + th('voces.inactiva') + '</small>' : '') + '</td><td>' +
        esc(v.comunidad || '—') + (v.variante ? '<br><small>' + esc(v.variante) + '</small>' : '') + '</td><td>' +
        chip(v.consentimiento || 'Pendiente', color(v.consentimiento)) + '</td><td class="ia" style="display:' + (d.solicitarIA ? '' : 'none') + '">' + chip(v.permiso_ia || 'Pendiente', color(v.permiso_ia)) + '</td><td>' + v.grabaciones +
        '</td><td><button class="peq" onclick="editar(' + i + ')">' + th('voces.editar') + '</button></td></tr>').join('') : '<tr><td colspan="7"><small>' + th('voces.ninguna') + '</small></td></tr>';
    }
    function editar(i) {
      actual = i == null ? {} : D.voces[i];
      $('fTitulo').textContent = i == null ? tt('voces.nueva') : tt('voces.editando', {codigo: actual.id_voz});
      $('id_comunidad').innerHTML = opciones([{v: '', t: '—'}].concat(D.comunidades.map(x => ({v: x.id, t: x.nombre + (x.variante ? ' · ' + x.variante : '')}))), actual.id_comunidad);
      $('id_idioma').innerHTML = opciones([{v: '', t: '—'}].concat(D.idiomas.map(x => ({v: x.id, t: x.nombre}))), actual.id_idioma);
      $('consentimiento').innerHTML = opciones(D.consentimientos.map(x => ({v: x, t: x})), actual.consentimiento || 'Pendiente');
      $('permiso_ia').innerHTML = opciones(D.permisos.map(x => ({v: x, t: x})), actual.permiso_ia || 'Pendiente');
      $('forma_permiso').innerHTML = opciones([{v: '', t: '—'}].concat(D.formas.map(x => ({v: x, t: x}))), actual.forma_permiso || '');
      CAMPOS.forEach(c => { if (!SELECTS.includes(c)) $(c).value = actual[c] || ''; });
      $('activo').checked = actual.activo !== 'No';
      $('form').style.display = 'block'; $('msg').innerHTML = ''; $('form').scrollIntoView();
    }
    function cerrar() { $('form').style.display = 'none'; }
    function guardar() {
      const p = { id_voz: actual.id_voz || '', activo: $('activo').checked, comunidad_privada: actual.comunidad_privada || '' };
      CAMPOS.forEach(c => { p[c] = $(c).value; });
      if (!D.solicitarIA) p.permiso_ia = '';   // apagado: no se pregunta ni se cambia
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; })
        .withSuccessHandler(d => { pintar(d); cerrar(); $('msg').innerHTML = '<div class="ok">' + th('voces.guardada') + '</div>'; }).vozGuardar(p);
    }
    google.script.run.withFailureHandler(e => { $('filas').innerHTML = '<tr><td colspan="7"><div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div></td></tr>'; }).withSuccessHandler(pintar).vocesListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(860).setHeight(700), tx('voces.titulo'));
}

function uiFuentes() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <div class="caja">
    <div class="tabs" role="tablist">
      <button role="tab" data-t="fuentes" onclick="tab('fuentes')">${txh('perm.tab.fuentes')}</button>
      <button role="tab" data-t="actores" onclick="tab('actores')">${txh('perm.tab.actores')}</button>
      <button role="tab" data-t="comunidades" onclick="tab('comunidades')">${txh('perm.tab.comunidades')}</button>
      <button role="tab" data-t="bloque" onclick="tab('bloque')">${txh('perm.tab.bloque')}</button>
      <button role="tab" data-t="corpus" onclick="tab('corpus')">${txh('perm.tab.corpus')}</button>
    </div>
    <div id="vista">${txh('comun.cargando')}</div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null, TAB = 'fuentes', ed = null;
    const err = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; };
    const ok = m => { $('msg').innerHTML = '<div class="ok">' + m + '</div>'; };
    const opt = (l, v, vacio) => (vacio ? '<option value="">—</option>' : '') + l.map(x => { const val = typeof x === 'object' ? x.v : x, t = typeof x === 'object' ? x.t : x;
      return '<option value="' + esc(val) + '"' + (String(val).toLowerCase() === String(v || '').toLowerCase() ? ' selected' : '') + '>' + esc(t) + '</option>'; }).join('');
    const COLOR = { si: 'verde', otorgado: 'verde', no: 'rojo', disputa: 'rojo', pendiente: 'ambar' };
    const chipP = v => '<span class="chip ' + (COLOR[v] || 'ambar') + '">' + esc(tt('perm.valor.' + v)) + '</span>';
    const nombreActor = id => (D.actores.find(a => a.id_actor === id) || {}).nombre || id;
    const nombreFuente = id => (D.fuentes.find(f => f.id_fuente === id) || {}).titulo || id;
    const nombreCom = id => (D.comunidades.find(c => c.id === id) || {}).nombre || id;
    const usosVisibles = () => D.listas.usos.filter(u => u !== 'ia' || D.solicitarIA);
    function campo(id, etq, html) { return '<div><label for="' + id + '">' + etq + '</label>' + html + '</div>'; }
    const txt = (id, v) => '<input type="text" id="' + id + '" value="' + esc(v || '') + '">';
    const sel = (id, l, v, vacio) => '<select id="' + id + '">' + opt(l, v, vacio) + '</select>';
    const area = (id, v) => '<textarea id="' + id + '" rows="2">' + esc(v || '') + '</textarea>';
    const val = id => $(id) ? $(id).value : '';
    const usosTabla = (pre, o) => '<table class="usos"><tr>' + usosVisibles().map(u => '<th>' + th('perm.uso.' + u) + '</th>').join('') + '</tr><tr>' +
      usosVisibles().map(u => '<td>' + sel(pre + u, D.listas.valores, o['uso_' + u] || 'Pendiente') + '</td>').join('') + '</tr></table>';
    function tab(t) { TAB = t; ed = null; document.querySelectorAll('.tabs button').forEach(b => b.setAttribute('aria-selected', b.dataset.t === t)); $('msg').innerHTML = ''; pintar(); }
    function pintar() {
      if (!D) return;
      if (TAB === 'fuentes') return vFuentes();
      if (TAB === 'actores') return vActores();
      if (TAB === 'comunidades') return vComunidades();
      if (TAB === 'bloque') return vBloque();
      vCorpus();
    }
    // ── Fuentes
    function vFuentes() {
      let h = '<p style="margin-top:0"><small>' + th('fuentes.intro') + '</small></p><div class="sec"><table><tr><th>' + th('fuentes.col.obra') + '</th><th>' + th('fuentes.col.base') + '</th><th>' + th('fuentes.col.permisos') + '</th><th></th></tr>' +
        D.fuentes.map((f, i) => '<tr><td><b>' + esc(f.titulo) + '</b>' + (f.autoria || f.anio ? '<br><small>' + esc([f.autoria, f.anio].filter(Boolean).join(', ')) + '</small>' : '') + (f.fiel ? '<br><span class="chip">🔒 ' + th('fuentes.fiel') + '</span>' : '') +
          '</td><td><small>' + esc(f.base_uso || '') + '</small></td><td>' + usosVisibles().map(u => '<small>' + th('perm.uso.' + u) + '</small> ' + chipP(f.permisos[u])).join('<br>') +
          '</td><td>' + (D.soyAdmin ? '<button class="peq" onclick="editarF(' + i + ')">' + th('voces.editar') + '</button>' : '') + '</td></tr>').join('') +
        '</table>' + (D.soyAdmin ? '<button class="prim" onclick="editarF(-1)">' + th('fuentes.nueva') + '</button>' : '') + '</div>';
      if (ed && ed.tipo === 'fuente') h += formFuente(ed.o);
      $('vista').innerHTML = h;
    }
    function editarF(i) { ed = { tipo: 'fuente', o: i < 0 ? { tipo: 'Diccionario', base_uso: 'Autorización expresa', publicar_credito: 'Sí' } : D.fuentes[i] }; vFuentes(); $('fF').scrollIntoView(); }
    function formFuente(f) {
      const auts = D.autorizaciones.filter(a => f.id_fuente && a.id_fuente === f.id_fuente);
      const base = f.base_uso || 'Autorización expresa', propia = /público|voces/i.test(base);
      return '<div class="sec" id="fF"><h3>' + esc(f.titulo || tt('fuentes.nueva')) + '</h3><div class="dos">' +
        campo('f_titulo', th('fuentes.f.titulo'), txt('f_titulo', f.titulo)) + campo('f_autoria', th('fuentes.f.autoria'), txt('f_autoria', f.autoria)) + campo('f_anio', th('fuentes.f.anio'), txt('f_anio', f.anio)) + '</div><div class="dos">' +
        campo('f_tipo', th('fuentes.f.tipo'), sel('f_tipo', D.listas.tiposFuente, f.tipo)) + campo('f_base_uso', th('fuentes.f.base') + ' <small>' + th('fuentes.f.base.d') + '</small>', sel('f_base_uso', D.listas.bases, base)) + '</div><div class="dos">' +
        campo('f_lenguas', th('fuentes.f.lenguas'), txt('f_lenguas', f.lenguas)) + campo('f_id_comunidad', th('fuentes.f.comunidad'), sel('f_id_comunidad', D.comunidades.map(c => ({v: c.id, t: c.nombre})), f.id_comunidad, true)) +
        campo('f_sistema_escritura', th('fuentes.f.escritura'), txt('f_sistema_escritura', f.sistema_escritura)) + '</div>' +
        campo('f_credito', th('fuentes.f.credito') + ' <small>' + th('fuentes.f.credito.d') + '</small>', txt('f_credito', f.credito)) + '<div class="dos">' +
        campo('f_enlace', th('fuentes.f.enlace'), txt('f_enlace', f.enlace)) + campo('f_fecha_consulta', th('fuentes.f.consulta'), txt('f_fecha_consulta', f.fecha_consulta)) + campo('f_licencia', th('fuentes.f.licencia'), txt('f_licencia', f.licencia)) + '</div>' +
        campo('f_condiciones', th('fuentes.f.condiciones') + ' <small>' + th('fuentes.f.condiciones.d') + '</small>', area('f_condiciones', f.condiciones)) +
        '<label style="font-weight:normal"><input type="checkbox" id="f_fidelidad"' + (f.fidelidad === 'Sí' ? ' checked' : '') + '> ' + th('fuentes.f.fidelidad') + '</label>' +
        '<label style="font-weight:normal"><input type="checkbox" id="f_publicar_credito"' + (f.publicar_credito !== 'No' ? ' checked' : '') + '> ' + th('fuentes.f.mostrarCredito') + '</label>' +
        '<div id="usosPropios" style="display:' + (propia ? '' : 'none') + '"><label>' + th('fuentes.f.usosPropios') + '</label>' + usosTabla('f_uso_', f) + '</div>' +
        campo('f_archivo', th('fuentes.f.archivo') + ' <small>' + th('fuentes.f.archivo.d') + '</small>', txt('f_archivo', f.archivo)) + campo('f_notas', th('fuentes.f.notas'), area('f_notas', f.notas)) +
        '<button class="prim" onclick="guardarF()">' + th('voces.guardar') + '</button><button onclick="ed=null;vFuentes()">' + th('voces.cancelar') + '</button>' +
        (f.id_fuente && !propia ? '<h3 style="margin-top:14px">' + th('autorizaciones.titulo') + '</h3><p style="margin:0"><small>' + th('autorizaciones.d') + '</small></p>' + tablaAuts(auts) +
          '<button onclick="editarA(null,\\'' + esc(f.id_fuente) + '\\',\\'\\')">' + th('autorizaciones.nueva') + '</button>' : '') +
        (ed.a ? formAut(ed.a) : '') + '</div>';
    }
    document.addEventListener('change', e => { if (e.target.id === 'f_base_uso') $('usosPropios').style.display = /público|voces/i.test(e.target.value) ? '' : 'none'; });
    function guardarF() {
      const p = { id_fuente: ed.o.id_fuente || '', estado: ed.o.estado || '', fidelidad: $('f_fidelidad').checked ? 'Sí' : 'No', publicar_credito: $('f_publicar_credito').checked ? 'Sí' : 'No' };
      ['titulo','autoria','anio','tipo','base_uso','lenguas','id_comunidad','sistema_escritura','credito','enlace','fecha_consulta','licencia','condiciones','archivo','notas'].forEach(c => { p[c] = val('f_' + c); });
      usosVisibles().forEach(u => { if ($('f_uso_' + u)) p['uso_' + u] = val('f_uso_' + u); });
      google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; const f = D.fuentes.find(x => x.titulo === p.titulo); ed = f ? { tipo: 'fuente', o: f } : null; vFuentes(); ok(th('fuentes.guardada')); }).fuenteGuardar(p);
    }
    // ── Autorizaciones (dentro de fuente o de comunidad)
    function tablaAuts(l) {
      if (!l.length) return '<p><small>' + th('autorizaciones.ninguna') + '</small></p>';
      return '<table><tr><th>' + th('autorizaciones.col.quien') + '</th><th>' + th('autorizaciones.col.permiso') + '</th><th>' + th('fuentes.col.permisos') + '</th><th></th></tr>' + l.map(a => '<tr><td><b>' + esc(nombreActor(a.id_actor)) + '</b><br><small>' + esc(a.rol) + '</small>' +
        (a.vigente === 'No' ? '<br><span class="chip rojo">' + th('autorizaciones.noVigente') + '</span>' : '') + (a.impugnada === 'Sí' ? '<br><span class="chip rojo">' + th('autorizaciones.impugnada') + '</span>' : '') +
        '</td><td><small>' + esc(a.tipo_permiso) + (a.fecha ? ' · ' + esc(a.fecha) : '') + '</small>' + (a.constancia ? '<br><a href="' + esc(a.constancia) + '" target="_blank">' + th('autorizaciones.verConstancia') + '</a>' : '<br><small>' + th('autorizaciones.sinConstancia') + '</small>') +
        '</td><td>' + usosVisibles().map(u => '<small>' + th('perm.uso.' + u) + ': ' + esc(a['uso_' + u] || 'Pendiente') + '</small>').join('<br>') +
        '</td><td><button class="peq" onclick="editarA(\\'' + esc(a.id_autorizacion) + '\\')">' + th('voces.editar') + '</button></td></tr>').join('') + '</table>';
    }
    function editarA(id, fuente, comunidad) {
      const a = id ? D.autorizaciones.find(x => x.id_autorizacion === id) : { id_fuente: fuente || '', id_comunidad: comunidad || '', rol: comunidad ? 'Comunidad' : 'Autoría', tipo_permiso: comunidad ? 'Acuerdo de asamblea' : 'Escrito', vigente: 'Sí' };
      ed.a = a; pintar(); $('fA').scrollIntoView();
    }
    function formAut(a) {
      return '<div class="sec" id="fA" style="background:#fff"><h3>' + th(a.id_autorizacion ? 'autorizaciones.editar' : 'autorizaciones.nueva') + '</h3><div class="dos">' +
        campo('a_id_actor', th('autorizaciones.f.actor') + ' <small>' + th('autorizaciones.f.actor.d') + '</small>', sel('a_id_actor', D.actores.map(x => ({v: x.id_actor, t: x.nombre + ' · ' + x.tipo})), a.id_actor, true)) +
        campo('a_rol', th('autorizaciones.f.rol'), sel('a_rol', D.listas.roles, a.rol)) + campo('a_tipo_permiso', th('autorizaciones.f.tipo'), sel('a_tipo_permiso', D.listas.tiposPermiso, a.tipo_permiso)) + '</div>' +
        '<label>' + th('autorizaciones.f.alcance') + '</label>' + usosTabla('a_uso_', a) + '<div class="dos">' +
        campo('a_constancia', th('autorizaciones.f.constancia'), txt('a_constancia', a.constancia)) + campo('a_fecha', th('autorizaciones.f.fecha'), '<input type="date" id="a_fecha" value="' + esc(a.fecha || '') + '">') + '</div>' +
        campo('a_condiciones', th('autorizaciones.f.condiciones'), area('a_condiciones', a.condiciones)) + campo('a_notas', th('fuentes.f.notas'), txt('a_notas', a.notas)) +
        '<label style="font-weight:normal"><input type="checkbox" id="a_vigente"' + (a.vigente !== 'No' ? ' checked' : '') + '> ' + th('autorizaciones.f.vigente') + '</label>' +
        '<label style="font-weight:normal"><input type="checkbox" id="a_impugnada"' + (a.impugnada === 'Sí' ? ' checked' : '') + '> ' + th('autorizaciones.f.impugnada') + '</label>' +
        '<button class="prim" onclick="guardarA()">' + th('voces.guardar') + '</button><button onclick="ed.a=null;pintar()">' + th('voces.cancelar') + '</button></div>';
    }
    function guardarA() {
      const a = ed.a, p = { id_autorizacion: a.id_autorizacion || '', id_fuente: a.id_fuente || '', id_comunidad: a.id_comunidad || '', vigente: $('a_vigente').checked ? 'Sí' : 'No', impugnada: $('a_impugnada').checked ? 'Sí' : 'No' };
      ['id_actor','rol','tipo_permiso','constancia','fecha','condiciones','notas'].forEach(c => { p[c] = val('a_' + c); });
      usosVisibles().forEach(u => { p['uso_' + u] = val('a_uso_' + u); });
      google.script.run.withFailureHandler(err).withSuccessHandler(d => { const keep = ed; D = d; ed = keep; ed.a = null;
        if (ed.o) ed.o = (ed.tipo === 'fuente' ? D.fuentes.find(x => x.id_fuente === ed.o.id_fuente) : D.comunidadesDet.find(x => x.id_comunidad === ed.o.id_comunidad)) || ed.o;
        pintar(); ok(th('fuentes.guardada')); }).autorizacionGuardar(p);
    }
    // ── Personas e instituciones
    function vActores() {
      let h = '<p style="margin-top:0"><small>' + th('actores.intro') + '</small></p><div class="sec"><table><tr><th>' + th('actores.col.nombre') + '</th><th>' + th('actores.col.tipo') + '</th><th>' + th('actores.col.autoriza') + '</th><th></th></tr>' +
        D.actores.map((a, i) => { const n = D.autorizaciones.filter(x => x.id_actor === a.id_actor);
          return '<tr><td><b>' + esc(a.nombre) + '</b>' + (a.credito ? '<br><small>' + esc(a.credito) + '</small>' : '') + '</td><td><small>' + esc(a.tipo) + '</small></td><td><small>' +
            n.map(x => esc(x.id_fuente ? nombreFuente(x.id_fuente) : nombreCom(x.id_comunidad)) + ' (' + esc(x.rol) + ')').join('<br>') + '</small></td><td><button class="peq" onclick="editarAc(' + i + ')">' + th('voces.editar') + '</button></td></tr>'; }).join('') +
        '</table><button class="prim" onclick="editarAc(-1)">' + th('actores.nuevo') + '</button></div>';
      if (ed && ed.tipo === 'actor') { const a = ed.o;
        h += '<div class="sec" id="fAc"><h3>' + esc(a.nombre || tt('actores.nuevo')) + '</h3><div class="dos">' + campo('ac_nombre', th('actores.f.nombre'), txt('ac_nombre', a.nombre)) + campo('ac_tipo', th('actores.f.tipo'), sel('ac_tipo', D.listas.tiposActor, a.tipo || 'Persona')) + '</div>' +
          campo('ac_credito', th('actores.f.credito') + ' <small>' + th('actores.f.credito.d') + '</small>', txt('ac_credito', a.credito)) + '<div class="dos">' +
          campo('ac_contacto', th('actores.f.contacto') + ' <small>' + th('actores.f.contacto.d') + '</small>', txt('ac_contacto', a.contacto)) + campo('ac_enlace', th('fuentes.f.enlace'), txt('ac_enlace', a.enlace)) + '</div>' +
          campo('ac_id_comunidad', th('actores.f.comunidad'), sel('ac_id_comunidad', D.comunidades.map(c => ({v: c.id, t: c.nombre})), a.id_comunidad, true)) + campo('ac_notas', th('fuentes.f.notas'), area('ac_notas', a.notas)) +
          '<button class="prim" onclick="guardarAc()">' + th('voces.guardar') + '</button><button onclick="ed=null;vActores()">' + th('voces.cancelar') + '</button></div>'; }
      $('vista').innerHTML = h;
    }
    function editarAc(i) { ed = { tipo: 'actor', o: i < 0 ? {} : D.actores[i] }; vActores(); $('fAc').scrollIntoView(); }
    function guardarAc() {
      const p = { id_actor: ed.o.id_actor || '', estado: ed.o.estado || '' };
      ['nombre','tipo','credito','contacto','enlace','id_comunidad','notas'].forEach(c => { p[c] = val('ac_' + c); });
      google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; ed = null; vActores(); ok(th('fuentes.guardada')); }).actorGuardar(p);
    }
    // ── Comunidades
    function vComunidades() {
      let h = '<p style="margin-top:0"><small>' + th('comunidades.intro') + '</small></p><div class="sec"><table><tr><th>' + th('comunidades.col.nombre') + '</th><th>' + th('comunidades.col.variante') + '</th><th>' + th('comunidades.col.consulta') + '</th>' +
        (D.solicitarIA ? '<th>' + th('comunidades.col.acuerdo') + '</th>' : '') + '<th></th></tr>' +
        D.comunidadesDet.map((c, i) => '<tr><td><b>' + esc(c.nombre_propio) + '</b>' + (c.nombre_otro ? '<br><small>' + esc(c.nombre_otro) + '</small>' : '') + '</td><td><small>' + esc([c.variante, c.region].filter(Boolean).join(' · ')) + '</small></td><td><small>' + esc(c.consulta || '—') + '</small></td>' +
          (D.solicitarIA ? '<td>' + chipP(c.colectivo) + '</td>' : '') + '<td><button class="peq" onclick="editarC(' + i + ')">' + th('voces.editar') + '</button></td></tr>').join('') +
        '</table><button class="prim" onclick="editarC(-1)">' + th('comunidades.nueva') + '</button></div>';
      if (ed && ed.tipo === 'comunidad') { const c = ed.o, auts = D.autorizaciones.filter(a => c.id_comunidad && a.id_comunidad === c.id_comunidad);
        h += '<div class="sec" id="fC"><h3>' + esc(c.nombre_propio || tt('comunidades.nueva')) + '</h3><div class="dos">' + campo('c_nombre_propio', th('comunidades.f.propio'), txt('c_nombre_propio', c.nombre_propio)) + campo('c_nombre_otro', th('comunidades.f.otro'), txt('c_nombre_otro', c.nombre_otro)) + '</div><div class="dos">' +
          campo('c_id_idioma', th('voces.f.lengua'), sel('c_id_idioma', D.idiomas.map(x => ({v: x.id, t: x.nombre})), c.id_idioma, true)) + campo('c_variante', th('voces.f.variante'), txt('c_variante', c.variante)) + campo('c_region', th('comunidades.f.region'), txt('c_region', c.region)) + '</div>' +
          campo('c_sistema_escritura', th('voces.f.escritura'), txt('c_sistema_escritura', c.sistema_escritura)) +
          campo('c_consulta', th('comunidades.f.consulta') + ' <small>' + th('comunidades.f.consulta.d') + '</small>', txt('c_consulta', c.consulta)) + campo('c_notas', th('fuentes.f.notas'), area('c_notas', c.notas)) +
          '<label style="font-weight:normal"><input type="checkbox" id="c_en_disputa"' + (c.en_disputa === 'Sí' ? ' checked' : '') + '> ' + th('comunidades.f.disputa') + '</label>' +
          '<button class="prim" onclick="guardarC()">' + th('voces.guardar') + '</button><button onclick="ed=null;vComunidades()">' + th('voces.cancelar') + '</button>' +
          (c.id_comunidad ? '<h3 style="margin-top:14px">' + th('comunidades.acuerdos') + '</h3><p style="margin:0"><small>' + th('comunidades.acuerdos.d') + '</small></p>' + tablaAuts(auts) +
            '<button onclick="editarA(null,\\'\\',\\'' + esc(c.id_comunidad) + '\\')">' + th('comunidades.acuerdoNuevo') + '</button>' : '') + (ed.a ? formAut(ed.a) : '') + '</div>'; }
      $('vista').innerHTML = h;
    }
    function editarC(i) { ed = { tipo: 'comunidad', o: i < 0 ? {} : D.comunidadesDet[i] }; vComunidades(); $('fC').scrollIntoView(); }
    function guardarC() {
      const p = { id_comunidad: ed.o.id_comunidad || '', estado: ed.o.estado || '', en_disputa: $('c_en_disputa').checked ? 'Sí' : 'No' };
      ['nombre_propio','nombre_otro','id_idioma','variante','region','sistema_escritura','consulta','notas'].forEach(c => { p[c] = val('c_' + c); });
      google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; const c = D.comunidadesDet.find(x => x.nombre_propio === (p.nombre_propio || p.nombre_otro)); ed = c ? { tipo: 'comunidad', o: c } : null; vComunidades(); ok(th('fuentes.guardada')); }).comunidadGuardar(p);
    }
    // ── Asignar en bloque
    function vBloque() {
      $('vista').innerHTML = '<p style="margin-top:0"><small>' + th('bloque.intro') + '</small></p><div class="sec"><h3>' + th('bloque.cuales') + '</h3><div class="dos">' +
        campo('b_lote', th('bloque.lote'), sel('b_lote', D.lotes.map(l => ({v: l.lote, t: l.lote + ' (' + l.n + ')'})), '', true)) +
        campo('b_fuente', th('bloque.fuenteActual'), sel('b_fuente', [{v: '__sin', t: tt('bloque.sinFuente')}].concat(D.fuentes.map(f => ({v: f.id_fuente, t: f.titulo}))), '', true)) + '</div><div class="dos">' +
        campo('b_idioma', th('voces.f.lengua'), sel('b_idioma', D.idiomas.map(x => ({v: x.id, t: x.nombre})), '', true)) +
        campo('b_categoria', th('bloque.categoria'), sel('b_categoria', D.categorias.map(c => ({v: c.id, t: c.nombre})), '', true)) +
        campo('b_creado', th('bloque.creadoPor'), txt('b_creado', '')) + '</div>' +
        '<h3 style="margin-top:10px">' + th('bloque.asignar') + '</h3><div class="dos">' + campo('b_destino', th('bloque.destino'), sel('b_destino', D.fuentes.map(f => ({v: f.id_fuente, t: f.titulo})), '', true)) +
        campo('b_ubic', th('bloque.ubicacion') + ' <small>' + th('comun.opcional') + '</small>', txt('b_ubic', '')) + '</div>' +
        '<button onclick="contar()">' + th('bloque.contar') + '</button><button class="prim" id="bAplicar" disabled onclick="aplicarB()">' + th('bloque.aplicar') + '</button><div id="bRes"></div></div>';
      ['b_lote','b_fuente','b_idioma','b_categoria','b_creado','b_destino'].forEach(i => $(i).addEventListener('input', () => { $('bAplicar').disabled = true; $('bRes').innerHTML = ''; }));
    }
    const filtroB = () => ({ lote: val('b_lote'), fuente: val('b_fuente'), idioma: val('b_idioma'), categoria: val('b_categoria'), creado_por: val('b_creado').trim() });
    function contar() { google.script.run.withFailureHandler(err).withSuccessHandler(r => { $('bRes').innerHTML = '<p>' + th('bloque.cuenta', {n: r.n}) + '</p>'; $('bAplicar').disabled = !r.n || !val('b_destino'); }).asignarFuenteBloque(filtroB(), '', '', false); }
    function aplicarB() { $('bAplicar').disabled = true; google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; vBloque(); ok(th('bloque.listo', {n: d.asignadas})); }).asignarFuenteBloque(filtroB(), val('b_destino'), val('b_ubic'), true); }
    // ── Corpus, interruptor y términos
    function vCorpus() {
      const c = D.corpus;
      let h = '<div class="sec"><h3>' + th('ia.titulo') + '</h3><p style="margin:0 0 6px"><small>' + th('ia.d') + '</small></p>' +
        '<label style="font-weight:normal"><input type="checkbox" id="iaSwitch"' + (D.solicitarIA ? ' checked' : '') + (D.soyAdmin ? '' : ' disabled') + ' onchange="interruptor(this.checked)"> ' + th('ia.solicitar') + '</label></div>';
      h += '<div class="sec"><h3>' + th('fuentes.corpus') + '</h3><p style="margin:0 0 6px"><small>' + th('fuentes.corpus.d') + '</small></p>' +
        (c.lenguas.length ? '<table><tr><th>' + th('fuentes.corpus.lengua') + '</th><th>' + th('fuentes.corpus.palabras') + '</th><th>' + th('fuentes.corpus.frases') + '</th><th>' + th('fuentes.corpus.oraciones') +
          '</th><th>' + th('fuentes.corpus.conFuente') + '</th>' + (D.solicitarIA ? '<th>' + th('fuentes.corpus.aptas') + '</th>' : '') + '</tr>' + c.lenguas.map(l => '<tr><td>' + esc(l.lengua) + '</td><td>' + l.Palabra + '</td><td>' + l.Frase + '</td><td>' + l['Oración'] +
          '</td><td>' + l.conFuente + '</td>' + (D.solicitarIA ? '<td>' + l.aptas + '</td>' : '') + '</tr>').join('') + '</table>' : '') +
        (c.pares.length ? '<p><b>' + th('fuentes.corpus.pares') + '</b> ' + c.pares.map(p => esc(p.par) + ': ' + p.n).join(' · ') + '</p>' : '') + '</div>';
      if (D.solicitarIA) {
        const P = c.pendientes;
        h += '<div class="sec"><h3>' + th('ia.pendientes') + '</h3>' +
          '<p><b>' + th('ia.pend.comunidades', {n: P.comunidades.length}) + '</b><br><small>' + esc(P.comunidades.join(' · ') || '—') + '</small></p>' +
          '<p><b>' + th('ia.pend.fuentes', {n: P.fuentes.length}) + '</b><br><small>' + esc(P.fuentes.join(' · ') || '—') + '</small></p>' +
          '<p><b>' + th('ia.pend.voces', {n: P.voces.length}) + '</b><br><small>' + esc(P.voces.join(' · ') || '—') + '</small></p>' +
          '<p><small>' + th('ia.motivos', {revision: c.motivos.revision || 0, fuente: c.motivos.fuente || 0, voz: c.motivos.voz || 0, comunidad: c.motivos.comunidad || 0, sinComunidad: c.motivos.sinComunidad || 0}) + '</small></p></div>';
      }
      h += '<div class="sec"><h3>' + th('terminos.vista') + '</h3><div class="dos"><div><label for="tCom">' + th('terminos.vistaComunidad') + '</label>' + sel('tCom', D.comunidades.map(x => ({v: x.id, t: x.nombre})), '', true) + '</div></div><div id="terminos"></div></div>';
      $('vista').innerHTML = h;
      $('tCom').onchange = verTerminos; verTerminos();
    }
    function verTerminos() {
      google.script.run.withFailureHandler(err).withSuccessHandler(t => { $('terminos').innerHTML = '<h3>' + esc(t.titulo) + '</h3><ol>' + t.puntos.map(p => '<li><b>' + esc(p.titulo) + '</b> ' + esc(p.texto) + '</li>').join('') + '</ol><p><b>' + esc(t.acepto) + '</b></p><p><small>' + th('terminos.version', {v: t.version}) + '</small></p>'; }).terminosTexto(val('tCom'));
    }
    function interruptor(v) { google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; pintar(); ok(th(v ? 'ia.activado' : 'ia.desactivado')); }).solicitarIAGuardar(v); }
    google.script.run.withFailureHandler(err).withSuccessHandler(d => { D = d; tab('fuentes'); }).permisosListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(900).setHeight(720), tx('fuentes.titulo'));
}
// ─────────────────────────────────────────────────────────────────────────────
// 19. LIMPIAR DATOS Y EXPORTAR POR LENGUA
// ─────────────────────────────────────────────────────────────────────────────
// Limpiar: siempre con vista previa, palabra de confirmación y un RESPALDO con todo lo que se borra
// (una hoja de cálculo en la carpeta RESPALDOS). Si el respaldo falla, no se borra nada.
// No se borran archivos de Drive (audios, imágenes): solo filas de las hojas.

const HOJAS_LIMPIABLES = () => { const S = CONFIG.SHEETS; return [S.TRADUCCIONES, S.PALABRAS, S.CATEGORIAS, S.TRAD_CATS, S.RELACIONES, S.IDIOMAS,
  S.FUENTES, S.ACTORES, S.AUTORIZACIONES, S.COMUNIDADES, S.VOCES, S.REGLAS, S.REPORT, S.MEDIOS, S.IMAGENES, S.PROPUESTAS]; };

function limpiezaOpciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const p = permisosListar();
  const idiomas = listarIdiomas().map(i => ({ id: i.id, nombre: (i.nombre_propio || i.nombre_completo) + (i.codigo_iso ? ' (' + i.codigo_iso + ')' : ''), textos: i.textos }));
  return { lotes: p.lotes, fuentes: p.fuentes.map(f => ({ id: f.id_fuente, nombre: f.titulo })), comunidades: p.comunidades, idiomas,
    hojas: HOJAS_LIMPIABLES().filter(h => ss.getSheetByName(h)), palabra: tx('limpiar.palabraConfirmar'),
    copia: _esCopia(ss), titulo: getAppConfig(ss).titulo || ss.getName(), archivo: ss.getName(), respaldos: _listaRespaldos(ss) };
}

/** Calcula qué filas se borrarían: { hoja: Set(índices de fila) }. */
function _planLimpieza(ss, op) {
  const S = CONFIG.SHEETS;
  const T = {};
  [S.TRADUCCIONES, S.PALABRAS, S.CATEGORIAS, S.TRAD_CATS, S.RELACIONES, S.IDIOMAS, S.FUENTES, S.AUTORIZACIONES].forEach(h => { T[h] = _tabla(ss, h); });
  const plan = {};
  const marcar = (h, i) => { (plan[h] = plan[h] || new Set()).add(i); };
  const filas = (h, cond) => { const t = T[h] || _tabla(ss, h); if (t) t.filas.forEach((f, i) => { if (cond(t, f)) marcar(h, i); }); };
  const tTr = T[S.TRADUCCIONES], tP = T[S.PALABRAS];
  const v = x => _k(x);
  // Nunca se adivina: si falta qué lote, fuente o lengua, no se borra nada
  if ((op.tipo === 'lote' && !_s(op.lote)) || (op.tipo === 'fuente' && !_s(op.id_fuente)) || (op.tipo === 'lengua' && !_s(op.id_idioma))) throw new Error(tx('limpiar.error.elige'));
  switch (op.tipo) {
    case 'lote': filas(S.TRADUCCIONES, (t, f) => _s(_get(t, f, 'lote')) === _s(op.lote)); break;
    case 'fuente':
      filas(S.TRADUCCIONES, (t, f) => v(_get(t, f, 'id_fuente')) === v(op.id_fuente));
      if (op.incluirFuente && v(op.id_fuente) !== FUENTE_EQUIPO) {
        filas(S.FUENTES, (t, f) => v(_get(t, f, 'id_fuente')) === v(op.id_fuente));
        filas(S.AUTORIZACIONES, (t, f) => v(_get(t, f, 'id_fuente')) === v(op.id_fuente));
      }
      break;
    case 'lengua':
      filas(S.TRADUCCIONES, (t, f) => v(_get(t, f, 'id_idioma')) === v(op.id_idioma));
      filas(S.TRAD_CATS, (t, f) => v(_get(t, f, 'id_idioma')) === v(op.id_idioma));
      if (op.incluirLengua) filas(S.IDIOMAS, (t, f) => v(_get(t, f, 'id_idioma')) === v(op.id_idioma));
      break;
    case 'huerfanas': {
      const con = new Set(tTr ? tTr.filas.map(f => v(_get(tTr, f, 'id_palabra'))) : []);
      filas(S.PALABRAS, (t, f) => _s(_get(t, f, 'id_palabra')) && !con.has(v(_get(t, f, 'id_palabra'))));
      break;
    }
    case 'categoriasVacias': {
      const usadas = new Set(tP ? tP.filas.map(f => v(_get(tP, f, 'id_categoria'))) : []);
      filas(S.CATEGORIAS, (t, f) => _s(_get(t, f, 'id_categoria')) && !usadas.has(v(_get(t, f, 'id_categoria'))));
      break;
    }
    case 'autorizaciones':
      filas(S.AUTORIZACIONES, (t, f) => op.id_fuente ? v(_get(t, f, 'id_fuente')) === v(op.id_fuente) : op.id_comunidad ? v(_get(t, f, 'id_comunidad')) === v(op.id_comunidad) : true);
      break;
    case 'contenido': {
      // Vaciar el contenido: la identidad, las lenguas, el equipo y la configuración se quedan
      [S.TRADUCCIONES, S.PALABRAS, S.CATEGORIAS, S.TRAD_CATS, S.RELACIONES, S.REPORT, S.MEDIOS, S.IMAGENES, S.PROPUESTAS].forEach(h => filas(h, () => true));
      if (op.fuentes) {
        filas(S.FUENTES, (t, f) => v(_get(t, f, 'id_fuente')) !== FUENTE_EQUIPO);
        [S.ACTORES, S.AUTORIZACIONES].forEach(h => filas(h, () => true));
      }
      if (op.voces) filas(S.VOCES, () => true);
      if (op.comunidades) filas(S.COMUNIDADES, () => true);
      break;
    }
    case 'enBlanco': {
      const conservar = new Set([S.REGLAS].filter(() => op.conservarAtajos).concat(op.conservarLenguas ? [S.IDIOMAS, S.COMUNIDADES] : []));
      HOJAS_LIMPIABLES().filter(h => !conservar.has(h)).forEach(h => filas(h, (t, f) => !(h === S.FUENTES && v(_get(t, f, 'id_fuente')) === FUENTE_EQUIPO)));
      if (!op.conservarEquipo) { const yo = _correoActual(); filas(S.EQUIPO, (t, f) => v(_get(t, f, 'correo')) !== yo); }
      filas(S.HISTORIAL, () => true);
      if (ss.getSheetByName(S.AUDIT)) filas(S.AUDIT, () => true);
      [S.AVISOS, S.COPIAS].forEach(h => { if (ss.getSheetByName(h)) filas(h, () => true); });   // avisos y copias eran del catálogo anterior
      break;
    }
    case 'hoja':
      if (!_s(op.hoja)) throw new Error(tx('limpiar.error.elige'));
      if (!HOJAS_LIMPIABLES().includes(op.hoja)) throw new Error(tx('limpiar.error.hoja'));
      filas(op.hoja, (t, f) => !(op.hoja === S.FUENTES && v(_get(t, f, 'id_fuente')) === FUENTE_EQUIPO));
      break;
    default: throw new Error(tx('limpiar.error.tipo'));
  }
  // Consecuencias: palabras que se quedan sin traducciones, categorías vacías, nombres de categoría y relaciones sueltas
  const trBorradas = plan[S.TRADUCCIONES] || new Set();
  if (tTr && tP && (op.huerfanas || op.tipo === 'hoja' && op.hoja === S.TRADUCCIONES) && trBorradas.size) {
    const quedan = new Set(), tocadas = new Set();
    tTr.filas.forEach((f, i) => { (trBorradas.has(i) ? tocadas : quedan).add(v(_get(tTr, f, 'id_palabra'))); });
    if (op.huerfanas) tP.filas.forEach((f, i) => { const id = v(_get(tP, f, 'id_palabra')); if (tocadas.has(id) && !quedan.has(id)) marcar(S.PALABRAS, i); });
  }
  const pBorradas = plan[S.PALABRAS] || new Set();
  if (op.vacias && tP && pBorradas.size && T[S.CATEGORIAS]) {
    const quedan = new Set(), tocadas = new Set();
    tP.filas.forEach((f, i) => { (pBorradas.has(i) ? tocadas : quedan).add(v(_get(tP, f, 'id_categoria'))); });
    T[S.CATEGORIAS].filas.forEach((f, i) => { const id = v(_get(T[S.CATEGORIAS], f, 'id_categoria')); if (id && tocadas.has(id) && !quedan.has(id)) marcar(S.CATEGORIAS, i); });
  }
  const cBorradas = new Set([...(plan[S.CATEGORIAS] || [])].map(i => v(_get(T[S.CATEGORIAS], T[S.CATEGORIAS].filas[i], 'id_categoria'))));
  if (cBorradas.size) filas(S.TRAD_CATS, (t, f) => cBorradas.has(v(_get(t, f, 'id_categoria'))));
  const idsP = new Set([...pBorradas].map(i => v(_get(tP, tP.filas[i], 'id_palabra'))));
  const idsT = new Set([...trBorradas].map(i => v(_get(tTr, tTr.filas[i], 'id_traduccion'))));
  if (idsP.size || idsT.size) filas(S.RELACIONES, (t, f) => [_get(t, f, 'id_origen'), _get(t, f, 'id_destino')].some(x => idsP.has(v(x)) || idsT.has(v(x))));
  if (idsP.size && ss.getSheetByName(S.IMAGENES)) filas(S.IMAGENES, (t, f) => idsP.has(v(_get(t, f, 'id_palabra'))));
  if (pBorradas.size && tTr) filas(S.TRADUCCIONES, (t, f) => idsP.has(v(_get(t, f, 'id_palabra')))); // nunca quedan traducciones sin palabra
  return plan;
}

const _resumenPlan = plan => Object.keys(plan).filter(h => plan[h].size).map(h => ({ hoja: h, n: plan[h].size }));

function limpiezaVista(op) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const plan = _planLimpieza(ss, op || {});
  const r = _resumenPlan(plan);
  return { hojas: r, total: r.reduce((a, x) => a + x.n, 0) };
}

/**
 * Google no deja borrar todas las filas que no están inmovilizadas ("No se pueden eliminar todas las filas…").
 * Pasa cuando la hoja no tiene filas vacías al final y se van a borrar todas las de datos:
 * antes de borrar se agregan filas vacías al final para que siempre quede al menos una.
 */
function _espacioParaBorrar(sh, n) {
  const libres = sh.getMaxRows() - n - sh.getFrozenRows();
  if (libres < 2) sh.insertRowsAfter(sh.getMaxRows(), 2 - libres);
}

function limpiezaAplicar(op, total, confirmacion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (_k(confirmacion) !== _k(tx('limpiar.palabraConfirmar'))) throw new Error(tx('limpiar.error.confirmar', { palabra: tx('limpiar.palabraConfirmar') }));
  // «Preparar una copia en blanco» no se hace en la hoja original; si no se sabe, exige escribir su nombre
  if (op && op.tipo === 'enBlanco' && _esCopia(ss) === false) throw new Error(tx('blanco.esOriginal', { nombre: getAppConfig(ss).titulo || ss.getName() }));
  if (op && op.tipo === 'enBlanco' && _esCopia(ss) !== true && _k(op.nombreOriginal) !== _k(getAppConfig(ss).titulo || ss.getName()))
    throw new Error(tx('blanco.error.original', { nombre: getAppConfig(ss).titulo || ss.getName() }));
  const lock = _lock();
  try {
    const plan = _planLimpieza(ss, op || {});
    const r = _resumenPlan(plan), n = r.reduce((a, x) => a + x.n, 0);
    if (n !== Number(total)) throw new Error(tx('limpiar.error.cambio'));
    if (!n) return { borradas: 0 };
    // 1. Respaldo (si falla, no se borra nada)
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const libro = SpreadsheetApp.create(tx('limpiar.respaldo.archivo', { fecha, catalogo: getAppConfig(ss).titulo || ss.getName() }));
    const primera = libro.getSheets()[0];
    primera.setName(tx('limpiar.respaldo.leeme'));
    primera.getRange(1, 1, 3, 1).setValues([[tx('limpiar.respaldo.t1', { fecha, quien: _correoActual() })], [tx('limpiar.respaldo.t2', { op: JSON.stringify(op) })], [tx('limpiar.respaldo.t3')]]);
    r.forEach(x => {
      const t = _tabla(ss, x.hoja);
      const datos = [t.sh.getRange(1, 1, 1, t.headers.length).getValues()[0]].concat([...plan[x.hoja]].sort((a, b) => a - b).map(i => t.filas[i]));
      libro.insertSheet(x.hoja).getRange(1, 1, datos.length, datos[0].length).setNumberFormat('@').setValues(datos);
    });
    // Al dejar en blanco también cambia la configuración (código, clave, carpetas, app): se guarda como estaba
    if (op.tipo === 'enBlanco') {
      const c = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
      if (c) { const d = c.getDataRange().getValues(); libro.insertSheet(CONFIG.SHEETS.CONFIG).getRange(1, 1, d.length, d[0].length).setNumberFormat('@').setValues(d); }
      const ti = _tabla(ss, CONFIG.SHEETS.IDIOMAS);   // las carpetas de audio de cada lengua también se olvidan
      if (ti && ti.filas.length) libro.insertSheet(HOJA_RESP_CARPETAS).getRange(1, 1, ti.filas.length + 1, 3).setNumberFormat('@')
        .setValues([['id_idioma', 'carpeta_audios', 'carpeta_videos']].concat(ti.filas.map(f => ['id_idioma', 'carpeta_audios', 'carpeta_videos'].map(c => _s(_get(ti, f, c))))));
    }
    DriveApp.getFileById(libro.getId()).moveTo(_subcarpeta(_carpetaDelLibro(ss), 'RESPALDOS'));
    // 2. Borrado por bloques contiguos, de abajo hacia arriba (no mueve las demás filas)
    const antes = {};
    r.forEach(x => {
      const sh = ss.getSheetByName(x.hoja);
      antes[x.hoja] = Math.max(0, sh.getLastRow() - 1);
      _espacioParaBorrar(sh, plan[x.hoja].size);
      const idx = [...plan[x.hoja]].sort((a, b) => b - a);
      let k = 0;
      while (k < idx.length) {
        let fin = idx[k], ini = fin;
        while (k + 1 < idx.length && idx[k + 1] === ini - 1) { k++; ini = idx[k]; }
        sh.deleteRows(ini + 2, fin - ini + 1);
        k++;
      }
    });
    // 3. Comprobar en la hoja que de verdad se borró (no basta con que no haya error)
    SpreadsheetApp.flush();
    const quedan = r.map(x => ({ hoja: x.hoja, esperadas: antes[x.hoja] - x.n, hay: Math.max(0, ss.getSheetByName(x.hoja).getLastRow() - 1) }))
      .filter(x => x.hay !== x.esperadas);
    if (quedan.length) throw new Error(tx('limpiar.error.noSeBorro', { hojas: quedan.map(x => x.hoja + ' (' + x.hay + ')').join(', '), respaldo: libro.getUrl() }));
    if (op.tipo === 'enBlanco') _identidadNueva(ss, op);
    _historial(ss, tx('limpiar.historialHoja'), JSON.stringify(op).slice(0, 200), { filas: '' }, { filas: r.map(x => x.hoja + ': ' + x.n).join(' · ') });
    _cfgMemo = null;
    try { publicarCatalogo(ss, true); } catch (e) { /* se publicará en la próxima hora */ }
    return { borradas: n, hojas: r, respaldo: libro.getUrl(), archivo: ss.getName(), enBlanco: op.tipo === 'enBlanco', titulo: getAppConfig(ss).titulo || ss.getName() };
  } finally { lock.releaseLock(); }
}

// ── Recuperar desde un respaldo ─────────────────────────────────────────────
// Solo AGREGA las filas del respaldo que ya no están (por su código); nunca borra ni reemplaza.
// La configuración (código, clave, carpetas, app) solo se recupera si se pide expresamente.

const HOJA_RESP_CARPETAS = 'CARPETAS_DE_LENGUAS';
const CONFIG_IDENTIDAD = ['ID_CATALOGO', 'API_TOKEN', 'URL_API', 'CLAVE_EQUIPO', 'APPSHEET_APP_ID', 'ID_CARPETA_RAIZ', 'ID_CARPETA_AUDIOS', 'ID_CARPETA_IMAGENES', 'ID_CARPETA_VIDEOS', 'ID_CARPETA_CATEGORIAS', 'NOMBRE_PROYECTO'];

function _listaRespaldos(ss) {
  const l = [];
  try {
    const it = _subcarpeta(_carpetaDelLibro(ss), 'RESPALDOS').getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext() && l.length < 300) { const f = it.next(); l.push({ id: f.getId(), nombre: f.getName(), t: f.getDateCreated().getTime() }); }
  } catch (e) { /* sin carpeta todavía */ }
  return l.sort((a, b) => b.t - a.t).slice(0, 40)
    .map(x => ({ id: x.id, nombre: x.nombre, fecha: Utilities.formatDate(new Date(x.t), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') }));
}

const _valorRestaurado = v => { const x = _s(v); return x === 'TRUE' || x === 'true' ? true : x === 'FALSE' || x === 'false' ? false : v; };

function _planRestaurar(ss, idRespaldo) {
  const S = CONFIG.SHEETS;
  if (!_listaRespaldos(ss).some(r => r.id === idRespaldo)) throw new Error(tx('limpiar.r.noEs'));
  const libro = SpreadsheetApp.openById(idRespaldo);
  const orden = _hojasRama().concat([S.REPORT, S.EQUIPO]).filter((h, i, a) => a.indexOf(h) === i);
  const hojas = []; let yaEstan = 0;
  orden.forEach(h => {
    const sh = libro.getSheetByName(h), t = _tabla(ss, h);
    if (!sh || !t || sh.getLastRow() < 2) return;
    const d = sh.getDataRange().getValues();
    const headers = d[0].map(x => _k(x)), idx = {};
    headers.forEach((x, i) => { if (x && !(x in idx)) idx[x] = i; });
    const r = { headers, idx };
    const clave = h === S.EQUIPO ? (tt, f) => _k(_get(tt, f, 'correo')) : (tt, f) => _claveFila(h, tt, f);
    const existen = new Set(t.filas.map(f => clave(t, f)).filter(Boolean));
    const filas = [];
    d.slice(1).forEach(f => {
      if (!f.some(x => _s(x))) return;
      const k = clave(r, f);
      if (k && existen.has(k)) { yaEstan++; return; }
      if (k) existen.add(k);
      const o = {}; headers.forEach((x, i) => { if (x) o[x] = _valorRestaurado(f[i]); });
      filas.push(o);
    });
    hojas.push({ hoja: h, t, filas });
  });
  let config = null;
  const c = libro.getSheetByName(S.CONFIG);
  if (c && c.getLastRow() >= 2) {
    const d = c.getDataRange().getValues(), h = d[0].map(x => _k(x)), valores = {};
    CONFIG_IDENTIDAD.forEach(k => { const i = h.indexOf(_k(k)); if (i > -1) valores[k] = _s(d[1][i]); });
    const carp = libro.getSheetByName(HOJA_RESP_CARPETAS);
    config = { valores, titulo: valores.NOMBRE_PROYECTO || '', carpetas: carp ? carp.getDataRange().getValues().slice(1) : [] };
  }
  return { hojas, yaEstan, config };
}

function limpiezaRestaurarVista(idRespaldo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const p = _planRestaurar(ss, idRespaldo);
  const hojas = p.hojas.map(x => ({ hoja: x.hoja, n: x.filas.length }));
  return { hojas, total: hojas.reduce((a, x) => a + x.n, 0), yaEstan: p.yaEstan, config: p.config ? { titulo: p.config.titulo } : null };
}

function limpiezaRestaurar(idRespaldo, conConfig, total) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lock = _lock();
  try {
    const p = _planRestaurar(ss, idRespaldo);
    const n = p.hojas.reduce((a, x) => a + x.filas.length, 0);
    if (n !== Number(total)) throw new Error(tx('limpiar.error.cambio'));
    const hojas = [];
    p.hojas.forEach(x => { if (x.filas.length) { _agregar(x.t, x.filas); hojas.push({ hoja: x.hoja, n: x.filas.length }); } });
    let config = false;
    if (conConfig && p.config) {
      const antes = getAppConfig(ss);
      _escribirConfig(ss, p.config.valores);
      const ti = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
      if (ti) p.config.carpetas.forEach(r => { const i = ti.filas.findIndex(f => _k(_get(ti, f, 'id_idioma')) === _k(r[0])); if (i > -1) _actualizar(ti, i, { carpeta_audios: _s(r[1]), carpeta_videos: _s(r[2]) }); });
      const pr = PropertiesService.getScriptProperties();
      [CONFIG.PROPS.TOKEN, CONFIG.PROPS.MANIFIESTO].forEach(k => { try { pr.deleteProperty(k); } catch (e) {} });
      _historial(ss, tx('limpiar.historialHoja'), 'CONFIG', { id: antes.idCatalogo, titulo: antes.titulo }, { id: p.config.valores.ID_CATALOGO, titulo: p.config.titulo });
      config = true;
    }
    if (hojas.length) _historial(ss, tx('limpiar.historialHoja'), tx('limpiar.r.historial'), { filas: '' }, { filas: hojas.map(x => x.hoja + ': ' + x.n).join(' · ') });
    _cfgMemo = null;
    try { publicarCatalogo(ss, true); } catch (e) { /* se publicará en la próxima hora */ }
    return { agregadas: n, hojas, config };
  } finally { lock.releaseLock(); }
}

// ── Copia en blanco: identidad propia ───────────────────────────────────────

/** Guarda el ID de esta hoja. En una copia, el ID guardado (que viaja con el script) ya no coincide. */
function _registrarHoja(ss) {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('HOJA_ID')) p.setProperty('HOJA_ID', ss.getId());
}
/** true = es una copia; false = es la hoja original; null = no se sabe. */
function _esCopia(ss) {
  const id = PropertiesService.getScriptProperties().getProperty('HOJA_ID');
  return id ? id !== ss.getId() : null;
}
/**
 * Después de dejar en blanco: código de catálogo y clave nuevos; se olvidan las carpetas, la app de AppSheet,
 * el directorio de voces y la publicación heredados (el asistente de Primeros pasos los crea de nuevo).
 */
function _identidadNueva(ss, op) {
  const cfg = getAppConfig(ss);
  _escribirConfig(ss, { ID_CATALOGO: Utilities.getUuid(), API_TOKEN: _nuevaClave(), APPSHEET_APP_ID: '', URL_API: '', CLAVE_EQUIPO: '',
    ID_CARPETA_RAIZ: '', ID_CARPETA_AUDIOS: '', ID_CARPETA_IMAGENES: '', ID_CARPETA_VIDEOS: '', ID_CARPETA_CATEGORIAS: '',
    NOMBRE_PROYECTO: _s(op.nombreNuevo) || cfg.titulo || ss.getName() });
  const t = _tabla(ss, CONFIG.SHEETS.IDIOMAS);
  if (t) t.filas.forEach((f, i) => { if (_s(_get(t, f, 'carpeta_audios')) || _s(_get(t, f, 'carpeta_videos'))) _actualizar(t, i, { carpeta_audios: '', carpeta_videos: '' }); });
  const p = PropertiesService.getScriptProperties();
  [CONFIG.PROPS.SNAPSHOT_FILE, CONFIG.PROPS.MANIFIESTO, CONFIG.PROPS.TOKEN, PROP_DIRECTORIO, 'CARPETA_PRIVADA_ID', 'IMAGENES_ULTIMO_REPORTE'].forEach(k => { try { p.deleteProperty(k); } catch (e) {} });
  p.setProperty('HOJA_ID', ss.getId());
  _cfgMemo = null;
}

// ── Exportar por lengua a una hoja que se puede editar y volver a traer ─────

/**
 * Crea una hoja DATOS_A_IMPORTAR con las lenguas elegidas, en el mismo formato que 📥 Traer lo que llenaron.
 * Traerla al mismo catálogo actualiza (por ID_PALABRA); traerla a otro catálogo agrega.
 */
function exportarHojaLenguas(op) {
  op = op || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lenguas = (op.lenguas || []).map(_k);
  if (!lenguas.length) throw new Error(tx('exportarHoja.error.lenguas'));
  const soloPublicado = op.alcance !== 'todo', ctx = _contextoPermisos(ss);
  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => lenguas.includes(_k(i.id_idioma)));
  const cod = {}; idiomas.forEach(i => { cod[_k(i.id_idioma)] = (_s(i.codigo_iso) || _s(i.id_idioma)).toUpperCase().replace(/[^A-Z0-9_]/g, ''); });
  const cats = {}; readSheet(ss, CONFIG.SHEETS.CATEGORIAS).forEach(c => { cats[_k(c.id_categoria)] = _s(c.nombre_categoria); });
  const tc = {}; readSheet(ss, CONFIG.SHEETS.TRAD_CATS).forEach(x => { tc[_k(x.id_categoria) + '|' + _k(x.id_idioma)] = _s(x.texto || x.nombre); });
  const fuentes = ctx.fuentes;
  const porPal = {}; let omitidas = 0;
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).forEach(t => {
    if (!lenguas.includes(_k(t.id_idioma)) || !_s(t.texto) || _retirado(t.estado)) return;
    if (soloPublicado && !_esPublico(t)) return;
    if (op.soloCompartibles !== false && _usoFuente(ctx, t.id_fuente, 'compartir') !== 'si') { omitidas++; return; }
    (porPal[_k(t.id_palabra)] = porPal[_k(t.id_palabra)] || {})[_k(t.id_idioma)] = t;
  });
  const H = ['ID_PALABRA', 'CATEGORIA'];
  idiomas.forEach(i => { const c = cod[_k(i.id_idioma)]; H.push('TRADUCCION_' + c, 'NOTA_' + c, 'TRAD_CAT_' + c, 'TIPO_' + c, 'ORIGINAL_' + c, 'FUENTE_' + c, 'UBICACION_' + c); });
  const filas = [H];
  readSheet(ss, CONFIG.SHEETS.PALABRAS).forEach(p => {
    const trs = porPal[_k(p.id_palabra)];
    if (!trs || (soloPublicado && !_esPublico(p))) return;
    const fila = [_s(p.id_palabra), cats[_k(p.id_categoria)] || ''];
    idiomas.forEach(i => {
      const t = trs[_k(i.id_idioma)] || {};
      const f = fuentes[_k(t.id_fuente)];
      fila.push(_s(t.texto), _s(t.nota_variante), tc[_k(p.id_categoria) + '|' + _k(i.id_idioma)] || '', _s(t.tipo_registro), _s(t.original),
        f && _k(f.id_fuente) !== FUENTE_EQUIPO ? _s(f.titulo) : '', _s(t.ubicacion_fuente));
    });
    filas.push(fila);
  });
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nombre = _s(op.nombre) || tx('exportarHoja.archivo', { lenguas: idiomas.map(i => _s(i.nombre_propio || i.nombre_completo)).join(', '), fecha });
  const libro = SpreadsheetApp.create(nombre);
  const sh = libro.getSheets()[0].setName('DATOS_A_IMPORTAR');
  sh.getRange(1, 1, filas.length, H.length).setNumberFormat('@').setValues(filas);
  sh.getRange(1, 1, 1, H.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('white');
  sh.setFrozenRows(1);
  // META: nombre de cada lengua, para que otro catálogo la reconozca o la cree con su nombre
  libro.insertSheet('META').getRange(1, 1, idiomas.length + 1, 2).setValues([['CODIGO_ISO', 'NOMBRE']].concat(idiomas.map(i => [cod[_k(i.id_idioma)], _s(i.nombre_propio || i.nombre_completo)])));
  DriveApp.getFileById(libro.getId()).moveTo(_carpetaExportaciones(ss));
  return { url: libro.getUrl(), nombre, palabras: filas.length - 1, omitidas };
}

/** Antes de traer una hoja: qué lenguas trae (por sus columnas TRADUCCION_X). */
function traerVerLenguas(url) {
  const ext = SpreadsheetApp.openByUrl(url);
  const sh = ext.getSheetByName('DATOS_A_IMPORTAR') || ext.getSheets()[0];
  const H = (sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0] || []).map(x => _s(x).toUpperCase());
  const meta = {};
  const m = ext.getSheetByName('META');
  if (m) m.getDataRange().getValues().slice(1).forEach(r => { if (_s(r[0])) meta[_s(r[0]).toUpperCase()] = _s(r[1]); });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const propias = {};
  readSheet(ss, CONFIG.SHEETS.IDIOMAS).forEach(i => { const n = _s(i.nombre_propio || i.nombre_completo); propias[_k(i.id_idioma)] = n;
    if (_s(i.codigo_internacional)) propias[_k(i.codigo_internacional)] = n; if (_s(i.codigo_iso)) propias[_k(i.codigo_iso)] = n; });
  return H.map(h => (h.match(/^TRADUCCION_([A-Z0-9_]+)$/) || [])[1]).filter(Boolean)
    .map(c => ({ codigo: c, nombre: meta[c] || propias[_k(c)] || c, existe: !!propias[_k(c)] }));
}

function uiLimpiar() {
  const tipos = [
    { grupo: 'limpiar.g.comun', items: ['contenido', 'lote', 'fuente', 'lengua'] },
    { grupo: 'limpiar.g.ordenar', items: ['huerfanas', 'categoriasVacias', 'autorizaciones'] },
    { grupo: 'limpiar.g.avanzado', items: ['hoja', 'enBlanco'] },
    { grupo: 'limpiar.g.recuperar', items: ['restaurar'] }
  ];
  const radios = tipos.map(g => `<div class="grupo">${txh(g.grupo)}</div>` + g.items.map(t =>
    `<label class="opcion"><input type="radio" name="tipo" value="${t}"${t === 'contenido' ? ' checked' : ''}><span><b>${txh('limpiar.t.' + t)}</b><small>${txh('limpiar.t.' + t + '.d')}</small></span></label>`).join('')).join('');
  const html = `${_CSS_FICHAS}${_tScript()}
  <style>
    .opcion{display:flex;gap:8px;align-items:flex-start;padding:6px 8px;border:1px solid #e2e8f0;border-radius:8px;margin:4px 0;background:#fff;cursor:pointer;font-weight:normal}
    .opcion small{display:block;color:#64748b}.opcion:has(input:checked){border-color:#475569;background:#f1f5f9}
    .grupo{font-size:12px;font-weight:bold;color:#475569;text-transform:uppercase;margin:10px 0 2px}
    .efecto{border-left:4px solid #475569;background:#f8fafc;padding:8px 10px;border-radius:6px;margin:8px 0}
    .efecto p{margin:4px 0}.efecto .b{color:#b91c1c}.efecto .c{color:#b45309}.efecto .k{color:#166534}
  </style>
  <div class="caja">
    <p style="margin-top:0">${txh('limpiar.intro')}</p>
    <details class="sec"><summary><b>${txh('limpiar.guia')}</b></summary><ul style="margin:6px 0 0 0;padding-left:18px">
      <li>${txh('limpiar.guia.1')}</li><li>${txh('limpiar.guia.2')}</li><li>${txh('limpiar.guia.3')}</li><li>${txh('limpiar.guia.4')}</li></ul></details>
    <div class="sec">
      <h3>${txh('limpiar.que')}</h3>
      ${radios}
    </div>
    <div class="sec" id="paso2">
      <h3 id="tituloSel"></h3>
      <div id="detalle"></div>
      <div id="extras" style="margin-top:6px">
        <label style="font-weight:normal"><input type="checkbox" id="huerfanas" checked> ${txh('limpiar.o.huerfanas')}</label>
        <label style="font-weight:normal"><input type="checkbox" id="vacias" checked> ${txh('limpiar.o.vacias')}</label>
      </div>
      <div class="efecto" id="efecto"></div>
      <button class="prim" id="bVer" onclick="ver()">${txh('limpiar.ver')}</button>
      <div id="vista"></div>
    </div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let O = null, ultimo = null;
    const opt = (l, vacio) => (vacio ? '<option value="">—</option>' : '') + l.map(x => '<option value="' + esc(x.v) + '">' + esc(x.t) + '</option>').join('');
    const tipo = () => document.querySelector('input[name="tipo"]:checked').value;
    const chk = (id, k, on) => '<label style="font-weight:normal"><input type="checkbox" id="' + id + '"' + (on ? ' checked' : '') + '> ' + th(k) + '</label>';
    const errMsg = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    function efecto() {
      const t = tipo(), p = [];
      const linea = (cls, rot, k) => p.push('<p class="' + cls + '"><b>' + th(rot) + '</b> ' + th(k) + '</p>');
      linea('b', 'limpiar.e.borra', 'limpiar.e.' + t + '.borra');
      if (t !== 'restaurar') linea('k', 'limpiar.e.conserva', 'limpiar.e.' + t + '.conserva');
      if (t === 'enBlanco') { linea('c', 'limpiar.e.cambia', 'limpiar.e.enBlanco.cambia'); linea('c', 'limpiar.e.ojo', 'limpiar.e.enBlanco.ojo'); }
      else if (t !== 'restaurar') linea('k', 'limpiar.e.noCambia', 'limpiar.e.noCambia.d');
      p.push('<p><small>' + th(t === 'restaurar' ? 'limpiar.e.restaurar.d' : 'limpiar.e.respaldo') + '</small></p>');
      $('efecto').innerHTML = p.join('');
    }
    function detalle() {
      const t = tipo(); ultimo = null; $('vista').innerHTML = ''; $('msg').innerHTML = '';
      $('tituloSel').innerHTML = th('limpiar.t.' + t);
      efecto();
      if (!O) { $('detalle').innerHTML = '<small>' + th('comun.cargando') + '</small>'; $('bVer').disabled = true; return; }
      $('bVer').disabled = false;
      $('bVer').innerHTML = th(t === 'restaurar' ? 'limpiar.r.ver' : 'limpiar.ver');
      let h = '';
      if (t === 'contenido') h = chk('d_fuentes', 'limpiar.d.cFuentes') + chk('d_voces', 'limpiar.d.cVoces') + chk('d_coms', 'limpiar.d.cComunidades');
      if (t === 'lote') h = O.lotes.length ? '<label for="d_lote">' + th('limpiar.d.lote') + '</label><select id="d_lote">' + opt(O.lotes.map(l => ({v: l.lote, t: tt('limpiar.d.loteOpcion', {lote: l.lote, n: l.n})})), true) + '</select>'
        : '<div class="error">' + th('limpiar.d.sinLotes') + '</div>';
      if (t === 'fuente') h = '<label for="d_fuente">' + th('limpiar.d.fuente') + '</label><select id="d_fuente">' + opt(O.fuentes.map(f => ({v: f.id, t: f.nombre})), true) + '</select>' + chk('d_incluirFuente', 'limpiar.d.incluirFuente');
      if (t === 'lengua') h = '<label for="d_lengua">' + th('limpiar.d.lengua') + '</label><select id="d_lengua">' + opt(O.idiomas.map(i => ({v: i.id, t: i.nombre + ' · ' + i.textos})), true) + '</select>' + chk('d_incluirLengua', 'limpiar.d.incluirLengua');
      if (t === 'autorizaciones') h = '<label for="d_aut">' + th('limpiar.d.aut') + '</label><select id="d_aut"><option value="">' + th('limpiar.d.autTodas') + '</option>' +
        opt(O.fuentes.map(f => ({v: 'f|' + f.id, t: '📚 ' + f.nombre})).concat(O.comunidades.map(c => ({v: 'c|' + c.id, t: '🏘️ ' + c.nombre})))) + '</select>';
      if (t === 'hoja') h = '<label for="d_hoja">' + th('limpiar.d.hoja') + '</label><select id="d_hoja">' + opt(O.hojas.map(x => ({v: x, t: x})), true) + '</select><div class="error">' + th('limpiar.d.hojaAviso') + '</div>';
      if (t === 'enBlanco') {
        if (O.copia === false) { h = '<div class="error">' + th('blanco.esOriginal', {nombre: O.titulo}) + '</div>'; $('bVer').disabled = true; }
        else h = (O.copia === true ? '<div class="ok">' + th('blanco.esCopia') + '</div>' : '<div class="error">' + th('blanco.noSe', {nombre: O.titulo}) + '</div>' +
            '<label for="d_original">' + th('blanco.escribeNombre', {nombre: O.titulo}) + '</label><input type="text" id="d_original" autocomplete="off">') +
          '<label for="d_nombreNuevo">' + th('blanco.nombreNuevo') + '</label><input type="text" id="d_nombreNuevo" placeholder="' + esc(O.titulo) + '">' +
          chk('d_lenguas', 'blanco.conservarLenguas', true) + chk('d_atajos', 'blanco.conservarAtajos', true) + chk('d_equipo', 'blanco.conservarEquipo');
      }
      if (t === 'restaurar') h = O.respaldos.length ? '<label for="d_resp">' + th('limpiar.r.cual') + '</label><select id="d_resp">' + opt(O.respaldos.map(r => ({v: r.id, t: r.fecha + ' · ' + r.nombre})), true) + '</select>'
        : '<div class="ok">' + th('limpiar.r.ninguno') + '</div>';
      $('detalle').innerHTML = h;
      $('extras').style.display = ['lote','fuente','lengua'].includes(t) ? '' : 'none';
      document.querySelectorAll('#detalle select, #detalle input, #extras input').forEach(e => e.onchange = () => { ultimo = null; $('vista').innerHTML = ''; });
    }
    const val = id => $(id) ? $(id).value : '', on = id => !!($(id) && $(id).checked);
    function op() {
      const t = tipo(), o = { tipo: t, huerfanas: ['lote','fuente','lengua'].includes(t) && on('huerfanas'), vacias: ['lote','fuente','lengua'].includes(t) && on('vacias') };
      if (t === 'contenido') { o.fuentes = on('d_fuentes'); o.voces = on('d_voces'); o.comunidades = on('d_coms'); }
      if (t === 'lote') o.lote = val('d_lote');
      if (t === 'fuente') { o.id_fuente = val('d_fuente'); o.incluirFuente = on('d_incluirFuente'); }
      if (t === 'lengua') { o.id_idioma = val('d_lengua'); o.incluirLengua = on('d_incluirLengua'); }
      if (t === 'autorizaciones') { const a = val('d_aut'); if (a.startsWith('f|')) o.id_fuente = a.slice(2); if (a.startsWith('c|')) o.id_comunidad = a.slice(2); }
      if (t === 'hoja') o.hoja = val('d_hoja');
      if (t === 'enBlanco') { o.conservarLenguas = on('d_lenguas'); o.conservarAtajos = on('d_atajos'); o.conservarEquipo = on('d_equipo'); o.nombreNuevo = val('d_nombreNuevo'); if ($('d_original')) o.nombreOriginal = val('d_original'); }
      if (t === 'restaurar') o.respaldo = val('d_resp');
      return o;
    }
    function ver() {
      if (!O) return;
      const o = op(); $('msg').innerHTML = '';
      if (o.tipo === 'restaurar') return verRestaurar(o);
      $('vista').innerHTML = '<p><small>' + th('comun.cargando') + '</small></p>';
      google.script.run.withFailureHandler(e => { $('vista').innerHTML = ''; errMsg(e); }).withSuccessHandler(r => {
        ultimo = { op: o, total: r.total };
        if (!r.total) { $('vista').innerHTML = '<div class="ok">' + th('limpiar.nada') + '</div>'; return; }
        $('vista').innerHTML = '<div class="error"><b>' + th('limpiar.seBorrarian', {n: r.total, archivo: O.archivo}) + '</b><ul>' + r.hojas.map(x => '<li>' + esc(x.hoja) + ': ' + x.n + '</li>').join('') + '</ul>' +
          th('limpiar.respaldoAviso') + '</div><label for="conf">' + th('limpiar.escribe', {palabra: O.palabra}) + '</label><input type="text" id="conf" autocomplete="off">' +
          '<button class="prim" style="background:#b91c1c" onclick="borrar()">' + th('limpiar.borrar') + '</button>';
      }).limpiezaVista(o);
    }
    function borrar() {
      if (!ultimo) return;
      $('msg').textContent = tt('limpiar.borrando');
      google.script.run.withFailureHandler(errMsg).withSuccessHandler(r => {
        $('vista').innerHTML = ''; ultimo = null;
        $('msg').innerHTML = '<div class="ok">' + th('limpiar.listo', {n: r.borradas, archivo: r.archivo}) +
          (r.hojas ? '<ul>' + r.hojas.map(x => '<li>' + esc(x.hoja) + ': ' + x.n + '</li>').join('') + '</ul>' : '') +
          (r.enBlanco ? '<p>' + th('blanco.despues', {nombre: r.titulo}) + '</p>' : '') +
          (r.respaldo ? '<p><a href="' + esc(r.respaldo) + '" target="_blank">' + th('limpiar.verRespaldo') + '</a><br><small>' + th('limpiar.queEsRespaldo') + '</small></p>' : '') + '</div>';
        cargarOpciones();
      }).limpiezaAplicar(ultimo.op, ultimo.total, $('conf').value);
    }
    function verRestaurar(o) {
      if (!o.respaldo) { errMsg({message: tt('limpiar.r.elige')}); return; }
      $('vista').innerHTML = '<p><small>' + th('comun.cargando') + '</small></p>';
      google.script.run.withFailureHandler(e => { $('vista').innerHTML = ''; errMsg(e); }).withSuccessHandler(r => {
        ultimo = { op: o, total: r.total };
        let h = r.total ? '<div class="ok"><b>' + th('limpiar.r.seAgregarian', {n: r.total}) + '</b><ul>' + r.hojas.filter(x => x.n).map(x => '<li>' + esc(x.hoja) + ': ' + x.n + '</li>').join('') + '</ul></div>' : '<div class="ok">' + th('limpiar.r.nada') + '</div>';
        if (r.yaEstan) h += '<p><small>' + th('limpiar.r.yaEstan', {n: r.yaEstan}) + '</small></p>';
        if (r.config) h += '<label style="font-weight:normal"><input type="checkbox" id="r_config"> ' + th('limpiar.r.config', {nombre: r.config.titulo}) + '</label><p><small>' + th('limpiar.r.config.d') + '</small></p>';
        if (r.total || r.config) h += '<button class="prim" onclick="restaurar()">' + th('limpiar.r.boton') + '</button>';
        $('vista').innerHTML = h;
      }).limpiezaRestaurarVista(o.respaldo);
    }
    function restaurar() {
      if (!ultimo) return;
      $('msg').textContent = tt('limpiar.r.trabajando');
      google.script.run.withFailureHandler(errMsg).withSuccessHandler(r => {
        $('vista').innerHTML = ''; ultimo = null;
        $('msg').innerHTML = '<div class="ok">' + th('limpiar.r.listo', {n: r.agregadas}) + (r.hojas.length ? '<ul>' + r.hojas.map(x => '<li>' + esc(x.hoja) + ': ' + x.n + '</li>').join('') + '</ul>' : '') +
          (r.config ? '<p>' + th('limpiar.r.configListo') + '</p>' : '') + '</div>';
        cargarOpciones();
      }).limpiezaRestaurar(ultimo.op.respaldo, on('r_config'), ultimo.total);
    }
    function cargarOpciones() { google.script.run.withFailureHandler(errMsg).withSuccessHandler(o => { O = o; detalle(); }).limpiezaOpciones(); }
    document.querySelectorAll('input[name="tipo"]').forEach(r => r.onchange = detalle);
    detalle(); cargarOpciones();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(680).setHeight(700), tx('limpiar.titulo'));
}
// ─────────────────────────────────────────────────────────────────────────────
// 20. COPIAS DE TRABAJO (RAMAS) Y PROPUESTAS DE CAMBIOS
// ─────────────────────────────────────────────────────────────────────────────
// Desde el catálogo principal se crea una copia de trabajo: conserva los códigos de todo, no publica en
// el visor y guarda una FOTO (huella de cada fila) de cómo estaba el principal al crearla.
// Quien administra el principal revisa luego lo que cambió en la copia (agregado, corregido, borrado)
// y decide qué se integra. Si esa misma fila también cambió en el principal desde la foto, se marca
// como conflicto y no se integra salvo que se elija expresamente.

const HOJA_FOTO = 'COPIA_FOTO';
const ESTADOS_COPIA = ['Abierta', 'Integrada en parte', 'Integrada', 'Descartada'];
const COLS_COPIAS = ['id_hoja', 'nombre', 'autor', 'fecha', 'url', 'estado', 'integrada_por', 'fecha_integracion', 'notas'];
const COLS_IGNORAR_EN_COPIAS = ['imagen_id', 'imagen_ref', 'imagen_publica'];  // se recalculan solos

/** Hojas que se comparan, en orden de dependencia (para agregar); borrar va al revés. */
function _hojasRama() {
  const S = CONFIG.SHEETS;
  return [S.IDIOMAS, S.COMUNIDADES, S.ACTORES, S.FUENTES, S.AUTORIZACIONES, S.VOCES, S.CATEGORIAS, S.TRAD_CATS, S.PALABRAS, S.TRADUCCIONES, S.TIPOS_REL, S.RELACIONES, S.IMAGENES, S.REGLAS];
}
/** Cómo se identifica una fila de cada hoja. */
function _claveFila(hoja, t, f) {
  const S = CONFIG.SHEETS;
  if (hoja === S.REGLAS) return [_get(t, f, 'id_idioma'), _get(t, f, 'buscar'), _get(t, f, 'aplica_a')].map(_k).join('|');
  if (hoja === S.TRAD_CATS) {
    const idc = t.headers.find(h => h.startsWith('id_') && h !== 'id_categoria' && h !== 'id_idioma');
    return (idc && _s(_get(t, f, idc))) || [_get(t, f, 'id_categoria'), _get(t, f, 'id_idioma')].map(_k).join('|');
  }
  const id = t.headers.find(h => h.startsWith('id_'));
  return id ? _s(_get(t, f, id)) : '';
}
const _valorFila = v => v instanceof Date ? v.toISOString() : _s(v);
function _huellaFila(t, f, cols) { return _hash(JSON.stringify(cols.map(c => _valorFila(_get(t, f, c))))); }
function _colsRama(t) { return t.headers.filter(h => h && !COLS_IGNORAR_EN_COPIAS.includes(h)); }

/** Configuración leída directo de la hoja (sin la memoria de getAppConfig, que es de la hoja activa). */
function _configCruda(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!sh) return {};
  const d = sh.getDataRange().getValues(), o = {};
  (d[0] || []).forEach((h, i) => { o[_s(h).toUpperCase()] = _s((d[1] || [])[i]); });
  return o;
}
function _esRama(ss) { return !!_configCruda(ss || SpreadsheetApp.getActiveSpreadsheet()).RAMA_DE; }

/** Foto de un catálogo: { hoja: { cols, filas: { clave: huella } } }. */
function _foto(ss) {
  const r = {};
  _hojasRama().forEach(h => {
    const t = _tabla(ss, h);
    if (!t) return;
    const cols = _colsRama(t), filas = {};
    t.filas.forEach(f => { const k = _claveFila(h, t, f); if (k) filas[k] = _huellaFila(t, f, cols); });
    r[h] = { cols, filas };
  });
  return r;
}
function _guardarFoto(ssDestino, foto) {
  let sh = ssDestino.getSheetByName(HOJA_FOTO);
  if (!sh) sh = ssDestino.insertSheet(HOJA_FOTO);
  const filas = [['hoja', 'clave', 'huella']];
  Object.keys(foto).forEach(h => {
    filas.push([h, '__cols', JSON.stringify(foto[h].cols)]);
    Object.keys(foto[h].filas).forEach(k => filas.push([h, k, foto[h].filas[k]]));
  });
  sh.clearContents();
  sh.getRange(1, 1, filas.length, 3).setNumberFormat('@').setValues(filas);
  try { sh.hideSheet(); } catch (e) {}
}
function _leerFoto(ss) {
  const sh = ss.getSheetByName(HOJA_FOTO);
  if (!sh) return null;
  const r = {};
  sh.getDataRange().getValues().slice(1).forEach(([h, k, v]) => {
    h = _s(h); if (!h) return;
    r[h] = r[h] || { cols: [], filas: {} };
    if (_s(k) === '__cols') r[h].cols = JSON.parse(_s(v)); else r[h].filas[_s(k)] = _s(v);
  });
  return r;
}

function _hojaCopias(ss) {
  let t = _tabla(ss, CONFIG.SHEETS.COPIAS);
  if (!t) { asegurarEstructura(ss); t = _tabla(ss, CONFIG.SHEETS.COPIAS); }
  return t;
}

/** Crea una copia de trabajo del catálogo principal. */
function ramaCrear(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (_esRama(ss)) throw new Error(tx('rama.error.desdeCopia'));
  nombre = _s(nombre).slice(0, 60);
  if (!nombre) throw new Error(tx('rama.error.nombre'));
  const lock = _lock();
  try {
    const cfg = getAppConfig(ss);
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const foto = _foto(ss);  // antes de copiar: así la foto coincide exactamente con lo copiado
    const archivo = DriveApp.getFileById(ss.getId()).makeCopy(tx('rama.archivo', { catalogo: cfg.titulo || ss.getName(), nombre }), _subcarpeta(_carpetaDelLibro(ss), 'COPIAS DE TRABAJO'));
    const copia = SpreadsheetApp.openById(archivo.getId());
    _escribirConfig(copia, { RAMA_DE: ss.getId(), RAMA_URL: ss.getUrl ? ss.getUrl() : '', RAMA_NOMBRE: nombre, RAMA_FECHA: fecha, RAMA_AUTOR: _correoActual(),
      APPSHEET_APP_ID: '', NOMBRE_PROYECTO: tx('rama.nombreCatalogo', { catalogo: cfg.titulo || ss.getName(), nombre }) });
    _guardarFoto(copia, foto);
    _cfgMemo = null;
    _agregar(_hojaCopias(ss), [{ id_hoja: archivo.getId(), nombre, autor: _correoActual(), fecha, url: copia.getUrl(), estado: 'Abierta' }]);
    _historial(ss, CONFIG.SHEETS.COPIAS, nombre, {}, { estado: 'Abierta' });
    return { url: copia.getUrl(), nombre };
  } finally { lock.releaseLock(); }
}

/** Lo que la ventana necesita, tanto en el principal como dentro de una copia. */
function ramasListar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const c = _configCruda(ss);
  if (c.RAMA_DE) {
    const r = _cambiosRama(ss, null);
    return { enCopia: true, nombre: c.RAMA_NOMBRE, fecha: c.RAMA_FECHA, autor: c.RAMA_AUTOR, principal: c.RAMA_URL, resumen: r.resumen };
  }
  _requerirAdmin(ss);
  const t = _hojaCopias(ss);
  return { enCopia: false, copias: t.filas.filter(f => _s(_get(t, f, 'id_hoja'))).map(f => _filaA(t, f, COLS_COPIAS)) };
}

/**
 * Compara una copia contra su foto (qué cambió en la copia) y contra el principal (si también cambió allá).
 * principal = null → solo cuenta lo que cambió en la copia.
 */
function _cambiosRama(copia, principal) {
  const foto = _leerFoto(copia);
  if (!foto) throw new Error(tx('rama.error.sinFoto'));
  const cambios = [], resumen = { agregar: 0, cambiar: 0, borrar: 0, conflictos: 0, yaIguales: 0 };
  const etiqueta = (t, f) => {
    for (const c of ['texto', 'nombre_categoria', 'nombre_propio', 'titulo', 'nombre', 'etiqueta', 'buscar', 'tipo']) { const v = _s(_get(t, f, c)); if (v) return v; }
    return '';
  };
  _hojasRama().forEach(h => {
    const tc = _tabla(copia, h), base = foto[h];
    if (!tc || !base) return;
    const cols = base.cols.filter(c => tc.headers.includes(c));
    const tp = principal ? _tabla(principal, h) : null;
    const enP = {};
    if (tp) tp.filas.forEach((f, i) => { const k = _claveFila(h, tp, f); if (k) enP[k] = i; });
    const huellaP = k => (tp && enP[k] != null) ? _huellaFila(tp, tp.filas[enP[k]], cols) : null;
    const vistas = new Set();
    tc.filas.forEach(f => {
      const k = _claveFila(h, tc, f); if (!k) return;
      vistas.add(k);
      const hc = _huellaFila(tc, f, cols), hb = base.filas[k];
      if (hb === hc) return;                                   // no cambió en la copia
      const tipo = hb ? 'cambiar' : 'agregar';
      const hp = principal ? huellaP(k) : null;
      if (principal && hp === hc) { resumen.yaIguales++; return; } // el principal ya lo tiene igual
      const conflicto = !!principal && (tipo === 'agregar' ? hp != null : hp !== hb);
      resumen[tipo]++; if (conflicto) resumen.conflictos++;
      if (!principal) return;
      const fila = {}; cols.forEach(c => { fila[c] = _valorFila(_get(tc, f, c)); });
      const antes = {}, campos = [];
      if (hp != null) { const fp = tp.filas[enP[k]]; cols.forEach(c => { const a = _valorFila(_get(tp, fp, c)); if (a !== fila[c]) { campos.push(c); antes[c] = a; } }); }
      cambios.push({ hoja: h, clave: k, tipo, conflicto, etiqueta: etiqueta(tc, f), fila, antes, campos });
    });
    Object.keys(base.filas).forEach(k => {
      if (k === '__cols' || vistas.has(k)) return;
      const hp = principal ? huellaP(k) : null;
      if (principal && hp == null) { resumen.yaIguales++; return; } // también se borró en el principal
      const conflicto = !!principal && hp !== base.filas[k];
      resumen.borrar++; if (conflicto) resumen.conflictos++;
      if (!principal) return;
      const fp = tp.filas[enP[k]], antes = {}; cols.forEach(c => { antes[c] = _valorFila(_get(tp, fp, c)); });
      cambios.push({ hoja: h, clave: k, tipo: 'borrar', conflicto, etiqueta: etiqueta(tp, fp), fila: {}, antes, campos: [] });
    });
  });
  return { cambios, resumen };
}

function _abrirCopia(ss, idHoja) {
  const t = _hojaCopias(ss);
  if (!t.filas.some(f => _s(_get(t, f, 'id_hoja')) === _s(idHoja))) throw new Error(tx('rama.error.noEsta'));
  try { return SpreadsheetApp.openById(_s(idHoja)); } catch (e) { throw new Error(tx('rama.error.acceso')); }
}

/** En el principal: qué propone una copia. */
function ramaComparar(idHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (_esRama(ss)) throw new Error(tx('rama.error.desdeCopia'));
  const r = _cambiosRama(_abrirCopia(ss, idHoja), ss);
  r.cambios.forEach((c, i) => { c.n = i; });
  return r;
}

/** Integra al principal los cambios elegidos (por hoja|clave). Respaldo antes, historial después. */
function ramaIntegrar(idHoja, elegidos, esperados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (_esRama(ss)) throw new Error(tx('rama.error.desdeCopia'));
  const lock = _lock();
  try {
    const copia = _abrirCopia(ss, idHoja);
    const { cambios } = _cambiosRama(copia, ss);
    const quiero = new Set((elegidos || []).map(String));
    const sel = cambios.filter(c => quiero.has(c.hoja + '|' + c.clave));
    if (sel.length !== Number(esperados)) throw new Error(tx('limpiar.error.cambio'));
    if (!sel.length) return { integrados: 0 };
    // 1. Respaldo de lo que había en el principal en esas filas
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const t0 = _hojaCopias(ss), fc = t0.filas.find(f => _s(_get(t0, f, 'id_hoja')) === _s(idHoja));
    const nombre = _s(_get(t0, fc, 'nombre'));
    const libro = SpreadsheetApp.create(tx('rama.respaldo', { nombre, fecha }));
    const filasResp = [['hoja', 'clave', 'cambio', 'antes', 'despues']].concat(sel.map(c => [c.hoja, c.clave, c.tipo, JSON.stringify(c.antes), JSON.stringify(c.fila)]));
    libro.getSheets()[0].getRange(1, 1, filasResp.length, 5).setNumberFormat('@').setValues(filasResp);
    DriveApp.getFileById(libro.getId()).moveTo(_subcarpeta(_carpetaDelLibro(ss), 'RESPALDOS'));
    // 2. Agregar y cambiar (en orden de dependencia); borrar al final y al revés
    const orden = _hojasRama();
    const rep = { agregar: 0, cambiar: 0, borrar: 0 };
    orden.forEach(h => {
      const deHoja = sel.filter(c => c.hoja === h && c.tipo !== 'borrar');
      if (!deHoja.length) return;
      const t = _tabla(ss, h);
      const idx = {}; t.filas.forEach((f, i) => { const k = _claveFila(h, t, f); if (k) idx[k] = i; });
      const nuevos = [];
      deHoja.forEach(c => {
        if (idx[c.clave] != null) { const o = {}; Object.keys(c.fila).forEach(k => { if (t.headers.includes(k)) o[k] = c.fila[k]; }); _actualizar(t, idx[c.clave], o); rep.cambiar++; }
        else { nuevos.push(c.fila); rep.agregar++; }
      });
      _agregar(t, nuevos);
    });
    orden.slice().reverse().forEach(h => {
      const deHoja = sel.filter(c => c.hoja === h && c.tipo === 'borrar');
      if (!deHoja.length) return;
      const t = _tabla(ss, h), claves = new Set(deHoja.map(c => c.clave));
      const idx = []; t.filas.forEach((f, i) => { if (claves.has(_claveFila(h, t, f))) idx.push(i); });
      if (idx.length) _espacioParaBorrar(t.sh, idx.length);
      idx.sort((a, b) => b - a).forEach(i => { t.sh.deleteRows(i + 2, 1); });
      rep.borrar += idx.length;
    });
    // 3. Registro
    const quedan = _cambiosRama(copia, ss).cambios.length;
    const i0 = t0.filas.indexOf(fc);
    _actualizar(t0, i0, { estado: quedan ? 'Integrada en parte' : 'Integrada', integrada_por: _correoActual(), fecha_integracion: fecha });
    _historial(ss, CONFIG.SHEETS.COPIAS, nombre, {}, { integrado: tx('rama.historial', rep) });
    _cfgMemo = null;
    try { publicarCatalogo(ss, true); } catch (e) {}
    return Object.assign({ integrados: sel.length, quedan, respaldo: libro.getUrl() }, rep);
  } finally { lock.releaseLock(); }
}

function ramaDescartar(idHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const t = _hojaCopias(ss), i = t.filas.findIndex(f => _s(_get(t, f, 'id_hoja')) === _s(idHoja));
  if (i < 0) throw new Error(tx('rama.error.noEsta'));
  _actualizar(t, i, { estado: 'Descartada' });
  _historial(ss, CONFIG.SHEETS.COPIAS, _get(t, t.filas[i], 'nombre'), {}, { estado: 'Descartada' });
  return ramasListar();
}

function uiRamas() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <div class="caja">
    <div id="vista">${txh('comun.cargando')}</div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null, C = null, actual = null;
    const err = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message})) + '</div>'; };
    const TIPO = { agregar: '➕', cambiar: '✏️', borrar: '🗑️' };
    function pintar(d) {
      D = d; C = null; $('msg').innerHTML = '';
      if (d.enCopia) {
        const r = d.resumen;
        $('vista').innerHTML = '<div class="sec"><h3>' + th('rama.enCopia', {nombre: d.nombre}) + '</h3><p>' + th('rama.enCopia.d', {fecha: d.fecha, autor: d.autor}) + '</p>' +
          '<p><b>' + th('rama.misCambios', r) + '</b></p><p><small>' + th('rama.enCopia.pasos') + '</small></p>' +
          (d.principal ? '<a href="' + esc(d.principal) + '" target="_blank">' + th('rama.abrirPrincipal') + '</a>' : '') + '</div>';
        return;
      }
      $('vista').innerHTML = '<p style="margin-top:0">' + th('rama.intro') + '</p><div class="sec"><h3>' + th('rama.nueva') + '</h3>' +
        '<div class="dos"><div><label for="nNombre">' + th('rama.nombre') + '</label><input type="text" id="nNombre" placeholder="' + esc(tt('rama.nombre.ej')) + '"></div></div>' +
        '<button class="prim" onclick="crear()">' + th('rama.crear') + '</button></div>' +
        '<div class="sec"><h3>' + th('rama.lista') + '</h3>' + (d.copias.length ? '<table><tr><th>' + th('rama.col.nombre') + '</th><th>' + th('rama.col.quien') + '</th><th>' + th('rama.col.estado') + '</th><th></th></tr>' +
          d.copias.map(c => '<tr><td><b>' + esc(c.nombre) + '</b><br><a href="' + esc(c.url) + '" target="_blank"><small>' + th('rama.abrir') + '</small></a></td><td><small>' + esc(c.autor) + '<br>' + esc(c.fecha) + '</small></td><td><small>' + esc(c.estado) +
            (c.integrada_por ? '<br>' + esc(c.integrada_por) : '') + '</small></td><td>' + (c.estado !== 'Descartada' ? '<button class="peq" onclick="revisar(\\'' + esc(c.id_hoja) + '\\')">' + th('rama.revisar') + '</button> <button class="peq" onclick="descartar(\\'' + esc(c.id_hoja) + '\\')">' + th('rama.descartar') + '</button>' : '') + '</td></tr>').join('') + '</table>'
          : '<p><small>' + th('rama.ninguna') + '</small></p>') + '</div><div id="revision"></div>';
    }
    function crear() { $('msg').textContent = tt('rama.creando'); google.script.run.withFailureHandler(err).withSuccessHandler(r => {
      google.script.run.withSuccessHandler(d => { pintar(d); $('msg').innerHTML = '<div class="ok">' + th('rama.creada', {nombre: r.nombre}) + ' <a href="' + esc(r.url) + '" target="_blank">' + th('rama.abrir') + '</a></div>'; }).ramasListar(); }).ramaCrear($('nNombre').value); }
    function descartar(id) { google.script.run.withFailureHandler(err).withSuccessHandler(pintar).ramaDescartar(id); }
    function revisar(id) {
      actual = id; $('revision').innerHTML = '<p>' + th('rama.comparando') + '</p>';
      google.script.run.withFailureHandler(err).withSuccessHandler(r => { C = r; pintarRevision(); }).ramaComparar(id);
    }
    function pintarRevision() {
      const r = C.resumen;
      if (!C.cambios.length) { $('revision').innerHTML = '<div class="ok">' + th('rama.sinCambios', {iguales: r.yaIguales}) + '</div>'; return; }
      const fila = c => {
        let det = '';
        if (c.tipo === 'cambiar') det = c.campos.map(k => '<small><b>' + esc(k) + ':</b> <s>' + esc(c.antes[k] || '∅') + '</s> → ' + esc(c.fila[k] || '∅') + '</small>').join('<br>');
        if (c.tipo === 'agregar') det = '<small>' + esc(Object.keys(c.fila).filter(k => c.fila[k] && !k.startsWith('id_')).slice(0, 4).map(k => k + ': ' + c.fila[k]).join(' · ')) + '</small>';
        if (c.tipo === 'borrar') det = '<small>' + th('rama.seBorra') + '</small>';
        return '<tr' + (c.conflicto ? ' style="background:#fef3c7"' : '') + '><td><input type="checkbox" class="sel" value="' + esc(c.hoja + '|' + c.clave) + '"' + (c.conflicto ? '' : ' checked') + ' aria-label="' + esc(c.etiqueta) + '"></td><td>' + TIPO[c.tipo] +
          '</td><td><small>' + esc(c.hoja) + '</small><br><b>' + esc(c.etiqueta || c.clave) + '</b>' + (c.conflicto ? '<br><span class="chip ambar">' + th('rama.conflicto') + '</span>' : '') + '</td><td>' + det + '</td></tr>';
      };
      $('revision').innerHTML = '<div class="sec"><h3>' + th('rama.propone', r) + '</h3>' + (r.conflictos ? '<p><small>' + th('rama.conflicto.d') + '</small></p>' : '') +
        (r.yaIguales ? '<p><small>' + th('rama.yaIguales', {n: r.yaIguales}) + '</small></p>' : '') +
        '<label style="font-weight:normal"><input type="checkbox" id="todos" checked onchange="document.querySelectorAll(\\'.sel\\').forEach(x => x.checked = this.checked)"> ' + th('rama.todos') + '</label>' +
        '<div style="max-height:320px;overflow:auto"><table>' + C.cambios.map(fila).join('') + '</table></div>' +
        '<button class="prim" onclick="integrar()">' + th('rama.integrar') + '</button></div>';
    }
    function integrar() {
      const sel = Array.from(document.querySelectorAll('.sel:checked')).map(x => x.value);
      if (!sel.length) { $('msg').innerHTML = '<div class="error">' + th('rama.eligeAlgo') + '</div>'; return; }
      $('msg').textContent = tt('rama.integrando');
      google.script.run.withFailureHandler(err).withSuccessHandler(r => {
        $('msg').innerHTML = '<div class="ok">' + th('rama.integrado', r) + (r.respaldo ? ' <a href="' + esc(r.respaldo) + '" target="_blank">' + th('limpiar.verRespaldo') + '</a>' : '') + '</div>';
        const aviso = $('msg').innerHTML;
        google.script.run.withSuccessHandler(d => { pintar(d); $('msg').innerHTML = aviso; }).ramasListar();
      }).ramaIntegrar(actual, sel, sel.length);
    }
    google.script.run.withFailureHandler(err).withSuccessHandler(pintar).ramasListar();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(860).setHeight(700), tx('rama.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 24. RELACIONES ENTRE PALABRAS (vocabulario de tipos, captura y publicación)
// ─────────────────────────────────────────────────────────────────────────────
// RELACIONES guarda vínculos entre palabras: «parte de», «sinónimo», «variante de»…
// Los tipos forman un vocabulario controlado (hoja TIPOS_RELACION) que cada catálogo ajusta desde
// 🕸️ Tipos de relación. Los básicos vienen de fábrica; su nombre vacío = el nombre traducido automáticamente
// (en el idioma de los menús y en el del visor). Cada vínculo se guarda una sola vez (origen → destino);
// la dirección contraria se lee con el «nombre inverso» («Parte de» / «Tiene como parte»).
// Esta misma hoja es la que se vuelve aristas al exportar a grafo.

const COLS_TIPOS_REL = ['id_tipo', 'nombre', 'nombre_inverso', 'simetrica', 'publica', 'sistema', 'orden', 'estado', 'nota'];
const TIPOS_REL_BASE = [
  { id: 'relacionada', clave: 'relacionada', sim: true },
  { id: 'sinonimo',    clave: 'sinonimo',    sim: true },
  { id: 'antonimo',    clave: 'antonimo',    sim: true },
  { id: 'tipo_de',     clave: 'tipo_de' },
  { id: 'parte_de',    clave: 'parte_de' },
  { id: 'variante_de', clave: 'variante_de' },
  { id: 'deriva_de',   clave: 'deriva_de' },
  { id: REL_OTRA_IDEA, clave: 'homonimo', sim: true, sistema: true },               // la anota 🔁 Palabras parecidas
  { id: REL_UNIDA,     clave: 'unida', sistema: true, privada: true, fija: true }    // la anota «es la misma idea»: solo para el historial
];
const _tipoRelBase = id => TIPOS_REL_BASE.find(b => _k(b.id) === _k(id));

/** Tipos de relación: los de la hoja, los básicos que falten y los que aparezcan escritos en RELACIONES sin definir. */
function _tiposRelacion(ss) {
  const m = new Map();
  if (ss.getSheetByName(CONFIG.SHEETS.TIPOS_REL)) readSheet(ss, CONFIG.SHEETS.TIPOS_REL).forEach((r, i) => {
    if (!_s(r.id_tipo) || m.has(_k(r.id_tipo))) return;
    const b = _tipoRelBase(r.id_tipo);
    m.set(_k(r.id_tipo), { id: _s(r.id_tipo), clave: b ? b.clave : '', nombre: _s(r.nombre), inverso: _s(r.nombre_inverso),
      simetrica: b ? !!b.sim : _siNo(r.simetrica), publica: b && b.fija ? false : (_s(r.publica) === '' ? !(b && b.privada) : _siNo(r.publica)),
      sistema: !!(b && b.sistema), base: !!b, orden: Number(r.orden) || (i + 1) * 10,
      activo: b && b.sistema ? true : _estadoCanonico(r.estado) !== 'retirado', nota: _s(r.nota), enHoja: true });
  });
  TIPOS_REL_BASE.forEach((b, i) => {
    if (!m.has(_k(b.id))) m.set(_k(b.id), { id: b.id, clave: b.clave, nombre: '', inverso: '', simetrica: !!b.sim, publica: !b.privada,
      sistema: !!b.sistema, base: true, orden: (i + 1) * 10, activo: true, nota: '', enHoja: false });
  });
  if (ss.getSheetByName(CONFIG.SHEETS.RELACIONES)) readSheet(ss, CONFIG.SHEETS.RELACIONES).forEach(r => {
    const t = _s(r.tipo);
    if (t && !m.has(_k(t))) m.set(_k(t), { id: t, clave: '', nombre: t, inverso: '', simetrica: false, publica: true, sistema: false, base: false,
      orden: 900, activo: true, nota: '', enHoja: false });
  });
  return [...m.values()].sort((a, b) => a.orden - b.orden);
}

/** Nombre que se lee: el que puso el catálogo, o el de fábrica traducido. inversa = leído desde la otra palabra. */
function _etiquetaTipoRel(t, inversa) {
  if (!t) return '';
  if (inversa && !t.simetrica) return t.inverso || (t.clave ? tx('visor.rel.' + t.clave + '.inv') : tx('visor.rel.inversaDe', { nombre: t.nombre || t.id }));
  return t.nombre || (t.clave ? tx('visor.rel.' + t.clave) : t.id);
}

/**
 * Deja la hoja TIPOS_RELACION completa (crea la hoja si falta, agrega los básicos y los tipos escritos a mano)
 * y lleva a su código los tipos escritos con otras palabras («sinónimo», «Parte de»…). No borra nada.
 */
function _asegurarTiposRelacion(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEETS.TIPOS_REL);
  if (!sh) { asegurarEstructura(ss); sh = ss.getSheetByName(CONFIG.SHEETS.TIPOS_REL); if (!sh) return { nuevos: 0, normalizadas: 0 }; }
  let normalizadas = 0;
  const tRel = _tabla(ss, CONFIG.SHEETS.RELACIONES);
  if (tRel && _col(tRel, 'tipo') > -1) {
    const definidos = new Set(readSheet(ss, CONFIG.SHEETS.TIPOS_REL).map(r => _k(r.id_tipo)).filter(Boolean));
    const equivalente = {};
    TIPOS_REL_BASE.forEach(b => {
      [b.id, b.id.replace(/_/g, ' '), tx('visor.rel.' + b.clave)].concat(b.sim ? [] : [tx('visor.rel.' + b.clave + '.inv')])
        .forEach(x => { equivalente[_plegar(x)] = { id: b.id, inv: x === tx('visor.rel.' + b.clave + '.inv') && !b.sim }; });
    });
    tRel.filas.forEach((f, i) => {
      const v = _s(_get(tRel, f, 'tipo'));
      if (!v || definidos.has(_k(v)) || _tipoRelBase(v)) return;
      const e = equivalente[_plegar(v)];
      if (!e) return;
      // «Tiene como parte» escrito a mano = «Parte de» leído al revés
      const c = e.inv ? { tipo: e.id, id_origen: _get(tRel, f, 'id_destino'), id_destino: _get(tRel, f, 'id_origen') } : { tipo: e.id };
      _actualizar(tRel, i, c);
      normalizadas++;
    });
  }
  const t = _tabla(ss, CONFIG.SHEETS.TIPOS_REL);
  const ya = new Set(t.filas.map(f => _k(_get(t, f, 'id_tipo'))));
  const faltan = _tiposRelacion(ss).filter(x => !ya.has(_k(x.id))).map(x => ({ id_tipo: x.id, nombre: x.base ? '' : x.nombre, nombre_inverso: x.inverso,
    simetrica: x.simetrica ? 'Sí' : 'No', publica: x.publica ? 'Sí' : 'No', sistema: x.sistema ? 'Sí' : 'No', orden: x.orden,
    estado: _etiquetaEstado('publicado'), nota: x.base ? '' : tx('rel.tipoEncontrado') }));
  if (faltan.length) _agregar(t, faltan);
  return { nuevos: faltan.length, normalizadas };
}

/** Al traer relaciones de otro catálogo: sus tipos que aquí no existen se agregan con el nombre que traen. */
function _tiposDesdePaquete(ss, tipos, usados) {
  if (!ss.getSheetByName(CONFIG.SHEETS.TIPOS_REL)) asegurarEstructura(ss);
  const t = _tabla(ss, CONFIG.SHEETS.TIPOS_REL);
  if (!t) return 0;
  const ya = new Set(t.filas.map(f => _k(_get(t, f, 'id_tipo'))));
  const porId = new Map((tipos || []).map(x => [_k(x.id), x]));
  const nuevos = [];
  (usados || []).map(_s).forEach(id => {
    if (!id || ya.has(_k(id)) || _tipoRelBase(id)) return;
    ya.add(_k(id));
    const x = porId.get(_k(id)) || {};
    nuevos.push({ id_tipo: id, nombre: _s(x.nombre) || id, nombre_inverso: x.simetrica ? '' : _s(x.inverso), simetrica: x.simetrica ? 'Sí' : 'No',
      publica: 'Sí', sistema: 'No', orden: 900, estado: _etiquetaEstado('publicado'), nota: tx('rel.tipoImportado') });
  });
  if (nuevos.length) _agregar(t, nuevos);
  return nuevos.length;
}

/** Relaciones vigentes de una palabra, leídas desde ella (con la dirección y el nombre que corresponde). */
function _relacionesDe(ss, idPalabra) {
  const tipos = new Map(_tiposRelacion(ss).map(t => [_k(t.id), t]));
  const textos = {}, principal = _k(getAppConfig(ss).idiomaPrincipal);
  readSheet(ss, CONFIG.SHEETS.TRADUCCIONES).forEach(t => {
    if (_estadoCanonico(t.estado) === 'retirado') return;
    const p = _k(t.id_palabra);
    if (!textos[p] || _k(t.id_idioma) === principal) textos[p] = _s(t.texto);
  });
  const yo = _k(idPalabra);
  return readSheet(ss, CONFIG.SHEETS.RELACIONES).filter(r => _s(r.id_relacion) && _estadoCanonico(r.estado) !== 'retirado' &&
      (_k(r.id_origen) === yo || _k(r.id_destino) === yo) && _k(r.id_origen) !== _k(r.id_destino))
    .map(r => {
      const entra = _k(r.id_origen) !== yo, otra = _s(entra ? r.id_origen : r.id_destino), t = tipos.get(_k(r.tipo));
      return { id: _s(r.id_relacion), tipo: _s(r.tipo), etiqueta: _etiquetaTipoRel(t, entra) || _s(r.tipo), entra, otra, texto: textos[_k(otra)] || otra,
        nota: _s(r.nota), estado: _estadoCanonico(r.estado) || 'publicado', publica: !!(t && t.publica && t.activo),
        sistema: !!(t && t.sistema), creado_por: _s(r.creado_por) };
    })
    .sort((a, b) => a.etiqueta.localeCompare(b.etiqueta, 'es') || a.texto.localeCompare(b.texto, 'es'));
}

/** Tipos que se pueden elegir al capturar: cada tipo con dirección aparece dos veces («Parte de» y «Tiene como parte»). */
function _opcionesRelacion(ss) {
  const l = [];
  _tiposRelacion(ss).filter(t => t.activo && !(_tipoRelBase(t.id) || {}).fija).forEach(t => {
    l.push({ v: t.id + '|>', t: _etiquetaTipoRel(t, false), publica: t.publica });
    if (!t.simetrica) l.push({ v: t.id + '|<', t: _etiquetaTipoRel(t, true), publica: t.publica });
  });
  return l;
}

/**
 * Agrega una relación desde la palabra `desde`. r = { desde, hacia, opcion: 'id_tipo|>' o 'id_tipo|<', nota }.
 * Quien administra la deja Publicada; quien captura, En revisión. Si ya existía retirada, la reactiva.
 */
function relacionAgregar(r) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  r = r || {};
  const [idTipo, dir] = _s(r.opcion).split('|');
  const tipo = _tiposRelacion(ss).find(t => _k(t.id) === _k(idTipo));
  if (!tipo || !tipo.activo || (_tipoRelBase(tipo.id) || {}).fija) throw new Error(tx('rel.error.tipo'));
  const desde = _s(r.desde), hacia = _s(r.hacia);
  if (!desde || !hacia) throw new Error(tx('rel.error.palabra'));
  if (_k(desde) === _k(hacia)) throw new Error(tx('rel.error.misma'));
  const pal = new Set(readSheet(ss, CONFIG.SHEETS.PALABRAS).map(p => _k(p.id_palabra)));
  if (!pal.has(_k(desde)) || !pal.has(_k(hacia))) throw new Error(tx('rel.error.palabra'));
  const origen = dir === '<' ? hacia : desde, destino = dir === '<' ? desde : hacia;
  const admin = _soyAdmin(ss);
  const lock = _lock();
  try {
    _asegurarTiposRelacion(ss);
    const t = _tablaConEstado(ss, CONFIG.SHEETS.RELACIONES);
    if (!t) throw new Error(tx('comun.faltaHoja', { hoja: CONFIG.SHEETS.RELACIONES }));
    const igual = f => _k(_get(t, f, 'tipo')) === _k(tipo.id) && ((_k(_get(t, f, 'id_origen')) === _k(origen) && _k(_get(t, f, 'id_destino')) === _k(destino)) ||
      (tipo.simetrica && _k(_get(t, f, 'id_origen')) === _k(destino) && _k(_get(t, f, 'id_destino')) === _k(origen)));
    const i = t.filas.findIndex(igual);
    const estado = _etiquetaEstado(admin ? 'publicado' : 'revisión');
    if (i > -1) {
      if (_estadoCanonico(_get(t, t.filas[i], 'estado')) !== 'retirado') throw new Error(tx('rel.error.yaExiste'));
      const c = { estado };
      if (_s(r.nota)) c.nota = _s(r.nota).slice(0, 300);
      _actualizar(t, i, c);
      _historial(ss, CONFIG.SHEETS.RELACIONES, _s(_get(t, t.filas[i], 'id_relacion')), { estado: _get(t, t.filas[i], 'estado') }, c);
    } else {
      const fila = { id_relacion: Utilities.getUuid(), id_origen: origen, tipo: tipo.id, id_destino: destino, nota: _s(r.nota).slice(0, 300),
        estado, creado_por: _correoActual() };
      _agregar(t, [fila]);
      _historial(ss, CONFIG.SHEETS.RELACIONES, fila.id_relacion, {}, { tipo: tipo.id, id_origen: origen, id_destino: destino });
    }
    return { relaciones: _relacionesDe(ss, desde), aviso: admin ? '' : tx('rel.aviso.revision') };
  } finally { lock.releaseLock(); }
}

/** Retira (no borra) una relación. Quien administra, cualquiera; quien captura, solo las suyas sin publicar. */
function relacionRetirar(idRelacion, desde) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = _lock();
  try {
    const t = _tablaConEstado(ss, CONFIG.SHEETS.RELACIONES);
    const i = t ? t.filas.findIndex(f => _k(_get(t, f, 'id_relacion')) === _k(idRelacion)) : -1;
    if (i < 0) throw new Error(tx('rel.error.noExiste'));
    const f = t.filas[i];
    if (!_soyAdmin(ss) && !(_k(_get(t, f, 'creado_por')) === _k(_correoActual()) && _estadoCanonico(_get(t, f, 'estado')) !== 'publicado'))
      throw new Error(tx('rel.error.soloAdmin'));
    const antes = { estado: _get(t, f, 'estado') }, despues = { estado: _etiquetaEstado('retirado') };
    _actualizar(t, i, despues);
    _historial(ss, CONFIG.SHEETS.RELACIONES, _s(idRelacion), antes, despues);
    return { relaciones: _relacionesDe(ss, desde) };
  } finally { lock.releaseLock(); }
}

/** Publica (o devuelve a revisión) una relación desde la ficha. Solo quien administra. */
function relacionEstado(idRelacion, canon, desde) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  if (!['publicado', 'revisión'].includes(canon)) throw new Error(tx('estados.error.estado'));
  const lock = _lock();
  try {
    const t = _tablaConEstado(ss, CONFIG.SHEETS.RELACIONES);
    const i = t.filas.findIndex(f => _k(_get(t, f, 'id_relacion')) === _k(idRelacion));
    if (i < 0) throw new Error(tx('rel.error.noExiste'));
    const antes = { estado: _get(t, t.filas[i], 'estado') }, despues = { estado: _etiquetaEstado(canon) };
    _actualizar(t, i, despues);
    _historial(ss, CONFIG.SHEETS.RELACIONES, _s(idRelacion), antes, despues);
    return { relaciones: _relacionesDe(ss, desde) };
  } finally { lock.releaseLock(); }
}

// ── 🕸️ Tipos de relación (ventana)

function tiposRelacionDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const admin = _soyAdmin(ss);
  let rep = { nuevos: 0, normalizadas: 0 };
  if (admin) { const lock = _lock(); try { rep = _asegurarTiposRelacion(ss); } finally { lock.releaseLock(); } }
  const usos = {};
  readSheet(ss, CONFIG.SHEETS.RELACIONES).forEach(r => { if (_estadoCanonico(r.estado) !== 'retirado') usos[_k(r.tipo)] = (usos[_k(r.tipo)] || 0) + 1; });
  return { soyAdmin: admin, normalizadas: rep.normalizadas,
    tipos: _tiposRelacion(ss).map(t => Object.assign({}, t, { fija: !!(_tipoRelBase(t.id) || {}).fija, usos: usos[_k(t.id)] || 0,
      autoNombre: t.clave ? tx('visor.rel.' + t.clave) : '', autoInverso: t.clave && !t.simetrica ? tx('visor.rel.' + t.clave + '.inv') : '' })) };
}

/** Guarda un tipo. d = { id ('' = nuevo), nombre, nombre_inverso, simetrica, publica, orden, activo, nota } */
function tipoRelacionGuardar(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  d = d || {};
  const lock = _lock();
  try {
    _asegurarTiposRelacion(ss);
    const t = _tabla(ss, CONFIG.SHEETS.TIPOS_REL);
    const nombre = _s(d.nombre).slice(0, 60), inverso = _s(d.nombre_inverso).slice(0, 60);
    let id = _s(d.id), i = id ? t.filas.findIndex(f => _k(_get(t, f, 'id_tipo')) === _k(id)) : -1;
    const b = id ? _tipoRelBase(id) : null;
    if (!id) {
      if (!nombre) throw new Error(tx('rel.error.nombre'));
      if (_tiposRelacion(ss).some(x => _plegar(_etiquetaTipoRel(x, false)) === _plegar(nombre))) throw new Error(tx('rel.error.nombreRepetido', { nombre }));
      const base = _sinAcentos(nombre).replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 30) || 'tipo';
      const usados = new Set(t.filas.map(f => _k(_get(t, f, 'id_tipo'))));
      id = base; for (let n = 2; usados.has(_k(id)) || _tipoRelBase(id); n++) id = base + '_' + n;
    } else if (i < 0) throw new Error(tx('rel.error.tipo'));
    else if (!b && !nombre) throw new Error(tx('rel.error.nombre'));
    const fila = {
      nombre, nombre_inverso: (b ? b.sim : d.simetrica) ? '' : inverso,
      simetrica: (b ? b.sim : d.simetrica) ? 'Sí' : 'No',
      publica: b && b.fija ? 'No' : (d.publica === false ? 'No' : 'Sí'),
      orden: Number(d.orden) || (i > -1 ? _get(t, t.filas[i], 'orden') : (t.filas.length + 1) * 10),
      estado: _etiquetaEstado(b && b.sistema ? 'publicado' : (d.activo === false ? 'retirado' : 'publicado')),
      nota: _s(d.nota).slice(0, 300)
    };
    if (i > -1) {
      const f = t.filas[i], antes = {};
      Object.keys(fila).forEach(k => { antes[k] = _get(t, f, k); });
      _actualizar(t, i, fila);
      _historial(ss, CONFIG.SHEETS.TIPOS_REL, id, antes, fila);
    } else {
      _agregar(t, [Object.assign({ id_tipo: id, sistema: 'No' }, fila)]);
      _historial(ss, CONFIG.SHEETS.TIPOS_REL, id, {}, fila);
    }
    return tiposRelacionDatos();
  } finally { lock.releaseLock(); }
}

function uiTiposRelacion() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <style>
    body{margin:0}.caja{padding:12px}
    table{width:100%;border-collapse:collapse;font-size:13px}th,td{border-bottom:1px solid #e2e8f0;padding:5px 4px;text-align:left;vertical-align:top}
    th{background:#f1f5f9;font-size:12px}td input[type=text]{width:100%;box-sizing:border-box;padding:5px;font-size:13px}
    td input[type=number]{width:56px;padding:4px}.gris{color:#94a3b8}.ret td{opacity:.55}
    .chip{display:inline-block;border-radius:10px;padding:0 7px;font-size:11px;background:#f1f5f9}
  </style>
  <div class="caja">
    <p style="margin-top:0">${txh('rel.tipos.intro')}</p>
    <div id="lista">${txh('comun.cargando')}</div>
    <div id="nuevo"></div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null;
    const err = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    function pintar(d) {
      D = d;
      const ro = !D.soyAdmin ? ' disabled' : '';
      $('lista').innerHTML = (D.normalizadas ? '<div class="ok">' + th('rel.tipos.normalizadas', {n: D.normalizadas}) + '</div>' : '') +
        '<table><thead><tr><th>' + th('rel.tipos.nombre') + '</th><th>' + th('rel.tipos.inverso') + '</th><th>' + th('rel.tipos.simetrica') + '</th><th>' + th('rel.tipos.publica') +
        '</th><th>' + th('rel.tipos.orden') + '</th><th>' + th('rel.tipos.usos') + '</th><th></th></tr></thead><tbody>' +
        D.tipos.map((t, i) => '<tr data-i="' + i + '"' + (t.activo ? '' : ' class="ret"') + '>' +
          '<td><input type="text" data-k="nombre" value="' + esc(t.nombre) + '" placeholder="' + esc(t.autoNombre || t.id) + '" aria-label="' + esc(tt('rel.tipos.nombre')) + '"' + ro + '>' +
            '<input type="text" data-k="nota" value="' + esc(t.nota) + '" placeholder="' + esc(tt('rel.tipos.nota')) + '" style="margin-top:3px;font-size:11px"' + ro + '>' +
            (t.sistema ? '<small class="gris">' + th(t.fija ? 'rel.tipos.fija' : 'rel.tipos.sistema') + '</small>' : '') + '</td>' +
          '<td>' + (t.simetrica ? '<small class="gris">' + th('rel.tipos.mismoNombre') + '</small>' : '<input type="text" data-k="nombre_inverso" value="' + esc(t.inverso) + '" placeholder="' + esc(t.autoInverso || tt('rel.tipos.inversoPh')) + '" aria-label="' + esc(tt('rel.tipos.inverso')) + '"' + ro + '>') + '</td>' +
          '<td><input type="checkbox" data-k="simetrica"' + (t.simetrica ? ' checked' : '') + (t.base || ro ? ' disabled' : '') + ' aria-label="' + esc(tt('rel.tipos.simetrica')) + '"></td>' +
          '<td><input type="checkbox" data-k="publica"' + (t.publica ? ' checked' : '') + (t.fija || ro ? ' disabled' : '') + ' aria-label="' + esc(tt('rel.tipos.publica')) + '"></td>' +
          '<td><input type="number" data-k="orden" value="' + esc(t.orden) + '"' + ro + ' aria-label="' + esc(tt('rel.tipos.orden')) + '"></td>' +
          '<td><span class="chip">' + t.usos + '</span></td>' +
          '<td>' + (D.soyAdmin ? '<button data-g="' + i + '">💾</button>' + (t.sistema ? '' : '<button data-a="' + i + '">' + th(t.activo ? 'rel.tipos.retirar' : 'rel.tipos.reactivar') + '</button>') : '') + '</td></tr>').join('') +
        '</tbody></table><p><small>' + th('rel.tipos.ayuda') + '</small></p>';
      $('nuevo').innerHTML = D.soyAdmin ? '<div class="tr" style="border:1px solid #e2e8f0;border-radius:8px;padding:8px;margin-top:8px"><b>＋ ' + th('rel.tipos.nuevo') + '</b>' +
        '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:4px"><input type="text" id="nNombre" placeholder="' + esc(tt('rel.tipos.nuevoNombre')) + '" style="flex:1;min-width:160px">' +
        '<input type="text" id="nInv" placeholder="' + esc(tt('rel.tipos.nuevoInverso')) + '" style="flex:1;min-width:160px">' +
        '<label style="font-weight:normal;display:inline-flex;gap:4px;align-items:center"><input type="checkbox" id="nSim"> ' + th('rel.tipos.simetrica') + '</label>' +
        '<label style="font-weight:normal;display:inline-flex;gap:4px;align-items:center"><input type="checkbox" id="nPub" checked> ' + th('rel.tipos.publica') + '</label>' +
        '<button class="prim" id="nGuardar">' + th('rel.tipos.agregar') + '</button></div><small>' + th('rel.tipos.nuevo.d') + '</small></div>' : '<p><small>' + th('rel.tipos.soloVer') + '</small></p>';
      if ($('nSim')) $('nSim').onchange = () => { $('nInv').disabled = $('nSim').checked; };
      if ($('nGuardar')) $('nGuardar').onclick = () => guardar({ id: '', nombre: $('nNombre').value, nombre_inverso: $('nInv').value, simetrica: $('nSim').checked, publica: $('nPub').checked, activo: true });
    }
    function fila(i) {
      const tr = document.querySelector('tr[data-i="' + i + '"]'), t = D.tipos[i], v = k => { const e = tr.querySelector('[data-k="' + k + '"]'); return e ? (e.type === 'checkbox' ? e.checked : e.value) : null; };
      return { id: t.id, nombre: v('nombre'), nombre_inverso: v('nombre_inverso') || '', simetrica: v('simetrica'), publica: v('publica'), orden: v('orden'), nota: v('nota'), activo: t.activo };
    }
    function guardar(d) {
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(err).withSuccessHandler(r => { pintar(r); $('msg').innerHTML = '<div class="ok">' + th('rel.tipos.guardado') + '</div>'; }).tipoRelacionGuardar(d);
    }
    $('lista').addEventListener('click', e => {
      const g = e.target.closest('[data-g]'), a = e.target.closest('[data-a]');
      if (g) guardar(fila(+g.dataset.g));
      if (a) { const d = fila(+a.dataset.a); d.activo = !D.tipos[+a.dataset.a].activo; guardar(d); }
    });
    google.script.run.withFailureHandler(err).withSuccessHandler(pintar).tiposRelacionDatos();
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(900).setHeight(640), tx('rel.tipos.titulo'));
}

// ─────────────────────────────────────────────────────────────────────────────
// 25. PROPUESTAS DE LA COMUNIDAD Y GALERÍAS DE IMÁGENES
// ─────────────────────────────────────────────────────────────────────────────
// Docentes, hablantes y familias proponen desde el visor: palabras, descripciones, correcciones,
// «así se dice en mi comunidad», imágenes y fotos de hojas trabajadas en clase. TODO llega como
// pendiente a la hoja PROPUESTAS y solo entra al catálogo cuando el equipo lo acepta en 📥 Propuestas.
// El crédito usa el registro de actores y fuentes (con el consentimiento que dio quien propone);
// el contacto nunca se publica. Las imágenes extra de cada palabra forman su galería (hoja IMAGENES);
// la principal sigue en PALABRAS, como siempre, para que juegos, lotería y AppSheet no cambien.

const COLS_PROPUESTAS = ['id_propuesta', 'codigo', 'fecha', 'tipo', 'estado', 'prioridad', 'id_palabra', 'id_idioma', 'texto', 'significado', 'id_idioma_significado',
  'descripcion', 'variante', 'categoria', 'nota', 'archivos', 'tipo_imagen', 'licencia', 'credito_nombre', 'credito_publico', 'rol', 'contacto', 'recopilo',
  'actividad', 'compartir', 'dispositivo', 'origen', 'respuesta', 'nota_interna', 'resuelta_por', 'fecha_resolucion', 'resultado'];
const COLS_IMAGENES = ['id_imagen', 'id_palabra', 'archivo', 'tipo', 'credito', 'licencia', 'alt', 'orden', 'estado', 'solo_interna', 'origen', 'nota', 'creado_por', 'fecha'];
const TIPOS_PROPUESTA = ['palabra', 'descripcion', 'correccion', 'variante', 'imagen', 'fotos', 'retiro'];
const ESTADOS_PROPUESTA = ['Pendiente', 'Aceptada', 'No incluida', 'Transcrita'];
const TIPOS_IMAGEN = ['dibujo', 'foto', 'ilustracion', 'otro'];
const LICENCIAS_APORTE = ['CC BY 4.0', 'CC BY-NC 4.0', 'Solo este catálogo'];
const ROLES_APORTE = ['docente', 'hablante', 'familia', 'otro'];
const BASE_APORTE = 'Aportación con consentimiento';
const _PROP_DEF = () => ({ activo: false, quien: 'todos', clave: '', maxImagenes: 7, maxPendDesc: 3, maxPendCorr: 3, porDia: 30, criterios: '', mostrarAportes: true, novedades: true });

function _propuestasAjustes(ss) { return Object.assign(_PROP_DEF(), _visorAjustes(ss).propuestas || {}); }

/** Lo que el visor necesita saber (va en el estado publicado): nunca la clave. */
function _propuestasPublicas(ss) {
  const p = _propuestasAjustes(ss);
  return { activo: !!p.activo, quien: p.quien === 'clave' ? 'clave' : 'todos', criterios: _s(p.criterios).slice(0, 1500), maxImagenes: Number(p.maxImagenes) || 7 };
}

function _hojaPropuestas(ss) {
  let t = _tabla(ss, CONFIG.SHEETS.PROPUESTAS);
  if (!t) { asegurarEstructura(ss); t = _tabla(ss, CONFIG.SHEETS.PROPUESTAS); }
  return t;
}
function _hojaImagenes(ss) {
  let t = _tabla(ss, CONFIG.SHEETS.IMAGENES);
  if (!t) { asegurarEstructura(ss); t = _tabla(ss, CONFIG.SHEETS.IMAGENES); }
  return t;
}
function _codigoPropuesta() { const A = 'ACDEFGHJKMNPQRTUVWXY3479'; let c = 'P-'; for (let i = 0; i < 5; i++) c += A[Math.floor(Math.random() * A.length)]; return c; }

/** Cuántas imágenes tiene (o tendrá) una palabra: la principal, la galería y las propuestas pendientes. */
function _conteoImagenes(ss) {
  const n = {};
  readSheet(ss, CONFIG.SHEETS.PALABRAS).forEach(p => { if (_s(p.imagen_referencia)) n[_k(p.id_palabra)] = 1; });
  if (ss.getSheetByName(CONFIG.SHEETS.IMAGENES)) readSheet(ss, CONFIG.SHEETS.IMAGENES).forEach(i => {
    if (_s(i.id_palabra) && _estadoCanonico(i.estado) !== 'retirado') n[_k(i.id_palabra)] = (n[_k(i.id_palabra)] || 0) + 1;
  });
  if (ss.getSheetByName(CONFIG.SHEETS.PROPUESTAS)) readSheet(ss, CONFIG.SHEETS.PROPUESTAS).forEach(p => {
    if (_k(p.tipo) === 'imagen' && _s(p.estado) === 'Pendiente') n[_k(p.id_palabra)] = (n[_k(p.id_palabra)] || 0) + 1;
  });
  return n;
}
/** Máximo de imágenes de una palabra: el de la palabra (columna imagenes_max; 0 = cerrada) o el general. */
function _maxImagenes(palabra, general) {
  const v = _s(palabra && palabra.imagenes_max);
  return v === '' ? general : Math.max(0, Number(v) || 0);
}

/** Datos de imagen (data:…;base64) → archivo privado en la carpeta de propuestas. */
function _guardarImagenPropuesta(ss, dataUrl, nombre) {
  const m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/.exec(_s(dataUrl));
  if (!m) throw new Error('imagen');
  if (m[2].length > 2.8 * 1024 * 1024) throw new Error('grande');   // ~2 MB
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nombre + (m[1] === 'image/png' ? '.png' : m[1] === 'image/webp' ? '.webp' : '.jpg'));
  const carpeta = _subcarpeta(_carpetaDelLibro(ss), 'PROPUESTAS_IMAGENES');
  return carpeta.createFile(blob).getId();
}

/**
 * Recibe propuestas del visor (doPost, accion «proponer»). d = { dispositivo, colab, web (trampa), quien: { nombre, publico,
 * rol, contacto, recopilo, compartir, acepto }, items: [ { tipo, id_palabra, id_idioma, texto, significado, id_idioma_significado,
 * descripcion, variante, categoria, nota, actividad, imagen, tipo_imagen, licencia, fotos: [] } ] }
 */
function _recibirPropuestas(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aj = _propuestasAjustes(ss);
  if (!aj.activo) return { status: 403, error: 'cerradas' };
  if (_s(d.web)) return { status: 200, codigo: _codigoPropuesta(), resultados: [] };          // trampa: se finge éxito y no se guarda nada
  const prioridad = _s(aj.clave) && _s(d.colab) === _s(aj.clave) ? 'alta' : '';   // con clave de colaboración: se revisa primero (nunca se publica sola)
  if (aj.quien === 'clave' && (!_s(aj.clave) || _s(d.colab) !== _s(aj.clave))) return { status: 403, error: 'clave' };
  const q = d.quien || {};
  if (!q.acepto) return { status: 400, error: 'consentimiento' };
  const items = (Array.isArray(d.items) ? d.items : []).slice(0, 40);
  if (!items.length) return { status: 400, error: 'vacia' };
  const disp = _s(d.dispositivo).replace(/[^\w-]/g, '').slice(0, 40) || 'sin';
  // Límites por día: por dispositivo y en total (protege la hoja y la cuenta de Drive)
  const cache = CacheService.getScriptCache(), hoy = _hoyISO(), kd = `PROP_${hoy}_${disp}`, kt = `PROP_${hoy}_TOTAL`;
  const usadosD = Number(cache.get(kd) || 0), usadosT = Number(cache.get(kt) || 0);
  if (usadosD + items.length > (Number(aj.porDia) || 30) || usadosT + items.length > 600) return { status: 429, error: 'limite' };
  const lock = _lock();
  try {
    const t = _hojaPropuestas(ss);
    const pal = {}; readSheet(ss, CONFIG.SHEETS.PALABRAS).forEach(p => { if (_s(p.id_palabra)) pal[_k(p.id_palabra)] = p; });
    const lenguas = new Set(readSheet(ss, CONFIG.SHEETS.IDIOMAS).map(i => _k(i.id_idioma)).filter(Boolean));
    const imgs = _conteoImagenes(ss), general = Number(aj.maxImagenes) || 7;
    const pend = {};
    t.filas.forEach(f => { if (_s(_get(t, f, 'estado')) === 'Pendiente') { const k = _k(_get(t, f, 'tipo')) + '|' + _k(_get(t, f, 'id_palabra')) + '|' + _k(_get(t, f, 'id_idioma')); pend[k] = (pend[k] || 0) + 1; } });
    const codigo = _codigoPropuesta(), fecha = _ahora(), filas = [], resultados = [];
    const txt = (v, n) => _s(v).replace(/\u0000/g, '').slice(0, n);
    const comun = { codigo, fecha, estado: 'Pendiente', prioridad, credito_nombre: txt(q.nombre, 120), credito_publico: q.publico ? 'Sí' : 'No',
      rol: ROLES_APORTE.includes(_k(q.rol)) ? _k(q.rol) : 'otro', contacto: txt(q.contacto, 160), recopilo: txt(q.recopilo, 120),
      compartir: q.compartir ? 'Sí' : 'No', dispositivo: disp, origen: 'visor' };
    items.forEach((it, i) => {
      const tipo = TIPOS_PROPUESTA.includes(_k(it.tipo)) ? _k(it.tipo) : '';
      const idP = _s(it.id_palabra), L = _s(it.id_idioma);
      const err = e => resultados.push({ i, ok: false, error: e });
      if (!tipo) return err('tipo');
      if (idP && !pal[_k(idP)]) return err('palabra');
      if (L && !lenguas.has(_k(L))) return err('lengua');
      const fila = Object.assign({ id_propuesta: Utilities.getUuid(), tipo, id_palabra: idP, id_idioma: L, texto: txt(it.texto, 300), significado: txt(it.significado, 300),
        id_idioma_significado: lenguas.has(_k(it.id_idioma_significado)) ? _s(it.id_idioma_significado) : '', descripcion: txt(it.descripcion, LARGO_DESCRIPCION),
        variante: txt(it.variante, 200), categoria: txt(it.categoria, 120), nota: txt(it.nota, 500), actividad: txt(it.actividad, 120) }, comun);
      if (tipo === 'palabra' && !fila.texto) return err('vacia');
      if (['descripcion', 'correccion', 'variante', 'imagen'].includes(tipo) && !idP) return err('palabra');
      if (tipo === 'descripcion' && (!fila.descripcion || !L)) return err('vacia');
      if (tipo === 'correccion' && !fila.texto && !fila.nota) return err('vacia');
      const kp = tipo + '|' + _k(idP) + '|' + _k(L);
      if (tipo === 'descripcion' && (pend[kp] || 0) >= (Number(aj.maxPendDesc) || 3)) return err('completa');
      if (tipo === 'correccion' && (pend[kp] || 0) >= (Number(aj.maxPendCorr) || 3)) return err('completa');
      if (tipo === 'imagen') {
        const max = _maxImagenes(pal[_k(idP)], general);
        if ((imgs[_k(idP)] || 0) >= max) return err('completa');        // el cupo se revisa también aquí (pudo llenarse mientras tanto)
        try { fila.archivos = _guardarImagenPropuesta(ss, it.imagen, `${codigo}_${i}`); } catch (e) { return err(e.message === 'grande' ? 'grande' : 'imagen'); }
        fila.tipo_imagen = TIPOS_IMAGEN.includes(_k(it.tipo_imagen)) ? _k(it.tipo_imagen) : 'otro';
        fila.licencia = LICENCIAS_APORTE.includes(_s(it.licencia)) ? _s(it.licencia) : LICENCIAS_APORTE[2];
        imgs[_k(idP)] = (imgs[_k(idP)] || 0) + 1;
      }
      if (tipo === 'fotos') {
        const ids = [];
        (Array.isArray(it.fotos) ? it.fotos : []).slice(0, 8).forEach((f, j) => { try { ids.push(_guardarImagenPropuesta(ss, f, `${codigo}_hoja_${j}`)); } catch (e) {} });
        if (!ids.length) return err('imagen');
        fila.archivos = ids.join(',');
      }
      pend[kp] = (pend[kp] || 0) + 1;
      filas.push(fila); resultados.push({ i, ok: true });
    });
    if (filas.length) _agregar(t, filas);
    cache.put(kd, String(usadosD + filas.length), 86400); cache.put(kt, String(usadosT + filas.length), 86400);
    return { status: 200, codigo: filas.length ? codigo : '', resultados };
  } finally { lock.releaseLock(); }
}

/** Estado de propuestas por código (para «Mis propuestas»): sin nombres ni contactos. */
function _estadoPropuestas(codigos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quiero = new Set(_s(codigos).toUpperCase().split(',').map(_s).filter(c => /^P-[A-Z0-9]{5}$/.test(c)).slice(0, 30));
  if (!quiero.size || !ss.getSheetByName(CONFIG.SHEETS.PROPUESTAS)) return { status: 200, propuestas: [] };
  const r = {};
  readSheet(ss, CONFIG.SHEETS.PROPUESTAS).forEach(p => {
    const c = _s(p.codigo).toUpperCase(); if (!quiero.has(c)) return;
    (r[c] = r[c] || []).push({ tipo: _k(p.tipo), texto: _s(p.texto || p.descripcion).slice(0, 60), estado: _s(p.estado), respuesta: _s(p.respuesta).slice(0, 500) });
  });
  return { status: 200, propuestas: Object.keys(r).map(c => ({ codigo: c, items: r[c] })) };
}

// ── Créditos: quien aporta queda como actor (si lo permitió) y como fuente con su consentimiento ──

/** Fuente de una aportación: una por persona (y por quien recopiló), reutilizada en todo lo que aporte. */
function _fuenteAportacion(ss, p) {
  const nombre = _s(p.credito_nombre), publico = _s(p.credito_publico) === 'Sí', recopilo = _s(p.recopilo);
  let idActor = '';
  if (nombre) idActor = _actorPorNombre(ss, nombre, true).id;
  let tFu = _tabla(ss, CONFIG.SHEETS.FUENTES);
  if (!tFu) { asegurarEstructura(ss); tFu = _tabla(ss, CONFIG.SHEETS.FUENTES); }
  const titulo = nombre ? tx('prop.fuente.de', { nombre }) + (recopilo ? ' · ' + tx('prop.fuente.recopilo', { nombre: recopilo }) : '') : tx('prop.fuente.anonima') + (recopilo ? ' · ' + tx('prop.fuente.recopilo', { nombre: recopilo }) : '');
  const compartir = _s(p.compartir) === 'Sí' ? 'Sí' : 'No';
  const i = tFu.filas.findIndex(f => _s(_get(tFu, f, 'titulo')) === titulo && _s(_get(tFu, f, 'base_uso')) === BASE_APORTE && _s(_get(tFu, f, 'uso_compartir')) === compartir
    && _s(_get(tFu, f, 'publicar_credito')) === (publico ? 'Sí' : 'No'));
  if (i > -1) return { idFuente: _s(_get(tFu, tFu.filas[i], 'id_fuente')), idActor: publico ? idActor : '' };
  const id = Utilities.getUuid();
  _agregar(tFu, [{ id_fuente: id, titulo, autoria: nombre, tipo: 'Otro', base_uso: BASE_APORTE, credito: publico ? [nombre, recopilo ? tx('prop.fuente.recopilo', { nombre: recopilo }) : ''].filter(Boolean).join(' · ') : '',
    publicar_credito: publico ? 'Sí' : 'No', uso_publicar: 'Sí', uso_didactico: 'Sí', uso_adaptar: 'Sí', uso_compartir: compartir, uso_ia: 'Pendiente',
    condiciones: tx('prop.fuente.condiciones'), notas: tx('prop.fuente.nota', { codigo: _s(p.codigo) }), estado: _etiquetaEstado('publicado') }]);
  return { idFuente: id, idActor: publico ? idActor : '' };
}

// ── Revisión (📥 Propuestas) ──

function propuestasDatos(verTodas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const t = _hojaPropuestas(ss);
  const idiomas = readSheet(ss, CONFIG.SHEETS.IDIOMAS).filter(i => _s(i.id_idioma)).map(i => ({ id: _s(i.id_idioma), nombre: _s(i.nombre_propio || i.nombre_completo) || _s(i.codigo_iso) }));
  const nomL = {}; idiomas.forEach(i => { nomL[_k(i.id)] = i.nombre; });
  const tTr = _tabla(ss, CONFIG.SHEETS.TRADUCCIONES), textos = {}, principal = _k(getAppConfig(ss).idiomaPrincipal);
  const porForma = {};
  tTr.filas.forEach(f => {
    if (_estadoCanonico(_get(tTr, f, 'estado')) === 'retirado') return;
    const p = _k(_get(tTr, f, 'id_palabra')), L = _k(_get(tTr, f, 'id_idioma')), tx0 = _s(_get(tTr, f, 'texto'));
    const o = textos[p] = textos[p] || {}; o[L] = { texto: tx0, descripcion: _s(_get(tTr, f, 'descripcion')) };
    _formas(tx0).forEach(fo => { const k = L + '|' + _claveForma(fo); (porForma[k] = porForma[k] || []).push(p); });
  });
  const refPal = id => { const o = textos[_k(id)] || {}; return (o[principal] || Object.values(o)[0] || {}).texto || id; };
  const pal = {}; readSheet(ss, CONFIG.SHEETS.PALABRAS).forEach(p => { pal[_k(p.id_palabra)] = p; });
  const imgs = _conteoImagenes(ss), aj = _propuestasAjustes(ss);
  const lista = t.filas.map(f => { const o = {}; t.headers.forEach((h, i) => { if (h) o[h] = f[i] instanceof Date ? f[i].toISOString() : _s(f[i]); }); return o; })
    .filter(p => p.id_propuesta && (verTodas || p.estado === 'Pendiente'))
    .map(p => Object.assign(p, {
      lengua: nomL[_k(p.id_idioma)] || '', lenguaSig: nomL[_k(p.id_idioma_significado)] || '',
      palabraTexto: p.id_palabra ? refPal(p.id_palabra) : '',
      actual: p.id_palabra && p.id_idioma ? ((textos[_k(p.id_palabra)] || {})[_k(p.id_idioma)] || null) : null,
      coincide: p.tipo === 'palabra' && p.id_idioma ? [...new Set(_formas(p.texto).flatMap(fo => porForma[_k(p.id_idioma) + '|' + _claveForma(fo)] || []))].slice(0, 5).map(id => ({ id, texto: refPal(id) })) : [],
      cupo: p.tipo === 'imagen' ? { n: imgs[_k(p.id_palabra)] || 0, max: _maxImagenes(pal[_k(p.id_palabra)], Number(aj.maxImagenes) || 7), tienePrincipal: !!_s((pal[_k(p.id_palabra)] || {}).imagen_referencia) } : null
    }))
    .sort((a, b) => (a.prioridad === 'alta' ? 0 : 1) - (b.prioridad === 'alta' ? 0 : 1) || String(a.fecha).localeCompare(String(b.fecha)));
  const palabras = Object.keys(textos).map(id => ({ id, texto: refPal(id) })).sort((a, b) => a.texto.localeCompare(b.texto, 'es'));
  const categorias = readSheet(ss, CONFIG.SHEETS.CATEGORIAS).filter(c => _s(c.id_categoria)).map(c => ({ id: _s(c.id_categoria), nombre: _s(c.nombre_categoria) }));
  return { propuestas: lista, idiomas, palabras, categorias, pendientes: t.filas.filter(f => _s(_get(t, f, 'estado')) === 'Pendiente').length };
}

/**
 * Resuelve una propuesta. accion: 'nueva' | 'sumar' | 'descripcion' | 'correccion' | 'imagen' | 'variante' | 'transcrita' | 'retirarCredito' | 'no'
 * o = { respuesta, nota_interna, publicar, id_destino, texto, descripcion, id_categoria, principal, alt, borrarFotos, modo }
 */
function propuestaResolver(id, accion, o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  o = o || {};
  const lock = _lock();
  try {
    const t = _hojaPropuestas(ss);
    const i = t.filas.findIndex(f => _k(_get(t, f, 'id_propuesta')) === _k(id));
    if (i < 0) throw new Error(tx('prop.error.noExiste'));
    const p = {}; t.headers.forEach((h, j) => { if (h) p[h] = _s(t.filas[i][j]); });
    if (p.estado !== 'Pendiente' && accion !== 'retirarCredito') throw new Error(tx('prop.error.yaResuelta'));
    const reglas = _cargarReglas(ss), yo = _correoActual();
    const estadoNuevo = _etiquetaEstado(o.publicar === false ? 'revisión' : 'publicado');
    let resultado = '', estado = 'Aceptada';
    const credito = () => _fuenteAportacion(ss, p);
    const tTr = () => { const sh = ss.getSheetByName(CONFIG.SHEETS.TRADUCCIONES); _asegurarColumnas(sh, ['descripcion']); return _tablaConEstado(ss, CONFIG.SHEETS.TRADUCCIONES); };
    const filaTr = (tt, idP, L) => tt.filas.findIndex(f => _k(_get(tt, f, 'id_palabra')) === _k(idP) && _k(_get(tt, f, 'id_idioma')) === _k(L));
    const norm = (v, L) => _normalizar(_s(v), L, reglas, yo, false);
    if (accion === 'no') estado = 'No incluida';
    else if (accion === 'nueva') {
      const c = credito(), tP = _tablaConEstado(ss, CONFIG.SHEETS.PALABRAS), tt = tTr();
      const idP = Utilities.getUuid();
      _agregar(tP, [{ id_palabra: idP, id_categoria: _s(o.id_categoria), estado: estadoNuevo, usar_en_juegos: 'Sí', creado_por: yo, origen: tx('prop.origen', { codigo: p.codigo }) }]);
      const nuevas = [{ id_traduccion: Utilities.getUuid(), id_palabra: idP, id_idioma: p.id_idioma, texto: norm(o.texto || p.texto, p.id_idioma), texto_captura: _s(o.texto || p.texto),
        nota_variante: p.variante ? tx('prop.seDiceEn', { lugar: p.variante }) : '', descripcion: _s(o.descripcion != null ? o.descripcion : p.descripcion).slice(0, LARGO_DESCRIPCION),
        id_fuente: c.idFuente, estado: estadoNuevo, creado_por: yo, origen: tx('prop.origen', { codigo: p.codigo }) }];
      if (p.significado && p.id_idioma_significado && _k(p.id_idioma_significado) !== _k(p.id_idioma))
        nuevas.push({ id_traduccion: Utilities.getUuid(), id_palabra: idP, id_idioma: p.id_idioma_significado, texto: norm(p.significado, p.id_idioma_significado), texto_captura: p.significado,
          id_fuente: c.idFuente, estado: estadoNuevo, creado_por: yo, origen: tx('prop.origen', { codigo: p.codigo }) });
      _agregar(tt, nuevas);
      _historial(ss, CONFIG.SHEETS.PALABRAS, idP, {}, { propuesta: p.codigo });
      resultado = idP;
    } else if (accion === 'sumar' || accion === 'variante' || accion === 'correccion') {
      const idP = _s(o.id_destino || p.id_palabra); if (!idP) throw new Error(tx('rel.error.palabra'));
      const tt = tTr(), L = p.id_idioma, k = filaTr(tt, idP, L), c = credito();
      const forma = norm(o.texto || p.texto, L);
      if (k < 0) {
        if (!forma) throw new Error(tx('prop.error.sinTexto'));
        const fila = { id_traduccion: Utilities.getUuid(), id_palabra: idP, id_idioma: L, texto: forma, texto_captura: _s(o.texto || p.texto), id_fuente: c.idFuente,
          nota_variante: p.variante ? tx('prop.seDiceEn', { lugar: p.variante }) : '', estado: estadoNuevo, creado_por: yo, origen: tx('prop.origen', { codigo: p.codigo }) };
        _agregar(tt, [fila]); resultado = fila.id_traduccion;
      } else {
        const f = tt.filas[k], antes = { texto: _get(tt, f, 'texto'), nota_variante: _get(tt, f, 'nota_variante') }, c2 = {};
        if (accion === 'correccion' && o.modo === 'reemplazar' && forma) { c2.texto = forma; if (_col(tt, 'texto_captura') > -1) c2.texto_captura = _s(o.texto || p.texto); }
        else if (forma && !_formas(_s(_get(tt, f, 'texto'))).some(x => _claveForma(x) === _claveForma(forma))) c2.texto = _s(_get(tt, f, 'texto')) + ', ' + forma;
        const nota = p.variante ? tx('prop.seDiceEn', { lugar: p.variante }) + (forma ? ': ' + forma : '') : (accion === 'correccion' ? _s(p.nota) : '');
        if (nota && !_s(_get(tt, f, 'nota_variante')).includes(nota)) c2.nota_variante = [_s(_get(tt, f, 'nota_variante')), nota].filter(Boolean).join(' · ');
        if (Object.keys(c2).length) { _actualizar(tt, k, c2); _historial(ss, CONFIG.SHEETS.TRADUCCIONES, _s(_get(tt, f, 'id_traduccion')), antes, Object.assign({ propuesta: p.codigo }, c2)); }
        resultado = _s(_get(tt, f, 'id_traduccion'));
      }
    } else if (accion === 'descripcion') {
      const tt = tTr(), k = filaTr(tt, p.id_palabra, p.id_idioma);
      if (k < 0) throw new Error(tx('prop.error.sinTraduccion'));
      const f = tt.filas[k], nueva = norm(o.descripcion != null ? o.descripcion : p.descripcion, p.id_idioma).slice(0, LARGO_DESCRIPCION);
      const antes = { descripcion: _get(tt, f, 'descripcion') };
      const val = o.modo === 'agregar' && _s(antes.descripcion) ? (_s(antes.descripcion) + ' ' + nueva).slice(0, LARGO_DESCRIPCION) : nueva;
      _actualizar(tt, k, { descripcion: val });
      _historial(ss, CONFIG.SHEETS.TRADUCCIONES, _s(_get(tt, f, 'id_traduccion')), antes, { descripcion: val, propuesta: p.codigo });
      resultado = _s(_get(tt, f, 'id_traduccion'));
    } else if (accion === 'imagen') {
      const c = credito(), archivo = _s(p.archivos);
      const carpeta = _carpetaImagenes(ss), file = DriveApp.getFileById(archivo);
      try { file.moveTo(carpeta); } catch (e) { /* se queda donde está */ }
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      const tP = _tabla(ss, CONFIG.SHEETS.PALABRAS), ip = tP.filas.findIndex(f => _k(_get(tP, f, 'id_palabra')) === _k(p.id_palabra));
      if (ip < 0) throw new Error(tx('medios.error.palabra'));
      const sinPrincipal = !_s(_get(tP, tP.filas[ip], 'imagen_referencia'));
      const fila = { id_imagen: Utilities.getUuid(), id_palabra: p.id_palabra, archivo, tipo: p.tipo_imagen, credito: c.idActor, licencia: p.licencia, alt: _s(o.alt).slice(0, 300),
        orden: Date.now() % 1e9, estado: _etiquetaEstado('publicado'), solo_interna: o.soloInterna ? 'Sí' : 'No', origen: tx('prop.origen', { codigo: p.codigo }), creado_por: yo, fecha: _ahora() };
      if ((o.principal || sinPrincipal) && !o.soloInterna) { _hacerPrincipal(ss, p.id_palabra, fila); resultado = 'principal:' + archivo; }
      else { _agregar(_hojaImagenes(ss), [fila]); resultado = fila.id_imagen; }
    } else if (accion === 'transcrita') {
      estado = 'Transcrita';
      if (o.borrarFotos) _s(p.archivos).split(',').map(_s).filter(Boolean).forEach(a => { try { DriveApp.getFileById(a).setTrashed(true); } catch (e) {} });   // quedan 30 días en la papelera de Drive
    } else if (accion === 'retirarCredito') {
      // Quien aportó pide quitar su nombre: sus fuentes dejan de mostrar crédito (no se borra nada)
      const tFu = _tabla(ss, CONFIG.SHEETS.FUENTES), nombre = _s(o.nombre || p.credito_nombre);
      let n = 0;
      if (tFu && nombre) tFu.filas.forEach((f, j) => { if (_s(_get(tFu, f, 'base_uso')) === BASE_APORTE && _plegar(_get(tFu, f, 'autoria')) === _plegar(nombre) && _s(_get(tFu, f, 'publicar_credito')) !== 'No') { _actualizar(tFu, j, { publicar_credito: 'No' }); n++; } });
      const tImg = ss.getSheetByName(CONFIG.SHEETS.IMAGENES) ? _tabla(ss, CONFIG.SHEETS.IMAGENES) : null, a = _actorPorNombre(ss, nombre, false).id;
      if (tImg && a) tImg.filas.forEach((f, j) => { if (_k(_get(tImg, f, 'credito')) === _k(a)) _actualizar(tImg, j, { credito: '' }); });
      const tP = _tabla(ss, CONFIG.SHEETS.PALABRAS);
      if (a) tP.filas.forEach((f, j) => { if (_k(_get(tP, f, 'imagen_credito')) === _k(a)) _actualizar(tP, j, { imagen_credito: '' }); });
      _historial(ss, CONFIG.SHEETS.PROPUESTAS, p.codigo, {}, { retiroCredito: nombre, fuentes: n });
      resultado = 'credito:' + n;
    } else throw new Error(tx('prop.error.accion'));
    const cambios = { estado, respuesta: _s(o.respuesta).slice(0, 500), nota_interna: _s(o.nota_interna).slice(0, 500), resuelta_por: yo, fecha_resolucion: _ahora(), resultado };
    if (accion === 'retirarCredito') delete cambios.estado;
    _actualizar(t, i, cambios);
    if (estado === 'Aceptada' && _propuestasAjustes(ss).novedades) { try { _avisoNovedades(ss); } catch (e) {} }
    return propuestasDatos(!!o.verTodas);
  } finally { lock.releaseLock(); }
}

// ── 📜 Constancia de colaboración (para quien aportó; sale del registro de propuestas aceptadas) ──

function constanciaDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const n = {};
  if (ss.getSheetByName(CONFIG.SHEETS.PROPUESTAS)) readSheet(ss, CONFIG.SHEETS.PROPUESTAS).forEach(p => {
    if (_s(p.estado) !== 'Aceptada' && _s(p.estado) !== 'Transcrita') return;
    [_s(p.credito_nombre), _s(p.recopilo)].filter(Boolean).forEach((nom, j) => {
      const k = _plegar(nom); n[k] = n[k] || { nombre: nom, total: 0, porTipo: {}, primera: '', ultima: '', recopilo: 0 };
      const o = n[k]; o.total++; if (j === 1) o.recopilo++;
      o.porTipo[_k(p.tipo)] = (o.porTipo[_k(p.tipo)] || 0) + 1;
      const f = String(p.fecha_resolucion || p.fecha).slice(0, 10);
      if (!o.primera || f < o.primera) o.primera = f; if (!o.ultima || f > o.ultima) o.ultima = f;
    });
  });
  const cfg = getAppConfig(ss);
  return { personas: Object.values(n).sort((a, b) => b.total - a.total), titulo: cfg.titulo || ss.getName(), hoy: _hoyISO(), firmantes: LS_CONSTANCIA(ss) };
}
function LS_CONSTANCIA(ss) { try { return PropertiesService.getDocumentProperties().getProperty('CONSTANCIA_FIRMAS') || ''; } catch (e) { return ''; } }
function constanciaGuardarFirmas(t) { _requerirAdmin(SpreadsheetApp.getActiveSpreadsheet()); PropertiesService.getDocumentProperties().setProperty('CONSTANCIA_FIRMAS', _s(t).slice(0, 500)); return true; }

function uiConstancia() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <style>body{margin:0}.caja{padding:12px}#doc{border:1px solid #cbd5e1;border-radius:8px;padding:36px 44px;margin-top:10px;background:#fff;font-family:Georgia,serif;font-size:15px;line-height:1.7}
    #doc h1{text-align:center;font-size:22px;margin:0 0 18px}#doc .firmas{display:flex;gap:40px;justify-content:center;margin-top:60px;flex-wrap:wrap}
    #doc .firma{border-top:1px solid #111;min-width:200px;text-align:center;padding-top:4px;font-size:13px;white-space:pre-line}#doc .pie{font-size:11px;color:#555;margin-top:30px;text-align:center}
    @media print{.no-imp{display:none!important}#doc{border:0;margin:0}body{background:#fff}}</style>
  <div class="caja">
    <div class="no-imp">
      <p><small>${txh('const.intro')}</small></p>
      <div class="dos"><div><label for="per">${txh('const.persona')}</label><select id="per"></select></div><div><label for="lug">${txh('const.lugar')}</label><input type="text" id="lug"></div></div>
      <div class="dos"><div><label for="de">${txh('const.desde')}</label><input type="date" id="de"></div><div><label for="a">${txh('const.hasta')}</label><input type="date" id="a"></div></div>
      <label for="fir">${txh('const.firmantes')}</label><textarea id="fir" style="height:60px" placeholder="${txh('const.firmantes.ph')}"></textarea>
      <label for="ext">${txh('const.texto')}</label><textarea id="ext" style="height:50px"></textarea>
      <button class="prim" onclick="armar()">${txh('const.ver')}</button> <button onclick="window.print()">🖨️ ${txh('const.imprimir')}</button>
      <div id="msg"></div>
    </div>
    <div id="doc"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null;
    google.script.run.withFailureHandler(e => $('msg').innerHTML = '<div class="error">' + esc(e.message) + '</div>').withSuccessHandler(d => {
      D = d;
      $('per').innerHTML = d.personas.length ? d.personas.map((p, i) => '<option value="' + i + '">' + esc(p.nombre) + ' (' + p.total + ')</option>').join('') : '<option value="">' + th('const.nadie') + '</option>';
      $('fir').value = d.firmantes; $('a').value = d.hoy;
      $('per').onchange = () => { const p = D.personas[+$('per').value]; if (p) { $('de').value = p.primera; $('a').value = p.ultima || D.hoy; } armar(); };
      $('per').onchange();
    }).constanciaDatos();
    function armar() {
      const p = D && D.personas[+$('per').value]; if (!p) { $('doc').innerHTML = '<p>' + th('const.nadie') + '</p>'; return; }
      google.script.run.constanciaGuardarFirmas($('fir').value);
      const tipos = Object.keys(p.porTipo).map(k => p.porTipo[k] + ' ' + tt('const.tipo.' + k)).join(', ');
      const firmas = $('fir').value.split(/\\n\\s*\\n|;/).map(x => x.trim()).filter(Boolean);
      $('doc').innerHTML = '<h1>' + th('const.titulo') + '</h1>' +
        '<p>' + th('const.cuerpo', {catalogo: D.titulo, nombre: p.nombre, n: p.total, desde: $('de').value, hasta: $('a').value}) + '</p>' +
        '<p>' + th('const.detalle', {tipos}) + (p.recopilo ? ' ' + th('const.recopilo', {n: p.recopilo}) : '') + '</p>' +
        ($('ext').value.trim() ? '<p>' + esc($('ext').value.trim()) + '</p>' : '') +
        '<p>' + th('const.agradece') + '</p>' +
        '<p style="text-align:right">' + esc($('lug').value || '') + (($('lug').value) ? ', ' : '') + esc(D.hoy) + '</p>' +
        '<div class="firmas">' + (firmas.length ? firmas : [tt('const.firmaEquipo', {catalogo: D.titulo})]).map(f => '<div class="firma">' + esc(f) + '</div>').join('') + '</div>' +
        '<div class="pie">' + th('const.pie', {catalogo: D.titulo}) + '</div>';
    }
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(900).setHeight(760), tx('const.titulo'));
}

/** Carpeta donde viven las imágenes públicas del catálogo (la configurada o una subcarpeta IMAGENES). */
function _carpetaImagenes(ss) {
  const cfg = getAppConfig(ss);
  try { if (_s(cfg.imagenes)) return DriveApp.getFolderById(_s(cfg.imagenes)); } catch (e) {}
  return _subcarpeta(_carpetaDelLibro(ss), 'IMAGENES');
}

/** La imagen elegida pasa a ser la principal; la principal anterior queda en la galería. */
function _hacerPrincipal(ss, idPalabra, fila) {
  const tP = _tabla(ss, CONFIG.SHEETS.PALABRAS);
  if (_col(tP, 'imagen_alt') < 0) { _asegurarColumnas(tP.sh, ['imagen_alt']); return _hacerPrincipal(ss, idPalabra, fila); }
  const ip = tP.filas.findIndex(f => _k(_get(tP, f, 'id_palabra')) === _k(idPalabra));
  const f = tP.filas[ip], tImg = _hojaImagenes(ss);
  const ant = { ref: _s(_get(tP, f, 'imagen_referencia')), id: _s(_get(tP, f, 'imagen_id')), credito: _s(_get(tP, f, 'imagen_credito')), alt: _s(_get(tP, f, 'imagen_alt')) };
  if (ant.ref) _agregar(tImg, [{ id_imagen: Utilities.getUuid(), id_palabra: idPalabra, archivo: ant.id || _idDrive(ant.ref) || ant.ref, tipo: 'otro', credito: ant.credito, alt: ant.alt,
    orden: Date.now() % 1e9, estado: _etiquetaEstado('publicado'), solo_interna: 'No', origen: tx('prop.antesPrincipal'), creado_por: _correoActual(), fecha: _ahora() }]);
  const c = { imagen_referencia: `https://drive.google.com/file/d/${fila.archivo}/view`, imagen_id: '', imagen_ref: '', imagen_credito: fila.credito || '', imagen_alt: fila.alt || '' };
  _actualizar(tP, ip, c);
  _historial(ss, CONFIG.SHEETS.PALABRAS, idPalabra, { imagen_referencia: ant.ref }, { imagen_referencia: c.imagen_referencia });
}

/** Galería de una palabra (para Editar y Propuestas): principal + imágenes extra. */
function galeriaDatos(idPalabra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = readSheet(ss, CONFIG.SHEETS.PALABRAS).find(x => _k(x.id_palabra) === _k(idPalabra));
  if (!p) throw new Error(tx('medios.error.palabra'));
  const act = _actores(ss);
  const extra = ss.getSheetByName(CONFIG.SHEETS.IMAGENES) ? readSheet(ss, CONFIG.SHEETS.IMAGENES).filter(i => _k(i.id_palabra) === _k(idPalabra) && _estadoCanonico(i.estado) !== 'retirado') : [];
  return { principal: _s(p.imagen_referencia) ? { archivo: _s(p.imagen_id) || _idDrive(p.imagen_referencia), credito: _creditoVisible(act, p.imagen_credito), alt: _s(p.imagen_alt) } : null,
    extra: extra.map(i => ({ id: _s(i.id_imagen), archivo: _s(i.archivo), tipo: _s(i.tipo), credito: _creditoVisible(act, i.credito), alt: _s(i.alt), soloInterna: _s(i.solo_interna) === 'Sí' })),
    max: _maxImagenes(p, Number(_propuestasAjustes(ss).maxImagenes) || 7), maxPropio: _s(p.imagenes_max) };
}
/** accion: 'principal' | 'interna' | 'publica' | 'retirar' | 'alt'. */
function galeriaAccion(idImagen, accion, valor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const lock = _lock();
  try {
    const t = _hojaImagenes(ss), i = t.filas.findIndex(f => _k(_get(t, f, 'id_imagen')) === _k(idImagen));
    if (i < 0) throw new Error(tx('prop.error.noExiste'));
    const f = t.filas[i], idP = _s(_get(t, f, 'id_palabra'));
    if (accion === 'principal') {
      const fila = { archivo: _s(_get(t, f, 'archivo')), credito: _s(_get(t, f, 'credito')), alt: _s(_get(t, f, 'alt')) };
      _actualizar(t, i, { estado: _etiquetaEstado('retirado'), nota: tx('prop.pasoPrincipal') });
      _hacerPrincipal(ss, idP, fila);
    } else if (accion === 'interna' || accion === 'publica') _actualizar(t, i, { solo_interna: accion === 'interna' ? 'Sí' : 'No' });
    else if (accion === 'retirar') _actualizar(t, i, { estado: _etiquetaEstado('retirado') });
    else if (accion === 'alt') _actualizar(t, i, { alt: _s(valor).slice(0, 300) });
    else throw new Error(tx('prop.error.accion'));
    _historial(ss, CONFIG.SHEETS.IMAGENES, idImagen, {}, { accion });
    return galeriaDatos(idP);
  } finally { lock.releaseLock(); }
}

/** Galerías publicadas para el visor: solo imágenes publicadas y no internas, con su crédito visible. */
function _galeriasPublicas(ss, visibles, actores) {
  const g = {};
  if (!ss.getSheetByName(CONFIG.SHEETS.IMAGENES)) return g;
  readSheet(ss, CONFIG.SHEETS.IMAGENES).filter(i => _s(i.archivo) && visibles.has(_k(i.id_palabra)) && _esPublico(i) && _s(i.solo_interna) !== 'Sí')
    .sort((a, b) => Number(a.orden || 0) - Number(b.orden || 0))
    .forEach(i => {
      const id = _idDrive(i.archivo) || (/^[-\w]{20,}$/.test(_s(i.archivo)) ? _s(i.archivo) : '');
      if (!id) return;
      (g[_k(i.id_palabra)] = g[_k(i.id_palabra)] || []).push({ url: `https://lh3.googleusercontent.com/d/${id}`, tipo: _s(i.tipo), credito: _s(i.credito) ? _creditoVisible(actores, i.credito) : '', alt: _s(i.alt) });
    });
  return g;
}

/** Quiénes han aportado (con permiso de mostrar su nombre) y cuántas aportaciones se aceptaron. */
function _aportesPublicos(ss) {
  if (!_propuestasAjustes(ss).mostrarAportes || !ss.getSheetByName(CONFIG.SHEETS.PROPUESTAS)) return [];
  const n = {};
  readSheet(ss, CONFIG.SHEETS.PROPUESTAS).forEach(p => {
    if (_s(p.estado) !== 'Aceptada' || _s(p.credito_publico) !== 'Sí' || !_s(p.credito_nombre)) return;
    const k = _plegar(p.credito_nombre); n[k] = n[k] || { nombre: _s(p.credito_nombre), n: 0 }; n[k].n++;
  });
  // Quien pidió retirar su crédito ya no aparece (sus fuentes quedaron sin crédito público)
  const sinCredito = new Set(readSheet(ss, CONFIG.SHEETS.FUENTES).filter(f => _s(f.base_uso) === BASE_APORTE && _s(f.publicar_credito) === 'No').map(f => _plegar(f.autoria)));
  return Object.keys(n).filter(k => !sinCredito.has(k)).map(k => n[k]).sort((a, b) => b.n - a.n).slice(0, 300);
}

/** Aviso automático de novedades (una franja que se renueva): «Esta semana se sumaron…». */
function _avisoNovedades(ss) {
  const hace7 = Date.now() - 7 * 864e5;
  const rec = readSheet(ss, CONFIG.SHEETS.PROPUESTAS).filter(p => _s(p.estado) === 'Aceptada' && new Date(p.fecha_resolucion).getTime() > hace7);
  if (!rec.length) return;
  const personas = new Set(rec.map(p => _plegar(p.credito_nombre) || _s(p.dispositivo))).size;
  const t = _hojaAvisos(ss), id = 'novedades-automatico';
  const fila = { tipo: 'novedad', texto: tx('prop.novedades', { n: rec.length, m: personas }), desde: _hoyISO(),
    hasta: Utilities.formatDate(new Date(Date.now() + 7 * 864e5), Session.getScriptTimeZone(), 'yyyy-MM-dd'), boton: tx('prop.novedades.boton'), enlace: 'buscar:',
    cerrable: 'Sí', orden: 5, estado: _etiquetaEstado('publicado') };
  const i = t.filas.findIndex(f => _s(_get(t, f, 'id_aviso')) === id);
  if (i > -1) _actualizar(t, i, fila); else _agregar(t, [Object.assign({ id_aviso: id, creado_por: 'automático', fecha: _ahora() }, fila)]);
  _publicarEstadoVisor(ss);
}

/** Ajustes de participación (pestaña 📥 Participación de la ventana del visor). */
function propuestasAjustesGuardar(o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _requerirAdmin(ss);
  const a = _visorAjustes(ss), p = Object.assign(_PROP_DEF(), a.propuestas || {});
  if (o.activo != null) p.activo = !!o.activo;
  if (o.quien) p.quien = o.quien === 'clave' ? 'clave' : 'todos';
  const num = (v, min, max, d) => { const n = Math.round(Number(v)); return isFinite(n) ? Math.min(max, Math.max(min, n)) : d; };
  if (o.maxImagenes != null) p.maxImagenes = num(o.maxImagenes, 1, 30, 7);
  if (o.maxPendDesc != null) p.maxPendDesc = num(o.maxPendDesc, 1, 20, 3);
  if (o.maxPendCorr != null) p.maxPendCorr = num(o.maxPendCorr, 1, 20, 3);
  if (o.porDia != null) p.porDia = num(o.porDia, 1, 200, 30);
  if (o.criterios != null) p.criterios = _s(o.criterios).slice(0, 1500);
  if (o.mostrarAportes != null) p.mostrarAportes = !!o.mostrarAportes;
  if (o.novedades != null) p.novedades = !!o.novedades;
  if (o.clave === 'nueva') p.clave = 'col' + _nuevaClave().slice(0, 12);
  if (o.clave === 'quitar') p.clave = '';
  if (p.quien === 'clave' && !p.clave) p.clave = 'col' + _nuevaClave().slice(0, 12);
  a.propuestas = p;
  _guardarVisorAjustes(ss, a);
  _historial(ss, 'CONFIG', 'PROPUESTAS', {}, { activo: p.activo, quien: p.quien });
  return conexionDatos();
}

/** Espacio de la cuenta de Drive (para avisar a tiempo, antes de que fallen los envíos). */
function _espacioDrive() {
  try { const u = DriveApp.getStorageUsed(), l = DriveApp.getStorageLimit(); return l > 0 ? { usado: u, limite: l, pct: Math.round(100 * u / l) } : null; } catch (e) { return null; }
}

function uiPropuestas() {
  const html = `${_CSS_FICHAS}${_tScript()}
  <style>
    body{margin:0}.caja{padding:10px}
    .cab{display:flex;gap:8px;align-items:center;flex-wrap:wrap;background:#f1f5f9;border-radius:8px;padding:6px 10px;margin-bottom:8px;font-size:13px}
    .pr{border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin:8px 0;background:#fff}.pr.alta{border-left:5px solid #0f766e}
    .pr h4{margin:0 0 4px;font-size:15px}.pr .meta{font-size:12px;color:#64748b}.pr .dato{margin:4px 0}
    .chip{display:inline-block;border-radius:10px;padding:0 7px;font-size:11px;background:#f1f5f9;margin-right:4px}
    .fila{display:flex;gap:6px;flex-wrap:wrap;align-items:center}.fila>input,.fila>select{flex:1;min-width:140px;width:auto}
    .thumbs img{max-height:160px;max-width:220px;border:1px solid #e2e8f0;border-radius:6px;margin:4px 4px 0 0;background:#f8fafc}
    .coinc{background:#fffbeb;border:1px solid #fde68a;border-radius:6px;padding:6px;margin:4px 0;font-size:12px}
    details summary{cursor:pointer}textarea{min-height:44px}
  </style>
  <div class="caja">
    <div class="cab"><b>📥 ${txh('prop.titulo')}</b><span id="cuenta"></span><button class="peq" onclick="google.script.run.uiConstancia()">📜 ${txh('const.boton')}</button>
      <label style="display:inline;font-weight:normal;margin-left:auto"><input type="checkbox" id="todas"> ${txh('prop.verTodas')}</label></div>
    <p><small>${txh('prop.intro')}</small></p>
    <div id="lista">${txh('comun.cargando')}</div>
    <div id="msg" role="status" aria-live="polite"></div>
  </div>
  <script>
    const $ = id => document.getElementById(id), esc = _escT;
    let D = null;
    const err = e => { $('msg').innerHTML = '<div class="error">' + esc(tt('comun.error', {m: e.message || e})) + '</div>'; };
    const opts = (l, v) => l.map(x => '<option value="' + esc(x.v) + '"' + (x.v === v ? ' selected' : '') + '>' + esc(x.t) + '</option>').join('');
    const thumb = id => '<img src="https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w400" alt="" loading="lazy">';
    function pintar(d) {
      D = d;
      $('cuenta').textContent = tt('prop.pendientes', {n: d.pendientes});
      if (!d.propuestas.length) { $('lista').innerHTML = '<div class="ok">' + th('prop.vacia') + '</div>'; return; }
      $('lista').innerHTML = d.propuestas.map(p => {
        const tipo = th('prop.tipo.' + p.tipo), quien = (p.credito_nombre ? esc(p.credito_nombre) + (p.credito_publico === 'Sí' ? ' 🙋' : ' 🔒') : th('prop.anonimo')) + ' · ' + th('prop.rol.' + (p.rol || 'otro')) + (p.recopilo ? ' · ' + th('prop.recopilo', {x: p.recopilo}) : '');
        let cuerpo = '';
        if (p.palabraTexto) cuerpo += '<div class="dato">' + th('prop.paraPalabra', {x: p.palabraTexto}) + (p.lengua ? ' · ' + esc(p.lengua) : '') + '</div>';
        if (p.texto) cuerpo += '<div class="dato"><b>' + esc(p.texto) + '</b>' + (p.lengua && !p.palabraTexto ? ' <span class="chip">' + esc(p.lengua) + '</span>' : '') + '</div>';
        if (p.actual && p.tipo !== 'descripcion') cuerpo += '<div class="dato"><small>' + th('prop.actual', {x: p.actual.texto}) + '</small></div>';
        if (p.significado) cuerpo += '<div class="dato">= ' + esc(p.significado) + (p.lenguaSig ? ' <span class="chip">' + esc(p.lenguaSig) + '</span>' : '') + '</div>';
        if (p.descripcion) cuerpo += '<div class="dato">📖 ' + esc(p.descripcion) + (p.actual && p.actual.descripcion ? '<br><small>' + th('prop.descActual', {x: p.actual.descripcion}) + '</small>' : '') + '</div>';
        if (p.variante) cuerpo += '<div class="dato">📍 ' + th('prop.seDice', {x: p.variante}) + '</div>';
        if (p.categoria) cuerpo += '<div class="dato">📁 ' + esc(p.categoria) + '</div>';
        if (p.nota) cuerpo += '<div class="dato">📝 ' + esc(p.nota) + '</div>';
        if (p.actividad) cuerpo += '<div class="dato">🧩 ' + th('prop.actividad', {x: p.actividad}) + '</div>';
        if (p.archivos) cuerpo += '<div class="thumbs">' + p.archivos.split(',').filter(Boolean).map(thumb).join('') + '</div>';
        if (p.cupo) cuerpo += '<div class="dato"><small>🖼️ ' + th('prop.cupo', {n: p.cupo.n, max: p.cupo.max}) + ' · ' + esc(p.tipo_imagen) + ' · ' + esc(p.licencia) + '</small></div>';
        if (p.coincide.length) cuerpo += '<div class="coinc">⚠️ ' + th('prop.coincide') + ' ' + p.coincide.map(c => '<b>' + esc(c.texto) + '</b>').join(' · ') + '</div>';
        if (p.contacto) cuerpo += '<details><summary><small>🔒 ' + th('prop.contacto') + '</small></summary><small>' + esc(p.contacto) + '</small></details>';
        if (p.estado !== 'Pendiente') return '<div class="pr"><h4>' + tipo + ' · <span class="chip">' + esc(p.estado) + '</span></h4><div class="meta">' + esc(p.codigo) + ' · ' + quien + '</div>' + cuerpo +
          (p.respuesta ? '<div class="dato"><small>💬 ' + esc(p.respuesta) + '</small></div>' : '') +
          (p.credito_nombre ? '<div class="fila"><button class="peq" data-acc="retirarCredito" data-id="' + esc(p.id_propuesta) + '">' + th('prop.acc.retirarCredito') + '</button></div>' : '') + '</div>';
        // Acciones según el tipo
        let acc = '';
        const pub = '<label style="display:inline;font-weight:normal"><input type="checkbox" class="o-pub" checked> ' + th('prop.publicarYa') + '</label>';
        if (p.tipo === 'palabra') acc = '<div class="fila"><select class="o-cat"><option value="">' + th('editar.sinCategoria') + '</option>' + opts(D.categorias.map(c => ({v: c.id, t: c.nombre})), (D.categorias.find(c => c.nombre.toLowerCase() === (p.categoria || '').toLowerCase()) || {}).id) + '</select>' + pub + '</div>' +
          '<div class="fila"><button class="prim" data-acc="nueva">＋ ' + th('prop.acc.nueva') + '</button><select class="o-dest"><option value="">' + th('prop.sumarA') + '</option>' + opts(D.palabras.map(x => ({v: x.id, t: x.texto})), (p.coincide[0] || {}).id) + '</select><button data-acc="sumar">' + th('prop.acc.sumar') + '</button></div>';
        else if (p.tipo === 'descripcion') acc = '<textarea class="o-desc">' + esc(p.descripcion) + '</textarea><div class="fila"><button class="prim" data-acc="descripcion" data-modo="reemplazar">✓ ' + th('prop.acc.descripcion') + '</button>' + (p.actual && p.actual.descripcion ? '<button data-acc="descripcion" data-modo="agregar">' + th('prop.acc.agregarDesc') + '</button>' : '') + '</div>';
        else if (p.tipo === 'correccion') acc = '<div class="fila"><input type="text" class="o-texto" value="' + esc(p.texto) + '">' + '<button class="prim" data-acc="correccion" data-modo="reemplazar">' + th('prop.acc.reemplazar') + '</button><button data-acc="correccion" data-modo="sumar">' + th('prop.acc.sumarForma') + '</button></div>';
        else if (p.tipo === 'variante') acc = p.texto ? '<div class="fila"><input type="text" class="o-texto" value="' + esc(p.texto) + '"><button class="prim" data-acc="variante">' + th('prop.acc.variante') + '</button><button data-acc="no" data-registro="1">' + th('prop.acc.soloRegistro') + '</button></div>' : '<div class="fila"><button class="prim" data-acc="no" data-registro="1">✓ ' + th('prop.acc.confirmacion') + '</button></div>';
        else if (p.tipo === 'imagen') acc = '<div class="fila"><input type="text" class="o-alt" placeholder="' + esc(tt('prop.alt')) + '"></div><div class="fila"><button class="prim" data-acc="imagen">✓ ' + th('prop.acc.imagen') + '</button><button data-acc="imagen" data-principal="1">⭐ ' + th('prop.acc.principal') + '</button><button data-acc="imagen" data-interna="1">🔒 ' + th('prop.acc.interna') + '</button></div><small>' + th('prop.alt.d') + '</small>';
        else if (p.tipo === 'fotos') acc = '<div class="fila"><label style="display:inline;font-weight:normal"><input type="checkbox" class="o-borrar" checked> ' + th('prop.borrarFotos') + '</label><button class="prim" data-acc="transcrita">✓ ' + th('prop.acc.transcrita') + '</button></div><small>' + th('prop.fotos.d') + '</small>';
        else if (p.tipo === 'retiro') acc = '<small>' + th('prop.retiro.d') + '</small><div class="fila"><input type="text" class="o-nombre" value="' + esc(p.credito_nombre) + '" placeholder="' + esc(tt('prop.nombreRetiro')) + '"><button class="prim" data-acc="retirarCredito">' + th('prop.acc.retirarCredito') + '</button><button data-acc="no" data-registro="1">' + th('prop.acc.atendida') + '</button></div>';
        return '<div class="pr' + (p.prioridad === 'alta' ? ' alta' : '') + '" data-id="' + esc(p.id_propuesta) + '"><h4>' + tipo + (p.prioridad === 'alta' ? ' <span class="chip">⭐ ' + th('prop.prioridad') + '</span>' : '') + '</h4>' +
          '<div class="meta">' + esc(p.codigo) + ' · ' + esc(String(p.fecha).slice(0, 16).replace('T', ' ')) + ' · ' + quien + (p.compartir === 'Sí' ? ' · 🌎' : '') + '</div>' + cuerpo + acc +
          '<details><summary><small>' + th('prop.respuesta') + '</small></summary><textarea class="o-resp" placeholder="' + esc(tt('prop.respuesta.ph')) + '"></textarea><textarea class="o-nota" placeholder="' + esc(tt('prop.notaInterna')) + '"></textarea></details>' +
          '<div class="fila"><button data-acc="no">✕ ' + th('prop.acc.no') + '</button></div></div>';
      }).join('');
    }
    $('lista').addEventListener('click', e => {
      const b = e.target.closest('[data-acc]'); if (!b) return;
      const card = b.closest('.pr'), id = b.dataset.id || card.dataset.id, q = s => card.querySelector(s);
      const o = { respuesta: (q('.o-resp') || {}).value || '', nota_interna: (q('.o-nota') || {}).value || '', publicar: q('.o-pub') ? q('.o-pub').checked : true, verTodas: $('todas').checked };
      let acc = b.dataset.acc;
      if (acc === 'nueva') o.id_categoria = q('.o-cat').value;
      if (acc === 'sumar') { o.id_destino = q('.o-dest').value; if (!o.id_destino) return err({message: tt('prop.eligeDestino')}); }
      if (acc === 'descripcion') { o.descripcion = q('.o-desc').value; o.modo = b.dataset.modo; }
      if (acc === 'correccion' || acc === 'variante') { o.texto = q('.o-texto').value; o.modo = b.dataset.modo; }
      if (acc === 'imagen') { o.alt = q('.o-alt').value; o.principal = !!b.dataset.principal; o.soloInterna = !!b.dataset.interna; }
      if (acc === 'transcrita') o.borrarFotos = q('.o-borrar').checked;
      if (acc === 'retirarCredito') { o.nombre = (q('.o-nombre') || {}).value || ''; if (!confirm(tt('prop.confirmaRetiro'))) return; }
      if (acc === 'no' && b.dataset.registro) o.respuesta = o.respuesta || tt('prop.gracias');
      $('msg').textContent = tt('comun.guardando');
      google.script.run.withFailureHandler(err).withSuccessHandler(d => { pintar(d); $('msg').innerHTML = '<div class="ok">' + th('prop.resuelta') + '</div>'; }).propuestaResolver(id, acc, o);
    });
    $('todas').onchange = () => google.script.run.withFailureHandler(err).withSuccessHandler(pintar).propuestasDatos($('todas').checked);
    google.script.run.withFailureHandler(err).withSuccessHandler(pintar).propuestasDatos(false);
  </script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(980).setHeight(760), tx('prop.titulo'));
}

// Textos en español (base de todas las traducciones). Agrupados por la parte del programa que los usa.
const T_ES = {
  "rama.nombreCatalogo": "{catalogo} · 🌿 {nombre}",
  "rama.error.directorio": "En una copia de trabajo no se guardan datos privados de contacto: agrégalos en el catálogo principal.",
  "rama.error.sinFoto": "Esta hoja no tiene la foto de cuando se creó la copia; no se puede comparar.",
  "rama.error.acceso": "No puedo abrir esa copia. Pide a quien la creó que te la comparta.",
  "rama.error.noEsta": "Esa copia no está registrada en este catálogo.",
  "rama.error.nombre": "Ponle un nombre a la copia.",
  "rama.error.desdeCopia": "Esto se hace desde el catálogo principal, no desde una copia de trabajo.",
  "rama.abrirPrincipal": "Abrir el catálogo principal",
  "rama.enCopia.pasos": "Cuando termines, avisa a quien administra el catálogo principal: desde 🌿 Copias de trabajo de allá revisa tus cambios y decide qué integrar. Los datos privados de las voces (contactos) solo se editan en el principal.",
  "rama.misCambios": "Cambios hasta ahora: ➕ {agregar} · ✏️ {cambiar} · 🗑️ {borrar}",
  "rama.enCopia.d": "Creada el {fecha} por {autor}. Aquí puedes probar y corregir con libertad: no se publica en el visor ni toca el catálogo principal.",
  "rama.enCopia": "🌿 Estás en la copia de trabajo «{nombre}»",
  "rama.archivo": "🌿 {catalogo} — {nombre}",
  "rama.respaldo": "Respaldo antes de integrar «{nombre}» — {fecha}",
  "rama.historial": "{agregar} nuevos, {cambiar} corregidos, {borrar} borrados",
  "rama.integrado": "✅ Integrado: {agregar} nuevo(s), {cambiar} corregido(s), {borrar} borrado(s). Quedan {quedan} cambio(s) sin integrar en esa copia.",
  "rama.integrando": "Guardando un respaldo e integrando…",
  "rama.eligeAlgo": "Marca al menos un cambio.",
  "rama.integrar": "✅ Integrar lo marcado al principal",
  "rama.seBorra": "Se borra en el principal",
  "rama.todos": "Marcar o desmarcar todos",
  "rama.yaIguales": "{n} cambio(s) ya están igual en el principal y no se muestran.",
  "rama.conflicto.d": "Las filas en amarillo también se cambiaron en el principal después de crear la copia. Vienen sin marcar: si las marcas, la versión de la copia reemplaza a la del principal.",
  "rama.conflicto": "También cambió en el principal",
  "rama.propone": "La copia propone: ➕ {agregar} nuevo(s) · ✏️ {cambiar} corregido(s) · 🗑️ {borrar} borrado(s)",
  "rama.sinCambios": "No hay nada nuevo que integrar. ({iguales} cambio(s) ya están igual en el principal.)",
  "rama.comparando": "Comparando la copia con el principal…",
  "rama.descartar": "Descartar",
  "rama.revisar": "Revisar cambios",
  "rama.abrir": "Abrir la copia",
  "rama.col.estado": "Estado",
  "rama.col.quien": "Quién y cuándo",
  "rama.col.nombre": "Copia",
  "rama.ninguna": "Todavía no hay copias de trabajo.",
  "rama.lista": "Copias de este catálogo",
  "rama.creada": "✅ Copia «{nombre}» creada en la carpeta COPIAS DE TRABAJO.",
  "rama.creando": "Creando la copia… (puede tardar un minuto)",
  "rama.crear": "🌿 Crear copia",
  "rama.nombre.ej": "ej. Corrección de plantas — Tutupika",
  "rama.nombre": "Nombre de la copia",
  "rama.nueva": "Crear una copia de trabajo",
  "rama.intro": "Una copia de trabajo sirve para probar o corregir sin tocar el catálogo principal. No publica en el visor. Cuando esté lista, aquí revisas qué cambió y eliges qué integrar.",
  "rama.titulo": "🌿 Copias de trabajo",
  "ayuda.estadoCopia": "En qué va esta copia de trabajo.",
  "menu.ramas": "🌿 Copias de trabajo",
  "blanco.error.original": "Esta parece ser la hoja original. Para dejarla en blanco escribe exactamente su nombre: {nombre}.",
  "blanco.conservarEquipo": "Conservar el equipo (si no, solo quedas tú como administrador)",
  "blanco.conservarAtajos": "Conservar los atajos de escritura",
  "blanco.conservarLenguas": "Conservar las lenguas y sus letras (y las comunidades)",
  "blanco.nombreNuevo": "Nombre del nuevo catálogo (opcional; si lo dejas vacío, se queda el actual)",
  "blanco.escribeNombre": "Si de verdad quieres dejar en blanco «{nombre}», escribe su nombre:",
  "blanco.noSe": "⚠️ No puedo saber si esta hoja es una copia o el original de «{nombre}». Si es el original y solo quieres quitar las palabras, usa «🧺 Vaciar el contenido»: no cambia la identidad del catálogo.",
  "blanco.esOriginal": "⚠️ Esta es la hoja ORIGINAL de «{nombre}». Esta opción es solo para copias: primero usa Archivo → Hacer una copia y abre la copia. Si lo que quieres es quitar todo menos las lenguas, usa «🧺 Vaciar el contenido».",
  "blanco.esCopia": "✅ Esta hoja es una copia de otro catálogo: se puede dejar en blanco sin tocar el original.",
  "limpiar.t.enBlanco": "🧾 Preparar una copia en blanco para otro catálogo",
  "limpiar.error.elige": "Elige primero cuál (importación, fuente, lengua u hoja).",
  "traer.sinLenguas": "No elegiste ninguna lengua para traer.",
  "traer.sinColumnas": "No encontré columnas TRADUCCION_… en esa hoja.",
  "traer.lenguaNuevaAviso": "(no está en este catálogo: se creará)",
  "traer.cualesLenguas": "Trae estas lenguas (quita las que no quieras):",
  "traer.verLenguas": "🔎 Ver qué lenguas trae",
  "inter.queLenguasTraer": "¿Qué lenguas traer?",
  "inter.queLenguas": "¿Qué lenguas incluir?",
  "exportarHoja.error.lenguas": "Elige al menos una lengua.",
  "exportarHoja.archivo": "Catálogo · {lenguas} · {fecha}",
  "exportarHoja.listo": "{n} palabra(s). Está en la carpeta EXPORTACIONES.",
  "exportarHoja.crear": "Crear la hoja",
  "exportarHoja.compartibles": "Solo lo que sus fuentes permiten compartir",
  "exportarHoja.todo": "Todas las palabras (también las que están en revisión)",
  "exportarHoja.d": "Crea una hoja con las lenguas que elijas, en el mismo formato que *📥 Traer lo que llenaron*. Sirve para compartir solo una parte del catálogo, o para corregir en bloque y volver a traerla: en este mismo catálogo actualiza las palabras (por su código); en otro, las agrega.",
  "exportarHoja.titulo": "📄 Hoja por lengua (para editar y volver a traer)",
  "limpiar.historialHoja": "LIMPIEZA",
  "limpiar.respaldo.t3": "Para recuperarlas, copia las filas de cada pestaña y pégalas al final de la hoja del mismo nombre en el catálogo.",
  "limpiar.respaldo.t2": "Qué se limpió: {op}",
  "limpiar.respaldo.t1": "Filas borradas el {fecha} por {quien}.",
  "limpiar.respaldo.leeme": "LEEME",
  "limpiar.respaldo.archivo": "Respaldo de limpieza — {catalogo} — {fecha}",
  "limpiar.error.tipo": "Elige qué quieres limpiar.",
  "limpiar.error.hoja": "Esa hoja no se puede limpiar desde aquí.",
  "limpiar.error.cambio": "Los datos cambiaron desde la vista previa. Vuelve a tocar «Ver qué se borraría».",
  "limpiar.error.confirmar": "Para borrar, escribe exactamente {palabra}.",
  "limpiar.verRespaldo": "Ver el respaldo",
  "limpiar.listo": "✅ Listo: se quitaron {n} fila(s) en el archivo *{archivo}*.",
  "limpiar.error.noSeBorro": "Google no terminó de borrar estas hojas: {hojas}. No se cambió nada más. Vuelve a intentarlo; lo que se alcanzó a borrar está en el respaldo: {respaldo}",
  "blanco.despues": "Esta hoja ya es un catálogo propio («{nombre}») con código y clave nuevos. AppSheet y el visor todavía no están conectados a ella: configúralos en 🚀 Primeros pasos.",
  "limpiar.borrando": "Guardando el respaldo y borrando…",
  "limpiar.borrar": "🗑️ Borrar",
  "limpiar.palabraConfirmar": "BORRAR",
  "limpiar.escribe": "Para confirmar, escribe {palabra}",
  "limpiar.respaldoAviso": "Antes de borrar se guarda un respaldo con todas estas filas.",
  "limpiar.seBorrarian": "En el archivo «{archivo}» se quitarían {n} fila(s):",
  "limpiar.nada": "No hay nada que borrar con esas opciones.",
  "limpiar.ver": "👀 Revisar antes de borrar",
  "limpiar.d.hojaAviso": "⚠️ Se borran TODAS las filas de esa hoja (los encabezados se quedan).",
  "limpiar.d.hoja": "Hoja",
  "limpiar.d.autTodas": "Todas las autorizaciones",
  "limpiar.d.aut": "¿Cuáles?",
  "limpiar.d.incluirLengua": "Borrar también la lengua",
  "limpiar.d.lengua": "Lengua (traducciones registradas)",
  "limpiar.d.incluirFuente": "Borrar también la ficha de la fuente y sus autorizaciones",
  "limpiar.d.fuente": "Fuente",
  "limpiar.d.sinLotes": "No hay importaciones registradas. Las importaciones hechas con versiones anteriores no guardaron su número: para quitarlas usa «Vaciar el contenido» o «Quitar lo que viene de una fuente».",
  "limpiar.d.lote": "Importación",
  "limpiar.o.vacias": "Borrar también las categorías que se queden sin palabras",
  "limpiar.o.huerfanas": "Borrar también las palabras que se queden sin ninguna traducción",
  "limpiar.t.hoja": "Vaciar una pestaña completa",
  "limpiar.t.autorizaciones": "Autorizaciones",
  "limpiar.t.huerfanas": "Palabras sin ninguna traducción",
  "limpiar.t.lengua": "🗣️ Quitar una lengua",
  "limpiar.t.fuente": "📚 Quitar lo que viene de una fuente",
  "limpiar.t.lote": "↩️ Deshacer una importación",
  "limpiar.que": "¿Qué quieres hacer?",
  "limpiar.intro": "Elige qué quieres hacer. Antes de borrar verás *qué se borra, qué se conserva y qué cambia*, y cuántas filas son. Todo lo que se quita se guarda en un respaldo.",
  "limpiar.guia": "¿No sabes cuál elegir?",
  "limpiar.guia.1": "*Importé algo por error:* «Deshacer una importación». Si esa importación no aparece, «Vaciar el contenido».",
  "limpiar.guia.2": "*Quiero empezar de nuevo en este catálogo, sin perder las lenguas:* «Vaciar el contenido».",
  "limpiar.guia.3": "*Quiero un catálogo nuevo para otra comunidad a partir de este:* primero Archivo → Hacer una copia; abre la copia y ahí usa «Preparar una copia en blanco».",
  "limpiar.guia.4": "*Borré algo que no debía:* «Recuperar desde un respaldo».",
  "limpiar.g.comun": "Lo más común",
  "limpiar.g.ordenar": "Ordenar",
  "limpiar.g.avanzado": "Avanzado",
  "limpiar.g.recuperar": "Deshacer",
  "limpiar.t.contenido": "🧺 Vaciar el contenido y conservar las lenguas",
  "limpiar.t.contenido.d": "Quita todas las palabras, traducciones y categorías. El catálogo sigue siendo el mismo.",
  "limpiar.t.lote.d": "Quita lo que entró en una misma importación.",
  "limpiar.t.fuente.d": "Por ejemplo, todo lo de un libro o diccionario.",
  "limpiar.t.lengua.d": "Sus traducciones y, si quieres, la lengua misma.",
  "limpiar.t.huerfanas.d": "Palabras que quedaron vacías.",
  "limpiar.t.categoriasVacias": "Categorías sin palabras",
  "limpiar.t.categoriasVacias.d": "Categorías que no usa ninguna palabra.",
  "limpiar.t.autorizaciones.d": "Las de una fuente, una comunidad o todas.",
  "limpiar.t.hoja.d": "Para quien conoce bien las pestañas del catálogo.",
  "limpiar.t.enBlanco.d": "Solo en una copia hecha con Archivo → Hacer una copia. La vacía y le da identidad propia.",
  "limpiar.t.restaurar": "♻️ Recuperar desde un respaldo",
  "limpiar.t.restaurar.d": "Vuelve a poner lo que se quitó en una limpieza.",
  "limpiar.e.borra": "Se quita:",
  "limpiar.e.conserva": "Se conserva:",
  "limpiar.e.cambia": "Cambia:",
  "limpiar.e.ojo": "Ojo:",
  "limpiar.e.noCambia": "No cambia:",
  "limpiar.e.noCambia.d": "el código y la clave del catálogo, la dirección del visor, las carpetas, la app de AppSheet ni la configuración. Los audios e imágenes de Drive no se borran.",
  "limpiar.e.respaldo": "Antes de quitar nada se guarda un *respaldo* en la carpeta RESPALDOS: una hoja aparte con las filas quitadas. No es un catálogo que funcione; sirve para volver a ponerlas con «♻️ Recuperar desde un respaldo».",
  "limpiar.e.contenido.borra": "palabras, traducciones, categorías y sus nombres en otras lenguas, relaciones entre palabras, galerías de imágenes, propuestas y reportes. También lo que marques arriba.",
  "limpiar.e.contenido.conserva": "lenguas, letras y atajos, equipo, configuración, menús traducidos e historial; fuentes, voces y comunidades si no las marcas.",
  "limpiar.e.lote.borra": "las traducciones que entraron en esa importación; si lo marcas, también las palabras que se queden sin traducciones y las categorías que se queden vacías.",
  "limpiar.e.lote.conserva": "todo lo demás.",
  "limpiar.e.fuente.borra": "las traducciones de esa fuente; si lo marcas, también su ficha y sus autorizaciones, las palabras que se queden sin traducciones y las categorías vacías.",
  "limpiar.e.fuente.conserva": "todo lo demás. La ficha «Captura del equipo» nunca se borra.",
  "limpiar.e.lengua.borra": "las traducciones en esa lengua y los nombres de categorías en esa lengua; si lo marcas, también la lengua misma.",
  "limpiar.e.lengua.conserva": "las demás lenguas y sus traducciones.",
  "limpiar.e.huerfanas.borra": "las palabras que no tienen ninguna traducción y sus relaciones.",
  "limpiar.e.huerfanas.conserva": "todo lo demás.",
  "limpiar.e.categoriasVacias.borra": "las categorías que no usa ninguna palabra y sus nombres en otras lenguas.",
  "limpiar.e.categoriasVacias.conserva": "todo lo demás.",
  "limpiar.e.autorizaciones.borra": "las autorizaciones elegidas. Lo que dependía de ellas puede dejar de publicarse.",
  "limpiar.e.autorizaciones.conserva": "fuentes, personas y comunidades.",
  "limpiar.e.hoja.borra": "todas las filas de la pestaña elegida (los encabezados se quedan). Si es TRADUCCIONES o PALABRAS, las relaciones sueltas también.",
  "limpiar.e.hoja.conserva": "las demás pestañas. La ficha «Captura del equipo» nunca se borra.",
  "limpiar.e.enBlanco.borra": "palabras, traducciones, categorías, relaciones, fuentes, personas, autorizaciones, voces, propuestas, galerías, avisos, reportes e historial; también lenguas, atajos o equipo si los desmarcas.",
  "limpiar.e.enBlanco.conserva": "los menús traducidos, la configuración del visor y lo que dejes marcado.",
  "limpiar.e.enBlanco.cambia": "esta hoja recibe un código y una clave nuevos, para que no se confunda con el catálogo del que salió. Se olvidan las carpetas, la app de AppSheet y el directorio de voces (siguen existiendo en Drive). El título solo cambia si escribes uno nuevo.",
  "limpiar.e.enBlanco.ojo": "no se crea otra hoja: se vacía *esta*. La dirección del visor es la misma, pero un visor que use la clave anterior ya no se conecta. Después configura la copia en 🚀 Primeros pasos.",
  "limpiar.e.restaurar.borra": "nada. Solo se agregan de vuelta las filas del respaldo que ya no están (se reconocen por su código); lo que ya existe se salta.",
  "limpiar.e.restaurar.d": "Si el respaldo es de «Preparar una copia en blanco», también trae la configuración anterior (código, clave, carpetas y app) y puedes elegir recuperarla.",
  "limpiar.d.cFuentes": "También fuentes, personas e instituciones y autorizaciones (la ficha del equipo se queda)",
  "limpiar.d.cVoces": "También las fichas de voces (los audios de Drive no se borran)",
  "limpiar.d.cComunidades": "También las comunidades",
  "limpiar.d.loteOpcion": "{lote} · {n} traducciones",
  "limpiar.queEsRespaldo": "El respaldo es una hoja aparte con lo que se quitó. No es un catálogo que funcione: para volver a poner esas filas usa «♻️ Recuperar desde un respaldo».",
  "limpiar.r.ver": "👀 Revisar qué se recuperaría",
  "limpiar.r.cual": "Respaldo",
  "limpiar.r.ninguno": "Todavía no hay respaldos en la carpeta RESPALDOS.",
  "limpiar.r.elige": "Elige primero un respaldo.",
  "limpiar.r.noEs": "Ese archivo no está en la carpeta RESPALDOS de este catálogo.",
  "limpiar.r.seAgregarian": "Se agregarían de vuelta {n} fila(s):",
  "limpiar.r.nada": "Todo lo de ese respaldo ya está en el catálogo.",
  "limpiar.r.yaEstan": "{n} fila(s) del respaldo ya están en el catálogo y se saltan.",
  "limpiar.r.config": "Recuperar también la configuración de «{nombre}» (código, clave, carpetas y app de AppSheet)",
  "limpiar.r.config.d": "Úsalo si dejaste en blanco una hoja por error: vuelve a conectarla con su visor, sus carpetas y su app.",
  "limpiar.r.boton": "♻️ Recuperar",
  "limpiar.r.trabajando": "Recuperando…",
  "limpiar.r.listo": "✅ Se agregaron de vuelta {n} fila(s).",
  "limpiar.r.configListo": "También se recuperó la configuración: código, clave, carpetas y app.",
  "limpiar.r.historial": "Recuperación desde un respaldo",
  "traer.avisoLenguaNueva": "⚠️ Se creó una lengua nueva para {codigos} porque ninguna lengua del catálogo tiene ese código. Si es una lengua que ya tenías, deshaz esta importación (🧹 Limpiar → Deshacer una importación), escribe ese código en su ficha (🗣️ Lenguas → Código internacional) y vuelve a importar.",
  "menu.medios": "🖼️ Imágenes y sonidos",
  "medios.titulo": "🖼️ Imágenes y sonidos",
  "medios.intro": "Elige las palabras por lo que dicen, no por su código. Una *imagen* vale para la palabra en todas las lenguas; un *sonido* es de la palabra en una lengua.",
  "medios.tab.palabra": "🔎 Por palabra",
  "medios.tab.banco": "🗂️ Banco sin asignar",
  "medios.tab.revisar": "✅ Por revisar",
  "medios.tab.revisarN": "✅ Por revisar ({n})",
  "medios.tab.ilustrar": "🎨 Quien ilustra",
  "medios.tab.rec": "📏 Recomendaciones",
  "medios.imagenes": "Imágenes",
  "medios.sonidos": "Sonidos",
  "medios.lengua": "Lengua",
  "medios.buscar": "Buscar palabra en cualquier lengua o categoría…",
  "medios.f.sinImagen": "Sin imagen",
  "medios.f.conImagen": "Con imagen",
  "medios.f.sinAudio": "Sin sonido",
  "medios.f.conAudio": "Con sonido",
  "medios.f.todas": "Todas",
  "medios.mostrando": "Mostrando {n} de {total}",
  "medios.e.conImagen": "🖼️ Con imagen",
  "medios.e.sinImagen": "Sin imagen",
  "medios.e.consulta": "🔊 Consulta",
  "medios.e.sinConsulta": "Sin sonido de consulta",
  "medios.e.archivo": "💾 Alta calidad",
  "medios.e.sinArchivo": "Sin alta calidad",
  "medios.e.pendientes": "⏳ {n} por revisar",
  "medios.ver": "Ver",
  "medios.oir": "Escuchar",
  "medios.sinTexto": "(sin texto)",
  "medios.enlace.imagen": "Enlace de Drive, nombre del archivo o dirección web de la imagen",
  "medios.enlace.audio": "Enlace de Drive o nombre del archivo de sonido",
  "medios.credito": "Crédito: persona o institución",
  "medios.calidad": "Calidad del sonido",
  "medios.calidad.consulta": "Consulta (ligera, se publica)",
  "medios.calidad.archivo": "Alta calidad (archivo, privada)",
  "medios.vozSin": "Voz: sin registrar",
  "medios.asignar": "Asignar",
  "medios.directo": "Aplicar de inmediato (soy quien revisa). Si lo desmarcas, queda como propuesta.",
  "medios.propuesta": "Lo que asignes queda como propuesta hasta que quien administra lo revise.",
  "medios.listo": "✅ Listo. Se verá en el visor en la próxima publicación (o toca 📡).",
  "medios.propuesto": "✅ Propuesta enviada: aparece en «Por revisar».",
  "medios.banco.d": "Para imágenes o sonidos que llegaron sin palabra (de un banco, de una sesión de grabación…): pega la carpeta de Drive donde están. Aquí solo aparecen los que todavía no están asignados.",
  "medios.banco.carpeta": "Enlace de la carpeta de Drive",
  "medios.banco.ver": "Ver archivos sin asignar",
  "medios.banco.cuenta": "{n} archivo(s) sin asignar en «{carpeta}»",
  "medios.banco.sugerencia": "¿Es esta palabra?",
  "medios.banco.buscar": "Escribe la palabra en cualquier lengua…",
  "medios.banco.reemplaza": "Ya tiene imagen: se reemplazará",
  "medios.banco.elige": "Primero elige a qué palabra corresponde.",
  "medios.banco.mas": "Ver más archivos",
  "medios.rev.nada": "No hay imágenes ni sonidos esperando revisión.",
  "medios.rev.d": "Revisa cada propuesta: al aceptarla se coloca en la palabra (la imagen) o en su traducción (el sonido).",
  "medios.rev.dCaptura": "Estas propuestas esperan a que quien administra las revise.",
  "medios.rev.todo": "✓ Aceptar las {n}",
  "medios.rev.aceptar": "Aceptar",
  "medios.rev.rechazar": "Rechazar",
  "medios.rev.aceptadas": "✅ {n} aceptada(s).",
  "medios.rev.rechazadas": "{n} rechazada(s).",
  "medios.rev.credito": "Crédito: {c}",
  "medios.rev.voz": "Voz: {v}",
  "medios.origen.ventana": "Ventana de imágenes y sonidos",
  "medios.origen.banco": "Banco sin asignar",
  "medios.origen.ilustracion": "Hoja de ilustración «{nombre}»",
  "medios.error.enlace": "Escribe el enlace o el nombre del archivo.",
  "medios.error.noEncontrado": "No encontré ese archivo en Drive. Revisa el enlace o que tengas acceso a él.",
  "medios.error.noImagen": "Ese archivo no es una imagen.",
  "medios.error.noAudio": "Ese archivo no es de sonido.",
  "medios.error.audioWeb": "Los sonidos deben estar en Drive (el visor no reproduce sonidos de otras páginas).",
  "medios.error.audioMax": "Ese sonido pesa {kb} KB: el visor solo reproduce hasta 5 MB. Súbelo como «Alta calidad» y agrega una versión ligera para consulta.",
  "medios.error.palabra": "No encontré esa palabra.",
  "medios.error.sinTraduccion": "Esa palabra todavía no está escrita en esa lengua: primero agrega su traducción.",
  "medios.error.carpeta": "No pude abrir esa carpeta. Pega el enlace de una carpeta de Drive a la que tengas acceso.",
  "medios.aviso.web": "Es una imagen de otra página: si la página la quita, deja de verse. Mejor súbela a Drive.",
  "medios.aviso.imagenPesada": "La imagen pesa {kb} KB (más de 1 MB): tardará en cargar con poca señal. Conviene reducirla.",
  "medios.aviso.imagenMejorable": "La imagen pesa {kb} KB; lo ideal es menos de 300 KB.",
  "medios.aviso.audioPesado": "El sonido de consulta pesa {kb} KB; lo ideal es menos de 1 MB (unos 60 KB por palabra).",
  "medios.aviso.archivoComprimido": "Para «Alta calidad» conviene WAV o FLAC; este archivo ya está comprimido.",
  "medios.aviso.soloArchivo": "{kb} KB: solo sirve como alta calidad.",
  "medios.rec.imagen.titulo": "📏 Recomendaciones para imágenes",
  "medios.rec.imagen.1": "*Formato:* JPG para fotos; PNG o WebP para dibujos de colores planos.",
  "medios.rec.imagen.2": "*Tamaño:* cuadrada o 4:3, de 800 a 1200 píxeles por lado (nunca más de 1600).",
  "medios.rec.imagen.3": "*Peso:* ideal de 100 a 300 KB; máximo 1 MB. Así cargan rápido con poca señal y caben muchas en Drive.",
  "medios.rec.imagen.4": "*Contenido:* un solo objeto, fondo claro y *sin letras escritas*, para que sirva en todas las lenguas.",
  "medios.rec.imagen.5": "*Nombre del archivo:* la palabra (por ejemplo, takie.jpg) ayuda a que el banco la sugiera sola. Anota siempre el crédito de quien la hizo.",
  "medios.rec.audio.titulo": "📏 Recomendaciones para sonidos",
  "medios.rec.audio.1": "*Consulta (la que se publica):* MP3 o M4A, mono, 64 a 96 kbps, 44.1 kHz. Una palabra pesa de 20 a 60 KB; máximo 5 MB.",
  "medios.rec.audio.2": "*Alta calidad (archivo, no se publica):* WAV o FLAC, mono, 48 kHz, 24 bits. Pesa unos 8 MB por minuto (WAV) o la mitad (FLAC).",
  "medios.rec.audio.3": "*Grabación:* lugar callado, celular a un palmo de la boca, sin música; deja medio segundo de silencio antes y después y recorta lo demás.",
  "medios.rec.audio.4": "*Una palabra o frase por archivo*, con su nombre (por ejemplo, takie_wix.mp3) para que el banco lo sugiera solo.",
  "medios.rec.audio.5": "*Voz:* elige la voz registrada para que se respete su consentimiento; un sonido sin voz registrada no se publica.",
  "medios.rec.calidades.titulo": "¿Por qué dos calidades?",
  "medios.rec.calidades": "La versión *ligera* es la que se escucha en el visor y en los juegos, aun con poca señal. La de *alta calidad* se guarda para el acervo y para trabajos futuros (estudios, nuevas grabaciones de referencia); es privada. Quien graba elige: si tiene poco espacio en su celular, puede guardar solo la ligera, y la de alta calidad se agrega después.",
  "ilustrar.crear": "Crear una hoja para quien ilustra",
  "ilustrar.crear.d": "Una hoja con las palabras escritas en cada lengua y una columna para pegar el enlace de cada imagen. Se guarda en COLABORACIONES y se comparte como las de traducción.",
  "ilustrar.todasPalabras": "Incluir también las palabras que ya tienen imagen (para mejorarlas)",
  "ilustrar.crearBoton": "🎨 Crear la hoja",
  "ilustrar.creada": "✅ Hoja lista con {n} palabra(s).",
  "ilustrar.mensaje": "Hola. Te comparto la hoja «{nombre}» para agregar las imágenes del catálogo: {url}",
  "ilustrar.traer": "Traer lo que llenaron",
  "ilustrar.traer.d": "Las imágenes llegan como propuestas a «Por revisar»; nada cambia en el catálogo hasta aceptarlas.",
  "ilustrar.traerBoton": "Traer",
  "ilustrar.ninguna": "Todavía no hay hojas para quien ilustra.",
  "ilustrar.noExiste": "ya no existe",
  "ilustrar.traidas": "✅ Llegaron {n} imagen(es) nuevas ({rep} ya estaban).",
  "ilustrar.conAvisos": "⚠️ {n} tienen avisos de tamaño: revísalas antes de aceptar.",
  "ilustrar.errorFila": "Fila {fila}: {error}",
  "ilustrar.error.hoja": "Esa hoja no tiene la pestaña DATOS_ILUSTRACION.",
  "ilustrar.enc.ID_PALABRA": "(código: no modificar)",
  "ilustrar.enc.CATEGORIA": "Categoría",
  "ilustrar.enc.PALABRA": "Palabra en cada lengua",
  "ilustrar.enc.ENLACE_IMAGEN": "👉 Enlace de la imagen en Drive",
  "ilustrar.enc.CREDITO": "👉 Crédito (quién la hizo)",
  "ilustrar.enc.COMENTARIOS": "👉 Comentarios",
  "ilustrar.inst.1": "🎨 Imágenes para «{catalogo}»",
  "ilustrar.inst.2": "1. En la pestaña DATOS_ILUSTRACION verás cada palabra escrita en las lenguas del catálogo. Una imagen sirve para todas las lenguas: dibuja la idea, sin letras.",
  "ilustrar.inst.3": "2. Sube cada imagen a Google Drive. Recomendado: JPG o PNG, cuadrada, de 800 a 1200 píxeles por lado y menos de 300 KB (máximo 1 MB). Un solo objeto y fondo claro.",
  "ilustrar.inst.4": "3. En Drive: clic derecho → Compartir → Copiar vínculo. Pégalo en la columna «Enlace de la imagen», en la fila de su palabra.",
  "ilustrar.inst.5": "4. En «Crédito» elige tu nombre de la lista (o escríbelo si no aparece), y escribe cualquier duda en «Comentarios».",
  "ilustrar.inst.6": "5. No cambies ni borres filas, ni la primera fila de títulos: el catálogo las usa para saber a qué palabra va cada imagen.",
  "ilustrar.inst.7": "Quien cuida el catálogo revisa cada imagen con la comunidad antes de publicarla. ¡Gracias!",
  "medios.categoria": "Categoría",
  "medios.todasCategorias": "Todas las categorías",
  "medios.f.conNota": "Con nota",
  "medios.notaInterna": "Nota interna (no se publica): p. ej. Imagen temporal",
  "medios.guardarNota": "Guardar nota",
  "medios.notaGuardada": "📝 Nota guardada.",
  "medios.notaBorrada": "Nota borrada.",
  "medios.actorCreado": "Agregada desde 🖼️ Imágenes y sonidos (crédito de imagen).",
  "medios.aviso.actorNuevo": "Se agregó «{nombre}» a 📚 Personas e instituciones. Revisa ahí su tipo y cómo se escribe su crédito.",
  "medios.nota.temporal": "Imagen temporal",
  "medios.nota.mejorar": "Mejorar calidad",
  "medios.nota.comunidad": "Revisar con la comunidad",
  "medios.nota.permiso": "Falta confirmar permiso",
  "medios.nota.regrabar": "Volver a grabar",
  "medios.nota.deIlustracion": "Comentario de quien ilustra: {c}",
  "ilustrar.categorias": "Categorías",
  "ilustrar.categorias.d": "(sin marcar = todas; así cada persona trabaja por temas)",
  "ilustrar.archivoCat": "Imágenes · {categorias} · {fecha}",
  "ilustrar.nadaCategorias": "En esas categorías no hay palabras que necesiten imagen.",
  "menu.editar": "✏️ Editar palabras",
  "menu.colaborar": "🤝 Colaborar",
  "menu.configurar": "🗂️ Configurar el catálogo",
  "menu.mantenimiento": "🔧 Mantenimiento",
  "panel.grupo.colaborar": "Colaborar",
  "panel.grupo.configurar": "Configurar el catálogo",
  "panel.grupo.mantenimiento": "Mantenimiento",
  "panel.editar": "✏️ Editar palabras",
  "panel.editar.d": "Buscar en cualquier lengua, corregir, agregar o cambiar en bloque",
  "panel.medios": "🖼️ Imágenes y sonidos",
  "panel.medios.d": "Asignar, revisar y trabajar con quien ilustra",
  "panel.hojaIlustrar": "🎨 Hoja para quien ilustra",
  "panel.voces": "🎙️ Voces y permisos",
  "panel.fuentes": "📚 Fuentes y corpus",
  "panel.coincidencias": "🔎 Palabras que se escriben igual",
  "panel.ramas": "🌿 Copias de trabajo",
  "panel.limpiar": "🧹 Limpiar o recuperar datos",
  "panel.limpiar.d": "Siempre con vista previa y respaldo",
  "editar.titulo": "✏️ Editar palabras",
  "editar.buscar": "Buscar en cualquier lengua…",
  "editar.lista": "Palabras",
  "editar.nueva": "Palabra nueva",
  "editar.elige": "Elige una palabra de la lista, o crea una nueva.",
  "editar.esRama": "Copia de trabajo",
  "editar.rolAdmin": "Administras: tus cambios se aplican tal cual.",
  "editar.rolCaptura": "Capturas: lo que corrijas de algo ya publicado vuelve a revisión.",
  "editar.sinCategoria": "Sin categoría",
  "editar.todosEstados": "Todos los estados",
  "editar.categoriaNueva": "Categoría nueva…",
  "editar.f.falta": "Qué le falta",
  "editar.f.nada": "Sin filtro de faltantes",
  "editar.f.sinLengua": "Sin {lengua}",
  "editar.f.sinFuente": "Con traducciones sin fuente",
  "editar.tieneImagen": "Tiene imagen",
  "editar.sinImagen": "Sin imagen (se agrega en 🖼️ Imágenes y sonidos)",
  "editar.sinTraduccion": "sin traducción",
  "editar.escribe": "Escribe la palabra en {lengua} (varias formas: sepáralas con coma)",
  "editar.seGuardaComo": "En el visor se ve como: {texto}",
  "editar.mas": "Nota, sílabas, fuente y estado",
  "editar.notaVariante": "Nota de variante (pública)",
  "editar.silabas": "Sílabas (ka-ki)",
  "editar.sinFuente": "Sin fuente",
  "editar.guardar": "Guardar",
  "editar.retirar": "Retirar la palabra",
  "editar.guardado": "✅ Guardado. Se verá en el visor en la próxima publicación.",
  "editar.sinCambios": "No había cambios que guardar.",
  "editar.coinciden": "Esta forma ya existe en otra palabra:",
  "editar.coincide": "«{forma}» ({lengua}) también está en: {palabras}",
  "editar.coinciden.d": "Si son la misma idea, conviene unirlas en 🔎 Palabras que se escriben igual. Si son distintas (homónimos), puedes guardar igual.",
  "editar.guardarIgual": "Guardar de todos modos",
  "editar.aviso.soloAdmin": "Solo quien administra publica o retira: quedó «En revisión».",
  "editar.aviso.aRevision": "Corregiste algo publicado: vuelve a «En revisión» hasta que alguien lo apruebe.",
  "editar.aviso.categoriaNueva": "Se creó la categoría «{nombre}».",
  "editar.aviso.vacio": "No se borra el texto en {lengua}: para quitarlo, cambia su estado a Retirado.",
  "editar.error.vacia": "Escribe la palabra al menos en una lengua.",
  "editar.error.categoria": "Elige una categoría que exista.",
  "editar.error.fuente": "Elige una fuente que exista.",
  "editar.error.accion": "Elige qué cambiar.",
  "editar.origen": "Edición desde la hoja",
  "editar.historialBloque": "EDICIÓN EN BLOQUE",
  "editar.bloque": "{n} palabra(s) seleccionada(s): cambiar en bloque",
  "editar.b.categoria": "Mover a la categoría",
  "editar.b.estado": "Cambiar el estado (y sus traducciones)",
  "editar.b.fuente": "Asignar la fuente a sus traducciones",
  "editar.b.juegos": "Usar en juegos",
  "editar.b.todasLenguas": "En todas las lenguas",
  "editar.b.ver": "👀 Ver cuántas cambian",
  "editar.b.vista": "Cambiarían {p} palabra(s) y {t} traducción(es).",
  "editar.b.listo": "✅ Listo: {p} palabra(s) y {t} traducción(es).",
  "editar.b.quitarSel": "Quitar la selección",
  "editar.publicado": "📡 Publicado: {palabras} palabras y {traducciones} traducciones en el visor.",
  "editar.error.lengua": "Esa lengua no existe en el catálogo. Cierra y vuelve a abrir la ventana.",
  "editar.reparadaSinLengua": "Retirada: se guardó sin lengua por un error de la ventana",
  "editar.reparadas": "Se corrigió un error anterior de esta ventana: {n} texto(s) se habían guardado sin lengua y quedaron retirados. Vuelve a escribir esos cambios en la lengua correcta.",
  "menu.avisos": "📣 Avisos en el visor",
  "panel.avisos": "📣 Avisos en el visor",
  "vtab.conexion": "Conexión y claves",
  "vtab.mant": "Mantenimiento",
  "vtab.visor": "Ajustes del visor",
  "conexion.para.publico": "Para el público",
  "conexion.para.equipo": "Para el equipo",
  "conexion.mensajeEquipo": "Acceso del equipo al catálogo «{catalogo}» (también funciona durante el mantenimiento). No lo compartas fuera del equipo:",
  "equipo.claveTitulo": "Clave del equipo",
  "equipo.claveD": "Con esta clave el equipo ve el catálogo aunque esté en mantenimiento, y ve los cambios en curso. Es distinta de la clave de consulta.",
  "equipo.sinClave": "Todavía no hay clave del equipo.",
  "equipo.confirmaNueva": "Quien tenga la clave del equipo anterior dejará de entrar durante el mantenimiento. ¿Seguir?",
  "equipo.confirmaQuitar": "Sin clave del equipo, nadie podrá ver el catálogo mientras esté cerrado por mantenimiento. ¿Seguir?",
  "equipo.lista": "Listo: la clave del equipo ya rige. Compártela solo con el equipo.",
  "equipo.quitada": "Listo: se quitó la clave del equipo.",
  "mant.como": "¿Qué ve el público durante el mantenimiento?",
  "mant.congelar": "Congelar la versión actual",
  "mant.congelar.d": "El público sigue viendo el catálogo como estaba antes de empezar. El equipo, con su clave, ve los cambios en curso. Para cargas y reorganizaciones grandes.",
  "mant.cerrar": "Cerrar temporalmente",
  "mant.cerrar.d": "El público ve una pantalla de mantenimiento en lugar de las palabras. Para retirar contenido por un señalamiento o cuando no conviene que se vea nada mientras tanto.",
  "mant.m.congelar": "congelado",
  "mant.m.cerrar": "cerrado",
  "mant.avisoPublico": "Avisar al público que está viendo la versión anterior (una franja arriba)",
  "mant.purgar": "Borrar también la copia guardada en los dispositivos",
  "mant.purgar.d": "Cada visor borra su copia al conectarse. Un celular que no se conecta conserva su copia hasta que lo haga: es una medida precautoria, no absoluta.",
  "mant.mensaje": "Mensaje para el público",
  "mant.mensaje.ph": "Estamos actualizando el catálogo. Gracias por su paciencia.",
  "mant.hasta": "Fecha estimada de regreso",
  "mant.hasta.ph": "Por ejemplo: lunes 3 de noviembre",
  "mant.imagen": "Imagen de la pantalla de mantenimiento",
  "mant.imagen.d": "Si la dejas vacía, se usa la imagen del catálogo (pestaña Ajustes del visor). Recomendado: 600 a 800 píxeles de ancho y menos de 200 KB.",
  "mant.imagen.ph": "Enlace de Drive o dirección https de la imagen",
  "mant.contacto": "Mostrar estos datos de contacto (de 🛠️ Datos del catálogo):",
  "mant.sinContacto": "No hay datos de contacto: agrégalos en 🛠️ Datos del catálogo y contacto.",
  "mant.c.nombre": "Nombre",
  "mant.c.correo": "Correo",
  "mant.c.whatsapp": "WhatsApp",
  "mant.c.telefono": "Teléfono",
  "mant.c.horario": "Horario",
  "mant.sinClaveEquipo": "⚠️ No hay clave del equipo: mientras esté cerrado, tampoco el equipo podrá ver el catálogo en el visor. Créala en la pestaña Conexión y claves.",
  "mant.encender": "Encender el mantenimiento ahora",
  "mant.encendido": "🚧 Mantenimiento encendido. Los visores lo verán en unos minutos.",
  "mant.activo": "En mantenimiento ({modo}) desde {desde} · {quien}",
  "mant.enLinea": "El catálogo está en línea.",
  "mant.terminar": "Terminar el mantenimiento y publicar",
  "mant.terminado": "✅ Mantenimiento terminado: se publicó lo último y los visores se actualizan solos.",
  "mant.programar": "Programar un mantenimiento",
  "mant.programar.d": "Se enciende y se apaga solo a la hora indicada (se revisa cada hora), con las opciones elegidas arriba.",
  "mant.inicio": "Empieza",
  "mant.fin": "Termina",
  "mant.crearAviso": "Anunciarlo con un aviso desde hoy:",
  "mant.textoAviso": "El catálogo entrará en mantenimiento pronto y durante ese tiempo podría no verse. Agradecemos su comprensión.",
  "mant.programarBoton": "Programar",
  "mant.programado": "Programado: {inicio} → {fin} ({modo})",
  "mant.programadoOk": "📅 Mantenimiento programado.",
  "mant.programaCancelado": "Se canceló el mantenimiento programado.",
  "mant.error.inicio": "Elige cuándo empieza.",
  "mant.error.fin": "El final debe ser después del inicio.",
  "visorAj.revision": "Revisar si hay cambios",
  "visorAj.revision.d": "Cada cuánto el visor abierto revisa si hay una versión nueva (solo con la pestaña visible y con conexión). Es una consulta muy pequeña.",
  "visorAj.min": "Cada {n} minutos",
  "visorAj.nunca": "No revisar (solo al abrir el visor)",
  "visorAj.identidad": "Imagen del catálogo",
  "visorAj.identidad.d": "Un logo o ilustración que da identidad. Se usa en la pantalla de mantenimiento y, si quieres, junto al título del visor. Recomendado: menos de 200 KB.",
  "visorAj.alt": "Descripción de la imagen (la leen los lectores de pantalla)",
  "visorAj.enTitulo": "Mostrarla también junto al título del visor",
  "visorAj.error.imagen": "La imagen debe ser un enlace de Drive o una dirección https.",
  "visorAj.avisos.d": "Una franja arriba del visor para anunciar novedades, pausas o algo importante.",
  "visorAj.abrirAvisos": "Abrir avisos",
  "visorAj.borrar": "Borrar las copias guardadas en los dispositivos",
  "visorAj.borrar.d": "Útil después de retirar algo sensible: cada visor borra su copia al conectarse y descarga lo actual. Un dispositivo sin conexión la conserva hasta que se conecte.",
  "visorAj.borrarBoton": "Borrar copias",
  "visorAj.borrar.confirma": "Todos los visores descargarán el catálogo completo otra vez al conectarse. ¿Seguir?",
  "visorAj.borrado": "🧽 Listo: los visores borrarán su copia al conectarse.",
  "avisos.titulo": "📣 Avisos en el visor",
  "avisos.intro": "Una franja arriba del visor para quienes consultan, sin necesidad de registrarse. Se muestra y desaparece sola según sus fechas.",
  "avisos.nuevo": "Nuevo aviso",
  "avisos.editar": "Editar",
  "avisos.tipo.info": "ℹ️ Información",
  "avisos.tipo.novedad": "✨ Novedad",
  "avisos.tipo.pausa": "⏸️ Pausa programada",
  "avisos.tipo.importante": "⚠️ Importante",
  "avisos.cerrable": "Se puede cerrar",
  "avisos.texto": "Texto del aviso",
  "avisos.texto.d": "Breve y claro. Puedes usar *negritas*. Si lo quieres en varias lenguas, escríbelas una tras otra.",
  "avisos.desde": "Se muestra desde",
  "avisos.hasta": "Hasta (incluido)",
  "avisos.boton": "Botón (opcional)",
  "avisos.boton.texto": "Texto del botón: p. ej. Ver las nuevas palabras",
  "avisos.boton.ninguno": "Sin botón",
  "avisos.boton.web": "Abre una página web",
  "avisos.boton.buscar": "Busca en el catálogo",
  "avisos.boton.categoria": "Abre una categoría",
  "avisos.boton.buscarQue": "Qué buscar",
  "avisos.visible": "Visible (si lo desmarcas, se guarda sin mostrarse)",
  "avisos.lista": "Avisos",
  "avisos.max": "Cuántos se ven a la vez:",
  "avisos.v.activo": "Se está mostrando",
  "avisos.v.futuro": "Empieza después",
  "avisos.v.vencido": "Ya terminó",
  "avisos.v.oculto": "Oculto",
  "avisos.desdeF": "desde {f}",
  "avisos.hastaF": "hasta {f}",
  "avisos.noCerrable": "no se puede cerrar",
  "avisos.quitar": "Quitar",
  "avisos.ninguno": "Todavía no hay avisos.",
  "avisos.guardado": "✅ Guardado. Los visores lo verán en unos minutos.",
  "avisos.quitado": "Aviso quitado.",
  "avisos.error.texto": "Escribe el texto del aviso.",
  "avisos.error.enlace": "El botón debe abrir una página https, una búsqueda o una categoría.",
  "avisos.error.fechas": "La fecha final debe ser después de la inicial.",
  "limpiar.titulo": "🧹 Limpiar o recuperar datos",
  "menu.herr.limpiar": "🧹 Limpiar o recuperar datos",
  "traer.loteSinNombre": "hoja sin nombre",
  "comoVa.esperanPermiso": "📚 {n} traducción(es) están Publicadas pero no salen en el visor porque su fuente todavía no autoriza publicar. Revísalo en 📚 Fuentes, permisos y comunidades.",
  "inter.noCompartibles": "{n} traducción(es) no se incluyeron porque su fuente no autoriza compartirlas con otros catálogos.",
  "coinc.error.fiel": "No se puede juntar: «{texto}» viene de una fuente que pide no alterar su contenido. Junta en la otra palabra o anótalas como ideas distintas.",
  "ayuda.uso": "Sí, No o Pendiente.",
  "ayuda.disputa": "Sí: hay constancias contradictorias o se impugna una representación.",
  "ayuda.tipoPermiso": "Cómo se dio el permiso.",
  "ayuda.rolAutorizacion": "Su papel respecto a la obra o la comunidad.",
  "ayuda.tipoActor": "Qué clase de persona o institución es.",
  "ayuda.siNo": "Sí o No.",
  "ayuda.baseUso": "Con qué derecho se usa esta fuente.",
  "lenguas.alfabetoFuente.d": "(norma o gramática registrada en 📚 Fuentes; no reemplaza otras escrituras)",
  "lenguas.alfabetoFuente": "Alfabeto de referencia",
  "lengua.error.codigoIntl": "El código internacional de {lengua} debe tener 3 letras (ISO 639-3).",
  "lengua.codigoIntl.d": "(opcional, ISO 639-3, solo para exportar, ej. hch)",
  "lengua.codigoIntl": "Código internacional",
  "voces.f.escritura": "Sistema de escritura",
  "voces.f.comunidad.d": "(si no está, agrégala en 📚 Fuentes… → Comunidades)",
  "voces.col.comunidad": "Comunidad",
  "gob.aviso": "El equipo de este catálogo no representa a ningún pueblo ni comunidad. Aquí se muestra qué acuerdos existen, con quién y para qué.",
  "terminos.sinConsulta": "la autoridad o asamblea que la comunidad defina",
  "terminos.sinComunidad": "la que indique cada persona",
  "terminos.acepto": "Al aceptar, confirmas que eres la persona que habla o escribe, o que leíste estos términos a quien habla y estuvo de acuerdo.",
  "terminos.responde": "Responsable: {responsable} · {contacto}. Comunidad a la que pertenecen estas grabaciones: {comunidad}. Se consulta con: {consulta}.",
  "terminos.responde.t": "Quién responde.",
  "terminos.equipo": "El equipo de este catálogo no representa a ningún pueblo ni comunidad. Si alguien dice hablar en nombre de tu comunidad y no estás de acuerdo, puedes decírnoslo: se registra y se revisa.",
  "terminos.equipo.t": "Quién habla por el pueblo.",
  "terminos.retiro": "Puedes pedir en cualquier momento que se quiten tu voz o tus textos. Lo que ya se haya usado en una herramienta creada antes no siempre puede deshacerse.",
  "terminos.retiro.t": "Puedes retirarte.",
  "terminos.nunca": "No se clonarán voces. No se venderán los datos. No se entregarán a empresas ni a gobiernos sin el acuerdo de la comunidad.",
  "terminos.nunca.t": "Lo que nunca se hará.",
  "terminos.ia": "Si lo autorizas, tus grabaciones y textos pueden usarse para crear materiales didácticos y herramientas digitales, incluida inteligencia artificial, que ayuden a la comunidad a usar su lengua y reducir la brecha digital, como lo promueven la Agenda 2030 y el Decenio Internacional de las Lenguas Indígenas. Para eso también se necesita el acuerdo de tu comunidad.",
  "terminos.ia.t": "Herramientas digitales.",
  "terminos.publica": "Solo lo que el equipo revisó. Tu voz se publica con tu nombre, sin tu nombre, o no se publica, según lo que elijas.",
  "terminos.publica.t": "Qué se publica.",
  "terminos.guarda": "Lo que escribes o grabas, la lengua, y tu comunidad, tu variante y tu región tal como tú las nombras. Si lo autorizas, también tu nombre. Tus datos de contacto nunca se publican.",
  "terminos.guarda.t": "Qué se guarda.",
  "terminos.para": "Este catálogo reúne palabras, frases y voces para que la comunidad las conserve, las enseñe y las use.",
  "terminos.para.t": "Para qué sirve.",
  "terminos.titulo": "Términos de uso de {catalogo}",
  "terminos.version": "Versión {v}. Si cambia el texto o el interruptor de IA, cambia la versión: así se sabe qué aceptó cada persona.",
  "terminos.vistaComunidad": "Ver como los leería alguien de",
  "terminos.vista": "📄 Términos de uso",
  "ia.motivos": "Traducciones que aún no cuentan — sin revisar: {revision} · por la fuente: {fuente} · por la voz: {voz} · por la comunidad: {comunidad} · sin comunidad: {sinComunidad}",
  "ia.pend.voces": "Voces sin respuesta: {n}",
  "ia.pend.fuentes": "Fuentes sin autorización para IA: {n}",
  "ia.pend.comunidades": "Comunidades sin acuerdo otorgado: {n}",
  "ia.pendientes": "Lo que falta para IA",
  "ia.desactivado": "Desactivado: las preguntas sobre IA ya no se muestran.",
  "ia.activado": "Activado: la pregunta ya aparece en los formularios y abajo verás lo que falta.",
  "ia.solicitar": "Solicitar autorización para IA",
  "ia.d": "Mientras esté apagado, nadie ve preguntas sobre IA y todo queda en Pendiente, en silencio. Apagado no significa autorizado: nada pasa a Sí por sí solo. Esto no limita ningún otro uso del catálogo.",
  "ia.titulo": "🤖 Autorización para inteligencia artificial",
  "bloque.historialAntes": "filtro {filtro}",
  "bloque.historialId": "{n} traducciones",
  "bloque.error.fuente": "Elige la fuente que se va a asignar.",
  "bloque.listo": "✅ Fuente asignada a {n} traducción(es).",
  "bloque.cuenta": "Se cambiarían *{n}* traducción(es).",
  "bloque.aplicar": "Asignar",
  "bloque.contar": "Contar",
  "bloque.ubicacion": "Ubicación en la fuente",
  "bloque.destino": "Fuente",
  "bloque.asignar": "Asignarles",
  "bloque.creadoPor": "Capturó (correo)",
  "bloque.categoria": "Categoría",
  "bloque.sinFuente": "Sin fuente",
  "bloque.fuenteActual": "Fuente actual",
  "bloque.lote": "Importación (lote)",
  "bloque.cuales": "¿Cuáles traducciones?",
  "bloque.intro": "Asigna una fuente a muchas traducciones a la vez. Primero elige cuáles (puedes combinar filtros) y cuenta; luego aplica.",
  "comunidades.error.nombre": "Escribe el nombre de la comunidad.",
  "comunidades.acuerdoNuevo": "➕ Registrar acuerdo",
  "comunidades.acuerdos.d": "Puede haber varios (asamblea general, autoridades, comisariado). El catálogo no decide quién representa: registra lo que existe. Si se contradicen, queda En disputa.",
  "comunidades.acuerdos": "🤝 Acuerdos de la comunidad",
  "comunidades.f.disputa": "En disputa (hay constancias contradictorias o se impugna una representación)",
  "comunidades.f.consulta.d": "(asamblea, autoridad tradicional… como cargo, no como persona)",
  "comunidades.f.consulta": "A quién se consulta",
  "comunidades.f.region": "Región",
  "comunidades.f.otro": "Otro nombre",
  "comunidades.f.propio": "Nombre propio",
  "comunidades.nueva": "➕ Agregar comunidad",
  "comunidades.col.acuerdo": "Acuerdo para IA",
  "comunidades.col.consulta": "Se consulta con",
  "comunidades.col.variante": "Variante y región",
  "comunidades.col.nombre": "Comunidad",
  "comunidades.intro": "Las comunidades a las que pertenecen las voces y los textos, con el nombre que ellas usan. Registrar palabras no necesita el permiso de nadie más que de quien habla; el acuerdo de la comunidad solo se pide para usos colectivos como la IA.",
  "actores.error.nombre": "Escribe el nombre.",
  "actores.f.comunidad": "Comunidad (si es una autoridad o asamblea)",
  "actores.f.contacto.d": "(oficina o correo de la institución; no uses datos personales)",
  "actores.f.contacto": "Contacto institucional",
  "actores.f.credito.d": "(ej. «Universidad Autónoma de Nayarit»)",
  "actores.f.credito": "Cómo dar su crédito",
  "actores.f.tipo": "Tipo",
  "actores.f.nombre": "Nombre",
  "actores.nuevo": "➕ Agregar persona o institución",
  "actores.col.autoriza": "Qué autoriza",
  "actores.col.tipo": "Tipo",
  "actores.col.nombre": "Nombre",
  "actores.intro": "Personas, universidades, instituciones públicas, organizaciones y comunidades que autorizan o son autoras. Se registran una vez y se ligan a varias fuentes.",
  "autorizaciones.error.objeto": "La autorización debe ser sobre una fuente o una comunidad.",
  "autorizaciones.error.actor": "Elige la persona o institución que autoriza.",
  "autorizaciones.f.impugnada": "Impugnada (alguien la cuestiona: se congelan los usos colectivos)",
  "autorizaciones.f.vigente": "Vigente",
  "autorizaciones.f.condiciones": "Condiciones",
  "autorizaciones.f.fecha": "Fecha",
  "autorizaciones.f.constancia": "Enlace a la constancia (carta, acta, grabación)",
  "autorizaciones.f.alcance": "Qué autoriza",
  "autorizaciones.f.tipo": "Tipo de permiso",
  "autorizaciones.f.rol": "Su papel",
  "autorizaciones.f.actor.d": "(si no está, agrégala en 👤 Personas e instituciones)",
  "autorizaciones.f.actor": "Persona o institución",
  "autorizaciones.sinConstancia": "Sin constancia",
  "autorizaciones.verConstancia": "Ver constancia",
  "autorizaciones.impugnada": "Impugnada",
  "autorizaciones.noVigente": "Ya no vigente",
  "autorizaciones.col.permiso": "Constancia",
  "autorizaciones.col.quien": "Quién",
  "autorizaciones.editar": "Autorización",
  "autorizaciones.nueva": "➕ Agregar autorización",
  "autorizaciones.ninguna": "Todavía no hay autorizaciones registradas: lo que dependa de ellas queda Pendiente.",
  "autorizaciones.d": "Una fila por cada persona o institución con derechos sobre la obra. Un uso queda permitido cuando *todas* lo autorizan.",
  "autorizaciones.titulo": "✍️ Quiénes autorizan",
  "fuentes.corpus.aptas": "Con todos los permisos para IA",
  "fuentes.f.usosPropios": "Qué permite esta fuente",
  "fuentes.f.mostrarCredito": "Mostrar el crédito en el visor",
  "fuentes.f.fidelidad": "🔒 No se puede alterar su contenido (no se le aplican atajos de escritura; las correcciones se agregan como variante)",
  "fuentes.f.condiciones.d": "(cópialas tal cual, ej. «uso didáctico sin alterar su sentido ni contenido»)",
  "fuentes.f.condiciones": "Condiciones de uso",
  "fuentes.f.consulta": "Fecha de consulta",
  "fuentes.f.enlace": "Enlace",
  "fuentes.f.credito.d": "(tal como debe aparecer en el visor)",
  "fuentes.f.escritura": "Sistema de escritura",
  "fuentes.f.comunidad": "Comunidad",
  "fuentes.f.base.d": "(con qué derecho se usa)",
  "fuentes.f.base": "Base de uso",
  "fuentes.fiel": "No alterar",
  "fuentes.col.permisos": "Qué permite",
  "fuentes.col.base": "Base de uso",
  "perm.valor.otorgado": "Otorgado",
  "perm.valor.disputa": "En disputa",
  "perm.valor.pendiente": "Pendiente",
  "perm.valor.no": "No",
  "perm.valor.si": "Sí",
  "perm.uso.ia": "Herramientas de IA",
  "perm.uso.compartir": "Compartir con otros catálogos",
  "perm.uso.adaptar": "Adaptar o corregir",
  "perm.uso.didactico": "Juegos y materiales",
  "perm.uso.publicar": "Publicar",
  "perm.tab.corpus": "📊 Corpus y términos",
  "perm.tab.bloque": "🧩 Asignar en bloque",
  "perm.tab.comunidades": "🏘️ Comunidades",
  "perm.tab.actores": "👤 Personas e instituciones",
  "perm.tab.fuentes": "📚 Fuentes",
  "traer.fuenteNueva": "Creada al traer una hoja. Completa la autoría, quiénes autorizan y qué permiten.",
  "traer.avisoFuente": "📚 Se registraron {n} fuente(s) nueva(s). Revisa su autorización y sus permisos en 📚 Fuentes y corpus.",
  "menu.voces": "🎙️ Voces y permisos",
  "menu.fuentes": "📚 Fuentes y corpus",
  "ayuda.tipoRegistro": "Palabra, Frase u Oración. Si lo dejas vacío, el catálogo lo deduce solo.",
  "ayuda.original": "Original: se pensó en esta lengua. Traducción: viene de otra lengua.",
  "ayuda.permisoIA": "Permiso para usar estos datos en herramientas de inteligencia artificial. Es distinto del permiso para publicar.",
  "ayuda.formaPermiso": "Cómo se dio el permiso.",
  "ayuda.tipoFuente": "Qué clase de obra o material es.",
  "voces.titulo": "🎙️ Voces y permisos",
  "voces.intro": "Aquí se registra cada persona que presta su voz. Cada voz recibe un código (V-001) que es lo que se anota en las grabaciones. *Su nombre y su contacto no se guardan en el catálogo*: van a un directorio privado que solo ven quienes administran. Al catálogo solo pasa el nombre si la persona aceptó que se publique.",
  "voces.col.codigo": "Código",
  "voces.col.quien": "Quién",
  "voces.col.publicar": "Publicar su voz",
  "voces.col.ia": "Uso en IA",
  "voces.col.grabaciones": "Grabaciones",
  "voces.nueva": "➕ Agregar una voz",
  "voces.editar": "Editar",
  "voces.editando": "Voz {codigo}",
  "voces.ninguna": "Todavía no hay voces registradas.",
  "voces.inactiva": "(ya no participa)",
  "voces.privado.d": "🔒 Esto va al directorio privado: no se guarda en el catálogo ni se publica.",
  "voces.f.nombre": "Nombre completo",
  "voces.f.telefono": "Teléfono",
  "voces.f.correo": "Correo",
  "voces.f.comunidad": "Comunidad a la que pertenece",
  "voces.f.notasPrivadas": "Notas privadas",
  "voces.f.etiqueta": "Cómo la reconoce el equipo",
  "voces.f.etiqueta.d": "(sin nombre completo: iniciales o una seña, ej. «R.C., abuela de Tateikíe»)",
  "voces.f.lengua": "Lengua",
  "voces.f.variante": "Variante",
  "voces.f.lugar": "Lugar que se puede mostrar",
  "voces.f.lugar.d": "(solo se muestra si aceptó publicar con su nombre)",
  "voces.f.publicar": "¿Acepta que su voz se publique en el visor?",
  "voces.f.ia": "¿Acepta que su voz se use para herramientas de IA?",
  "voces.f.ia.d": "(es otro permiso; si no se ha preguntado, deja Pendiente)",
  "voces.f.forma": "Cómo dio el permiso",
  "voces.f.fecha": "Cuándo",
  "voces.f.evidencia": "Dónde está la constancia",
  "voces.f.evidencia.d": "(enlace al audio o al documento del permiso)",
  "voces.f.notas": "Notas",
  "voces.f.activo": "Sigue participando",
  "voces.guardar": "💾 Guardar",
  "voces.cancelar": "Cancelar",
  "voces.guardada": "✅ Guardado.",
  "voces.abrirDirectorio": "🔒 Abrir el directorio privado",
  "voces.abrirDirectorio.d": "(solo quienes administran pueden abrirlo)",
  "voces.directorio.archivo": "🔒 Directorio de voces — {catalogo} (privado)",
  "voces.directorio.carpeta": "🔒 Privado — {catalogo}",
  "voces.error.soloAdmin": "Solo quienes administran pueden ver y cambiar esto.",
  "voces.error.nombre": "Escribe al menos el nombre o cómo la reconoce el equipo.",
  "voces.error.sinDirectorio": "No puedes abrir el directorio privado de voces. Pide a quien lo creó que te comparta el archivo, o revisa que tu cuenta esté como administradora en 👥 Equipo.",
  "fuentes.titulo": "📚 Fuentes, permisos y comunidades",
  "fuentes.intro": "Cada obra o recurso de donde vienen las palabras. El permiso no se marca palabra por palabra: se calcula con la *base de uso* de la fuente y con quienes la autorizan. Cambias una autorización y se actualizan todas sus traducciones.",
  "fuentes.col.obra": "Obra",
  "fuentes.nueva": "➕ Agregar una fuente",
  "fuentes.f.titulo": "Título",
  "fuentes.f.autoria": "Autoría",
  "fuentes.f.anio": "Año",
  "fuentes.f.tipo": "Tipo",
  "fuentes.f.lenguas": "Lenguas",
  "fuentes.f.licencia": "Licencia",
  "fuentes.f.archivo": "Archivo de la obra",
  "fuentes.f.archivo.d": "(enlace en Drive; no se publica)",
  "fuentes.f.notas": "Notas",
  "fuentes.f.credito": "Cómo dar el crédito",
  "fuentes.guardada": "✅ Guardado.",
  "fuentes.equipo.titulo": "Captura del equipo",
  "fuentes.equipo.nota": "Todo lo que se captura sin indicar otra fuente.",
  "fuentes.error.titulo": "Escribe el título de la fuente.",
  "fuentes.corpus": "📊 Cómo va el corpus",
  "fuentes.corpus.d": "Solo es un conteo: no se exporta ni se comparte nada.",
  "fuentes.corpus.lengua": "Lengua",
  "fuentes.corpus.palabras": "Palabras",
  "fuentes.corpus.frases": "Frases",
  "fuentes.corpus.oraciones": "Oraciones",
  "fuentes.corpus.conFuente": "Con fuente",
  "fuentes.corpus.pares": "Pares alineados:",
  "plantilla.error.correos": "Escribe al menos un correo (ej. nombre@gmail.com).",
  "plantilla.error.enlace": "No se pudo abrir la hoja a cualquiera con el enlace; quizá tu escuela u organización no lo permite. Agrega los correos uno por uno.",
  "plantilla.error.noEsta": "No se encontró esa hoja en la carpeta COLABORACIONES.",
  "plantilla.mensaje": "Hola. Te comparto esta hoja para ayudar a escribir palabras: {nombre}\n{url}\nEn la primera pestaña vienen las instrucciones. ¡Gracias!",
  "plantilla.copiaManual": "El enlace ya está seleccionado abajo: cópialo con Ctrl+C (o mantén presionado en el teléfono).",
  "plantilla.copiado": "Copiado. Ya puedes pegarlo donde quieras.",
  "plantilla.enlace": "Enlace de la hoja",
  "plantilla.copiar": "📋 Copiar el enlace",
  "plantilla.correo": "✉️ Correo",
  "plantilla.whatsapp": "💬 WhatsApp",
  "plantilla.enviar": "Mandar el enlace",
  "plantilla.permisoEnlace": "Listo: quien tenga el enlace ya puede escribir. Ahora mándalo.",
  "plantilla.permisoCorreos": "Listo: {n} persona(s) ya pueden escribir. Ahora mándales el enlace.",
  "plantilla.abrirEnlace": "✔️ Permitir que escriba quien tenga el enlace",
  "plantilla.darPermiso": "✔️ Dar permiso",
  "plantilla.quien.enlace.d": "Es lo más fácil, pero si alguien reenvía el enlace, otras personas también podrán escribir. Lo que llenen no entra al catálogo hasta que lo traigas y lo revises.",
  "plantilla.quien.enlace": "Cualquier persona que reciba el enlace",
  "plantilla.quien.correos": "Solo las personas que yo agregue",
  "plantilla.quien": "¿Quién puede escribir en ella?",
  "plantilla.lista": "✅ La hoja está lista",
  "plantilla.nombre.d": "Así la verá quien la reciba. Puedes cambiarlo.",
  "plantilla.nombre": "Nombre de la hoja",
  "menu.coincidencias": "🔎 Palabras que se escriben igual",
  "comoVa.coincidencias": "🔎 {n} escritura(s) aparecen en más de una palabra. Puede ser una palabra con varios sentidos o la misma palabra capturada dos veces.",
  "comoVa.coincBoton": "Revisarlas",
  "coinc.titulo": "🔎 Palabras que se escriben igual",
  "coinc.intro": "Aquí aparecen las escrituras que están en más de una palabra, en la misma lengua. Para cada una decide:\n*Son ideas distintas* si la misma palabra nombra cosas diferentes (se deja así y no vuelve a aparecer).\n*Es la misma idea* si se capturó dos veces: se juntan en la palabra que marques. No se borra nada: lo que sobra queda como Retirado.",
  "coinc.ninguna": "🎉 No hay escrituras repetidas por revisar.",
  "coinc.grupo": "«{forma}» en {lengua} aparece en {n} palabras",
  "coinc.queda": "Juntar aquí",
  "coinc.sinImagen": "Sin imagen",
  "coinc.distintas": "Son ideas distintas",
  "coinc.unir": "Es la misma idea: juntar",
  "coinc.unir.d": "Al juntar, las traducciones pasan a la palabra marcada; si ya tenía esa lengua, las formas nuevas se agregan a su texto.",
  "coinc.anotado": "✅ Anotado: «{forma}» nombra ideas distintas. Ya no aparecerá aquí.",
  "coinc.unidas": "✅ Listo: {movidas} traducción(es) pasaron, {sumadas} se agregaron a una que ya estaba, {repetidas} estaban repetidas. {palabras} palabra(s) quedaron como Retirado.",
  "coinc.error.pocas": "Hacen falta al menos dos palabras.",
  "coinc.error.noEsta": "Una de esas palabras ya no está. Cierra esta ventana y vuelve a abrirla.",
  "panel.sinNombre": "(sin nombre)",
  "menu.titulo": "💠 CATÁLOGO",
  "menu.panel": "⚙️ Panel de control",
  "menu.comoVa": "📊 ¿Cómo va el catálogo?",
  "menu.estados": "📋 Revisar y publicar palabras",
  "menu.publicar": "📡 Mostrar los cambios en el visor",
  "menu.lenguas": "🗣️ Lenguas",
  "menu.equipo": "👥 Equipo",
  "menu.atajos": "✍️ Atajos de escritura",
  "menu.idioma": "🌐 Idioma de los menús",
  "menu.falta": "📝 Lo que falta",
  "menu.falta.lista": "📝 Lista de traducciones y audios que faltan",
  "menu.falta.guardar": "✅ Guardar lo que se llenó en la lista",
  "menu.falta.imagenes": "📷 Revisar las imágenes",
  "menu.falta.ilustrar": "🎨 Hoja para quien ilustra",
  "menu.compartir.hoja": "📄 Hoja para que alguien ayude a traducir",
  "menu.compartir.traer": "📥 Traer lo que llenaron en esa hoja",
  "menu.compartir.comunidades": "🌎 Compartir con otras comunidades",
  "menu.compartir.copia": "💾 Descargar una copia de seguridad",
  "menu.herr.primeros": "🚀 Primeros pasos (asistente)",
  "menu.herr.datos": "🛠️ Datos del catálogo y contacto",
  "menu.herr.direccion": "🔗 Visor: conexión y mantenimiento",
  "menu.herr.reparar": "🧩 Revisar y reparar el catálogo",
  "menu.herr.codigos": "🆔 Completar códigos que faltan",
  "menu.herr.aplicarAtajos": "🔄 Aplicar los atajos a todo lo escrito",
  "menu.herr.modoCaptura": "🧷 Escribir con atajos desde la app",
  "menu.herr.visor": "🔃 El visor no muestra los cambios",
  "inicio.mensaje": "👋 Para empezar: menú 💠 CATÁLOGO → 🔧 Herramientas → 🚀 Primeros pasos",
  "comun.abriendo": "⏳ Abriendo…",
  "comun.guardando": "⏳ Guardando…",
  "comun.cargando": "Cargando…",
  "comun.error": "❌ {m}",
  "panel.titulo": "⚙️ Panel del catálogo",
  "panel.catalogo": "Catálogo",
  "panel.grupo.diario": "Trabajo diario",
  "panel.comoVa": "📊 ¿Cómo va el catálogo?",
  "panel.comoVa.d": "Cuántas palabras hay, qué falta y qué juegos ya funcionan",
  "panel.estados": "📋 Revisar y publicar palabras",
  "panel.estados.d": "Elegir qué palabras se ven y cuáles todavía no",
  "panel.publicar": "📡 Mostrar los cambios en el visor",
  "panel.publicar.d": "Se hace sola cada hora; aquí, al momento",
  "panel.lenguas": "🗣️ Lenguas",
  "panel.lenguas.d": "Agregar una lengua o cambiar su nombre y sus letras",
  "panel.equipo": "👥 Equipo",
  "panel.equipo.d": "Quién captura y quién revisa",
  "panel.atajos": "✍️ Atajos de escritura",
  "panel.atajos.d": "Por ejemplo: escribir + para que salga ɨ",
  "panel.falta": "📝 Lista de traducciones y audios que faltan",
  "panel.hojaAyuda": "📄 Hoja para que alguien ayude a traducir",
  "panel.traer": "📥 Traer lo que llenaron en esa hoja",
  "panel.comunidades": "🌎 Compartir con otras comunidades",
  "panel.copia": "💾 Descargar una copia de seguridad",
  "panel.datos": "🛠️ Datos del catálogo y contacto",
  "panel.datos.d": "Nombre, a quién escribir, créditos",
  "panel.idioma": "🌐 Idioma de los menús",
  "panel.idioma.d": "Ver los menús en otra lengua o traducirlos",
  "panel.direccion": "🔗 Visor: conexión y mantenimiento",
  "panel.primeros": "🚀 Primeros pasos (asistente)",
  "panel.reparar": "🧩 Revisar y reparar el catálogo",
  "panel.reparar.d": "Si algo dejó de funcionar. No borra nada.",
  "panel.hojasInternas": "👁️ Mostrar u ocultar las hojas internas",
  "panel.nuevo.titulo": "Catálogo nuevo",
  "panel.nuevo.texto": "Para empezar, toca *🚀 Primeros pasos* (abajo). Te hace unas preguntas y prepara todo; no tienes que escribir en las tablas.",
  "panel.resumen.cuentas": "🗣️ {lenguas} lengua(s) · 📚 {palabras} palabra(s) · 📝 {traducciones} traducción(es)",
  "panel.resumen.estados": "✅ {publicadas} se ven en el visor · 🔍 {revision} por revisar · ✏️ {borrador} en borrador",
  "panel.resumen.retiradas": " · 🗄️ {n} retirada(s)",
  "panel.resumen.publicado": "📡 El visor se actualizó: {fecha}",
  "panel.resumen.nunca": "nunca",
  "panel.resumen.conClave": "🔒 Para consultar se necesita una clave",
  "panel.resumen.sinClave": "🔓 Cualquiera con el enlace puede consultar",
  "panel.resumen.captura": "🧷 Los atajos de escritura funcionan desde la app",
  "panel.hojas.ocultas": "Listo: las hojas internas quedaron ocultas.",
  "panel.hojas.visibles": "Las hojas internas están visibles. Es mejor no escribir en ellas: usa este panel.",
  "atajos.titulo": "✍️ Atajos de escritura",
  "atajos.intro": "Un atajo cambia solo lo que se escribe. Sirve cuando el teclado no tiene una letra: por ejemplo, escribir *+* para que quede *ɨ*. Se aplica a lo que ya está escrito y a lo que se escriba después.",
  "atajos.buscar": "Cuando se escriba",
  "atajos.reemplazar": "Que se cambie por",
  "atajos.lengua": "Lengua",
  "atajos.todas": "Todas las lenguas",
  "atajos.paraQuien": "¿Para quién?",
  "atajos.todos": "Para todo el equipo",
  "atajos.mio": "Solo para lo que yo escriba",
  "atajos.mayus": "Distinguir mayúsculas y minúsculas",
  "atajos.mayus.d": "(así “O9” y “o9” son atajos distintos)",
  "atajos.regex": "Avanzado, para especialistas: es un patrón de texto (expresión regular)",
  "atajos.probar": "🔍 Ver qué cambiaría",
  "atajos.guardar": "💾 Guardar el atajo",
  "atajos.guardados": "Atajos guardados",
  "atajos.revisando": "⏳ Revisando…",
  "atajos.cambiaria": "Cambiaría *{n}* palabra(s). Ejemplos:",
  "atajos.sinCorreo": "(no se pudo saber tu correo)",
  "atajos.ninguno": "Todavía no hay atajos.",
  "atajos.col.escribe": "Se escribe",
  "atajos.col.queda": "Queda",
  "atajos.col.lengua": "Lengua",
  "atajos.col.para": "Para",
  "atajos.patron": "(patrón)",
  "atajos.apagar": "Apagar",
  "atajos.encender": "Encender",
  "atajos.aplicando": "⏳ Aplicando…",
  "atajos.vacio": "(vacío)",
  "atajos.paraTodos": "Todo el equipo",
  "atajos.soloYo": "Solo yo",
  "atajos.todasCorto": "Todas",
  "atajos.sinLengua": "(sin lengua)",
  "atajos.error.lengua": "Elige una lengua.",
  "atajos.error.correo": "No se pudo saber tu correo. Guarda el atajo para todo el equipo.",
  "atajos.error.vacio": "Escribe qué se va a teclear (por ejemplo +).",
  "atajos.error.iguales": "Lo que se teclea y lo que debe quedar son iguales: el atajo no cambiaría nada.",
  "atajos.error.patron": "El patrón avanzado no está bien escrito: {m}",
  "atajos.error.noExiste": "No se encontró ese atajo.",
  "atajos.error.inestable": "Este atajo solo se puede encender si los atajos funcionan desde la app (🔧 Herramientas → Escribir con atajos desde la app).",
  "atajos.aviso.inestableCaptura": "Ojo: este atajo vuelve a cambiar su propio resultado (por ejemplo, “a” → “aa”). Funciona, pero revisa que sea lo que quieres.",
  "atajos.aviso.inestable": "Este atajo vuelve a cambiar su propio resultado (por ejemplo, “a” → “aa” terminaría en “aaaa…”). Para poder usarlo, primero activa 🔧 Herramientas → Escribir con atajos desde la app.",
  "atajos.aviso.sinAutor": "Los atajos “solo para mí” necesitan saber quién escribió cada palabra. Actívalo en 🔧 Herramientas → Escribir con atajos desde la app.",
  "atajos.aviso.nada": "Por ahora no cambia nada de lo escrito. Se usará con lo que se escriba después.",
  "atajos.guardado": "Atajo guardado. Palabras corregidas: {n}.",
  "atajos.nota": "Creado desde el panel por {quien} ({fecha})",
  "atajos.notaAsistente": "Atajo definido en el asistente",
  "atajos.error.formato": "No entiendo el atajo \"{atajo}\". Escríbelo así: lo que se teclea = lo que debe quedar (ej. +=ɨ).",
  "captura.titulo": "Escribir con atajos desde la app",
  "captura.texto": "Con esto, lo que el equipo escriba en la app con atajos (por ejemplo +) se guarda ya corregido (ɨ), y también se guarda lo que escribió tal cual.\n\nQué va a cambiar en la hoja TRADUCCIONES:\n• Aparece la columna texto_captura: ahí queda lo que se escribió.\n• La columna texto se llena sola, ya corregida. Si alguien la cambia a mano, se vuelve a calcular.\n\nDespués, en la app de AppSheet, alguien del equipo debe:\n1. Volver a leer la tabla TRADUCCIONES (botón Regenerate).\n2. Escribir en texto_captura y dejar texto sin editar.\n\n¿Continuar?",
  "captura.listo": "✅ Listo: los atajos ya funcionan desde la app.\nColumnas nuevas: {columnas}\nPalabras corregidas: {n}.",
  "captura.yaEstaban": "ninguna (ya estaban)",
  "atajos.aplicados": "✅ Listo: se aplicaron los atajos a todo lo escrito. Palabras corregidas: {n}.",
  "comun.faltaHoja": "No se encuentra la hoja {hoja}. Usa 🔧 Herramientas → 🧩 Revisar y reparar el catálogo.",
  "comun.faltanColumnas": "A la hoja {hoja} le faltan columnas ({columnas}). Usa 🔧 Herramientas → 🧩 Revisar y reparar el catálogo.",
  "comun.sinPreparar": "El catálogo todavía no está preparado. Usa 🔧 Herramientas → 🚀 Primeros pasos.",
  "mesa.titulo": "📝 Lo que falta",
  "mesa.encabezado": "📝 Lista de lo que falta",
  "mesa.intro": "Se arma una hoja con cada palabra a la que le falta traducción o audio. Ahí puedes escribir las traducciones, o tocar *📱 Abrir en la app* para grabar.",
  "mesa.lenguas": "¿Qué lenguas quieres trabajar hoy?",
  "mesa.buscar": "¿Qué buscar?",
  "mesa.sinTraducir": "📝 Palabras sin traducir",
  "mesa.sinAudio": "🔊 Palabras sin audio",
  "mesa.armar": "Armar la lista ➔",
  "mesa.buscando": "⏳ Buscando…",
  "mesa.sinReferencia": "(sin texto en la lengua de referencia)",
  "mesa.tarea.traduccion": "📝 FALTA TRADUCCIÓN",
  "mesa.tarea.audio": "🔊 FALTA AUDIO",
  "mesa.nadaFalta": "🎉 No falta nada en las lenguas que elegiste.",
  "mesa.col.codigo": "CÓDIGO",
  "mesa.col.referencia": "PALABRA DE REFERENCIA",
  "mesa.col.actual": "TEXTO ACTUAL",
  "mesa.col.codigoLengua": "CÓDIGO LENGUA",
  "mesa.col.lengua": "LENGUA",
  "mesa.col.falta": "QUÉ FALTA",
  "mesa.col.nueva": "ESCRIBE AQUÍ LA TRADUCCIÓN",
  "mesa.col.nota": "NOTA",
  "mesa.col.estado": "ESTADO",
  "mesa.col.abrir": "GRABAR O EDITAR",
  "mesa.valor.pendiente": "PENDIENTE",
  "mesa.valor.guardado": "GUARDADO",
  "mesa.valor.omitir": "OMITIR",
  "mesa.abrirApp": "📱 Abrir en la app",
  "mesa.sinApp": "(falta el código de la app en Datos del catálogo)",
  "mesa.primeroLista": "⚠️ Primero arma la lista: 💠 CATÁLOGO → 📝 Lo que falta → Lista de traducciones y audios que faltan.",
  "mesa.preguntaTitulo": "¿Mostrar ya las traducciones nuevas?",
  "mesa.preguntaTexto": "SÍ = se ven en el visor desde ahora.\nNO = quedan \"por revisar\" y no se ven hasta que alguien las revise.\n\nLas traducciones que ya existían no cambian.",
  "mesa.guardadoListo": "✅ Listo, se guardó lo que escribieron en la lista.\nTraducciones nuevas ({estado}): {nuevas}\nTraducciones corregidas: {corregidas}",
  "imagenes.col.codigo": "CÓDIGO",
  "imagenes.col.palabra": "PALABRA",
  "imagenes.col.categoria": "CATEGORÍA",
  "imagenes.col.tiene": "¿TIENE IMAGEN?",
  "imagenes.col.agregar": "AGREGAR IMAGEN",
  "imagenes.si": "✅ Sí",
  "imagenes.falta": "❌ Falta",
  "ilustrar.todas": "🎉 Todas las palabras que se ven en el visor ya tienen imagen.",
  "ilustrar.archivo": "Imágenes pendientes · {fecha}",
  "plantilla.titulo": "📄 Hoja para ayudar a traducir",
  "plantilla.encabezado": "📄 Hoja para que alguien ayude a traducir",
  "plantilla.intro": "Se crea una hoja aparte que puedes compartir con alguien que no está en el equipo. Cuando la llene, la traes con *📥 Traer lo que llenaron*.",
  "plantilla.paraQue": "¿Para qué es?",
  "plantilla.nuevas": "Para escribir palabras nuevas (hoja en blanco)",
  "plantilla.faltantes": "Para completar lo que les falta a las palabras que ya están",
  "plantilla.lenguas": "¿En qué lenguas?",
  "plantilla.otraLengua": "¿Va a escribir en una lengua que todavía no está? (opcional)",
  "plantilla.otraNombre": "Nombre (ej. Cora)",
  "plantilla.otraCodigo": "Código, ej. crn",
  "plantilla.otraCodigo.d": "Código corto de la lengua (ISO). Si no lo sabes, escribe unas letras del nombre.",
  "plantilla.correos": "Correo de quien la va a llenar (si son varias personas, sepáralos con coma):",
  "plantilla.correos.d": "Si son varias, sepáralas con coma. Cada persona necesita una cuenta de Google (por ejemplo, Gmail) para escribir.",
  "plantilla.crear": "Crear la hoja ➔",
  "plantilla.preparando": "⏳ Preparando la hoja…",
  "plantilla.abrir": "✅ Abrir la hoja",
  "plantilla.abrirN": "✅ Abrir la hoja ({n} palabras)",
  "plantilla.carpeta": "📁 Queda guardada en la carpeta COLABORACIONES, junto al catálogo.",
  "plantilla.archivo.nuevas": "Palabras nuevas · {fecha}",
  "plantilla.archivo.faltantes": "Completar traducciones · {fecha}",
  "plantilla.hojaInstrucciones": "INSTRUCCIONES",
  "plantilla.inst.1": "Cómo llenar esta hoja",
  "plantilla.inst.2": "1. Ve a la pestaña DATOS_A_IMPORTAR.",
  "plantilla.inst.3": "2. Escribe cada palabra en la columna TRADUCCION de su lengua (por ejemplo TRADUCCION_HCH). Si se dice de varias formas, sepáralas con coma: rancho, casa.",
  "plantilla.inst.4": "3. Si quieres explicar algo (otra forma de decirla, de dónde es), escríbelo en NOTA. En DESCRIPCION puedes contar, en esa misma lengua, qué es o para qué se usa (es opcional).",
  "plantilla.inst.5": "4. En CATEGORIA puedes poner el grupo de la palabra (animales, plantas, casa…).",
  "plantilla.inst.6": "5. No cambies ni borres la columna ID_PALABRA ni los títulos de la primera fila.",
  "plantilla.inst.7": "6. Las columnas AUDIO e IMG_STATUS solo informan: no hace falta llenarlas.",
  "plantilla.inst.8": "Cuando termines, avisa a quien te compartió la hoja.",
  "plantilla.inst.9": "Gracias por ayudar a cuidar la lengua.",
  "plantilla.sinBase": "(sin texto de referencia)",
  "plantilla.audioOk": "✅ OK",
  "plantilla.audioDrive": "🔍 En Drive",
  "traer.titulo": "📥 Traer lo que llenaron",
  "traer.intro": "Pega el enlace de la hoja que alguien llenó para ayudar a traducir (la que se creó con *📄 Hoja para que alguien ayude a traducir*). No se repite nada que ya esté.",
  "traer.enlace": "Pega aquí el enlace de la hoja",
  "traer.comoQuedan": "¿Cómo deben quedar las palabras nuevas?",
  "traer.revision": "Por revisar (todavía no se ven en el visor)",
  "traer.borrador": "Borrador (no se ven)",
  "traer.publicado": "Ya se ven en el visor",
  "traer.nota": "Lo que ya estaba en el catálogo no cambia de estado.",
  "traer.boton": "📥 Traer las palabras",
  "traer.faltaEnlace": "Primero pega el enlace de la hoja.",
  "traer.trayendo": "⏳ Trayendo las palabras… puede tardar un minuto.",
  "traer.vacia": "Esa hoja todavía no tiene palabras escritas.",
  "traer.avisoBase": "No se guardó la columna CONCEPTO_BASE porque el catálogo no tiene una lengua de referencia. Elígela en 🛠️ Datos del catálogo.",
  "traer.avisoCat": "No se guardaron los nombres de las categorías en otras lenguas (columnas TRAD_CAT). Usa 🧩 Revisar y reparar el catálogo y vuelve a intentarlo.",
  "traer.listo": "✅ Listo",
  "traer.resumen": "Palabras nuevas: {palabras}\nTraducciones nuevas: {nuevas} · corregidas: {corregidas}\nCategorías nuevas: {categorias} · nombres de categoría en otras lenguas: {nombresCat}\nLenguas nuevas: {lenguas}",
  "traer.lenguaNueva": "Lengua {codigo}",
  "publicar.no.titulo": "⚠️ Hay traducciones que todavía no se ven:",
  "publicar.no.estado": "{n} traducción(es) no están en «Publicado». Cámbialas en 📋 Revisar y publicar palabras (marca «cambiar también sus traducciones»).",
  "publicar.no.palabra": "{n} traducción(es) son de palabras que no están en «Publicado».",
  "publicar.no.fuente": "{n} traducción(es) de la fuente «{fuente}» esperan permiso para publicarse. Regístralo en 📚 Fuentes y corpus.",
  "publicar.no.fuenteBorrada": "{n} traducción(es) apuntan a una fuente que ya no existe (se borró). Asígnales su fuente en 📚 Fuentes y corpus → Asignar en bloque.",
  "publicar.listo": "✅ Listo: el visor ya muestra los cambios.\n\nEn el visor hay:\nPalabras: {palabras}\nTraducciones: {traducciones}\nLenguas: {lenguas}\n\nSolo se ven las palabras marcadas como Publicado. Quien ya tenía el visor abierto lo ve al recargarlo.",
  "direccion.etiqueta": "Dirección de la aplicación web",
  "direccion.sinComprobar": "⚠️ Esta dirección la sugiere Google y a veces no es la correcta. Tócala en «Probar»; si falla, copia la de Implementar → Gestionar implementaciones y guárdala aquí.",
  "direccion.pega": "Pega aquí la dirección de tu aplicación web (Implementar → Gestionar implementaciones → Aplicación web → URL) y guárdala.",
  "direccion.probar": "🔍 Probar",
  "direccion.claveTitulo": "Clave",
  "direccion.dondeEsta": "¿Dónde encuentro la dirección o cómo la creo?",
  "direccion.error.biblioteca": "Esa es la dirección de una *biblioteca*, no de la aplicación web. En Implementar → Gestionar implementaciones elige la de tipo «Aplicación web» y copia su URL (termina en /exec).",
  "direccion.error.dev": "Esa dirección (termina en /dev) solo sirve para pruebas de quien programa. Usa la que termina en /exec.",
  "direccion.error.formato": "Esa no parece la dirección de una aplicación web: debe verse como https://script.google.com/macros/s/…/exec",
  "direccion.prueba.ok": "✅ La dirección responde bien: el visor verá {palabras} palabra(s) publicadas.",
  "direccion.prueba.otroCatalogo": "⚠️ Pero responde con OTRO catálogo (por ejemplo, el original de una copia). Esta hoja necesita su propia implementación.",
  "direccion.prueba.acceso": "La dirección pide iniciar sesión: en Gestionar implementaciones → ✏️ pon «Quién tiene acceso: Cualquier persona» y Versión: Nueva.",
  "direccion.prueba.html": "La dirección no respondió como un catálogo (puede ser una versión vieja del programa). En Gestionar implementaciones → ✏️ elige Versión: Nueva e Implementar.",
  "direccion.prueba.noExiste": "Esa dirección es de una implementación que ya no existe o está archivada. En Implementar → Gestionar implementaciones copia la dirección de la implementación ACTIVA (termina en /exec) y pégala aquí.",
  "direccion.prueba.sinDoGet": "Esa dirección es de otro proyecto o de una versión sin el programa del catálogo. Revisa que sea la implementación de ESTE archivo y elige Versión: Nueva.",
  "direccion.prueba.autorizar": "Falta autorizar el programa: abre Extensiones → Apps Script, ejecuta cualquier función (por ejemplo onOpen), acepta los permisos y vuelve a probar.",
  "direccion.prueba.detalle": "Lo que respondió Google (código {codigo}): «{texto}»",
  "conexion.guardadoPeroPrueba": "Se guardó, pero al probar la dirección del catálogo:",
  "direccion.prueba.otraClave": "La dirección responde, pero con otra clave: probablemente es la de otro catálogo, o la implementación es vieja (Versión: Nueva).",
  "direccion.prueba.sinRespuesta": "No hubo respuesta de esa dirección. Revisa que esté completa.",
  "direccion.prueba.error": "La dirección respondió con un error: {error}",
  "conexion.intro": "Para ver este catálogo en el visor se necesitan *la dirección* y, si la hay, *la clave*. Aquí están las dos, y abajo un mensaje listo para compartirlas.",
  "conexion.visor": "Dirección del visor",
  "conexion.visor.d": "Dónde está publicado el visor (por ejemplo, en GitHub Pages). Si la guardas, el mensaje incluye un enlace que abre el catálogo directo, sin escribir nada.",
  "conexion.guardar": "💾 Guardar",
  "conexion.guardado": "Guardado.",
  "conexion.claveEs": "El catálogo pide esta clave:",
  "conexion.claveCrear": "Poner una clave",
  "conexion.claveNueva": "Cambiar por una clave nueva",
  "conexion.claveQuitar": "Quitar la clave",
  "conexion.confirmaNueva": "Quien tenga la clave anterior (o el enlace con ella) dejará de entrar hasta que le compartas la nueva. ¿Seguir?",
  "conexion.confirmaQuitar": "Cualquier persona con la dirección podrá ver lo publicado. ¿Seguir?",
  "conexion.siCambiar": "Sí, cambiar",
  "conexion.cancelar": "Cancelar",
  "conexion.claveLista": "Listo: la clave nueva ya rige. Comparte otra vez los datos de conexión.",
  "conexion.claveQuitada": "Listo: el catálogo ya no pide clave.",
  "conexion.compartir": "Compartir el acceso",
  "conexion.compartir.d": "Un mensaje con todo lo necesario para abrir el catálogo en el visor.",
  "conexion.copiar": "Copiar",
  "conexion.correo": "Correo",
  "conexion.copiado": "Mensaje copiado.",
  "conexion.asunto": "Acceso al catálogo «{catalogo}»",
  "conexion.enlaceDirecto": "Enlace directo:",
  "conexion.avisoClave": "🔒 El mensaje incluye la clave: compártelo solo con quien deba ver el catálogo. Si se filtra, cámbiala aquí arriba.",
  "conexion.sinDireccion": "Primero escribe y guarda la dirección de la aplicación web (paso 1).",
  "conexion.mensaje": "Para ver el catálogo «{catalogo}»:",
  "conexion.mensaje.enlace": "Abre este enlace: {enlace}",
  "conexion.mensaje.oManual": "O abre {visor}, toca 🔌 y escribe:",
  "conexion.mensaje.abreVisor": "Abre el visor y toca 🔌. Escribe:",
  "conexion.mensaje.direccion": "Dirección: {url}",
  "conexion.mensaje.clave": "Clave: {clave}",
  "conexion.mensaje.sinClave": "Clave: no hace falta.",
  "conexion.error.visor": "La dirección del visor debe empezar con https:// (por ejemplo, https://usuario.github.io/catalogo/).",
  "conexion.error.accion": "Elige poner, cambiar o quitar la clave.",
  "direccion.titulo": "🔗 Visor: conexión y mantenimiento",
  "direccion.sinClave": "Este catálogo no pide clave: con la dirección basta.",
  "direccion.actualizar": "*Para quien actualice el programa:* en Extensiones → Apps Script → Implementar → Gestionar implementaciones → ✏️ → Versión: Nueva. Así la dirección no cambia.",
  "direccion.paso1": "Menú *Extensiones → Apps Script*.",
  "direccion.paso2": "Arriba a la derecha: *Implementar → Nueva implementación*.",
  "direccion.paso3": "En el engrane ⚙️ elige *Aplicación web*.",
  "direccion.paso4": "“Ejecutar como”: *Yo*. “Quién tiene acceso”: *Cualquier persona*.",
  "direccion.paso5": "*Implementar*, acepta los permisos y copia la *URL de la aplicación web* (termina en */exec*). No uses la de «Biblioteca».",
  "codigos.listo": "✅ Listo: {n} fila(s) no tenían código y ya lo tienen.",
  "codigos.bien": "✅ Todo bien: todas las filas tienen su código.",
  "reparar.listo": "✅ Listo: el catálogo está revisado. No se borró ni se movió nada.",
  "reparar.enOrden": "Todo estaba en orden.",
  "reparar.hojas": "Hojas nuevas: {hojas}",
  "reparar.columnas": "Columnas nuevas: {columnas}",
  "reparar.appsheet": "⚠️ Aviso para quien administra la app de AppSheet: abre esas tablas y toca Regenerate para que la app vea las columnas nuevas.",
  "reparar.equipo": "👥 {correo} quedó como administrador(a) del equipo.",
  "reparar.cadaHora": "Cada hora el catálogo se revisa y el visor se actualiza solo.",
  "reparar.visorListo": "✅ Listo: el visor se volvió a preparar desde cero. Si todavía no ves los cambios, recarga el visor.",
  "comun.noTermino": "❌ No se pudo terminar: {m}",
  "copia.error": "❌ No se pudo hacer la copia de seguridad. Intenta de nuevo en unos minutos.",
  "copia.texto": "Se guardó una copia de todo el catálogo en la carpeta *RESPALDOS* de tu Drive. Se abre con Excel o con Hojas de cálculo.",
  "copia.abrir": "💾 Abrir la copia",
  "copia.titulo": "💾 Copia de seguridad",
  "asis.titulo": "🚀 Primeros pasos",
  "asis.intro": "Responde estas preguntas y el catálogo queda listo. *No tienes que escribir en las tablas.* Todo se puede cambiar después.",
  "asis.yaHay": "Ya hay {n} lengua(s). Las que agregues aquí se suman, sin repetirse.",
  "asis.nombre": "Nombre del catálogo",
  "asis.nombre.ej": "Ej. Glosario de la Sierra",
  "asis.frase": "Una frase que lo describa",
  "comun.opcional": "(opcional)",
  "asis.cuantas": "¿Cuántas lenguas vas a registrar ahora?",
  "asis.cuantas.d": "(puedes agregar más después)",
  "asis.referencia": "Lengua de referencia",
  "asis.referencia.d": "(opcional). Se usa para nombrar las palabras en estas ventanas; ninguna lengua es obligatoria",
  "asis.grupos": "Grupos de palabras",
  "asis.grupos.d": "(opcional, separados por coma: animales, plantas, cuerpo…)",
  "asis.opciones": "Opciones",
  "asis.carpetas": "Crear carpetas en Drive para guardar audios y videos de cada lengua",
  "asis.clave": "Pedir una clave para consultar",
  "asis.clave.d": "(si no, cualquiera con el enlace ve lo publicado)",
  "asis.ocultar": "Ocultar las hojas internas que no hace falta tocar",
  "asis.appsheet": "Código de tu app de AppSheet",
  "asis.appsheet.d": "(opcional; si no lo tienes, déjalo vacío)",
  "asis.preparar": "🚀 Preparar el catálogo",
  "asis.lengua": "Lengua {n}",
  "lengua.nombrePropio": "Nombre propio",
  "lengua.nombrePropio.d": "(como la comunidad nombra su lengua)",
  "lengua.nombreOtra": "Nombre en otra lengua",
  "lengua.nombreOtra.d": "(opcional, ej. en español)",
  "lengua.codigo": "Código de la lengua",
  "lengua.codigo.d": "(el que usa la comunidad, 3 letras, ej. wix)",
  "lengua.variante": "Variante o comunidad",
  "lengua.atajos": "Atajos de escritura",
  "lengua.atajos.d": "(opcional). Si el teclado no tiene una letra: lo que se teclea = lo que debe quedar, separados por ;  Ej. +=ɨ; O9=Ö",
  "comun.ninguna": "Ninguna",
  "asis.preparando": "⏳ Preparando todo… puede tardar un minuto.",
  "asis.listo": "✅ Catálogo listo: {nombre}",
  "asis.abrirCarpeta": "📁 Abrir la carpeta del catálogo en Drive",
  "asis.siguientes": "Siguientes pasos",
  "asis.demas": "Lo demás (lenguas, equipo, atajos) está en el menú 💠 CATÁLOGO → ⚙️ Panel de control.",
  "asis.paso.direccion": "Dale al catálogo su dirección en internet (una sola vez): en Extensiones → Apps Script → Implementar → Nueva implementación → Aplicación web, con «Ejecutar como: Yo» y «Quién tiene acceso: Cualquier persona». Copia la dirección que termina en /exec y pégala en ⚙️ Panel → 🔗 Visor: conexión y mantenimiento. Para actualizar el programa después, usa Gestionar implementaciones → ✏️ → Versión: Nueva (la dirección no cambia).",
  "asis.paso.appsheet": "Si usas la app de AppSheet, en ⚙️ Panel → 🛠️ Datos del catálogo pega el código de tu app (está en la dirección de la app). Así los botones “Abrir en la app” funcionan.",
  "asis.paso.equipo": "Agrega a tu equipo en ⚙️ Panel → 👥 Equipo.",
  "asis.paso.clave": "La clave para consultar el catálogo es: {clave}. Compártela solo con quien deba ver el visor (también está en Datos del catálogo).",
  "asis.resumen.lenguas": "{n} lengua(s)",
  "asis.resumen.categorias": "{n} categoría(s) nuevas",
  "asis.resumen.atajos": "{n} atajo(s) de escritura",
  "asis.resumen.carpetas": "carpetas en Drive listas",
  "asis.resumen.sinCarpetas": "sin carpetas en Drive",
  "asis.error.nombre": "Escribe el nombre del catálogo.",
  "asis.error.lengua": "Escribe al menos una lengua (basta con su nombre).",
  "lengua.error.nombre": "Escribe al menos un nombre para cada lengua.",
  "lengua.error.codigo": "{lengua}: el código de la lengua lleva 3 letras (ej. hch). Si no lo sabes, déjalo vacío.",
  "lengua.error.glotto": "{lengua}: el código de Glottolog lleva 4 letras y 4 números (ej. huic1243). Si no lo sabes, déjalo vacío.",
  "lengua.laLengua": "la lengua",
  "lenguas.titulo": "🗣️ Lenguas del catálogo",
  "lenguas.agregar": "Agregar una lengua",
  "lenguas.editar": "Editar: {lengua}",
  "lenguas.glotto": "Código de Glottolog",
  "lenguas.glotto.d": "(opcional, para especialistas)",
  "lenguas.region": "Región o lugar",
  "lenguas.letras": "Letras de la lengua",
  "lenguas.letras.d": "(separadas por espacios; las letras dobles van juntas: a ch e i ɨ k kw ꞌ … Las usan los juegos de letras)",
  "lenguas.sugerir": "✨ Sugerir",
  "lenguas.sugerir.d": "Propone las letras que aparecen en las palabras ya escritas",
  "lenguas.atajosNuevos": "Atajos de escritura nuevos",
  "lenguas.atajosNuevos.d": "(ej. +=ɨ; O9=Ö. Los que ya existen se ven en ✍️ Atajos de escritura)",
  "lenguas.notas": "Notas",
  "lenguas.estado": "Estado",
  "lenguas.estado.publicado": "Se ve en el visor",
  "lenguas.estado.borrador": "Borrador (no se ve)",
  "lenguas.estado.revision": "Por revisar (no se ve)",
  "lenguas.estado.retirado": "Retirada (no se ve, pero no se borra nada)",
  "lenguas.guardar": "💾 Guardar lengua",
  "lenguas.nueva": "Nueva",
  "lenguas.primeroGuardar": "Primero guarda la lengua. Después podrás pedir la sugerencia de letras.",
  "lenguas.revisandoPalabras": "⏳ Revisando las palabras…",
  "lenguas.revisaPropuesta": "Revisa la propuesta: agrega las letras dobles (ch, tz, kw…) y quita lo que no sea una letra de la lengua. Mejor si lo revisas con la comunidad.",
  "lenguas.ninguna": "Todavía no hay lenguas.",
  "lenguas.col.lengua": "Lengua",
  "lenguas.col.palabras": "Palabras",
  "lenguas.col.audios": "Audios",
  "lenguas.col.carpetas": "Carpetas",
  "lenguas.audios": "audios",
  "lenguas.videos": "videos",
  "lenguas.editarBoton": "Editar",
  "lenguas.guardada": "Lengua guardada.",
  "ajustes.titulo": "🛠️ Datos del catálogo y contacto",
  "ajustes.frase": "Una frase que lo describa",
  "ajustes.referencia.d": "(opcional; se usa para nombrar las palabras en estas ventanas)",
  "ajustes.appsheet.d": "(está en la dirección de la app, después de appId=. Sirve para los botones “Abrir en la app”)",
  "ajustes.contacto": "📬 A quién escribir",
  "ajustes.contacto.d": "(se ve en el visor: usa datos del proyecto, no personales)",
  "ajustes.responsable": "Persona o equipo responsable",
  "ajustes.correo": "Correo",
  "ajustes.telefono": "Teléfono para llamadas",
  "ajustes.whatsapp": "WhatsApp",
  "ajustes.whatsapp.d": "(con la clave del país, ej. 52 para México: 523111234567)",
  "ajustes.horario": "Horario de atención",
  "ajustes.invitacion": "Mensaje para invitar a colaborar",
  "ajustes.invitacion.ej": "¿Conoces palabras que no están? ¿Quieres grabar tu voz o compartir imágenes? Escríbenos.",
  "ajustes.acerca": "📖 Acerca del catálogo",
  "ajustes.acerca.d": "Sobre la lengua, la comunidad y el proyecto",
  "ajustes.creditos": "Créditos y agradecimientos",
  "ajustes.creditos.d": "(solo personas que dieron su consentimiento)",
  "ajustes.licencia": "Cómo se puede usar este catálogo",
  "ajustes.licencia.d": "(ej. \"Uso educativo y comunitario; no comercial\")",
  "ajustes.clave": "Clave para consultar el visor",
  "ajustes.clave.igual": "Dejar como está",
  "ajustes.clave.nueva": "Crear una clave nueva (quien tenga la anterior ya no podrá entrar)",
  "ajustes.clave.quitar": "Quitar la clave (cualquiera con el enlace podrá ver)",
  "ajustes.clave.es": "La clave es: {clave}",
  "ajustes.clave.sin": "Sin clave: cualquiera con el enlace puede ver lo publicado.",
  "ajustes.guardar": "💾 Guardar",
  "ajustes.carpetas": "Carpetas en Drive",
  "ajustes.carpetas.crear": "📁 Crear las carpetas que falten (una por lengua)",
  "ajustes.carpeta.noEsta": "{carpeta}: todavía no está",
  "ajustes.carpeta.archivos": "Archivos",
  "ajustes.carpeta.audios": "Audios",
  "ajustes.carpeta.imagenes": "Imágenes",
  "ajustes.carpeta.videos": "Videos",
  "ajustes.carpeta.categorias": "Categorías",
  "ajustes.guardado": "Guardado. El visor ya muestra los cambios.",
  "ajustes.revisandoCarpetas": "⏳ Revisando carpetas…",
  "ajustes.carpetasListas": "Carpetas listas.",
  "ajustes.error.nombre": "El catálogo necesita un nombre.",
  "ajustes.error.correo": "El correo de contacto parece estar mal escrito. Revísalo (ej. nombre@gmail.com).",
  "estado.borrador": "Borrador",
  "estado.revisión": "En revisión",
  "estado.publicado": "Publicado",
  "estado.retirado": "Retirado",
  "estado.todas": "Todas",
  "estados.titulo": "📋 Revisar y publicar palabras",
  "estados.intro": "Solo las palabras en *Publicado* se ven en el visor. Cambia el estado de una palabra en su lista, o marca varias y usa “Con las marcadas”.",
  "estados.buscar": "Buscar palabra o categoría…",
  "estados.conMarcadas": "Con las marcadas:",
  "estados.elegir": "Elegir qué hacer…",
  "estados.juegosSi": "🎮 Que aparezcan en juegos",
  "estados.juegosNo": "🚫 Que no aparezcan en juegos",
  "estados.tambienTrads": "cambiar también sus traducciones",
  "estados.aplicar": "Aplicar",
  "estados.col.palabra": "Palabra",
  "estados.col.categoria": "Categoría",
  "estados.col.traducciones": "Traducciones",
  "estados.col.juegos": "🎮 Juegos",
  "estados.col.estado": "Estado",
  "estados.selTodas": "Seleccionar todas",
  "estados.sel": "Seleccionar",
  "estados.noSeVen": "({n} traducción(es) todavía no se ven)",
  "estados.vacia": "No hay palabras en esta vista.",
  "estados.cuenta": "Mostrando {n} de {total} · {sel} seleccionada(s)",
  "estados.mas": "Mostrar más",
  "estados.error.elegir": "Elige qué hacer con las palabras marcadas.",
  "estados.error.ninguna": "Selecciona al menos una palabra.",
  "estados.error.estado": "Elige uno de los estados de la lista.",
  "estados.sinTexto": "(sin texto aún)",
  "estados.listo": "Listo: {n} palabra(s) quedaron en \"{estado}\". El visor ya muestra los cambios.",
  "estados.listoTrads": "Listo: {n} palabra(s) quedaron en \"{estado}\" (y {nt} traducción(es)). El visor ya muestra los cambios.",
  "estados.juegosAhora": "Listo: {n} palabra(s) ahora aparecen en los juegos.",
  "estados.juegosYaNo": "Listo: {n} palabra(s) ya no aparecen en los juegos.",
  "estados.ayudaValidacion": "Elige: Borrador, En revisión, Publicado o Retirado. Solo \"Publicado\" se muestra en el catálogo.",
  "juego.loteria": "Lotería",
  "juego.loteriaA": "Lotería con audio",
  "juego.memorama": "Memorama",
  "juego.charadas": "Charadas",
  "juego.escucha": "¿Cuál escuchaste?",
  "juego.dictado": "Dictado",
  "juego.letras": "Ahorcado, sopa y arma la palabra",
  "juego.basta": "Basta",
  "comoVa.titulo": "📊 ¿Cómo va el catálogo?",
  "comoVa.calculando": "Calculando…",
  "comoVa.ilustrar": "🎨 Lista de palabras sin imagen (para quien ilustra)",
  "comoVa.falta": "📝 Lista de traducciones y audios que faltan",
  "comoVa.estados": "📋 Revisar y publicar palabras",
  "comoVa.lenguas": "🗣️ Lenguas y sus letras",
  "comoVa.preparandoLista": "⏳ Preparando la lista…",
  "comoVa.listaHecha": "✅ {n} palabra(s) sin imagen.",
  "comoVa.abrirLista": "Abrir la lista",
  "comoVa.compartirLista": "y compartirla con quien ilustra. Queda en la carpeta COLABORACIONES.",
  "comoVa.todasImagen": "🎉 Todas las palabras publicadas tienen imagen.",
  "comoVa.seVen": "se ven en el visor (de {total})",
  "comoVa.porRevisar": "por revisar",
  "comoVa.borrador": "en borrador",
  "comoVa.sinImagen": "se ven pero no tienen imagen",
  "comoVa.imagenPerdida": "⚠️ {n} palabra(s) tienen una imagen que ya no se encuentra en Drive. Quizá se borró o se movió: vuelve a subirla desde la app.",
  "comoVa.imagenNoPublica": "⚠️ {n} imagen(es) no se pudieron hacer visibles en internet, quizá por una regla de tu escuela u organización. El visor no podrá mostrarlas.",
  "comoVa.imagenPendiente": "⏳ {n} imagen(es) se terminarán de preparar en la próxima hora, solas.",
  "comoVa.sinConsentimiento": "🎙️ {n} audio(s) publicados sin consentimiento registrado. Se publican sin nombre; conviene preguntar a cada voz y anotarlo en la columna consentimiento.",
  "comoVa.noPublicos": "🔒 {n} audio(s) se quedan solo en la comunidad o tienen el consentimiento pendiente: no se publican.",
  "comoVa.sinHablante": "🎙️ {n} audio(s) publicados sin el nombre de quien habla (columna hablante).",
  "comoVa.sinLenguas": "Todavía no hay lenguas visibles. Agrégalas en 🌐 Lenguas.",
  "comoVa.col.lengua": "Lengua",
  "comoVa.col.palabras": "Palabras",
  "comoVa.col.audio": "Con audio",
  "comoVa.col.imagen": "Con imagen",
  "comoVa.col.juegos": "Juegos y actividades en el visor",
  "comoVa.faltanLetras": "faltan sus letras",
  "comoVa.faltan": "faltan {n}",
  "comoVa.listo": "Listo",
  "comoVa.faltanLetrasAyuda": "Faltan las letras de la lengua (en 🌐 Lenguas)",
  "comoVa.tieneNecesita": "Tiene {tiene}, necesita {min}",
  "comoVa.faltanLetrasCorto": "faltan letras",
  "comoVa.deCuantos": "{tiene} de {min}",
  "comoVa.nota": "Solo cuenta lo que se ve en el visor. ✅ = ya funciona. ⏳ = cuánto tiene y cuánto necesita para funcionar.",
  "ayuda.juegos": "No = esta palabra no aparece en juegos (por ejemplo, palabras rituales o delicadas). Vacío = Sí.",
  "ayuda.consentimiento": "Público con nombre: el audio se publica y se da crédito a la voz (nombre y lugar). Público sin nombre: se publica sin decir quién habla. Solo comunidad: queda en el catálogo, pero no se publica. Pendiente: aún no se pide el consentimiento; no se publica.",
  "ayuda.rol": "Administrador: revisa, publica y borra. Capturista: agrega y edita.",
  "ayuda.activo": "No = conserva el registro pero sin permisos.",
  "rol.Administrador": "Administrador",
  "rol.Capturista": "Capturista",
  "equipo.titulo": "👥 Equipo",
  "equipo.intro": "*Capturistas* agregan y editan palabras en la app. *Administradores* además revisan, publican y borran. Hasta {limite} personas en total con AppSheet gratuito.",
  "equipo.agregar": "Agregar o cambiar a una persona",
  "equipo.correo": "Correo",
  "equipo.nombre": "Nombre (opcional)",
  "equipo.rol": "Rol",
  "equipo.guardar": "Guardar",
  "equipo.share": "*Último paso, en AppSheet* (lo hace quien creó la app): abre la app → botón *Share* (Compartir) → escribe el mismo correo. Sin eso, la persona no puede abrir la app.",
  "equipo.formulas": "Fórmulas de la app plantilla (solo para quien prepara la app de AppSheet)",
  "equipo.formula.permisos": "PALABRAS y TRADUCCIONES → Are updates allowed? (ícono 🧪). Quien no esté activo aquí solo puede ver:",
  "equipo.formula.estado": "Columna estado → Editable? (ícono 🧪) — solo administradores cambian el estado:",
  "equipo.formula.inicial": "Columna estado → Initial value:",
  "equipo.formula.tabla": "Agrega la tabla EQUIPO a la app como solo lectura (Read-only).",
  "equipo.nadie": "Todavía no hay nadie en el equipo.",
  "equipo.col.persona": "Persona",
  "equipo.col.rol": "Rol",
  "equipo.col.activo": "Activo",
  "equipo.tu": "(tú)",
  "equipo.duenio": "· dueño(a)",
  "equipo.quitar": "Quitar",
  "equipo.guardado": "✅ Guardado.",
  "equipo.agregaShare": "Ahora agrégalo también en AppSheet → Share.",
  "equipo.quitaShare": "Quítalo también en AppSheet → Share.",
  "equipo.confirmarQuitar": "¿Quitar a {correo} del equipo?",
  "equipo.quitando": "⏳ Quitando…",
  "equipo.error.correo": "El correo parece estar mal escrito. Revísalo (ej. nombre@gmail.com).",
  "equipo.error.ultimoCambiar": "Siempre debe quedar al menos una persona que administre. Agrega a otra antes de cambiar esta.",
  "equipo.error.ultimoQuitar": "Siempre debe quedar al menos una persona que administre. Agrega a otra antes de quitar esta.",
  "equipo.aviso.limite": "Hay {n} personas activas. AppSheet gratuito permite {limite} usuarios: desactiva a alguien o quien sobre no podrá entrar a la app.",
  "equipo.aviso.compartida": "Ya puede abrir y editar esta hoja.",
  "equipo.aviso.retirada": "Ya no puede editar esta hoja.",
  "equipo.aviso.noCompartir": "No se pudo compartir la hoja: {m}",
  "equipo.aviso.noDejar": "No se pudo dejar de compartir la hoja: {m}",
  "equipo.aviso.noCarpeta": "No se pudo compartir la carpeta del catálogo: {m}",
  "equipo.aviso.noDejarCarpeta": "No se pudo dejar de compartir la carpeta del catálogo: {m}",
  "equipo.notaAlta": "Agregado el {fecha}",
  "equipo.notaInstalo": "Instaló el catálogo",
  "inter.titulo": "🤝 Compartir con otras comunidades",
  "inter.enviar": "📤 Enviar nuestras palabras a otra comunidad",
  "inter.enviar.d": "Se crea un archivo con las palabras. Se lo mandas a la otra comunidad y ella lo recibe aquí abajo, en su propio catálogo.",
  "inter.soloVisibles": "Solo las que se ven en el visor",
  "inter.todas": "Todas (también las que están en borrador o por revisar)",
  "inter.medios": "Incluir imágenes y audios",
  "inter.medios.d": "(solo los audios con permiso público de quien habla. Las imágenes y audios quedarán visibles para quien tenga el archivo)",
  "inter.crear": "Crear el archivo",
  "inter.recibir": "📥 Recibir palabras de otra comunidad",
  "inter.recibir.d": "Elige el archivo que te mandaron. Primero verás qué se agregaría; nada cambia hasta que toques “Agregar”. Nunca se repite una palabra que ya esté.",
  "inter.unir": "Si una palabra se escribe igual que una nuestra en la misma lengua, tratarla como la misma",
  "inter.quedan": "Las palabras nuevas quedan:",
  "inter.quedan.revision": "Por revisar (recomendado)",
  "inter.quedan.borrador": "En borrador",
  "inter.quedan.publicado": "Visibles en el visor",
  "inter.ver": "👀 Ver qué se agregaría",
  "inter.agregar": "Agregar",
  "inter.investigacion": "🔬 Para investigación y especialistas",
  "inter.investigacion.d": "(formatos de bases de datos de grafos)",
  "inter.investigacion.texto": "Archivos para programas de análisis lingüístico. Si no sabes qué son, no los necesitas.",
  "inter.soloPublicado": "Solo lo publicado",
  "inter.todo": "Todo",
  "inter.todos": "Todos",
  "inter.carpeta": "Los archivos se guardan en tu Drive, en la carpeta EXPORTACIONES, junto a esta hoja.",
  "inter.preparando": "⏳ Preparando el archivo… con audios puede tardar unos minutos.",
  "inter.resumenArchivo": "{palabras} palabras, {traducciones} traducciones, {imagenes} imágenes",
  "inter.resumenAudios": ", {n} audios",
  "inter.audiosTiempo": "⚠️ {n} audios no alcanzaron a entrar. Vuelve a crear el archivo para incluirlos.",
  "inter.comparteArchivo": "Comparte ese archivo con la otra comunidad (desde Drive, por correo o por WhatsApp).",
  "inter.archivoListo": "Archivo listo. Toca “Ver qué se agregaría”.",
  "inter.eligeArchivo": "Elige primero un archivo.",
  "inter.agregando": "⏳ Agregando…",
  "inter.revisando": "⏳ Revisando el archivo…",
  "inter.seAgregaron": "✅ Listo, se agregaron — de {origen}",
  "inter.seAgregaria": "Esto se agregaría — de {origen}",
  "inter.lenguas": "Lenguas: {n} que ya tenemos",
  "inter.lenguasNuevas": ", {n} nueva(s): {nombres}",
  "inter.categorias": "Categorías nuevas: {n}",
  "inter.palabras": "Palabras nuevas: {n} (ya teníamos {ya})",
  "inter.traducciones": "Traducciones nuevas: {n} (iguales a las nuestras: {iguales})",
  "inter.relaciones": "Relaciones entre palabras: {n}",
  "menu.constancia": "📜 Constancia de colaboración",
  "const.boton": "Constancias",
  "const.titulo": "Constancia de colaboración",
  "const.intro": "Una constancia para quien aportó al catálogo (docentes, hablantes, familias). Sale de las propuestas aceptadas. Revisa el texto, imprímela o guárdala como PDF y fírmala.",
  "const.persona": "Persona",
  "const.lugar": "Lugar",
  "const.desde": "Desde",
  "const.hasta": "Hasta",
  "const.firmantes": "Quién firma (separa con una línea en blanco o ;)",
  "const.firmantes.ph": "Nombre\nCargo en el equipo del catálogo",
  "const.texto": "Texto adicional (opcional)",
  "const.ver": "Ver constancia",
  "const.imprimir": "Imprimir o guardar PDF",
  "const.nadie": "Todavía no hay aportaciones aceptadas con nombre.",
  "const.cuerpo": "El equipo del catálogo *{catalogo}* hace constar que *{nombre}* colaboró con *{n}* aportación(es) que fueron revisadas y aceptadas, entre el {desde} y el {hasta}.",
  "const.detalle": "Aportaciones: {tipos}.",
  "const.recopilo": "De ellas, {n} las recopiló de otras personas de su comunidad.",
  "const.agradece": "Agradecemos su contribución al cuidado y fortalecimiento de la lengua.",
  "const.firmaEquipo": "Equipo del catálogo {catalogo}",
  "const.pie": "Constancia generada con el sistema del catálogo {catalogo}.",
  "const.tipo.palabra": "palabra(s)",
  "const.tipo.descripcion": "descripción(es)",
  "const.tipo.correccion": "corrección(es)",
  "const.tipo.variante": "variante(s)",
  "const.tipo.imagen": "imagen(es)",
  "const.tipo.fotos": "hoja(s) trabajada(s)",
  "const.tipo.retiro": "aviso(s)",
  "menu.propuestas": "📥 Propuestas de la comunidad",
  "panel.propuestas": "📥 Propuestas de la comunidad",
  "panel.propuestas.d": "Lo que proponen docentes, hablantes y familias desde el visor. Nada entra sin tu revisión.",
  "panel.propuestasPend": "Hay {n} propuesta(s) esperando revisión",
  "panel.sinCorreo": "Falta el correo de contacto (🛠️ Datos del catálogo). Es la alternativa para quien no puede enviar desde el visor.",
  "panel.drive": "Tu Drive va en {pct} % de su espacio. Revisa fotos de hojas ya transcritas o amplía el espacio.",
  "vtab.part": "Participación",
  "part.titulo": "Propuestas desde el visor",
  "part.d": "Docentes, hablantes y familias pueden proponer palabras, descripciones, correcciones e imágenes. Todo llega como pendiente a 📥 Propuestas y solo entra si lo aceptas.",
  "part.activo": "Recibir propuestas desde el visor",
  "part.sinCorreo": "Antes de abrir las propuestas, pon un correo de contacto en 🛠️ Datos del catálogo: es la alternativa siempre visible.",
  "part.quien": "Quién puede proponer",
  "part.quien.todos": "Cualquier persona con acceso al visor",
  "part.quien.clave": "Solo quien tenga la clave de colaboración",
  "part.clave.d": "La clave de colaboración se reparte a docentes de confianza. Con ella, sus propuestas aparecen primero para revisar (nunca se publican solas).",
  "part.claveEs": "Clave de colaboración:",
  "part.sinClave": "Sin clave de colaboración.",
  "part.confirmaClave": "Quien tenga la clave anterior ya no podrá usarla. ¿Crear una nueva?",
  "part.limites": "Límites",
  "part.maxImagenes": "Imágenes por palabra (máximo)",
  "part.porDia": "Propuestas por dispositivo al día",
  "part.maxPendDesc": "Descripciones pendientes por palabra",
  "part.maxPendCorr": "Correcciones pendientes por palabra",
  "part.limites.d": "Las pendientes también cuentan: así no llegan decenas de propuestas de lo mismo mientras revisas. En ✏️ Editar palabras puedes dar más (o cerrar) el cupo de imágenes de una palabra.",
  "part.criterios": "Cómo decide el equipo",
  "part.criterios.d": "Se muestra en el formulario de propuestas, para que se sepa qué esperar.",
  "part.criterios.ph": "Ej. Revisamos con hablantes de dos comunidades. Respetamos las variantes. Respondemos en dos semanas.",
  "part.aportes": "Mostrar en el visor quiénes han aportado (solo quien lo permitió)",
  "part.novedades": "Avisar en el visor las novedades de la semana (cuando aceptes propuestas)",
  "part.abrirPropuestas": "Abrir propuestas",
  "part.piloto": "Sugerencia: empieza con «Solo quien tenga la clave» y dos o tres docentes durante unas semanas; ajusta los límites y después ábrelo a todos.",
  "prop.titulo": "Propuestas de la comunidad",
  "prop.intro": "Revisa cada propuesta: acéptala (como nueva, sumada a una palabra, como descripción, imagen…) o no la incluyas. Puedes dejar un mensaje que quien propuso verá con su código.",
  "prop.verTodas": "Ver también las resueltas",
  "prop.pendientes": "{n} pendiente(s)",
  "prop.vacia": "No hay propuestas pendientes. 🎉",
  "prop.tipo.palabra": "Palabra nueva",
  "prop.tipo.descripcion": "Descripción",
  "prop.tipo.correccion": "Corrección",
  "prop.tipo.variante": "¿Así se dice en mi comunidad?",
  "prop.tipo.imagen": "Imagen",
  "prop.tipo.fotos": "Fotos de hojas",
  "prop.tipo.retiro": "Pide retirar una aportación",
  "prop.anonimo": "Sin nombre",
  "prop.rol.docente": "docente",
  "prop.rol.hablante": "hablante",
  "prop.rol.familia": "familia",
  "prop.rol.otro": "otra persona",
  "prop.recopilo": "recopiló {x}",
  "prop.paraPalabra": "Para: *{x}*",
  "prop.actual": "Ahora dice: {x}",
  "prop.descActual": "Descripción actual: {x}",
  "prop.seDice": "Así se dice en: {x}",
  "prop.actividad": "Actividad: {x}",
  "prop.cupo": "Imágenes de esta palabra: {n} de {max}",
  "prop.coincide": "Ya hay palabras que se escriben igual:",
  "prop.contacto": "Contacto (privado)",
  "prop.prioridad": "clave de colaboración",
  "prop.publicarYa": "Publicar ya",
  "prop.sumarA": "— sumar a una palabra existente —",
  "prop.acc.nueva": "Aceptar como palabra nueva",
  "prop.acc.sumar": "Sumar",
  "prop.acc.descripcion": "Aceptar descripción",
  "prop.acc.agregarDesc": "Agregar a la actual",
  "prop.acc.reemplazar": "Reemplazar el texto",
  "prop.acc.sumarForma": "Sumar como otra forma",
  "prop.acc.variante": "Sumar como forma de esa comunidad",
  "prop.acc.soloRegistro": "Solo registrar",
  "prop.acc.confirmacion": "Registrar confirmación",
  "prop.acc.imagen": "Aceptar en la galería",
  "prop.acc.principal": "Aceptar como principal",
  "prop.acc.interna": "Guardar solo interna",
  "prop.acc.transcrita": "Marcar como transcrita",
  "prop.acc.retirarCredito": "Retirar su nombre de los créditos",
  "prop.acc.atendida": "Marcar como atendida",
  "prop.acc.no": "No incluir",
  "prop.alt": "Texto alternativo: qué se ve en la imagen (para lectores de pantalla)",
  "prop.alt.d": "Una frase corta, por ejemplo «Coyote dibujado con crayón junto a un nopal». Revisa que no haya rostros de menores ni nada que la comunidad reserve.",
  "prop.borrarFotos": "Mandar las fotos a la papelera de Drive (quedan 30 días)",
  "prop.fotos.d": "Transcribe lo que sirva con ➕ en Editar palabras o pide a quien recopiló que lo capture con «Proponer varias». Después marca la propuesta como transcrita.",
  "prop.retiro.d": "Quien aportó pide retirar su nombre o su aportación. Retirar el nombre quita su crédito de todo lo que aportó; para retirar una palabra o imagen concreta, retírala en ✏️ Editar palabras.",
  "prop.nombreRetiro": "Nombre con que aparece",
  "prop.confirmaRetiro": "Se quitará ese nombre de los créditos públicos de todo lo que aportó. ¿Continuar?",
  "prop.respuesta": "Mensaje para quien propuso y nota interna",
  "prop.respuesta.ph": "Lo verá con su código (opcional). Ej. ¡Gracias! La sumamos como variante de tu comunidad.",
  "prop.notaInterna": "Nota interna (solo el equipo)",
  "prop.gracias": "¡Gracias! Quedó registrado.",
  "prop.resuelta": "✅ Listo.",
  "prop.eligeDestino": "Elige la palabra a la que se suma.",
  "prop.error.noExiste": "No se encontró esa propuesta (¿alguien la resolvió?).",
  "prop.error.yaResuelta": "Esa propuesta ya fue resuelta.",
  "prop.error.accion": "Acción no reconocida.",
  "prop.error.sinTexto": "Falta el texto.",
  "prop.error.sinTraduccion": "Esa palabra no tiene texto en esa lengua; primero agrégalo.",
  "prop.fuente.de": "Aportación de {nombre}",
  "prop.fuente.anonima": "Aportaciones de la comunidad",
  "prop.fuente.recopilo": "recopiló {nombre}",
  "prop.fuente.condiciones": "Aportado desde el visor con consentimiento para usarse en el catálogo y sus materiales educativos.",
  "prop.fuente.nota": "Creada al aceptar la propuesta {codigo}.",
  "prop.origen": "propuesta {codigo}",
  "prop.seDiceEn": "Se dice así en {lugar}",
  "prop.antesPrincipal": "antes era la imagen principal",
  "prop.pasoPrincipal": "pasó a ser la imagen principal",
  "prop.novedades": "Esta semana se sumaron {n} aportaciones al catálogo gracias a {m} persona(s). ¡Gracias!",
  "prop.novedades.boton": "Ver el catálogo",
  "ayuda.estadoPropuesta": "Pendiente hasta que el equipo la revise en 📥 Propuestas.",
  "ayuda.tipoImagen": "Dibujo, foto, ilustración u otro.",
  "editar.galeria": "Imágenes: {n} de {max}",
  "editar.gal.vacia": "Sin imágenes extra en la galería.",
  "editar.gal.principal": "Hacer principal",
  "editar.gal.interna": "Dejar solo interna",
  "editar.gal.publica": "Mostrar en el visor",
  "editar.gal.retirarC": "¿Retirar esta imagen de la galería? (no se borra de Drive)",
  "editar.gal.max": "Máximo para esta palabra:",
  "editar.gal.max.d": "vacío = el general · 0 = no recibir más",
  "editar.f.sinDescripcion": "Sin descripción en {lengua}",
  "editar.descripcion": "Descripción en {lengua}: qué es, para qué se usa (opcional)",
  "editar.descripcion.d": "Se escribe en la misma lengua de la palabra. Se ve en la ficha del visor y puede servir de pista en los materiales para imprimir.",
  "editar.descripcion.tiene": "Tiene descripción",
  "visorAj.leyenda": "Leyenda en lo impreso",
  "visorAj.leyenda.d": "Aparece al pie de todo lo que se imprime desde el visor, junto con el nombre del catálogo y su licencia.",
  "visorAj.leyenda.ph": "Ej. Material comunitario, prohibida su venta",
  "panel.credito": "Sistema diseñado y desarrollado por {autor} · Código abierto (AGPL-3.0) · © 2025",
  "menu.tiposRel": "🕸️ Tipos de relación",
  "panel.tiposRel": "🕸️ Tipos de relación",
  "panel.tiposRel.d": "Cómo se pueden relacionar las palabras: parte de, sinónimo, variante…",
  "panel.sinDireccion.t": "⚠️ Falta un paso para que el visor muestre este catálogo",
  "panel.sinDireccion.1": "En Extensiones → Apps Script: Implementar → Nueva implementación → Aplicación web, con «Ejecutar como: Yo» y «Quién tiene acceso: Cualquier persona». Se hace una sola vez por catálogo.",
  "panel.sinDireccion.2": "Copia la dirección que termina en /exec y pégala en 🔗 Visor: conexión y mantenimiento (botón de abajo). Guarda y prueba.",
  "panel.sinDireccion.3": "Cuando actualices el programa no crees otra: usa Gestionar implementaciones → ✏️ → Versión: Nueva. Así la dirección no cambia.",
  "rel.titulo": "Relaciones con otras palabras",
  "rel.guardaPrimero": "Guarda la palabra primero; después podrás relacionarla con otras.",
  "rel.ninguna": "Sin relaciones todavía.",
  "rel.tipo": "Tipo de relación",
  "rel.buscaPalabra": "Busca la otra palabra (en cualquier lengua)…",
  "rel.nota": "Nota (opcional, se ve en el visor)",
  "rel.agregar": "Relacionar",
  "rel.sinResultados": "No hay palabras que coincidan.",
  "rel.ayudaAdmin": "Se publica al momento. Los tipos se ajustan en Configurar ▸ 🕸️ Tipos de relación; los marcados con 🔒 no se muestran en el visor.",
  "rel.ayudaCaptura": "Queda en revisión hasta que alguien que administra la publique. Los tipos con 🔒 no se muestran en el visor.",
  "rel.agregada": "✅ Relación guardada.",
  "rel.actualizada": "✅ Listo.",
  "rel.publicar": "Publicar",
  "rel.retirar": "Quitar",
  "rel.retirar.d": "Queda Retirada (no se borra) y deja de verse.",
  "rel.noPublica": "solo interna",
  "rel.noPublica.d": "Este tipo de relación no se muestra en el visor.",
  "rel.aviso.revision": "Quedó en revisión: se verá cuando alguien que administra la publique.",
  "rel.error.tipo": "Ese tipo de relación no existe o está retirado.",
  "rel.error.palabra": "Elige la otra palabra de la lista.",
  "rel.error.misma": "Una palabra no se puede relacionar consigo misma.",
  "rel.error.yaExiste": "Esa relación ya existe.",
  "rel.error.noExiste": "No se encontró esa relación (¿alguien la cambió?). Vuelve a abrir la palabra.",
  "rel.error.soloAdmin": "Solo quien administra puede quitar relaciones ya publicadas.",
  "rel.error.nombre": "Escribe el nombre del tipo de relación.",
  "rel.error.nombreRepetido": "Ya hay un tipo llamado «{nombre}».",
  "rel.tipoEncontrado": "Encontrado escrito en RELACIONES: revisa su nombre.",
  "rel.tipoImportado": "Traído con relaciones de otro catálogo.",
  "rel.tipos.titulo": "🕸️ Tipos de relación",
  "rel.tipos.intro": "Con estos tipos se relacionan las palabras. Cada catálogo los ajusta: cambia nombres, agrega los suyos y decide cuáles se ven en el visor. Un nombre vacío usa el de fábrica, traducido al idioma de cada persona.",
  "rel.tipos.nombre": "Nombre",
  "rel.tipos.inverso": "Leído desde la otra palabra",
  "rel.tipos.inversoPh": "Ej. Tiene como parte",
  "rel.tipos.mismoNombre": "(igual en ambos sentidos)",
  "rel.tipos.simetrica": "Igual en ambos sentidos",
  "rel.tipos.publica": "Se ve en el visor",
  "rel.tipos.orden": "Orden",
  "rel.tipos.usos": "En uso",
  "rel.tipos.nota": "Nota interna (para el equipo)",
  "rel.tipos.sistema": "La usa 🔎 Palabras que se escriben igual. Puedes cambiar su nombre y si se ve.",
  "rel.tipos.fija": "Interna: recuerda qué palabras se unieron. No se muestra ni se elige.",
  "rel.tipos.retirar": "Retirar",
  "rel.tipos.reactivar": "Reactivar",
  "rel.tipos.ayuda": "Retirar un tipo no borra sus relaciones: dejan de verse y de poder elegirse. Cada tipo se identifica por un código que no cambia aunque cambies el nombre (así lo leen AppSheet y el grafo).",
  "rel.tipos.nuevo": "Nuevo tipo",
  "rel.tipos.nuevoNombre": "Nombre (ej. Se usa en)",
  "rel.tipos.nuevoInverso": "Leído desde la otra palabra (ej. Lleva)",
  "rel.tipos.agregar": "Agregar",
  "rel.tipos.nuevo.d": "Si la relación se lee igual desde las dos palabras (como «sinónimo»), marca «Igual en ambos sentidos».",
  "rel.tipos.soloVer": "Solo quien administra puede cambiar los tipos.",
  "rel.tipos.guardado": "✅ Guardado.",
  "rel.tipos.normalizadas": "Se ordenaron {n} relación(es) que tenían el tipo escrito a mano.",
  "inter.variantes": "📝 {n} palabra(s) se escriben distinto aquí y allá; no se cambió nada nuestro. Vale la pena platicarlo con la comunidad: {lista}",
  "inter.ambiguas": "⚠️ {n} palabra(s) se parecían a varias nuestras; se agregan aparte para revisarlas: {lista}",
  "inter.quedanComo": "Lo nuevo queda como \"{estado}\".",
  "inter.generando": "⏳ Generando…",
  "inter.grafoListo": "✅ {nodos} nodos y {aristas} aristas.",
  "inter.error.leer": "Ese archivo no se puede leer. Pide a la otra comunidad que lo vuelva a crear y enviar.",
  "inter.error.formato": "Ese archivo no es de un catálogo de palabras. Revisa que sea el que te mandó la otra comunidad.",
  "inter.error.mismo": "Ese archivo salió de este mismo catálogo: no hay nada nuevo que agregar.",
  "inter.otroCatalogo": "otro catálogo",
  "interfaz.titulo": "🌐 Idioma de los menús",
  "interfaz.intro": "Los menús, ventanas y avisos de esta hoja pueden verse en otra lengua. Cada persona elige la suya; lo que no esté traducido se ve en el idioma de respaldo.",
  "interfaz.elegir": "1. ¿En qué idioma quieres ver los menús?",
  "interfaz.paraMi": "Usar para mí",
  "interfaz.paraTodos": "Usar también como idioma de todo el catálogo",
  "interfaz.paraTodos.d": "Es el que verá quien no haya elegido otro.",
  "interfaz.actual": "Ahora ves los menús en: *{nombre}*. Idioma de todo el catálogo: *{pred}*.",
  "interfaz.avance": "{nombre} · {n} de {total} textos traducidos",
  "interfaz.cambiado": "Listo. El menú ya cambió; las ventanas que abras desde ahora saldrán en ese idioma.",
  "interfaz.traducir": "2. Traducir los menús",
  "interfaz.traducir.d": "Escribe cada texto en la lengua que elijas. Respeta lo que va entre llaves, como {n}: ahí el programa pone un número o un nombre. Lo que va entre asteriscos *así* se ve en negritas.",
  "interfaz.cual": "Lengua a traducir",
  "interfaz.nueva": "➕ Agregar una lengua para los menús",
  "interfaz.codigo": "Código corto (ej. hch, nah, mix-ox)",
  "interfaz.nombre": "Nombre (ej. Wixárika)",
  "interfaz.crear": "Agregar",
  "interfaz.soloFaltan": "Ver solo lo que falta traducir",
  "interfaz.buscar": "Buscar un texto…",
  "interfaz.col.espanol": "Español",
  "interfaz.col.traduccion": "Traducción",
  "interfaz.guardar": "💾 Guardar traducciones",
  "interfaz.guardadas": "Se guardaron {n} texto(s).",
  "interfaz.sinCambios": "No hay cambios por guardar.",
  "interfaz.mostrando": "Mostrando {n} de {total}",
  "interfaz.mas": "Mostrar más",
  "interfaz.compartir": "3. Compartir una traducción",
  "interfaz.compartir.d": "Crea un archivo con la traducción para que otra comunidad la instale en su propio catálogo, o instala aquí la que te compartieron.",
  "interfaz.exportar": "📤 Crear archivo de esta traducción",
  "interfaz.exportado": "Archivo listo: {n} de {total} textos traducidos.",
  "interfaz.instalar": "📥 Instalar una traducción",
  "interfaz.reemplazar": "Si un texto ya está traducido aquí, cambiarlo por el del archivo",
  "interfaz.reemplazar.d": "(si no lo marcas, solo se llena lo que falta y no se toca lo que esta comunidad ya tradujo)",
  "interfaz.ver": "👀 Ver qué cambiaría",
  "interfaz.instalarBoton": "Instalar",
  "interfaz.previa": "{nombre}: {nuevos} texto(s) nuevos, {iguales} iguales, {distintos} distintos a los de aquí ({reemplazados} se cambiarían).",
  "interfaz.previaExtra": "Le faltan {faltan} texto(s) de esta versión: se verán en el idioma de respaldo. {desconocidos} texto(s) son de otra versión y se ignoran.",
  "interfaz.instalada": "✅ Traducción instalada: {nombre}.",
  "interfaz.eligeArchivo": "Elige primero el archivo de la traducción.",
  "interfaz.nadaQueTraducir": "Primero agrega una lengua para los menús (arriba).",
  "interfaz.error.noInstalado": "Esa traducción no está instalada en este catálogo.",
  "interfaz.error.codigo": "El código debe tener de 2 a 8 letras, y puede llevar una variante después de un guion (ej. hch, nah, mix-ox).",
  "interfaz.error.nombre": "Escribe el nombre de la lengua.",
  "interfaz.error.yaExiste": "Ese código ya está en uso.",
  "interfaz.error.archivo": "Ese archivo no es una traducción de los menús del catálogo.",
  "interfaz.parte": "Mostrar textos de",
  "interfaz.parte.todo": "Todo",
  "interfaz.parte.hoja": "La hoja de cálculo (menús y ventanas)",
  "interfaz.parte.visor": "El visor web (lo que ve el público)",
  "interfaz.visor.d": "Las traducciones de los textos del visor se publican con el catálogo: al publicar, el visor las ofrece en su botón 🌐. También se pueden compartir con el archivo y cada persona la instala en su celular.",
  "interfaz.error.espanol": "El español ya viene incluido: no hace falta instalarlo."
};

/**
 * Textos del visor web (index.html). Se guardan en la misma hoja TEXTOS para traducirlos
 * junto con los menús; al publicar, las traducciones viajan con el catálogo (clave "interfaz").
 * Deben coincidir con V_ES del visor: los textos nuevos del visor se agregan aquí.
 */
const T_VISOR = {
  "visor.a11y": "Accesibilidad",
  "visor.a11y.contraste": "Alto contraste",
  "visor.a11y.mayor": "Texto más grande",
  "visor.a11y.menor": "Texto más pequeño",
  "visor.a11y.normal": "Tamaño de texto normal",
  "visor.acerca": "Acerca de este catálogo",
  "visor.act.abrir": "Abrir",
  "visor.act.abrirArchivo": "Abrir archivo",
  "visor.act.agregar": "Agregar a Mis actividades",
  "visor.act.ahorcado": "Ahorcado",
  "visor.act.ahorcado.d": "Adivina la palabra letra por letra, con el alfabeto de la lengua.",
  "visor.act.arma": "Arma la palabra",
  "visor.act.arma.d": "Ordena las letras o sílabas.",
  "visor.act.armaSilabas": "Arma con sílabas (recortable)",
  "visor.act.armaSilabas.d": "Recorta las sílabas y pégalas en orden para formar la palabra.",
  "visor.act.aviso.1": "No se guardan en internet ni en el catálogo; nadie más las ve. Pueden desaparecer si:",
  "visor.act.aviso.c1": "se borran los datos o el historial del navegador, o se usa una app de limpieza;",
  "visor.act.aviso.c2": "usas una ventana privada o de incógnito (al cerrarla se borra todo);",
  "visor.act.aviso.c3": "abres el visor en otro navegador, otro teléfono u otra computadora (ahí no estarán);",
  "visor.act.aviso.c4": "en algunos teléfonos, sobre todo iPhone, pasan varias semanas sin abrir el visor.",
  "visor.act.aviso.recup": "*Para recuperarlas:* toca *📂 Abrir archivo*, elige ese archivo y tus actividades vuelven a aparecer. Sirve también en otro equipo o para compartirlas con otra maestra o maestro.",
  "visor.act.aviso.resp": "*Te recomendamos guardar un respaldo:* toca *💾 Guardar en archivo*. Se descarga un archivo pequeño a tu teléfono o computadora; guárdalo en una carpeta que ubiques, o mándatelo por correo o WhatsApp.",
  "visor.act.aviso.t": "Tus actividades se guardan solo en este navegador",
  "visor.act.basta": "Basta",
  "visor.act.basta.d": "Una letra, varias categorías: ¿cuántas palabras conoces?",
  "visor.act.bastaImp": "Basta en papel",
  "visor.act.bastaImp.d": "La tabla para jugar en grupo; el docente canta la letra.",
  "visor.act.borrar.c": "¿Borrar esta actividad de este dispositivo?",
  "visor.act.borrarDisp": "Borrar lo guardado sin conexión",
  "visor.act.borrarDisp.d": "Quita de este dispositivo las imágenes y audios guardados con 📥 Llevar sin conexión. Tus actividades y listas no se borran.",
  "visor.act.borrarOff": "Borrar imágenes y audios guardados",
  "visor.act.borrarOff.c": "¿Borrar las imágenes y audios guardados en este dispositivo?",
  "visor.act.charadas": "Charadas",
  "visor.act.charadas.d": "Mímica, dibujos, sonidos o descripciones por equipos.",
  "visor.act.copiaDe": "Copia de {x}",
  "visor.act.cotejo": "Lista de cotejo",
  "visor.act.cotejo.d": "Tabla en blanco para seguir el avance del grupo a mano. El visor no guarda datos de estudiantes.",
  "visor.act.crucigrama": "Crucigrama",
  "visor.act.crucigrama.d": "Con imágenes, otra lengua o descripciones como pistas; o deja el espacio para tus propias pistas.",
  "visor.act.descargado": "Se descargó el archivo. Guárdalo en un lugar que ubiques.",
  "visor.act.dictado": "Dictado",
  "visor.act.dictado.d": "Escucha y escribe.",
  "visor.act.duplicar": "Duplicar",
  "visor.act.enlaceLargo": "Esta actividad es muy larga para un enlace (tiene muchas pistas propias). Compártela con 💾 Guardar en archivo.",
  "visor.act.enquelengua": "¿En qué lengua lo dicen?",
  "visor.act.enquelengua.d": "Descubre la diversidad de lenguas que te rodean.",
  "visor.act.error.formato": "Ese archivo no es un respaldo de actividades del visor.",
  "visor.act.error.grande": "El archivo es demasiado grande para ser un respaldo de actividades.",
  "visor.act.escucha": "¿Cuál escuchaste?",
  "visor.act.escucha.d": "Escucha a una persona hablante y elige la imagen.",
  "visor.act.escuchaImp": "Escucha y señala (papel)",
  "visor.act.escuchaImp.d": "El docente dice las palabras (o las reproduce con un QR) y el grupo anota el orden.",
  "visor.act.etiquetas": "Etiquetas para el entorno",
  "visor.act.etiquetas.d": "Para pegar en objetos de la casa o el aula: la lengua a la vista todos los días.",
  "visor.act.faltanPalabras": "{n} palabra(s) de esas listas ya no están en el catálogo; el resto funciona igual.",
  "visor.act.glosario": "Glosario de la unidad",
  "visor.act.glosario.d": "Las palabras de tu lista, en orden, con imagen y si quieres otra lengua o la descripción.",
  "visor.act.guardar": "Guardar",
  "visor.act.guardarArchivo": "Guardar en archivo",
  "visor.act.hojas": "Hojas de trabajo",
  "visor.act.hojas.d": "Une con una línea, completa la letra, escribe el nombre, dibuja la palabra.",
  "visor.act.importado": "Se agregaron {act} actividad(es), {lis} lista(s) y {pis} pista(s) propia(s). Lo que ya tenías se conservó.",
  "visor.act.inicialImp": "¿Con qué letra empieza? (papel)",
  "visor.act.inicialImp.d": "Encierra o escribe la primera letra de cada palabra.",
  "visor.act.irCatalogo": "Ir a ese catálogo",
  "visor.act.lista": "Tus actividades",
  "visor.act.lleno": "El navegador no tiene más espacio para guardar. Guarda un respaldo en archivo y borra actividades que ya no uses.",
  "visor.act.loteria": "Lotería",
  "visor.act.loteria.d": "En pantalla, a distancia, impresa o descargada para jugar sin internet.",
  "visor.act.loteriaImp": "Lotería para imprimir",
  "visor.act.loteriaImp.d": "Cartas (8 por hoja) y tablas 4×4 en carta o media carta.",
  "visor.act.memorama": "Memorama",
  "visor.act.memorama.d": "Encuentra los pares: imagen, palabra o audio, en una o dos lenguas.",
  "visor.act.memoramaImp": "Memorama recortable",
  "visor.act.memoramaImp.d": "Pares para recortar: imagen y palabra, dos lenguas o dos imágenes.",
  "visor.act.misActividades": "Mis actividades",
  "visor.act.misActividades.d": "Lo que has preparado para imprimir, guardado en este dispositivo.",
  "visor.act.misListas": "Mis listas de palabras",
  "visor.act.misListas.d": "Elige las palabras de tu unidad y úsalas en todos los materiales.",
  "visor.act.nadaQueGuardar": "Todavía no hay nada que guardar.",
  "visor.act.nuevoNombre": "Nombre de la actividad:",
  "visor.act.ordena": "Ordena las letras",
  "visor.act.ordena.d": "Letras revueltas para formar la palabra.",
  "visor.act.otroCatalogo": "Este archivo es del catálogo «{x}». Ábrelo con ese catálogo.",
  "visor.act.preguntaCasa": "Pregunta en casa",
  "visor.act.preguntaCasa.d": "Para buscar palabras y descripciones con la familia, y sumarlas al catálogo.",
  "visor.act.queLetra": "¿Con qué letra empieza?",
  "visor.act.queLetra.d": "Mira la imagen, escucha y elige la letra con la que empieza.",
  "visor.act.quees": "¿Qué es?",
  "visor.act.quees.d": "Lee la descripción en la lengua y elige la palabra que describe.",
  "visor.act.quitarGuardada": "Quitar de guardadas",
  "visor.act.recibida": "Te compartieron una actividad",
  "visor.act.recientes": "Lo que preparas aparece aquí solo. Se conservan las {n} más recientes; las que marques con ⭐ se quedan siempre.",
  "visor.act.recoleccion": "Hoja de recolección",
  "visor.act.recoleccion.d": "Para asambleas y talleres: palabras, quién las compartió y su permiso.",
  "visor.act.recordar": "Hace tiempo que no guardas un respaldo de tus actividades guardadas.",
  "visor.act.renombrar": "Nombre",
  "visor.act.repaso": "Repaso con tarjetas",
  "visor.act.repaso.d": "Las que te cuestan aparecen más seguido. Se guarda tu avance.",
  "visor.act.silabasImp": "Cuenta las sílabas",
  "visor.act.silabasImp.d": "Aplaude y colorea una palma por cada sílaba. Solo con palabras que tienen sílabas registradas.",
  "visor.act.sinRespaldo": "Todavía no has guardado un respaldo en archivo.",
  "visor.act.sopa": "Sopa de letras",
  "visor.act.sopa.d": "Encuentra palabras escondidas.",
  "visor.act.sopaImp": "Sopa de letras",
  "visor.act.sopaImp.d": "Con palabras cruzadas, pistas a tu elección y hoja de respuestas.",
  "visor.act.tarjetas": "Tarjetas de vocabulario",
  "visor.act.tarjetas.d": "Imagen y palabra, con otra lengua opcional.",
  "visor.act.titulo": "Mis actividades",
  "visor.act.trazo": "Hoja de trazo",
  "visor.act.trazo.d": "Una letra por hoja para repasar, con palabras del catálogo que la contienen.",
  "visor.act.ultimoRespaldo": "Último respaldo: hace {n} día(s).",
  "visor.act.unirse": "Unirme a una lotería",
  "visor.act.unirse.d": "Con el nombre y los 4 números de la partida.",
  "visor.act.vacia": "Todavía no hay actividades. Prepara una hoja en Imprimir y aparecerá aquí.",
  "visor.ah.dificil": "Difícil: sin pistas",
  "visor.ah.dificultad": "Dificultad",
  "visor.ah.empieza": "Empieza con *{letra}*",
  "visor.ah.era": "Era:",
  "visor.ah.errores": "Errores: {n}/6",
  "visor.ah.facil": "Fácil: pistas sin perder puntos",
  "visor.ah.medio": "Intermedio: cada pista resta",
  "visor.ah.pistaN": "Pista {n}",
  "visor.ah.teclado": "El teclado usa el alfabeto de la lengua, con sus letras dobles (por ejemplo, ch o kw) como una sola letra.",
  "visor.ah.valdra": "la palabra valdrá {pts} pts",
  "visor.alAzar": "Al azar",
  "visor.alguien": "alguien",
  "visor.arma.casi": "Casi. Quedan las piezas que ya están bien.",
  "visor.arma.espacio": "Espacio {n}",
  "visor.arma.letras": "Letras",
  "visor.arma.piezas": "Piezas",
  "visor.arma.silabas": "Sílabas (si están registradas)",
  "visor.asiSeEscribe": "Así se escribe:",
  "visor.audio.conexion": "El audio necesita conexión a internet.",
  "visor.audio.error": "No se pudo reproducir el audio.",
  "visor.aviso.tipo.importante": "Importante",
  "visor.aviso.tipo.info": "Información",
  "visor.aviso.tipo.novedad": "Novedad",
  "visor.aviso.tipo.pausa": "Pausa",
  "visor.avisos": "Avisos",
  "visor.avisos.anterior": "Aviso anterior",
  "visor.avisos.cerrar": "Cerrar este aviso",
  "visor.avisos.n": "{n} de {total}",
  "visor.avisos.siguiente": "Aviso siguiente",
  "visor.basta.anotalas": "Anótalas y compártelas con quien cuida el catálogo.",
  "visor.basta.celular": "Cada quien en su celular",
  "visor.basta.columnas": "Columnas (de 3 a 6)",
  "visor.basta.columnasCorto": "Columnas",
  "visor.basta.como": "¿Cómo juegan?",
  "visor.basta.compartelas": "Compártelas con quien cuida el catálogo para que las revise con la comunidad:",
  "visor.basta.compartida": "Misma partida en varios celulares",
  "visor.basta.compartida.d": "Todos escriben el mismo nombre y los mismos 4 números: así les salen las mismas letras, sin internet.",
  "visor.basta.d": "Sale una letra. Escribe una palabra que empiece con ella en cada columna. Quien termina primero dice «¡Basta!».",
  "visor.basta.delCatalogo": "Palabras del catálogo",
  "visor.basta.error.cols": "Elige entre 3 y 6 columnas (ahora: {n}).",
  "visor.basta.error.vacias": "No hay palabras en esas categorías.",
  "visor.basta.esta": "Está en el catálogo",
  "visor.basta.fin": "¡Fin del Basta!",
  "visor.basta.grito": "¡BASTA!",
  "visor.basta.libre": "Columna libre (cualquier palabra)",
  "visor.basta.libreCorto": "Libre",
  "visor.basta.msg": "Hola. Jugando Basta en {lengua} salieron estas palabras que no están en el catálogo «{catalogo}»:",
  "visor.basta.msgAsunto": "Palabras nuevas desde Basta",
  "visor.basta.noEmpieza": "No empieza con la letra.",
  "visor.basta.noEsta": "No está en el catálogo. ¿El grupo la acepta? (🟡 50 si alguien más la repitió)",
  "visor.basta.nuevas": "Escribiste {n} palabra(s) que el catálogo aún no tiene",
  "visor.basta.papel": "Con papel: esta pantalla solo da la letra y el tiempo",
  "visor.basta.papelPuntos": "Sumen sus puntos en papel: 100 por palabra válida, 50 si alguien más la repitió.",
  "visor.basta.partida": "Partida *{x}*",
  "visor.basta.pendientes": "Decidan las palabras pendientes.",
  "visor.basta.siguiente": "Siguiente letra",
  "visor.basta.sinLetra": "El catálogo aún no tiene palabras aquí con esta letra.",
  "visor.basta.suma": "Esta ronda: *{ronda}* · Total: *{total}*",
  "visor.basta.tiempo": "Tiempo por ronda",
  "visor.borrar": "Borrar",
  "visor.buscar": "Buscar en cualquier lengua…",
  "visor.buscarCorto": "Buscar…",
  "visor.buscarPalabra": "Buscar palabra",
  "visor.cambiarCatalogo": "Cambiar de catálogo",
  "visor.cancelar": "Cancelar",
  "visor.cargando": "Cargando…",
  "visor.cartaN": "Carta {n}",
  "visor.catalogoPalabras": "Catálogo de palabras",
  "visor.categoria": "Categoría",
  "visor.cerrar": "Cerrar",
  "visor.char.acierto": "Acierto",
  "visor.char.audio": "Al acertar, escuchar la palabra (con internet)",
  "visor.char.celular": "¿Cómo se usa el celular?",
  "visor.char.corrige": "¿Hubo un error? Corrige los puntos:",
  "visor.char.decir": "Para acertar hay que decir la palabra en esta lengua",
  "visor.char.describe": "Dilo sin decirlo",
  "visor.char.describe.d": "Explícalo con otras palabras, sin decir la palabra.",
  "visor.char.dibujo": "Dibújalo",
  "visor.char.dibujo.d": "En papel o pizarrón, sin letras ni números.",
  "visor.char.dicen": "Para acertar, el equipo dice la palabra en *{lengua}*.",
  "visor.char.error.formas": "Elige al menos una forma de jugar.",
  "visor.char.error.pocas": "Hay {n} palabras con esas categorías; se necesitan 10.",
  "visor.char.fin": "¡Fin de las charadas!",
  "visor.char.formas": "Formas de jugar",
  "visor.char.frente": "*En la frente:* quien adivina pone el celular en su frente; botones enormes en los bordes, sin mirar",
  "visor.char.instr.frente": "Quien adivina se pone el celular en la frente con la pantalla hacia su equipo. El equipo actúa.",
  "visor.char.instr.presenta": "Quien actúa toma el celular y no deja que su equipo vea la pantalla.",
  "visor.char.instr.tv": "Quien actúa se pone frente a la pantalla. Mantén presionado «Ver» (o la barra espaciadora) para mirar la palabra.",
  "visor.char.lengua": "Lengua del juego",
  "visor.char.listos": "¡Listos!",
  "visor.char.mimica": "Mímica",
  "visor.char.mimica.d": "Solo con el cuerpo, sin hablar ni hacer sonidos.",
  "visor.char.pasar": "Pasar",
  "visor.char.presenta": "*Presentador:* quien actúa ve la palabra y toca ✓ o ⏭",
  "visor.char.ruleta": "Si eliges varias, cada tarjeta toca una al azar (ruleta).",
  "visor.char.significado": "Mostrar también el significado en",
  "visor.char.siguienteEquipo": "Siguiente equipo",
  "visor.char.sinSensor": "Sin permiso para inclinación: usa los bordes.",
  "visor.char.sonido": "Suena como",
  "visor.char.sonido.d": "Solo con sonidos, sin palabras.",
  "visor.char.tiempo": "Tiempo por turno",
  "visor.char.tiempoFin": "¡Tiempo! {equipo}: {n} ✓",
  "visor.char.tilt": "Además, inclinar el celular: hacia abajo = ✓ acierto, hacia arriba = ⏭ pasar",
  "visor.char.tilt.d": "Funciona en muchos celulares Android. Siempre puedes tocar los bordes.",
  "visor.char.turnoN": "Turno {n} de {total}",
  "visor.char.turnos": "Turnos por equipo",
  "visor.char.tv": "*Pantalla grande:* proyector o computadora; la palabra se ve solo al mantener presionado (teclas ← y →)",
  "visor.char.ver": "Mantén presionado para ver (barra espaciadora)",
  "visor.col.asunto": "Colaborar con «{catalogo}»",
  "visor.col.atiende": "Atiende: *{x}*.",
  "visor.col.c1": "Una palabra que no está, o que se dice distinto en tu comunidad.",
  "visor.col.c2": "Tu voz: grabar cómo se pronuncian las palabras (siempre con tu consentimiento).",
  "visor.col.c3": "Imágenes o dibujos para las palabras que no tienen.",
  "visor.col.c4": "Correcciones de escritura o de significado.",
  "visor.col.como": "¿Cómo puedes aportar?",
  "visor.col.correo": "Correo: {x}.",
  "visor.col.horario": "Horario: {x}.",
  "visor.col.intro": "Este catálogo crece con lo que aporta la comunidad: palabras que faltan, otras formas de decirlas, grabaciones de voz, imágenes y correcciones.",
  "visor.col.llamar": "Llamar",
  "visor.col.revision": "Las aportaciones se reciben por los medios de contacto; el equipo las revisa con la comunidad antes de publicarlas.",
  "visor.col.sinContacto": "Quien administra este catálogo todavía no ha publicado datos de contacto.",
  "visor.col.telefono": "Teléfono: {x}.",
  "visor.col.texto": "Hola. Quiero colaborar con el catálogo «{catalogo}»:",
  "visor.compartir": "Compartir este catálogo",
  "visor.compartirEnlace": "Compartir enlace",
  "visor.con.actualizando": "Actualizando…",
  "visor.con.aldia": "Al día",
  "visor.con.clave": "Este catálogo pide una clave, o la clave cambió. Escribe la clave correcta.",
  "visor.con.conectado": "Conectado",
  "visor.con.guardada": "Sin conexión: última versión guardada",
  "visor.con.inesperada": "Respuesta inesperada del catálogo.",
  "visor.con.nuevaVersion": "Se actualizó el catálogo",
  "visor.con.sinConexion": "Sin conexión: puedes seguir consultando lo guardado",
  "visor.continuar": "Continuar",
  "visor.copiaEnlace": "Copia este enlace:",
  "visor.copiar": "Copiar",
  "visor.correcto": "¡Correcto! +{pts}",
  "visor.correo": "Correo",
  "visor.creditos": "Créditos y agradecimientos",
  "visor.cuatroNumeros": "Cuatro números",
  "visor.cx.cancelar": "Seguir en el catálogo actual",
  "visor.cx.clave": "Clave (opcional)",
  "visor.cx.clave.d": "Este dispositivo recuerda la clave de cada catálogo por separado.",
  "visor.cx.conectar": "Conectar",
  "visor.cx.d": "Pega la dirección del catálogo (termina en /exec). Si el catálogo pide clave, escríbela también: quien lo administra la encuentra en la hoja, pestaña CONFIG, columna API_TOKEN.",
  "visor.cx.direccion": "Dirección",
  "visor.cx.error.direccion": "Esa dirección no parece de un catálogo: debe empezar con https://script.google.com/",
  "visor.cx.olvidar": "Olvidar este catálogo en este dispositivo",
  "visor.cx.predeterminado": "Volver al catálogo de este visor",
  "visor.cx.recientes": "Catálogos que ya abriste aquí",
  "visor.cx.sub": "Conecta un catálogo",
  "visor.cx.titulo": "Conectar un catálogo",
  "visor.dia": "Palabra del día · {lengua}",
  "visor.dic.casi": "¡Casi! Se escribe:",
  "visor.dic.d": "Escucha y escribe. Puedes usar los atajos de teclado de la lengua. No importan mayúsculas ni acentos; vale cualquiera de las formas registradas.",
  "visor.dic.escribe": "Escribe lo que escuchaste",
  "visor.dic.exacto": "¡Exacto!",
  "visor.dic.img": "Mostrar la imagen como pista",
  "visor.dic.seEscribe": "Se escribe:",
  "visor.editar": "Editar",
  "visor.elegirCategorias": "Elegir categorías (opcional)",
  "visor.empezar": "Empezar",
  "visor.enlaceCopiado": "Enlace copiado",
  "visor.eql.es": "Es {lengua}",
  "visor.eql.otras": "La misma palabra en otras lenguas:",
  "visor.eql.pregunta": "¿En qué lengua está escrita esta palabra?",
  "visor.eql.si": "¡Sí!",
  "visor.equipo": "Equipo",
  "visor.equipos": "Equipos",
  "visor.error.partida": "Escribe el nombre de la partida y sus 4 números.",
  "visor.esc.d": "Escucha la voz de una persona hablante y toca la imagen correcta. No hace falta leer.",
  "visor.esc.texto": "Mostrar también la palabra escrita",
  "visor.escucha": "Escucha",
  "visor.escuchar": "Escuchar",
  "visor.escucharOtraVez": "Escuchar otra vez",
  "visor.escucharPalabra": "Escuchar {palabra}",
  "visor.fuente": "Fuente: {credito}",
  "visor.fuentes": "Fuentes",
  "visor.galeria": "{n} imágenes",
  "visor.galeria.desliza": "desliza para ver más",
  "visor.general": "General",
  "visor.glosario": "Glosario",
  "visor.gob.comunidades": "Comunidades",
  "visor.gob.consulta": "Se consulta con: {x}",
  "visor.guardar": "Guardar",
  "visor.hoja.calcar": "Palabra punteada para calcar",
  "visor.hoja.completa": "Completa la palabra",
  "visor.hoja.completa.inst": "Escribe las letras que faltan.",
  "visor.hoja.cual": "¿Qué hoja?",
  "visor.hoja.cuantas": "Cuántas y para quién",
  "visor.hoja.desc": "{n} hoja(s) · {tipo}",
  "visor.hoja.desc2": "{tipo} · código {c} · {v} versión(es) × {n} hoja(s)",
  "visor.hoja.dibuja": "Dibuja la palabra",
  "visor.hoja.dibuja.inst": "Lee la palabra y haz un dibujo.",
  "visor.hoja.docente": "Hoja del docente",
  "visor.hoja.escribe": "Escribe el nombre",
  "visor.hoja.escribe.inst": "Escribe el nombre de cada imagen.",
  "visor.hoja.escribe.inst.pista": "Escribe la palabra que corresponde a cada pista.",
  "visor.hoja.hojas": "Hojas (de cada versión)",
  "visor.hoja.instruccion": "Instrucción",
  "visor.hoja.instruccion.d": "(puedes escribirla en tu lengua o dejarla vacía)",
  "visor.hoja.para": "¿Para quién?",
  "visor.hoja.para.ambas": "Ambas",
  "visor.hoja.para.d": "La hoja del docente es la misma, ya respondida, para calificar.",
  "visor.hoja.para.docente": "Del docente (con respuestas)",
  "visor.hoja.para.estudiante": "Del estudiante",
  "visor.hoja.pocas": "Hay pocas palabras con imagen para esa selección.",
  "visor.hoja.preparar": "Preparar hojas",
  "visor.hoja.reabierta": "Con las mismas opciones sale igual que la vez anterior (código *{c}*). Si cambias algo, saldrá con un código nuevo.",
  "visor.hoja.respuestas": "Incluir hoja de respuestas",
  "visor.hoja.segunda": "Segunda lengua (opcional)",
  "visor.hoja.titulo": "Hojas para imprimir",
  "visor.hoja.unaVersion": "Una",
  "visor.hoja.une": "Une con una línea",
  "visor.hoja.une.inst": "Une cada imagen con su palabra.",
  "visor.hoja.une.inst.pista": "Une cada pista con su palabra.",
  "visor.hoja.version": "Versión {v}",
  "visor.hoja.versiones": "Versiones",
  "visor.hoja.versiones.d": "Con varias versiones, cada fila del salón puede tener una distinta para que no se copien.",
  "visor.imagenAmpliada": "Imagen ampliada",
  "visor.imagenDe": "Imagen: {credito}",
  "visor.imp.abierta": "Se abrió la vista de impresión: {x}.",
  "visor.imp.ajustes": "Ajustes de impresión",
  "visor.imp.ajustes.d": "Se recuerdan en este dispositivo para la próxima vez.",
  "visor.imp.bloqueada": "El navegador bloqueó la ventana de impresión. Permite las ventanas emergentes para este sitio.",
  "visor.imp.boton": "Imprimir o guardar PDF",
  "visor.imp.codigo": "Hoja {c}",
  "visor.imp.creditos": "Incluir créditos del visor en el pie",
  "visor.imp.editables": "Las pistas con borde azul se pueden escribir aquí antes de imprimir.",
  "visor.imp.horizontal": "Imprime en *Carta*, *horizontal*, escala 100 %.",
  "visor.imp.horizontalCorto": "horizontal, escala 100 %",
  "visor.imp.letra": "Letra",
  "visor.imp.letra.escolar": "Escolar",
  "visor.imp.letra.grande": "Grande",
  "visor.imp.letra.normal": "Normal",
  "visor.imp.licencia": "Licencia: {x}",
  "visor.imp.margen": "Márgenes",
  "visor.imp.margen.minimo": "Mínimos",
  "visor.imp.margen.seguro": "Seguros (impresoras que recortan)",
  "visor.imp.mayus": "Letras de las palabras",
  "visor.imp.mayus.como": "Como están en el catálogo",
  "visor.imp.mayus.mayus": "MAYÚSCULAS",
  "visor.imp.mayus.minus": "minúsculas",
  "visor.imp.nHojas": "{n} hoja(s)",
  "visor.imp.papel": "Papel",
  "visor.imp.papel.a4": "A4",
  "visor.imp.papel.carta": "Carta",
  "visor.imp.papel.oficio": "Oficio",
  "visor.imp.tinta": "Tinta",
  "visor.imp.tinta.bn": "Blanco y negro (ahorra tinta)",
  "visor.imp.tinta.color": "Color",
  "visor.imp.vertical": "Imprime en *Carta*, *vertical*, escala 100 %.",
  "visor.imp.verticalCorto": "vertical, escala 100 %",
  "visor.imp2.armaSilabas": "Arma con sílabas",
  "visor.imp2.armaSilabas.d": "Cuatro palabras por hoja: las sílabas aparecen desordenadas para recortar y pegar.",
  "visor.imp2.armaSilabas.idea1": "Formar palabras nuevas con las sílabas sobrantes.",
  "visor.imp2.armaSilabas.idea2": "Trabajar en pareja: uno arma y el otro lee.",
  "visor.imp2.armaSilabas.idea3": "Guardar las piezas en un sobre para otro día.",
  "visor.imp2.armaSilabas.inst": "Recorta las sílabas y pégalas en orden.",
  "visor.imp2.ayuda.dobles": "Marcar letras dobles (kw, ts…)",
  "visor.imp2.ayuda.inicial": "Primera letra",
  "visor.imp2.ayuda.largo": "Número de letras",
  "visor.imp2.ayuda.lista": "Lista de palabras a la vista",
  "visor.imp2.ayuda.una": "Una letra ya puesta",
  "visor.imp2.ayudas": "Ayudas",
  "visor.imp2.banco": "Palabras:",
  "visor.imp2.basta": "Basta en papel",
  "visor.imp2.basta.blanco": "Encabezados en blanco (para escribirlos a mano)",
  "visor.imp2.basta.cats": "Categorías (columnas)",
  "visor.imp2.basta.d": "La tabla clásica con las categorías del catálogo (o en blanco). La hoja del docente trae letras para cantar y ejemplos del catálogo.",
  "visor.imp2.basta.filas": "Rondas",
  "visor.imp2.basta.idea1": "Canta letras del alfabeto de la lengua, incluidas las dobles.",
  "visor.imp2.basta.idea2": "Juega por equipos: cada equipo con una hoja.",
  "visor.imp2.basta.idea3": "Al terminar, lean en voz alta y comenten si en otras casas se dice distinto.",
  "visor.imp2.basta.inst": "Escribe una palabra para cada columna con la letra que se cante.",
  "visor.imp2.basta.letra": "Letra",
  "visor.imp2.basta.puntos": "Puntos",
  "visor.imp2.basta.total": "Total",
  "visor.imp2.cot.columnas": "Columnas",
  "visor.imp2.cot.filas": "Renglones",
  "visor.imp2.cot.hab.ph": "Separado por comas: lee, escribe, pronuncia…",
  "visor.imp2.cot.habilidades": "Lo que yo escriba",
  "visor.imp2.cot.nombre": "Nombre",
  "visor.imp2.cot.nota": "✓ lo logra · ~ en proceso · ○ todavía no (o la marca que prefieras).",
  "visor.imp2.cot.palabras": "Palabras de la selección",
  "visor.imp2.cotejo": "Lista de cotejo",
  "visor.imp2.cotejo.d": "Para anotar a mano. El visor no pide ni guarda nombres.",
  "visor.imp2.cotejo.idea1": "Usarla en varias sesiones con distinto color.",
  "visor.imp2.cotejo.idea2": "Compartirla con la familia en reuniones.",
  "visor.imp2.cotejo.idea3": "Escribir tus propias columnas en la lengua.",
  "visor.imp2.cotejo.inst": "Lista de cotejo",
  "visor.imp2.crucigrama": "Crucigrama",
  "visor.imp2.crucigrama.d": "Las palabras se cruzan por una letra común. Elige la pista y las ayudas; con niveles salen tres versiones para el mismo grupo.",
  "visor.imp2.crucigrama.idea1": "Pista con imagen y primera letra: para primeros grados.",
  "visor.imp2.crucigrama.idea2": "Pista con la descripción en la lengua: lectura de comprensión sin traducir.",
  "visor.imp2.crucigrama.idea3": "Pista en blanco: escribe tus propias adivinanzas antes de imprimir.",
  "visor.imp2.crucigrama.inst": "Completa el crucigrama.",
  "visor.imp2.cuantas": "Palabras por hoja",
  "visor.imp2.empiezaCon": "empieza con «{l}»",
  "visor.imp2.esc.docente": "Lee las palabras en este orden (o escanea el QR para escucharlas). La flecha indica el número de la imagen.",
  "visor.imp2.esc.qr": "QR con el audio de cada palabra (si lo tiene)",
  "visor.imp2.escucha": "Escucha y señala",
  "visor.imp2.escucha.d": "Hoja con imágenes numeradas; la hoja del docente trae el orden para leer y un QR con el audio de cada palabra.",
  "visor.imp2.escucha.idea1": "Que una persona mayor diga las palabras.",
  "visor.imp2.escucha.idea2": "Repetir con otra variante de la lengua.",
  "visor.imp2.escucha.idea3": "Al terminar, cada quien dice una palabra y el grupo la señala.",
  "visor.imp2.escucha.inst": "Escucha y escribe en cada cuadro en qué orden se dijo la palabra.",
  "visor.imp2.glo.desc": "Incluir la descripción (si la hay)",
  "visor.imp2.glo.otra": "Mostrar también en",
  "visor.imp2.glo.solo": "Solo en la lengua elegida",
  "visor.imp2.glosario": "Glosario de la unidad",
  "visor.imp2.glosario.d": "Ordenado alfabéticamente; doce palabras por hoja.",
  "visor.imp2.glosario.idea1": "Entregarlo al inicio de un tema para el cuaderno.",
  "visor.imp2.glosario.idea2": "Pedir que agreguen dibujos o ejemplos propios.",
  "visor.imp2.glosario.idea3": "Enviarlo a casa para repasar en familia.",
  "visor.imp2.glosario.inst": "Glosario",
  "visor.imp2.horizontales": "Horizontales",
  "visor.imp2.ideas": "Ideas de uso",
  "visor.imp2.ideas.d": "Son solo ideas: tú sabes qué le sirve a tu grupo.",
  "visor.imp2.ini.elegir": "Encerrar entre tres letras",
  "visor.imp2.ini.escribir": "Escribir la letra",
  "visor.imp2.ini.modo": "Forma de responder",
  "visor.imp2.inicial": "¿Con qué letra empieza?",
  "visor.imp2.inicial.d": "Cada palabra con su pista y tres letras para elegir, o un cuadro para escribirla.",
  "visor.imp2.inicial.idea1": "Decir la palabra en voz alta antes de elegir la letra.",
  "visor.imp2.inicial.idea2": "Pedir otra palabra que empiece con la misma letra.",
  "visor.imp2.inicial.idea3": "Usar las letras que el grupo ya conoce (en «Filtros»).",
  "visor.imp2.inicial.inst": "¿Con qué letra empieza cada palabra?",
  "visor.imp2.mem.dosImagenes": "Dos imágenes distintas de la misma palabra",
  "visor.imp2.mem.dosLenguas": "La palabra en dos lenguas",
  "visor.imp2.mem.imagenEspacio": "Imagen y carta en blanco",
  "visor.imp2.mem.imagenPalabra": "Imagen y palabra",
  "visor.imp2.mem.pares": "Cada par",
  "visor.imp2.mem.reverso": "Incluir el reverso (para imprimir del otro lado)",
  "visor.imp2.memorama": "Memorama recortable",
  "visor.imp2.memorama.d": "16 cartas (8 pares) por hoja, con líneas para recortar y, si quieres, el reverso para imprimir del otro lado.",
  "visor.imp2.memorama.idea1": "Imagen y palabra: relacionar lo que se ve con lo que se lee.",
  "visor.imp2.memorama.idea2": "Dos lenguas: para quien aprende la lengua como segunda lengua.",
  "visor.imp2.memorama.idea3": "Imagen y espacio: el grupo escribe la palabra en su carta antes de jugar.",
  "visor.imp2.memorama.inst": "",
  "visor.imp2.nivel.dificil": "Difícil",
  "visor.imp2.nivel.facil": "Fácil",
  "visor.imp2.nivel.medio": "Medio",
  "visor.imp2.niveles": "Niveles para grupos multigrado (fácil, medio y difícil con las mismas palabras)",
  "visor.imp2.niveles.d": "Cada nivel trae sus propias ayudas; las de arriba se usan cuando no hay niveles.",
  "visor.imp2.opciones": "Opciones",
  "visor.imp2.ordena": "Ordena las letras",
  "visor.imp2.ordena.d": "Cada palabra con sus letras revueltas (con sus letras dobles juntas) y espacio para escribirla.",
  "visor.imp2.ordena.idea1": "Con primera letra fija: para quien empieza.",
  "visor.imp2.ordena.idea2": "Recorta las letras y que el grupo las acomode en la mesa.",
  "visor.imp2.ordena.idea3": "Sin pista: adivinar la palabra solo con sus letras.",
  "visor.imp2.ordena.inst": "Ordena las letras y escribe la palabra.",
  "visor.imp2.pc.ambas": "Las dos",
  "visor.imp2.pc.dibuja": "Espacio para dibujar",
  "visor.imp2.pc.dibujaAqui": "Dibuja aquí una de tus palabras",
  "visor.imp2.pc.donde": "Dónde se dice así",
  "visor.imp2.pc.grupo": "Grupo",
  "visor.imp2.pc.guia": "Incluir la guía para el docente (editable)",
  "visor.imp2.pc.guia.1": "*Antes, en el salón:* presenta el tema y llena la hoja junto con el grupo usando una palabra de ejemplo.",
  "visor.imp2.pc.guia.2": "Explica que van a preguntar a quien sepa en casa (abuelas, abuelos, mamá, papá, vecinos) y que esa persona decide si quiere aparecer con su nombre.",
  "visor.imp2.pc.guia.3": "*Preguntas para hacer en casa:* ¿Cómo se dice…? ¿Para qué se usa? ¿Quién te lo enseñó? ¿En otros lugares se dice distinto?",
  "visor.imp2.pc.guia.4": "*Al regresar:* cada quien lee su palabra en voz alta, el grupo la escribe en el pizarrón y comentan si en otras casas se dice de otra manera. Esa conversación es en sí misma una clase de lengua.",
  "visor.imp2.pc.guia.5": "*Para enviarlas al catálogo:* captúralas con «Proponer varias» (con el crédito de quien las enseñó, si lo permitió) o toma fotos de las hojas con «Enviar fotos de hojas». Recorta antes el talón con el nombre del alumno.",
  "visor.imp2.pc.guia.6": "Esta guía es una sugerencia: cambia lo que necesites antes de imprimir (toca el texto para editarlo).",
  "visor.imp2.pc.guia.t": "Guía para el docente · Pregunta en casa",
  "visor.imp2.pc.investiga": "Qué significan algunas",
  "visor.imp2.pc.investiga.t": "Pregunta qué significan estas palabras y escríbelo en la lengua",
  "visor.imp2.pc.invitacion": "Invitación para la familia (puedes escribirla en la lengua)",
  "visor.imp2.pc.invitacion.pred": "Esta hoja es parte del catálogo {catalogo} de nuestra comunidad. Les invitamos a compartir palabras para sumarlas, si así lo desean. Gracias por enseñar.",
  "visor.imp2.pc.loQueQuieras": "lo que quieras",
  "visor.imp2.pc.nuevas": "Palabras nuevas",
  "visor.imp2.pc.nuevas.t": "Palabras nuevas sobre {tema}",
  "visor.imp2.pc.palabra": "Palabra",
  "visor.imp2.pc.permiso": "Pueden mencionar mi nombre como quien compartió estas palabras:",
  "visor.imp2.pc.qr": "Código QR para enviar las palabras desde un teléfono",
  "visor.imp2.pc.qr.d": "Con un teléfono, escanea este código para enviar las palabras al catálogo.",
  "visor.imp2.pc.queEs": "Qué es o qué significa",
  "visor.imp2.pc.quienEnseno": "Firma o nombre de quien las enseñó:",
  "visor.imp2.pc.sinPropuestas": "El catálogo no recibe propuestas por ahora: las palabras se pueden mandar al equipo por correo.",
  "visor.imp2.pc.talon": "Nombre (se queda en el salón)",
  "visor.imp2.pc.tema": "Tema",
  "visor.imp2.pc.temaLibre": "Libre (cualquier palabra)",
  "visor.imp2.pc.tipo": "¿Qué se busca?",
  "visor.imp2.pc.yaEstan": "Ya están en el catálogo:",
  "visor.imp2.pocas": "Hacen falta al menos {n} palabras que cumplan lo elegido. Prueba con otras categorías, otra lista u otra pista.",
  "visor.imp2.preguntaCasa": "Pregunta en casa",
  "visor.imp2.preguntaCasa.d": "Para llevar a casa: buscar con la familia palabras que no están en el catálogo, o investigar qué significan algunas. Al regresar, compártanlas con «Proponer».",
  "visor.imp2.preguntaCasa.idea1": "Antes, muestra en clase cómo llenar la hoja con una palabra de ejemplo.",
  "visor.imp2.preguntaCasa.idea2": "Al regresar, cada quien lee su palabra y el grupo comenta si en otras casas se dice distinto.",
  "visor.imp2.preguntaCasa.idea3": "Envía las palabras con «Proponer varias» o fotos de las hojas; recorta antes el talón con el nombre.",
  "visor.imp2.preguntaCasa.inst": "Pregunta en tu casa y escribe lo que te enseñen.",
  "visor.imp2.rec.evento": "Reunión o taller (opcional)",
  "visor.imp2.rec.evento.ph": "Ej. Asamblea de San Andrés, marzo",
  "visor.imp2.rec.fecha": "Fecha:",
  "visor.imp2.rec.lugar": "Lugar:",
  "visor.imp2.rec.nota": "Nada se publica sin revisión del equipo del catálogo. Si alguien quiere retirar su nombre, puede pedirlo en cualquier momento.",
  "visor.imp2.rec.permiso": "¿Su nombre en el catálogo?",
  "visor.imp2.rec.quien": "Quién la compartió",
  "visor.imp2.recoleccion": "Hoja de recolección",
  "visor.imp2.recoleccion.d": "Para reuniones comunitarias y talleres sin dispositivos: las mismas preguntas del formulario de propuestas, en papel.",
  "visor.imp2.recoleccion.idea1": "Una persona anota mientras otras comparten.",
  "visor.imp2.recoleccion.idea2": "Pregunta siempre si se puede mencionar el nombre de quien compartió.",
  "visor.imp2.recoleccion.idea3": "Después, captura todo con «Proponer varias» o envía fotos de la hoja.",
  "visor.imp2.recoleccion.inst": "Palabras compartidas en la reunión",
  "visor.imp2.silabas": "Cuenta las sílabas",
  "visor.imp2.silabas.d": "Usa las sílabas registradas en el catálogo. La hoja del docente muestra la división.",
  "visor.imp2.silabas.idea1": "Aplaudir en grupo antes de colorear.",
  "visor.imp2.silabas.idea2": "Comparar palabras largas y cortas.",
  "visor.imp2.silabas.idea3": "Pedir a alguien de la familia que diga la palabra y contar juntos.",
  "visor.imp2.silabas.inst": "Aplaude cada palabra y colorea una palma por cada sílaba.",
  "visor.imp2.sopa": "Sopa de letras",
  "visor.imp2.sopa.d": "Las palabras se cruzan y se reparten en varias direcciones. La lista puede ser la palabra, su imagen, otra lengua, una descripción o un espacio para dictar.",
  "visor.imp2.sopa.idea1": "Con imágenes en la lista: para quienes empiezan a leer.",
  "visor.imp2.sopa.idea2": "Con la palabra en otra lengua: para practicar vocabulario.",
  "visor.imp2.sopa.idea3": "Con espacios en blanco: el docente dicta y el grupo escribe antes de buscar.",
  "visor.imp2.sopa.inst": "Encuentra las palabras.",
  "visor.imp2.trazo": "Hoja de trazo",
  "visor.imp2.trazo.d": "Elige las letras: cada hoja repasa una, con renglones libres y palabras que la contienen.",
  "visor.imp2.trazo.idea1": "Trazar primero en el aire o en la arena.",
  "visor.imp2.trazo.idea2": "Buscar la letra en carteles o libros de la comunidad.",
  "visor.imp2.trazo.idea3": "Elegir la letra «Escolar» en los ajustes de impresión.",
  "visor.imp2.trazo.inst": "Repasa la letra y luego escríbela tú.",
  "visor.imp2.trazo.letras": "Letras (una por hoja; hasta 12)",
  "visor.imp2.trazo.palabras": "Agregar palabras del catálogo que tienen la letra",
  "visor.imp2.verticales": "Verticales",
  "visor.imprimir.intro": "Materiales en tamaño carta para escuelas y comunidades, también donde no hay dispositivos ni internet.",
  "visor.interfaz.avance": "{n} de {total} textos",
  "visor.interfaz.boton": "Idioma de los menús",
  "visor.interfaz.delCatalogo": "del catálogo",
  "visor.interfaz.enDispositivo": "instalada en este dispositivo",
  "visor.interfaz.error.archivo": "Ese archivo no es una traducción de los menús del catálogo.",
  "visor.interfaz.error.espacio": "No hay espacio en este dispositivo para guardar la traducción.",
  "visor.interfaz.error.vacio": "Ese archivo no trae textos para este visor.",
  "visor.interfaz.instalada": "Traducción instalada: {nombre}",
  "visor.interfaz.instalar": "Instalar una traducción",
  "visor.interfaz.instalar.d": "Si alguien te compartió el archivo de una traducción de los menús (.json), ábrelo aquí. Se guarda solo en este dispositivo.",
  "visor.interfaz.instalarBoton": "Instalar y usar",
  "visor.interfaz.intro": "Elige en qué idioma ves los menús y botones. Es distinto de la lengua de las palabras. Lo que no esté traducido se ve en español.",
  "visor.interfaz.previa": "*{nombre}*: {n} de {total} textos traducidos.",
  "visor.interfaz.quitar": "Quitar",
  "visor.interfaz.titulo": "Idioma de los menús",
  "visor.interfaz.usar": "Usar este idioma",
  "visor.jugador": "Jugador",
  "visor.jugadores": "Jugadores",
  "visor.jugar.intro": "Juegos con las mismas palabras, imágenes y voces del catálogo. Al terminar, cada juego te lleva de vuelta a las palabras.",
  "visor.jugarBoton": "Jugar",
  "visor.lengua": "Lengua",
  "visor.lenguasCatalogo": "Lenguas en este catálogo",
  "visor.licencia": "Condiciones de uso",
  "visor.listas.borrar": "¿Borrar esta lista? Las actividades que la usan dejarán de encontrar sus palabras.",
  "visor.listas.d": "Arma la lista de palabras de tu unidad y úsala en todos los materiales, para que el vocabulario sea el mismo en todo. Se guardan en este dispositivo (y en el respaldo de Mis actividades).",
  "visor.listas.editar": "Editar lista",
  "visor.listas.elegidas": "{n} elegida(s)",
  "visor.listas.faltaNombre": "Ponle un nombre a la lista.",
  "visor.listas.faltan": "Elige al menos una palabra.",
  "visor.listas.guardada": "Lista guardada",
  "visor.listas.n": "{n} palabra(s)",
  "visor.listas.nombre": "Nombre de la lista",
  "visor.listas.nombre.ph": "Ej. La milpa · 3.º",
  "visor.listas.nueva": "Nueva lista",
  "visor.listas.palabras": "Palabras",
  "visor.listas.soloElegidas": "Ver solo las elegidas",
  "visor.listas.titulo": "Mis listas de palabras",
  "visor.listas.vacia": "Todavía no tienes listas.",
  "visor.listo": "Listo",
  "visor.lot.archivo": "Descargar archivo para compartir (Bluetooth, WhatsApp): cantar, jugar y verificar sin internet",
  "visor.lot.audio": "Sonar la voz al cantar (solo cartas con audio)",
  "visor.lot.automatico": "Automático cada",
  "visor.lot.avanzarSolo": "Avanzar solo cada",
  "visor.lot.avanzoYo": "Avanzo yo",
  "visor.lot.cambiarCartas": "Cambiar cartas",
  "visor.lot.cantarAqui": "Cantar aquí (tablas impresas o descargadas)",
  "visor.lot.cantidad": "Número de cartas ({min} a {max}; la tradicional tiene {rec})",
  "visor.lot.cantor": "Cantor",
  "visor.lot.cargar": "Cargar una configuración guardada",
  "visor.lot.carta": "Carta (1 por hoja)",
  "visor.lot.cartaCantor": "Carta del cantor",
  "visor.lot.cartasCantar": "Cartas para cantar (8 por hoja)",
  "visor.lot.casi": "Casi: aún no han salido *{x}*. Quité esos frijoles; sigue jugando.",
  "visor.lot.categorias": "Categorías",
  "visor.lot.colores": "En color: ya salieron. Opacas: aún no.",
  "visor.lot.comoEligen": "¿Cómo se eligen las cartas?",
  "visor.lot.completarAzar": "Completar al azar",
  "visor.lot.completarOtras": "Completar con palabras de otras categorías",
  "visor.lot.confirmaRonda": "¿Empezar una ronda nueva? Las tablas siguen siendo las mismas; cambia el orden de las cartas.",
  "visor.lot.contador": "Ronda {ronda} · carta {n} de {total}",
  "visor.lot.creando": "Creando la partida…",
  "visor.lot.crear": "Crear partida «{nombre} + 4 números»",
  "visor.lot.d.carta": "{n} tablas en carta · {h} hojas",
  "visor.lot.d.cartas": "{n} cartas · {h} hojas",
  "visor.lot.d.medias": "{n} tablas en media carta · {h} hojas",
  "visor.lot.deA": "(de {min} a {max})",
  "visor.lot.distancia": "Jugar a distancia (cada quien en su celular)",
  "visor.lot.elegidas": "Elegidas:",
  "visor.lot.eligeCartas": "Elige las cartas",
  "visor.lot.enCurso": "Tienes una partida en curso: *{x}*",
  "visor.lot.entrando": "Entrando…",
  "visor.lot.entrar": "Entrar",
  "visor.lot.error.archivo": "Ese archivo no es una configuración de lotería.",
  "visor.lot.error.cantidad": "El número de cartas va de {min} a {max}.",
  "visor.lot.error.cargar": "No se pudo cargar: {x}",
  "visor.lot.error.crear": "No se pudo crear la partida: {x}.\\nPuedes cantar aquí sin internet o descargar el archivo.",
  "visor.lot.error.elige": "Elige de {min} a {max} cartas; tienes {n}.",
  "visor.lot.error.entrar": "No se pudo entrar: {x}",
  "visor.lot.error.formas": "Elige al menos una forma de ganar.",
  "visor.lot.error.nombre": "El nombre del juego necesita letras o números.",
  "visor.lot.error.otroCatalogo": "Esta partida usa cartas de otro catálogo o que ya no están publicadas.",
  "visor.lot.error.script": "El script del catálogo necesita actualizarse.",
  "visor.lot.error.tabla": "No se pudo generar esa tabla.",
  "visor.lot.esperaSiguiente": "Ronda terminada. Espera la siguiente ronda.",
  "visor.lot.esperando": "Esperando jugadores…",
  "visor.lot.faltan": "Faltan cartas",
  "visor.lot.faltan.audio": "Con «Sonar la voz» solo entran cartas con audio.",
  "visor.lot.faltan.d": "Con estas condiciones hay *{hay}* cartas y pediste *{pediste}*.",
  "visor.lot.formasGanar": "Formas de ganar",
  "visor.lot.formato": "Tamaño de las tablas",
  "visor.lot.frijol": "Toca una imagen para poner o quitar tu frijol 🫘. Al completar una forma de ganar, la tabla se revisa sola.",
  "visor.lot.gana": "¡Esta tabla gana!",
  "visor.lot.ganar": "Ganar: {formas}",
  "visor.lot.ganaste": "¡Ganaste: {forma}!",
  "visor.lot.ganoOtro": "{nombre} acaba de ganar: {forma}",
  "visor.lot.grita": "¡{forma}! Grita «¡Lotería!»",
  "visor.lot.guardada": "Configuración guardada en tu dispositivo. Para usarla otra vez: Lotería → «Cargar una configuración guardada».",
  "visor.lot.guardar": "Guardar configuración",
  "visor.lot.hay": "En esta lengua hay {n} palabras con imagen. Prueba otra lengua o pide al equipo más imágenes y palabras.",
  "visor.lot.hayConAudio": "En esta lengua hay {n} palabras con imagen y audio. Prueba otra lengua o pide al equipo más imágenes y palabras.",
  "visor.lot.imprimir": "Imprimir (hoja carta, margen de 5 mm)",
  "visor.lot.invitacion": "¡Juguemos lotería! Partida: {partida}. Entra aquí: {enlace}",
  "visor.lot.juegoN": "Juego {x}",
  "visor.lot.juegoSemilla": "Juego *{x}*. Las tablas impresas, descargadas y en pantalla de este juego son las mismas: guarda la configuración para volver a usarlas.",
  "visor.lot.jugadorTabla": "{nombre} (tabla {n})",
  "visor.lot.jugarCon": "Jugar con {n} cartas",
  "visor.lot.lasElijo": "Las elijo yo",
  "visor.lot.lengua": "Lengua de las cartas",
  "visor.lot.lengua.d": "Solo entran palabras con imagen y escritas en esta lengua.",
  "visor.lot.listo": "Archivo listo ({mb} MB). Compártelo por Bluetooth, WhatsApp o memoria USB: se abre en el navegador del celular sin internet.",
  "visor.lot.local": "Sin conexión · juego *{x}*: se juega con tablas impresas o descargadas de este juego.",
  "visor.lot.marcada": "{x}, marcada",
  "visor.lot.masCats": "Elegir más categorías",
  "visor.lot.media": "Media carta (2 por hoja)",
  "visor.lot.miTabla": "Mi tabla",
  "visor.lot.nCartas": "*{n} cartas*",
  "visor.lot.noGana": "Todavía no completa ninguna forma de ganar.",
  "visor.lot.nombreJuego": "Nombre del juego",
  "visor.lot.nombreJuego.d": "(sale en las tablas y sirve para unirse)",
  "visor.lot.nuevaRonda": "Nueva ronda (mismas tablas)",
  "visor.lot.nuevaRondaCorto": "Nueva ronda",
  "visor.lot.numTabla": "Número de tabla",
  "visor.lot.off.azar": "Al azar",
  "visor.lot.off.cambiar": "Cambiar de tabla",
  "visor.lot.off.cantar": "Cantar",
  "visor.lot.off.cartas": "{n} cartas · funciona sin internet",
  "visor.lot.off.colores": "En color: ya salieron (ronda {ronda}, {n} cartas).",
  "visor.lot.off.confirma": "¿Empezar una ronda nueva?",
  "visor.lot.off.frijol": "Toca una imagen para poner o quitar tu frijol 🫘. Si completas, grita «¡Lotería!» y di tu número de tabla: quien canta la revisa en «Verificar».",
  "visor.lot.off.numero": "Escribe el número de tu tabla. Pídelo a quien canta (cada persona usa uno distinto), o toca «Al azar».",
  "visor.lot.off.quitar": "Quitar frijoles",
  "visor.lot.off.retomar": "o retomar la ronda",
  "visor.lot.off.revisa": "Revisa la tabla de quien grita «¡Lotería!» con las cartas que ya cantaste en este celular.",
  "visor.lot.off.revisar": "Revisar",
  "visor.lot.off.usar": "Usar esta tabla",
  "visor.lot.off.verificar": "Verificar",
  "visor.lot.otra": "Otra lotería",
  "visor.lot.otraLengua": "Cada quien puede ver su tabla en otra lengua",
  "visor.lot.paraUnirse": "Para unirse: *{x}*",
  "visor.lot.perdidas": "{n} carta(s) ya no están disponibles y se quitaron. Las tablas cambian: vuelve a imprimirlas.",
  "visor.lot.pide": "Pide a quien canta el nombre de la partida y sus 4 números.",
  "visor.lot.preparando": "Preparando imágenes ligeras…",
  "visor.lot.primera": "Esperando la primera carta…",
  "visor.lot.primero": "¡Bien! Completaste {forma}, pero {nombre} lo logró primero.",
  "visor.lot.quedan": "Quedan {n} cartas válidas. Elige más o usa «Completar al azar».",
  "visor.lot.quitarTodas": "Quitar todas",
  "visor.lot.revisen": "Si dice que completó y no aparece ✅, revisen estas que aún no salen: {x}…",
  "visor.lot.rondaNueva": "¡Nueva ronda! Quité tus frijoles.",
  "visor.lot.rondaX": "ronda {x}",
  "visor.lot.salio": "Salió ({n} de {total}):",
  "visor.lot.salir": "¿Salir de esta lotería?",
  "visor.lot.siguiente": "Siguiente carta",
  "visor.lot.sinImagen": "{n} carta(s) sin imagen incrustada: se verán con su número y nombre.",
  "visor.lot.sinInternet": "Sin internet",
  "visor.lot.sinMarcar": "(sin marcar = todas)",
  "visor.lot.sinSenal": "Sin conexión. Escribe en qué carta va el cantor para revisar tu tabla:",
  "visor.lot.tablaN": "Tabla {n}",
  "visor.lot.tablas4": "Tablas 4×4",
  "visor.lot.tablasDos": "Tablas:",
  "visor.lot.terminada": "Ronda terminada",
  "visor.lot.terminar": "Terminar ronda",
  "visor.lot.terminarTodos": "¿Terminar esta ronda para todos?",
  "visor.lot.toca": "Toca *Siguiente carta* para empezar.",
  "visor.lot.todas": "Ya salieron todas las cartas",
  "visor.lot.tuNombre": "Tu nombre (opcional)",
  "visor.lot.tuNombre.d": "Así te anunciarán si ganas",
  "visor.lot.unirmePartida": "Unirme a una partida",
  "visor.lot.var.centro": "Centro (pozo)",
  "visor.lot.var.esquinas": "Cuatro esquinas",
  "visor.lot.var.linea": "Línea (chorro)",
  "visor.lot.var.llena": "Tabla llena",
  "visor.lot.verEn": "Ver mi tabla en:",
  "visor.lot.verificar": "Verificar una tabla",
  "visor.lot.verificar.d": "Ronda {ronda} · {n} cartas cantadas. Escribe el número de la tabla (está arriba a la derecha).",
  "visor.lot.verso": "Espacio en cada carta para escribir un verso o adivinanza en la lengua",
  "visor.lot.volverCantar": "Volver a cantar",
  "visor.lot.volverPartida": "Volver a esa partida",
  "visor.lot.yaSalieron": "Cartas que ya salieron ({n})",
  "visor.mant.borrado": "Por precaución, se borró la copia que este dispositivo tenía guardada.",
  "visor.mant.congelado": "Estamos actualizando el catálogo: por ahora ves la versión anterior.",
  "visor.mant.contacto": "¿Tienes dudas? Escríbenos:",
  "visor.mant.d": "Estamos actualizando el catálogo. Vuelve pronto: gracias por tu paciencia.",
  "visor.mant.equipo": "Modo equipo: el catálogo está en mantenimiento ({modo}). Tú ves los cambios en curso.",
  "visor.mant.estado": "En mantenimiento",
  "visor.mant.hasta": "Volvemos: {fecha}",
  "visor.mant.m.cerrar": "el público ve la pantalla de mantenimiento",
  "visor.mant.m.congelar": "el público ve la versión anterior",
  "visor.mant.titulo": "En mantenimiento",
  "visor.mem.cara.audio": "Audio",
  "visor.mem.cara.img": "Imagen",
  "visor.mem.cara.txt": "Palabra escrita",
  "visor.mem.caraA": "Cara A",
  "visor.mem.caraB": "Cara B",
  "visor.mem.completado": "¡Memorama completado!",
  "visor.mem.contador": "Pares: *{n}/{total}*",
  "visor.mem.d": "Cada par une dos caras de la misma palabra. Prueba «palabra en una lengua» con «palabra en otra lengua» para conocer la diversidad.",
  "visor.mem.iguales": "Las dos caras son iguales. Cambia el tipo o la lengua de una de ellas.",
  "visor.mem.pares": "Pares",
  "visor.mem.pocas": "Solo hay {n} palabras con esas dos caras. Prueba otra combinación.",
  "visor.mem.turnos": "Varios jugadores: si aciertas, sigues; si fallas, pasa el turno.",
  "visor.mostrarMas": "Mostrar más",
  "visor.muyBien": "¡Muy bien!",
  "visor.nadaFiltros": "Nada con estos filtros.",
  "visor.ninguna": "Ninguna",
  "visor.no": "No",
  "visor.nombreN": "Nombre {n}",
  "visor.nombrePartida": "Nombre de la partida",
  "visor.off.audios": "También los audios",
  "visor.off.bajando": "Guardando…",
  "visor.off.boton": "Llevar sin conexión",
  "visor.off.corto": "Sin conexión",
  "visor.off.d": "Guarda en este dispositivo las imágenes (y, si quieres, los audios) de «{x}»: {n} palabra(s). Así puedes imprimir o proyectar donde no hay internet.",
  "visor.off.estado": "Imágenes ya guardadas: {g} de {t}.",
  "visor.off.fallos": "{n} no se pudieron guardar; vuelve a intentarlo con mejor señal.",
  "visor.off.faltanImg": "Sin conexión: {n} imagen(es) no se verán. Usa 📥 Llevar sin conexión cuando tengas internet.",
  "visor.off.listo": "Listo: {i} imagen(es) y {a} audio(s) guardados en este dispositivo.",
  "visor.off.nota": "Necesita internet ahora. Se guarda solo en este dispositivo y puede borrarse si se limpian los datos del navegador.",
  "visor.off.sinRed": "Ahora no hay conexión. Conéctate y vuelve a intentar.",
  "visor.off.titulo": "Llevar sin conexión",
  "visor.opcion": "Opción",
  "visor.opciones": "Opciones",
  "visor.otraLengua": "Otra lengua",
  "visor.otraPartida": "Otra partida",
  "visor.otraVez": "Otra vez",
  "visor.palabra": "Palabra",
  "visor.palabrasJuego": "Palabras de este juego",
  "visor.pie.apoyo": "Con el apoyo de {x}",
  "visor.pie.codigo": "Código abierto (AGPL-3.0)",
  "visor.pie.contenido": "Contenido de la comunidad",
  "visor.pie.credito": "Visor diseñado y desarrollado por {autor}",
  "visor.pie.licencia": "compartido bajo {licencia}",
  "visor.pista": "Pista",
  "visor.pista.descripcion": "Descripción del catálogo",
  "visor.pista.enLengua": "¿En qué lengua?",
  "visor.pista.escribeAqui": "✍️ Escribe aquí la pista…",
  "visor.pista.espacio": "Espacio para escribir",
  "visor.pista.espacio.d": "Deja líneas para que tú o el grupo escriban la pista a mano: una instrucción, una adivinanza, lo que necesites.",
  "visor.pista.hayDesc": "{n} palabra(s) tienen descripción en {lengua}.",
  "visor.pista.imagen": "Imagen",
  "visor.pista.imagenEspacio": "Imagen + espacio",
  "visor.pista.lengua": "Otra lengua",
  "visor.pista.ninguna": "Sin pista",
  "visor.pista.noHayDesc": "Todavía no hay descripciones en {lengua}; se usará lo que elijas abajo.",
  "visor.pista.palabra": "La palabra escrita",
  "visor.pista.propia": "Pista propia",
  "visor.pista.propia.d": "En la vista previa podrás escribir o cambiar la pista de cada palabra (parte de la descripción del catálogo, si existe). Se guarda en este dispositivo y no cambia el catálogo.",
  "visor.pista.proponer": "Proponer al catálogo",
  "visor.pista.proponer.d": "Enviar esta pista al equipo como descripción en la lengua (la revisan antes de agregarla).",
  "visor.pista.segunda": "Segundo renglón",
  "visor.pista.siNoHay": "Si una palabra no tiene descripción, usar:",
  "visor.pista.titulo": "Pista de cada palabra",
  "visor.pista.volver": "Volver a la del catálogo",
  "visor.prop.acepto": "Acepto que el equipo revise esto y, si lo acepta, lo use en el catálogo y sus materiales educativos, con el crédito que elegí.",
  "visor.prop.actividad": "Actividad o código de la hoja (opcional)",
  "visor.prop.actividad.ph": "Ej. Pregunta en casa · La milpa · 7K3P",
  "visor.prop.ahoraDice": "Ahora dice:",
  "visor.prop.alternativa": "¿No se puede enviar o prefieres el correo?",
  "visor.prop.alternativa.d": "Escribe al equipo. Si son fotos, adjúntalas al mensaje.",
  "visor.prop.anota": "Anota este código o tómale captura: con él ves en «Mis propuestas» qué pasó.",
  "visor.prop.aportes": "Quiénes han aportado ({n})",
  "visor.prop.asunto": "Propuesta para el catálogo {catalogo}",
  "visor.prop.asunto.corr": "Corrección de «{palabra}»",
  "visor.prop.asunto.desc": "Descripción de «{palabra}»",
  "visor.prop.asunto.fotos": "Fotos de hojas {act}",
  "visor.prop.asunto.img": "Imagen para «{palabra}»",
  "visor.prop.asunto.palabras": "Palabras para el catálogo {act}",
  "visor.prop.asunto.perdi": "Pregunta sobre una aportación",
  "visor.prop.asunto.retiro": "Retiro de una aportación {codigo}",
  "visor.prop.asunto.var": "Cómo se dice «{palabra}» en mi comunidad",
  "visor.prop.ayuda": "¿Sabes algo de esta palabra?",
  "visor.prop.categoria": "Categoría (casa, animales, fiesta…)",
  "visor.prop.cerradas": "Este catálogo no recibe propuestas desde el visor por ahora. Puedes escribir al equipo (Colabora).",
  "visor.prop.clave": "Clave de colaboración",
  "visor.prop.comoSeEscribe": "Cómo debería escribirse",
  "visor.prop.compartir": "También pueden compartirlo con otras comunidades que usan este sistema.",
  "visor.prop.completa": "«{palabra}» ya tiene las imágenes que el equipo pidió. ¿Nos ayudas con estas?",
  "visor.prop.completa.t": "Esta palabra ya tiene suficientes imágenes",
  "visor.prop.contacto": "Correo o teléfono (privado, solo para el equipo)",
  "visor.prop.contacto.ph": "Nunca se publica",
  "visor.prop.corr": "Proponer una corrección",
  "visor.prop.corr.corto": "Corrección",
  "visor.prop.criterios": "Cómo decide el equipo",
  "visor.prop.cuerpoCorreo": "Hola, quiero proponer para el catálogo {catalogo}:",
  "visor.prop.cupo": "{n} de {max} imágenes",
  "visor.prop.desc": "Proponer una descripción",
  "visor.prop.desc.corto": "Descripción",
  "visor.prop.desc.otra": "Otra descripción",
  "visor.prop.descActual": "Descripción actual:",
  "visor.prop.descartar": "Descartar",
  "visor.prop.descartar.c": "¿Descartar este envío? No se mandará.",
  "visor.prop.descripcion": "Descripción",
  "visor.prop.descripcion.ph": "En la misma lengua: qué es, para qué se usa…",
  "visor.prop.donde": "Dónde ayudar",
  "visor.prop.donde.d": "Lo que le falta al catálogo, por categoría. Toca una palabra para proponer lo que le hace falta.",
  "visor.prop.donde.d2": "Palabras que necesitan imagen o descripción.",
  "visor.prop.donde.desc": "Sin descripción en {lengua} ({n})",
  "visor.prop.donde.img": "Necesitan imagen ({n})",
  "visor.prop.donde.imprimir": "💡 En Imprimir, «Dibuja la palabra» y «Pregunta en casa» pueden elegir justo estas palabras.",
  "visor.prop.dondeSeDice": "Dónde se dice así (comunidad o región)",
  "visor.prop.elegirFotos": "Elegir fotos (hasta 8)",
  "visor.prop.elegirImg": "Tomar foto o elegir imagen",
  "visor.prop.enCola": "Esperando conexión",
  "visor.prop.enLengua": "Lengua de la descripción",
  "visor.prop.enviando": "Enviando…",
  "visor.prop.enviar": "Enviar",
  "visor.prop.err.cerradas": "El catálogo dejó de recibir propuestas.",
  "visor.prop.err.clave": "La clave de colaboración no es correcta.",
  "visor.prop.err.completa": "La {n} ya está completa (mientras enviabas se llenó su cupo). Puedes proponerla para otra palabra.",
  "visor.prop.err.consentimiento": "Falta aceptar el uso.",
  "visor.prop.err.grande": "La imagen {n} es demasiado grande.",
  "visor.prop.err.imagen": "La imagen {n} no se pudo leer. Prueba con otra.",
  "visor.prop.err.lengua": "La {n} tiene una lengua que el catálogo no conoce.",
  "visor.prop.err.limite": "Límite del día alcanzado.",
  "visor.prop.err.mant": "En mantenimiento.",
  "visor.prop.err.palabra": "La palabra {n} ya no está en el catálogo.",
  "visor.prop.err.red": "Sin conexión.",
  "visor.prop.err.tipo": "Algo no se reconoció.",
  "visor.prop.err.vacia": "La {n} estaba vacía.",
  "visor.prop.esMia": "La hice yo o tengo permiso de quien la hizo",
  "visor.prop.espera": "Sin conexión: quedó guardado en este dispositivo y se enviará solo cuando haya internet.",
  "visor.prop.esperaLimite": "Por hoy ya se enviaron muchas propuestas desde aquí: quedó guardado y se enviará mañana.",
  "visor.prop.esperaMant": "El catálogo está en mantenimiento: quedó guardado y se enviará cuando termine.",
  "visor.prop.estado.Aceptada": "aceptada ✅",
  "visor.prop.estado.Noincluida": "no incluida",
  "visor.prop.estado.Pendiente": "pendiente",
  "visor.prop.estado.Transcrita": "transcrita ✅",
  "visor.prop.faltaAcepto": "Para enviar, marca que aceptas que el equipo lo revise y lo use.",
  "visor.prop.faltaComunidad": "Escribe tu comunidad o región.",
  "visor.prop.faltaMia": "Para enviar una imagen, confirma que la hiciste tú o que tienes permiso.",
  "visor.prop.fotos": "Enviar fotos de hojas",
  "visor.prop.fotos.d": "Para hojas trabajadas en clase o en casa (por ejemplo «Pregunta en casa»): el equipo las transcribe.",
  "visor.prop.fotos.d2": "Hojas trabajadas en clase o en casa.",
  "visor.prop.fotos.menores": "Antes de la foto, recorta o tapa el nombre de niñas y niños.",
  "visor.prop.fotos.resumen": "{n} foto(s) de hojas",
  "visor.prop.img": "Proponer una imagen",
  "visor.prop.img.corto": "Imagen",
  "visor.prop.intro": "Si conoces palabras, formas de decirlas, descripciones o tienes dibujos y fotos, puedes proponerlos. El equipo los revisa antes de agregarlos.",
  "visor.prop.lenguaOpc": "— lengua —",
  "visor.prop.lic.0": "Libre, con crédito (CC BY)",
  "visor.prop.lic.1": "Libre sin fines de lucro (CC BY-NC)",
  "visor.prop.lic.2": "Solo en este catálogo",
  "visor.prop.licencia": "Cómo se puede usar",
  "visor.prop.mas": "Más datos (opcional)",
  "visor.prop.masDatos": "Descripción, dónde se dice así, categoría",
  "visor.prop.menores": "No escribas nombres ni datos de niñas y niños.",
  "visor.prop.mias": "Mis propuestas",
  "visor.prop.mias.d": "Cada envío tiene un código. Aquí ves si el equipo ya lo revisó y su mensaje. Los códigos se guardan solo en este dispositivo.",
  "visor.prop.mias.d2": "Qué pasó con lo que enviaste.",
  "visor.prop.mias.vacia": "Todavía no has enviado propuestas desde este dispositivo.",
  "visor.prop.nadaFalta": "¡Nada por aquí!",
  "visor.prop.noHay": "Todavía no está en esta lengua.",
  "visor.prop.nombre": "Nombre para el crédito",
  "visor.prop.nombre.ph": "Ej. Mtra. Rosa Carrillo, Esc. Primaria de San Andrés",
  "visor.prop.notaFotos": "Nota para el equipo (opcional)",
  "visor.prop.otraPalabra": "Otra palabra",
  "visor.prop.palabra": "Palabra",
  "visor.prop.palabras": "Proponer palabras",
  "visor.prop.palabras.d": "Escribe una o varias. Solo la palabra es necesaria; lo demás ayuda al equipo.",
  "visor.prop.palabras.d2": "Una o varias, con lo que sepas de ellas.",
  "visor.prop.pegar": "Pegar una lista",
  "visor.prop.pegar.ph": "Una por renglón. Si quieres, con su significado: palabra = significado",
  "visor.prop.pegar.usar": "Usar la lista",
  "visor.prop.perdiste": "¿Perdiste tu código o quieres retirar algo? Escribe al equipo: di qué aportaste y cómo aparece tu nombre.",
  "visor.prop.porque": "¿Por qué? (opcional)",
  "visor.prop.publico": "Pueden mostrar mi nombre en los créditos",
  "visor.prop.quien": "Quién propone",
  "visor.prop.recibida": "¡Recibido! El equipo lo revisará.",
  "visor.prop.recopilo": "Lo recopiló",
  "visor.prop.recopilo.ph": "Si lo envías por otra persona (por ejemplo, la abuela de un alumno), escribe aquí tu nombre y arriba el de ella.",
  "visor.prop.reintentar": "Enviar ahora",
  "visor.prop.retirarEsta": "Pedir retiro",
  "visor.prop.retiro": "Retirar mi nombre o una aportación",
  "visor.prop.retiro.codigo": "Código (si lo tienes)",
  "visor.prop.retiro.d": "Puedes pedir que se quite tu nombre de los créditos o algo que aportaste. El equipo lo atiende; no se necesita el código, pero ayuda.",
  "visor.prop.retiro.que": "¿Qué quieres que se retire?",
  "visor.prop.rol": "Soy",
  "visor.prop.rol.docente": "Docente",
  "visor.prop.rol.familia": "Familia",
  "visor.prop.rol.hablante": "Hablante",
  "visor.prop.rol.otro": "Otra persona",
  "visor.prop.rostros": "Mejor sin rostros de personas, y nunca de niñas o niños. Nada que la comunidad reserve (ceremonias, lugares sagrados).",
  "visor.prop.significa": "Qué significa (en otra lengua, opcional)",
  "visor.prop.sinDatos": "Sin datos todavía.",
  "visor.prop.sinRed": "Sin conexión: el estado se verá cuando haya internet.",
  "visor.prop.tamano": "Lista para enviar ({kb} KB).",
  "visor.prop.tipo.correccion": "Corrección",
  "visor.prop.tipo.descripcion": "Descripción",
  "visor.prop.tipo.dibujo": "Dibujo",
  "visor.prop.tipo.foto": "Foto",
  "visor.prop.tipo.fotos": "Fotos de hojas",
  "visor.prop.tipo.ilustracion": "Ilustración",
  "visor.prop.tipo.imagen": "Imagen",
  "visor.prop.tipo.otro": "Otro",
  "visor.prop.tipo.palabra": "Palabra",
  "visor.prop.tipo.retiro": "Retiro",
  "visor.prop.tipo.variante": "En mi comunidad",
  "visor.prop.tipoImg": "¿Qué es?",
  "visor.prop.titulo": "Proponer al catálogo",
  "visor.prop.tuComunidad": "Tu comunidad o región",
  "visor.prop.vacio": "No hay nada que enviar todavía.",
  "visor.prop.var": "¿Así se dice en tu comunidad?",
  "visor.prop.var.como": "¿Cómo se dice ahí?",
  "visor.prop.var.corto": "En mi comunidad",
  "visor.prop.var.d": "No se trata de cuál es «la correcta»: así sabemos en qué comunidades se usa cada forma.",
  "visor.prop.var.igual": "Sí, se dice igual",
  "visor.prop.var.otra": "Se dice diferente",
  "visor.puntos": "Puntos",
  "visor.puntosN": "{n} puntos",
  "visor.qe.d": "Aparece una descripción escrita por la comunidad. Elige a qué palabra se refiere.",
  "visor.ql.d": "Aparece una imagen (con su audio si lo hay) y cuatro letras del alfabeto de la lengua.",
  "visor.rel.antonimo": "Antónimo",
  "visor.rel.deriva_de": "Deriva de",
  "visor.rel.deriva_de.inv": "Da origen a",
  "visor.rel.homonimo": "Misma forma, otro significado",
  "visor.rel.inversaDe": "{nombre} (desde la otra palabra)",
  "visor.rel.parte_de": "Parte de",
  "visor.rel.parte_de.inv": "Tiene como parte",
  "visor.rel.relacionada": "Relacionada",
  "visor.rel.sinonimo": "Sinónimo",
  "visor.rel.tipo_de": "Es un tipo de",
  "visor.rel.tipo_de.inv": "Incluye",
  "visor.rel.unida": "Unida a",
  "visor.rel.unida.inv": "Reúne a",
  "visor.rel.variante_de": "Variante de",
  "visor.rel.variante_de.inv": "Tiene como variante",
  "visor.relacionadas": "Relacionadas",
  "visor.rep.avance": "*Tu avance:* {vistas} palabras vistas · {aprendidas} bien aprendidas",
  "visor.rep.borrar": "¿Borrar tu avance en esta lengua?",
  "visor.rep.d": "Ves la imagen (o escuchas, o lees en otra lengua) y recuerdas la palabra. Las que te cuestan vuelven pronto; las que ya sabes, después de unos días. Tu avance se guarda solo en este dispositivo.",
  "visor.rep.deCero": "Empezar de cero",
  "visor.rep.empezar": "Empezar ({n} tarjetas)",
  "visor.rep.error": "No hay suficientes palabras con ese frente. Prueba otro.",
  "visor.rep.frente": "Al frente de la tarjeta",
  "visor.rep.lengua": "Lengua que quiero aprender",
  "visor.rep.no": "Todavía no",
  "visor.rep.si": "La sabía",
  "visor.rep.ver": "Ver la palabra",
  "visor.req.audio": "palabras con audio",
  "visor.req.cats": "categorías",
  "visor.req.desc": "palabras con descripción",
  "visor.req.disponible": "✓ Disponible ({tiene} {que})",
  "visor.req.img": "palabras con imagen",
  "visor.req.imgAudio": "palabras con imagen y audio",
  "visor.req.largas": "palabras de 3 letras o más",
  "visor.req.lenguas": "lenguas con 10 palabras o más",
  "visor.req.necesita": "Necesita {min} {que}; hay {tiene}",
  "visor.req.silabas": "palabras con sílabas registradas",
  "visor.req.textos": "palabras",
  "visor.req.toast": "{actividad}: {req} en {lengua}. Prueba otra lengua.",
  "visor.resultados": "Resultados",
  "visor.revisar": "Revisar",
  "visor.rondaN": "Ronda {n}/{total}",
  "visor.rondas": "Rondas",
  "visor.rondas.resultado": "{bien} de {total}",
  "visor.salirActividad": "¿Salir de la actividad?",
  "visor.saltar": "Saltar al contenido",
  "visor.sec.aprender": "Aprender",
  "visor.sec.colabora": "Colabora",
  "visor.sec.explorar": "Explorar",
  "visor.sec.imprimir": "Imprimir",
  "visor.sec.jugar": "Jugar",
  "visor.secciones": "Secciones",
  "visor.sel.audio": "Solo palabras con audio",
  "visor.sel.azar": "Al azar",
  "visor.sel.conocidas": "Solo con letras que ya conocen",
  "visor.sel.conocidas.d": "Marca las letras que el grupo ya trabajó: solo se usarán palabras escritas con ellas. Sin marcar, se usan todas.",
  "visor.sel.filtros": "Más filtros",
  "visor.sel.largo": "Largo de las palabras",
  "visor.sel.largo.12": "Hasta 12 letras",
  "visor.sel.largo.5": "Cortas (hasta 5 letras)",
  "visor.sel.largo.8": "Hasta 8 letras",
  "visor.sel.largo.todos": "Todas",
  "visor.sel.lista": "De mi lista",
  "visor.sel.titulo": "Palabras",
  "visor.si": "Sí",
  "visor.siguiente": "Siguiente",
  "visor.siguientePalabra": "Siguiente palabra",
  "visor.silabas": "Sílabas: {s}",
  "visor.sinCategorias": "Sin categorías.",
  "visor.sinConexion": "sin conexión",
  "visor.sinLimite": "Sin límite",
  "visor.sinNombre": "Sin nombre",
  "visor.sinRegistro": "sin registro en {lengua}",
  "visor.sinResultados": "Sin resultados.",
  "visor.sinResultados.comparte": "¿Conoces esta palabra? Compártela en *Colabora*.",
  "visor.sobreVisor": "Sobre el visor",
  "visor.sobreVisor.d": "El programa es libre y gratuito; los contenidos de este catálogo pertenecen a su comunidad.",
  "visor.sopa.chica": "Chica",
  "visor.sopa.completada": "¡Sopa completada!",
  "visor.sopa.d": "Desliza el dedo (o el ratón) sobre las letras. Vale en cualquier dirección.",
  "visor.sopa.dificil": "Difícil",
  "visor.sopa.dificil.d": "En diagonal y también al revés.",
  "visor.sopa.encontradas": "Encontradas: *{n}/{total}*",
  "visor.sopa.facil": "Fácil",
  "visor.sopa.facil.d": "Solo de izquierda a derecha y de arriba abajo.",
  "visor.sopa.grande": "Grande",
  "visor.sopa.mediana": "Mediana",
  "visor.sopa.nivel": "Dificultad",
  "visor.sopa.normal": "Normal",
  "visor.sopa.normal.d": "También en diagonal.",
  "visor.sopa.pocas": "Pocas palabras de este largo en esta lengua. Prueba un tamaño mayor u otra lengua.",
  "visor.sopa.tamano": "Tamaño",
  "visor.subtitulo": "{palabras} palabras · {lenguas} lenguas",
  "visor.sug.asunto": "Sugerencia sobre «{palabra}»",
  "visor.sug.d": "Escríbele a quien cuida este catálogo:",
  "visor.sug.texto": "Hola. Sobre la palabra «{palabra}» ({lengua}) del catálogo «{catalogo}» [ref. {id}]:",
  "visor.sug.titulo": "¿Sabes otra forma de decirlo o falta algo?",
  "visor.tecla": "tecla {t}",
  "visor.terminar": "Terminar",
  "visor.terminarJuego": "Terminar juego",
  "visor.todas": "Todas",
  "visor.todasCategorias": "Todas las categorías",
  "visor.tu": "Tú",
  "visor.turnoDe": "Turno de *{nombre}*",
  "visor.vacio": "(vacío)",
  "visor.verFuente": "ver",
  "visor.verPalabrasEn": "Ver palabras en",
  "visor.verTodas": "Ver en todas las lenguas",
  "visor.version": "Visor {v}",
  "visor.versionCatalogo": "catálogo {v}",
  "visor.versionPublicacion": "publicación {v}",
  "visor.volver": "Volver",
  "visor.voz": "Voz: {voz}",
  "visor.yMas": "y {n} más"
};
Object.assign(T_ES, T_VISOR);
