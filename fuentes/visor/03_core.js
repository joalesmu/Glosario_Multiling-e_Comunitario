<script>
/* =====================================================================
   VISOR DEL CATÁLOGO MULTILINGÜE · 1.1
   Ninguna lengua es obligatoria: cada palabra (concepto) se muestra en la
   lengua elegida y, si no existe, en la siguiente disponible.
   ===================================================================== */

/* Dirección del script del catálogo y su clave. Vacías = visor genérico: el catálogo se toma de
   ?api=…&token=… en el enlace, o se pide al abrir (🔌). Si se escriben, ese catálogo es el
   predeterminado de este visor, pero desde 🔌 se puede abrir otro y volver. */
const API_URL   = "";
const API_TOKEN = "";
const VERSION_VISOR = '1.2';
/* Textos de la interfaz en español (base de todas las traducciones). Se generan con el catálogo: no editar a mano. */
const V_ES = /*V_ES*/{};

/* =============================================
   UTILIDADES
   ============================================= */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const LS = {
    get(k, def = null) { try { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); } catch (e) { return def; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} }
};
function hash32(str) {
    let h = 1779033703 ^ str.length;
    for (let i = 0; i < str.length; i++) { h = Math.imul(h ^ str.charCodeAt(i), 3432918353); h = h << 13 | h >>> 19; }
    h = Math.imul(h ^ h >>> 16, 2246822507); h = Math.imul(h ^ h >>> 13, 3266489909);
    return (h ^= h >>> 16) >>> 0;
}
/** Azar reproducible: la misma semilla da el mismo resultado en cualquier dispositivo. */
function azar(semilla) {
    let a = hash32(String(semilla));
    return () => { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
}
function barajar(arr, r = Math.random) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(r() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
}
const elegir = (arr, r = Math.random) => arr[Math.floor(r() * arr.length)];
function azarSeguro(n) { const v = new Uint32Array(1); crypto.getRandomValues(v); return v[0] % n; }
function cuatroDigitos() { return String(azarSeguro(10000)).padStart(4, '0'); }
function fmtTiempo(s) { return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`; }
function descargar(nombre, contenido, tipo = 'text/html') {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([contenido], { type: tipo }));
    a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
function placeholder(letra) {
    const l = String(letra || '?').slice(0, 2);
    const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='300' height='300'><rect width='300' height='300' rx='16' fill='#94a3b8'/><text x='150' y='195' font-family='system-ui,sans-serif' font-size='140' font-weight='900' fill='white' text-anchor='middle'>${l.replace(/[<&>'"]/g, '')}</text></svg>`;
    // Sin comillas simples ni dobles: se puede usar con seguridad dentro de atributos HTML
    return 'data:image/svg+xml,' + encodeURIComponent(svg).replace(/'/g, '%27');
}

/* =============================================
   IDIOMA DE LA INTERFAZ
   La base es V_ES (español, al final de este bloque). Encima se ponen, en este orden:
   1) la traducción predeterminada del catálogo, 2) la que eligió esta persona.
   Las traducciones llegan publicadas por el catálogo o instaladas en este dispositivo con un
   archivo "catalogo-lenguas-interfaz" (el mismo que se crea en la hoja: 🗣️ Idioma de los menús).
   Lo que falte se ve en el idioma de respaldo. Nunca se interpreta HTML de una traducción.
   ============================================= */
const FORMATO_INTERFAZ = 'catalogo-lenguas-interfaz';
const I18n = {
    codigo: '', activo: 'es', dic: null, instaladas: {}, delCatalogo: {}, predeterminada: '',
    init() { this.instaladas = LS.get('interfaz_instaladas', {}) || {}; this.codigo = LS.get('interfaz_lengua', '') || ''; this.armar(); },
    textosDe(c) {
        if (!c || c === 'es') return null;
        const a = (this.delCatalogo[c] || {}).textos || {}, b = (this.instaladas[c] || {}).textos || {};
        return Object.assign({}, a, b);   // lo instalado en este dispositivo manda sobre lo publicado
    },
    total() { return Object.keys(V_ES).length; },
    disponibles() {
        const lista = [{ codigo: 'es', nombre: 'Español', n: this.total(), origen: 'base' }];
        const cods = new Set([...Object.keys(this.delCatalogo), ...Object.keys(this.instaladas)]);
        cods.delete('es');
        cods.forEach(c => {
            const t = this.textosDe(c) || {};
            lista.push({ codigo: c, nombre: (this.instaladas[c] || this.delCatalogo[c] || {}).nombre || c, n: Object.keys(V_ES).filter(k => t[k]).length,
                origen: this.instaladas[c] ? 'dispositivo' : 'catalogo' });
        });
        return lista;
    },
    armar() {
        const quiero = this.codigo || this.predeterminada || 'es';
        const d = Object.assign({}, V_ES);
        // Quien eligió español lo ve todo en español; si eligió otra lengua, lo que falte sale en la del catálogo y luego en español
        (quiero === 'es' ? [] : [this.predeterminada, quiero]).forEach(c => { const t = this.textosDe(c); if (t) Object.keys(V_ES).forEach(k => { if (t[k]) d[k] = String(t[k]); }); });
        this.dic = d;
        this.activo = this.disponibles().some(x => x.codigo === quiero) ? quiero : 'es';
        document.documentElement.lang = /^[a-z]{2,3}$/.test(this.activo) ? this.activo : 'es';
        this.aplicarFijos();
    },
    /** Textos fijos de la página: data-i="clave" y data-i-attr="atributo:clave|atributo:clave". */
    aplicarFijos() {
        document.querySelectorAll('[data-i]').forEach(el => { el.textContent = _i(el.dataset.i); });
        document.querySelectorAll('[data-i-attr]').forEach(el => el.dataset.iAttr.split('|').forEach(p => { const [a, k] = p.split(':'); if (a && k) el.setAttribute(a, _i(k)); }));
    },
    desdeCatalogo(interfaz) {
        this.delCatalogo = {};
        ((interfaz && interfaz.idiomas) || []).forEach(x => { if (x && x.codigo) this.delCatalogo[String(x.codigo)] = { nombre: String(x.nombre || x.codigo), textos: x.textos || {} }; });
        this.predeterminada = (interfaz && interfaz.predeterminado && this.delCatalogo[interfaz.predeterminado]) ? interfaz.predeterminado : '';
        this.armar();
    },
    elegir(c) { this.codigo = c === 'es' && !this.predeterminada ? '' : c; LS.set('interfaz_lengua', this.codigo); this.armar(); },
    /** Lee un archivo de traducción. aplicar=false solo dice qué haría. */
    instalar(texto, aplicar) {
        let p;
        try { p = JSON.parse(texto); } catch (e) { throw new Error(_i('visor.interfaz.error.archivo')); }
        if (!p || p.formato !== FORMATO_INTERFAZ || !p.textos || typeof p.textos !== 'object') throw new Error(_i('visor.interfaz.error.archivo'));
        const codigo = String(p.codigo || '').toLowerCase().replace(/[^a-z0-9-]/g, '');
        if (!/^[a-z]{2,8}(-[a-z0-9]{1,8})?$/.test(codigo) || codigo === 'es') throw new Error(_i('visor.interfaz.error.archivo'));
        const textos = {};
        Object.keys(p.textos).forEach(k => { if (k in V_ES && String(p.textos[k] || '').trim()) textos[k] = String(p.textos[k]).slice(0, 1500); });
        const rep = { codigo, nombre: String(p.nombre || codigo).slice(0, 60), n: Object.keys(textos).length, total: this.total() };
        if (!rep.n) throw new Error(_i('visor.interfaz.error.vacio'));
        if (aplicar) {
            this.instaladas[codigo] = { nombre: rep.nombre, textos, fecha: new Date().toISOString().slice(0, 10) };
            if (!LS.set('interfaz_instaladas', this.instaladas)) throw new Error(_i('visor.interfaz.error.espacio'));
            this.elegir(codigo);
        }
        return rep;
    },
    quitar(c) { delete this.instaladas[c]; LS.set('interfaz_instaladas', this.instaladas); if (this.codigo === c) this.codigo = ''; LS.set('interfaz_lengua', this.codigo); this.armar(); }
};
const _rellenarI = (s, v, e) => !v ? s : s.replace(/\{(\w+)\}/g, (m, x) => (x in v) ? (e ? esc(v[x]) : String(v[x] ?? '')) : m);
/** Texto plano de la interfaz (sin marcas). */
function _i(k, v) { const s = (I18n.dic || V_ES)[k]; return _rellenarI(s == null ? k : String(s), v, false).replace(/\*([^*\n]+)\*/g, '$1'); }
/** Texto de la interfaz para HTML: escapado, *negritas* y saltos de línea. */
function _ih(k, v) { const s = (I18n.dic || V_ES)[k]; return _rellenarI(esc(s == null ? k : String(s)), v, true).replace(/\*([^*\n]+)\*/g, '<b>$1</b>').replace(/\n/g, '<br>'); }

/** Ventana 🌐: elegir el idioma de los menús e instalar traducciones en este dispositivo. */
const Interfaz = {
    abrir() {
        const l = I18n.disponibles();
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🌐 ${_ih('visor.interfaz.titulo')}</h2>
            <p class="nota">${_ih('visor.interfaz.intro')}</p>
            <div class="chips" style="flex-direction:column;align-items:stretch" id="if-lista">${l.map(x => `<label class="chip"><input type="radio" name="if-l" value="${esc(x.codigo)}" ${x.codigo === I18n.activo ? 'checked' : ''}>
                <b>${esc(x.nombre)}</b> <span class="nota">${_ih('visor.interfaz.avance', { n: x.n, total: I18n.total() })}${x.origen === 'dispositivo' ? ' · ' + _ih('visor.interfaz.enDispositivo') : x.origen === 'catalogo' ? ' · ' + _ih('visor.interfaz.delCatalogo') : ''}</span>
                ${x.origen === 'dispositivo' ? `<button class="btn btn-secondary" style="min-height:30px;padding:2px 10px;margin-left:auto" data-quitar="${esc(x.codigo)}">${_ih('visor.interfaz.quitar')}</button>` : ''}</label>`).join('')}</div>
            <div class="acciones"><button class="btn btn-primary" id="if-usar">${_ih('visor.interfaz.usar')}</button></div>
            <details style="margin-top:10px"><summary class="t" style="cursor:pointer">📥 ${_ih('visor.interfaz.instalar')}</summary>
                <p class="nota">${_ih('visor.interfaz.instalar.d')}</p>
                <input type="file" id="if-arch" accept=".json,application/json" aria-label="${esc(_i('visor.interfaz.instalar'))}">
                <div id="if-msg" role="status" aria-live="polite"></div></details></div>`);
        $('if-usar').addEventListener('click', () => { const v = document.querySelector('input[name="if-l"]:checked'); if (v) { I18n.elegir(v.value); UI.cerrarModal(); UI.refrescar(); Datos.cabecera(); } });
        $('if-lista').addEventListener('click', e => { const b = e.target.closest('[data-quitar]'); if (!b) return; e.preventDefault(); I18n.quitar(b.dataset.quitar); this.abrir(); UI.refrescar(); Datos.cabecera(); });
        $('if-arch').addEventListener('change', e => {
            const f = e.target.files[0]; if (!f) return;
            const r = new FileReader();
            r.onload = () => {
                try {
                    const rep = I18n.instalar(r.result, false);
                    $('if-msg').innerHTML = `<div class="caja">${_ih('visor.interfaz.previa', rep)}<div class="acciones"><button class="btn btn-primary" id="if-inst">${_ih('visor.interfaz.instalarBoton')}</button></div></div>`;
                    $('if-inst').addEventListener('click', () => { try { I18n.instalar(r.result, true); UI.cerrarModal(); UI.refrescar(); Datos.cabecera(); UI.toast(_i('visor.interfaz.instalada', rep)); } catch (x) { $('if-msg').innerHTML = `<div class="error-box">${esc(x.message)}</div>`; } });
                } catch (x) { $('if-msg').innerHTML = `<div class="error-box">${esc(x.message)}</div>`; }
            };
            r.readAsText(f);
        });
    }
};

/* =============================================
   ACCESIBILIDAD
   ============================================= */
const A11y = {
    escalas: [0.85, 1, 1.15, 1.35, 1.6], idx: 1,
    init() {
        this.idx = LS.get('a11y_escala', 1);
        this.aplicar();
        if (LS.get('a11y_contraste', false)) { document.body.classList.add('high-contrast'); $('btn-contrast').classList.add('active'); }
    },
    aplicar() { document.documentElement.style.setProperty('--fs', this.escalas[this.idx] + 'rem'); LS.set('a11y_escala', this.idx); UI.ajustarCabecera(); },
    smaller() { if (this.idx > 0) { this.idx--; this.aplicar(); } },
    larger() { if (this.idx < this.escalas.length - 1) { this.idx++; this.aplicar(); } },
    reset() { this.idx = 1; this.aplicar(); },
    toggleContrast() {
        const on = document.body.classList.toggle('high-contrast');
        LS.set('a11y_contraste', on); $('btn-contrast').classList.toggle('active', on);
    }
};

/* =============================================
   TEXTO: escritura, alfabetos y comparación
   ============================================= */
const Texto = {
    _tokens: new Map(), _alfa: new Map(),
    _seg: (typeof Intl !== 'undefined' && Intl.Segmenter) ? new Intl.Segmenter('und', { granularity: 'grapheme' }) : null,
    /** Atajos de escritura de la lengua (p. ej. + → ɨ), los mismos que usa el equipo al registrar. */
    atajos(s, L) {
        let t = String(s ?? '');
        const lengua = Datos.langById.get(String(L));
        (lengua?.atajos || []).forEach(([b, r, mayus]) => {
            if (!b) return;
            t = mayus ? t.split(b).join(r) : t.replace(new RegExp(b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'giu'), () => r);
        });
        return t.normalize('NFC');
    },
    /** Comparación tolerante: minúsculas, sin acentos, apóstrofos unificados. */
    pliegue(s) {
        return String(s ?? '').normalize('NFC').toLowerCase().normalize('NFD').replace(/\p{M}/gu, '')
            .replace(/[’'ʼꞌ´`]/g, "'").replace(/\s+/g, ' ').trim();
    },
    /** Letras del alfabeto declarado de la lengua (con dígrafos), ordenadas de la más larga a la más corta. */
    tokens(L) {
        if (this._tokens.has(L)) return this._tokens.get(L);
        const a = (Datos.langById.get(String(L))?.alfabeto || '').normalize('NFC').toLowerCase().split(/\s+/).filter(Boolean);
        const t = [...new Set(a)].sort((x, y) => y.length - x.length);
        this._tokens.set(L, t);
        return t;
    },
    /** Divide una palabra en sus letras reales: "chɨkwa" con alfabeto [ch, kw…] → ch·ɨ·kw·a. */
    grafemas(s, L) {
        const t = String(s ?? '').normalize('NFC').toLowerCase();
        const toks = L ? this.tokens(L) : [];
        const out = [];
        let i = 0;
        while (i < t.length) {
            const m = toks.find(k => t.startsWith(k, i));
            if (m) { out.push({ g: m, letra: true }); i += m.length; continue; }
            let g;
            if (this._seg) { const r = this._seg.segment(t.slice(i, i + 8))[Symbol.iterator]().next().value; g = r ? r.segment : t[i]; }
            else { g = t[i]; let j = i + 1; while (j < t.length && /\p{M}/u.test(t[j])) g += t[j++]; }
            out.push({ g, letra: /[\p{L}Ꞌꞌʼ]/u.test(g) });
            i += g.length;
        }
        return out;
    },
    letras(s, L) { return this.grafemas(s, L).filter(x => x.letra).map(x => x.g); },
    /** Alfabeto para teclados y juegos: el declarado o, si no hay, el que aparece en las palabras. */
    alfabeto(L) {
        if (this._alfa.has(L)) return this._alfa.get(L);
        let a = this.tokens(L).slice();
        if (a.length) {
            const orden = (Datos.langById.get(String(L))?.alfabeto || '').normalize('NFC').toLowerCase().split(/\s+/).filter(Boolean);
            a = [...new Set(orden)];
        } else {
            const vistos = new Set();
            Datos.words.forEach(w => { const t = w.t[L]; if (t) this.letras(t.texto).forEach(g => vistos.add(g)); });
            a = [...vistos].sort((x, y) => x.localeCompare(y, 'es'));
        }
        this._alfa.set(L, a);
        return a;
    },
    /** Atributo lang para lectores de pantalla y tipografía (solo con código ISO válido). */
    lang(L) { const iso = Datos.langById.get(String(L))?.iso; return /^[a-z]{2,3}$/.test(iso || '') ? ` lang="${iso}"` : ''; },
    limpiarCaches() { this._tokens.clear(); this._alfa.clear(); }
};

/* =============================================
   DATOS DEL CATÁLOGO (con copia para usar sin conexión)
   ============================================= */
/** Catálogos abiertos en este dispositivo: dirección, su clave y su título (para volver a ellos). */
const Catalogos = {
    limpiar(u) { return String(u || '').trim().split('?')[0].split('#')[0]; },
    /** Lee lo que la persona pega: la dirección del catálogo o un enlace de invitación del visor (…?api=…&token=…). */
    desdeEnlace(texto) {
        const s = String(texto || '').trim(), m = s.match(/https?:\/\/\S+/); if (!m) return { url: '', token: '' };
        let u; try { u = new URL(m[0]); } catch (e) { return { url: '', token: '' }; }
        const api = u.searchParams.get('api');
        if (api) return { url: this.limpiar(api), token: (u.searchParams.get('token') || '').trim() };
        return { url: this.limpiar(m[0]), token: '' };
    },
    lista() { const l = LS.get('glosario_catalogos', []); return Array.isArray(l) ? l : []; },
    clave(url) {
        const c = this.lista().find(x => x.url === url);
        if (c) return c.token || '';
        if (url && url === this.limpiar(API_URL)) return API_TOKEN;
        if (url && url === LS.get('glosario_url', '') && LS.get('glosario_token', '')) return LS.get('glosario_token', '');   // versiones anteriores del visor
        return '';
    },
    recordar(url, token, titulo) {
        if (!url) return;
        const l = this.lista().filter(x => x.url !== url), previo = this.lista().find(x => x.url === url) || {};
        l.unshift({ url, token: token || '', titulo: titulo || previo.titulo || '' });
        LS.set('glosario_catalogos', l.slice(0, 12)); LS.del('glosario_token');
    },
    olvidar(url) { LS.set('glosario_catalogos', this.lista().filter(x => x.url !== url)); LS.del('cat:' + url); Copia.borrar(url); if (LS.get('glosario_url', '') === url) LS.del('glosario_url'); }
};

/* =============================================
   ESTADO DEL VISOR: avisos, mantenimiento, identidad y revisión periódica
   Llega en la consulta pequeña «manifiesto»; se recuerda en el dispositivo para verse sin conexión.
   ============================================= */
const Estado = {
    rol: 'publico', cerrado: false, e: null, mant: null, timer: null, ultima: 0, fallos: 0, idx: 0, pendiente: false, purgado: false,
    ICONOS: { info: 'ℹ️', novedad: '✨', pausa: '⏸️', importante: '⚠️' }, PRIORIDAD: { importante: 0, pausa: 1, novedad: 2, info: 3 },
    llave() { return 'estado:' + Datos.base; },
    recordado() { return Datos.base ? LS.get(this.llave(), null) : null; },
    /** m: respuesta del catálogo (manifiesto o mantenimiento) o lo recordado (offline = true). */
    aplicar(m, offline) {
        if (!m) return;
        const est = m.estado || {}, mant = m.mantenimiento || est.mant || null;
        if (m.rol || est.rol) this.rol = m.rol || est.rol;
        if (!offline) {
            // Borrado precautorio: si el catálogo lo pidió, la copia de este dispositivo se va
            const p = String(est.purga || ''), k = 'purga:' + Datos.base;
            if (p && LS.get(k, '') !== p) { if (LS.get('cat:' + Datos.base) || Copia.hay(Datos.base)) this.purgado = true; LS.del('cat:' + Datos.base); Copia.borrar(Datos.base); LS.set(k, p); }
        }
        this.e = est; this.mant = mant;
        this.cerrado = !!(mant && mant.modo === 'cerrar' && this.rol !== 'equipo');
        if (!offline) LS.set(this.llave(), { estado: est, mantenimiento: mant, rol: this.rol });
        this.pintarIdentidad();
        this.pintarAvisos();
    },
    /** Texto escrito por el equipo: escapado, con *negritas* y saltos de línea. */
    texto(s) { return esc(String(s || '')).replace(/\*([^*\n]+)\*/g, '<b>$1</b>').replace(/\n/g, '<br>'); },
    hoy() { const d = new Date(), z = n => String(n).padStart(2, '0'); return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate()); },
    avisos() {
        const e = this.e || {}, hoy = this.hoy(), cerrados = LS.get('avisos_cerrados:' + Datos.base, []) || [];
        const l = (e.avisos || []).filter(a => (!a.desde || a.desde <= hoy) && (!a.hasta || a.hasta >= hoy) && (!a.cerrable || a.tipo === 'importante' || !cerrados.includes(a.id)));
        const m = this.mant;
        if (m && this.rol === 'equipo') l.unshift({ id: '_equipo', tipo: 'importante', texto: _i('visor.mant.equipo', { modo: _i('visor.mant.m.' + m.modo) }), cerrable: false });
        else if (m && m.modo === 'congelar' && m.avisoPublico !== false && !cerrados.includes('_mant')) l.unshift({ id: '_mant', tipo: 'pausa', texto: m.mensaje || _i('visor.mant.congelado'), cerrable: true, propio: !m.mensaje });
        return l.sort((a, b) => (this.PRIORIDAD[a.tipo] ?? 3) - (this.PRIORIDAD[b.tipo] ?? 3));
    },
    pintarAvisos() {
        const c = $('avisos'); if (!c) return;
        const l = this.cerrado ? [] : this.avisos(), n = Math.min(3, Math.max(1, Number((this.e || {}).avisosMax) || 1));
        if (!l.length) { c.innerHTML = ''; c.hidden = true; return; }
        if (this.idx >= l.length) this.idx = 0;
        const ver = l.slice(this.idx, this.idx + n);
        c.hidden = false;
        c.innerHTML = ver.map(a => `<div class="franja t-${esc(a.tipo)}"><span class="ic" role="img" aria-label="${esc(_i('visor.aviso.tipo.' + a.tipo))}">${this.ICONOS[a.tipo] || 'ℹ️'}</span>
            <div class="tx">${a.propio ? _ih('visor.mant.congelado') : this.texto(a.texto)}</div>
            ${a.boton && a.enlace ? `<button class="btn btn-secondary" data-aviso-ir="${esc(a.id)}">${esc(a.boton)}</button>` : ''}
            ${a.cerrable && a.tipo !== 'importante' ? `<button class="cerrar" data-aviso-cerrar="${esc(a.id)}" aria-label="${esc(_i('visor.avisos.cerrar'))}" title="${esc(_i('visor.avisos.cerrar'))}">✕</button>` : ''}</div>`).join('') +
            (l.length > n ? `<div class="nav-avisos"><button data-aviso-mover="-1" aria-label="${esc(_i('visor.avisos.anterior'))}">◀</button><span>${_ih('visor.avisos.n', { n: Math.floor(this.idx / n) + 1, total: Math.ceil(l.length / n) })}</span><button data-aviso-mover="1" aria-label="${esc(_i('visor.avisos.siguiente'))}">▶</button></div>` : '');
        this._lista = l; this._n = n;
    },
    clic(e) {
        const b = e.target.closest('button'); if (!b) return;
        if (b.dataset.avisoCerrar) {
            const k = 'avisos_cerrados:' + Datos.base, l = LS.get(k, []) || [];
            if (!l.includes(b.dataset.avisoCerrar)) l.push(b.dataset.avisoCerrar);
            LS.set(k, l.slice(-60)); this.pintarAvisos();
        }
        if (b.dataset.avisoMover) { const pag = Math.ceil(this._lista.length / this._n), p = Math.floor(this.idx / this._n); this.idx = ((p + Number(b.dataset.avisoMover) + pag) % pag) * this._n; this.pintarAvisos(); }
        if (b.dataset.avisoIr) {
            const a = (this._lista || []).find(x => x.id === b.dataset.avisoIr); if (!a) return;
            const en = String(a.enlace || '');
            if (/^https:\/\//i.test(en)) window.open(en, '_blank', 'noopener');
            else if (en.startsWith('buscar:')) { Explorar.q = en.slice(7); Explorar.cat = 'all'; UI.cerrarModal(); UI.ir('explorar'); }
            else if (en.startsWith('categoria:')) { Explorar.q = ''; Explorar.cat = en.slice(10); UI.cerrarModal(); UI.ir('explorar'); }
        }
    },
    pintarIdentidad() {
        const id = (this.e || {}).identidad || {}, img = $('app-logo'); if (!img) return;
        if (id.imagen && id.enTitulo) { img.src = id.imagen; img.alt = id.alt || ''; img.classList.remove('oculto'); }
        else img.classList.add('oculto');
    },
    /** Pantalla de mantenimiento («cerrado»): imagen, mensaje, regreso y contacto si se eligió. */
    pantalla() {
        const m = this.mant || {}, e = this.e || {}, c = m.contacto || {};
        if (!Datos.cargado) {
            $('app-title').textContent = e.titulo || _i('visor.glosario');
            $('app-sub').textContent = e.subtitulo || '';
            $('sel-lengua').innerHTML = '';
        }
        const tel = String(c.whatsapp || '').replace(/[^\d]/g, '');
        const contacto = (c.correo || c.whatsapp || c.telefono) ? `<div class="caja"><b>${_ih('visor.mant.contacto')}</b>
            ${c.nombre ? `<p class="nota">${_ih('visor.col.atiende', { x: c.nombre })}</p>` : ''}
            <div class="fila" style="justify-content:center">
            ${tel ? `<a class="btn btn-primary" href="https://wa.me/${esc(tel)}" target="_blank" rel="noopener">💬 WhatsApp</a>` : ''}
            ${c.correo ? `<a class="btn btn-secondary" href="mailto:${esc(c.correo)}">✉️ ${_ih('visor.correo')}</a>` : ''}
            ${c.telefono ? `<a class="btn btn-secondary" href="tel:${esc(String(c.telefono).replace(/[^\d+]/g, ''))}">📞 ${_ih('visor.col.llamar')}</a>` : ''}</div>
            ${c.horario ? `<p class="nota">${_ih('visor.col.horario', { x: c.horario })}</p>` : ''}</div>` : '';
        $('vista').innerHTML = `<section class="mant" aria-labelledby="mant-t">
            ${m.imagen ? `<img class="mant-img" src="${esc(m.imagen)}" alt="${esc(m.alt || '')}" onerror="this.remove()">` : ''}
            <h2 id="mant-t">🚧 ${_ih('visor.mant.titulo')}</h2>
            <p class="mant-msg">${m.mensaje ? this.texto(m.mensaje) : _ih('visor.mant.d')}</p>
            ${m.hasta ? `<p><b>${_ih('visor.mant.hasta', { fecha: m.hasta })}</b></p>` : ''}
            ${contacto}
            ${this.purgado ? `<p class="nota">🧽 ${_ih('visor.mant.borrado')}</p>` : ''}
        </section>`;
        this.pintarAvisos();
    },
    /** Revisión periódica: solo con la pestaña visible y conexión; con margen al azar y espera creciente si falla. */
    minutos() { const r = (this.e || {}).revision; return r == null ? 10 : Number(r) || 0; },
    programar() {
        clearTimeout(this.timer);
        const min = this.minutos(); if (!min || !Datos.base) return;
        const espera = Math.min(3600000, min * 60000 * Math.min(6, Math.pow(2, this.fallos)) * (0.85 + Math.random() * 0.3));
        this.timer = setTimeout(() => this.revisar(), espera);
    },
    async revisar() {
        if (document.hidden || !navigator.onLine) { this.pendiente = true; return; }
        if ($('modal').style.display === 'flex' || Juego.activo) { clearTimeout(this.timer); this.timer = setTimeout(() => this.revisar(), 60000); return; }   // no interrumpir una actividad
        this.ultima = Date.now(); this.pendiente = false;
        const ok = await Datos.actualizar(true);
        this.fallos = ok ? 0 : Math.min(this.fallos + 1, 4);
        this.programar();
    },
    alVolver() {
        const min = this.minutos();
        if (!document.hidden && min && Datos.base && (this.pendiente || Date.now() - this.ultima > min * 60000)) this.revisar();
    }
};

const Datos = {
    base: '', token: '', cfg: {}, langs: [], langById: new Map(), cats: new Map(), words: [], byId: new Map(),
    lengua: '', principal: '', version: '', cargado: false, _conteos: new Map(),

    async iniciar() {
        const p = new URLSearchParams(location.search);
        // Orden: ?api=… del enlace → el que se eligió en 🔌 en este dispositivo → el escrito en este archivo
        this.base = Catalogos.limpiar(p.get('api') || LS.get('glosario_url', '') || API_URL);
        // Cada catálogo guarda su propia clave: nunca se manda la clave de uno a otro
        this.token = p.has('token') ? p.get('token').trim() : Catalogos.clave(this.base);
        if (p.get('api') || p.has('token')) {
            Catalogos.recordar(this.base, this.token);
            if (p.get('api')) LS.set('glosario_url', this.base);   // al recargar sigue en este catálogo
            const q = new URLSearchParams(location.search); q.delete('api'); q.delete('token');   // la clave no se queda a la vista; ?loteria=… sí
            history.replaceState(null, '', location.pathname + (q.toString() ? '?' + q : '') + location.hash);
        }
        if (!this.base) return UI.pantallaConexion();

        // Lo recordado primero (avisos, mantenimiento): si el catálogo estaba cerrado, no se muestra la copia
        const rec = Estado.recordado();
        if (rec) Estado.aplicar({ estado: rec.estado, mantenimiento: rec.mantenimiento, rol: rec.rol }, true);
        const copia = Estado.cerrado ? null : await Copia.leer(this.base);
        if (Estado.cerrado) Estado.pantalla();
        else if (copia && copia.data) { this.cargar(copia.data, copia.version); UI.refrescar(); UI.estadoConexion(_i(navigator.onLine ? 'visor.con.actualizando' : 'visor.con.guardada'), !navigator.onLine); }
        await this.actualizar(false);
        Estado.ultima = Date.now();
        Estado.programar();
    },

    /** Pregunta si hay versión nueva (consulta pequeña) y solo entonces descarga el catálogo. Devuelve si pudo consultar. */
    async actualizar(silencioso) {
        try {
            const m = await this.api({ action: 'manifiesto' }, 8000).catch(() => null);
            if (m && m.status === 401) { Copia.borrar(this.base); this.cargado = false; throw new Error(_i('visor.con.clave')); }   // sin la clave correcta no se muestra ni lo guardado
            if (m && (m.status === 200 || m.status === 503)) Estado.aplicar(m);
            if (Estado.cerrado) { this.cargado = false; Estado.pantalla(); UI.estadoConexion(_i('visor.mant.estado'), true, '🚧'); return !!m; }
            const copia = await Copia.leer(this.base), version = (m && m.version) || '';
            if (m && version && copia && copia.version === version) {
                if (!this.cargado) { this.cargar(copia.data, copia.version); UI.refrescar(); }   // p. ej. al volver de la pantalla de mantenimiento
                UI.estadoConexion(_i('visor.con.aldia'));
                return true;
            }
            if (!m && this.cargado) { UI.estadoConexion(_i('visor.con.guardada'), true); return false; }
            const r = await this.api({ action: 'glosario' }, 20000);
            if (r.status === 503) { Estado.aplicar(r); this.cargado = false; Estado.pantalla(); return true; }
            if (r.status === 401) { Copia.borrar(this.base); this.cargado = false; throw new Error(_i('visor.con.clave')); }
            if (r.status !== 200 || !r.data) throw new Error(r.error || _i('visor.con.inesperada'));
            const eraCargado = this.cargado, antes = this.version, y = window.scrollY;
            this.cargar(r.data, r.version || version);
            Catalogos.recordar(this.base, this.token, this.cfg.titulo);
            if (!(await Copia.guardar(this.base, { version: r.version || version, data: r.data, t: Date.now() }))) console.warn('Sin espacio para guardar la copia sin conexión.');
            UI.estadoConexion(_i('visor.con.aldia'));
            UI.refrescar();
            if (silencioso && eraCargado && antes !== this.version) { window.scrollTo(0, y); UI.toast(_i('visor.con.nuevaVersion')); }
            return true;
        } catch (e) {
            if (this.cargado) { UI.estadoConexion(_i('visor.con.guardada'), true); return false; }
            if (Estado.cerrado) return false;
            UI.pantallaConexion(e.message, true);
            return false;
        }
    },

    async api(params, espera = 15000) {
        const u = new URL(this.base);
        Object.entries(params).forEach(([k, v]) => u.searchParams.set(k, v));
        if (this.token) u.searchParams.set('token', this.token);
        const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), espera);
        try { const r = await fetch(u.toString(), { signal: ctl.signal }); return await r.json(); }
        finally { clearTimeout(t); }
    },
    async post(datos) {
        const r = await fetch(this.base, { method: 'POST', body: JSON.stringify(Object.assign({ token: this.token }, datos)) });
        return r.json();
    },

    cargar(d, version) {
        this.version = version || '';
        this.cfg = d.config || {};
        I18n.desdeCatalogo(d.interfaz);
        this.fuentes = new Map((d.fuentes || []).map(f => [String(f.id_fuente), f]));
        this.gobernanza = d.gobernanza || null;
        this.aportes = Array.isArray(d.aportes) ? d.aportes : [];
        this.langs = (d.idiomas || []).map(i => ({
            id: String(i.id_idioma), nombre: String(i.nombre_propio || i.nombre_completo || i.nombre || i.id_idioma),
            otro: String(i.nombre_completo || i.nombre || ''), iso: String(i.codigo_iso || '').toLowerCase(),
            variante: String(i.variante || ''), alfabeto: String(i.alfabeto || ''), atajos: Array.isArray(i.atajos) ? i.atajos : []
        }));
        this.langById = new Map(this.langs.map(l => [l.id, l]));
        this.cats = new Map();
        (d.categorias || []).forEach(c => {
            const id = String(c.id_categoria || '').trim();
            if (id) this.cats.set(id, { id, nombre: String(c.nombre_categoria || c.nombre || _i('visor.sinNombre')), trads: c.traducciones || {} });
        });
        const map = new Map();
        (d.palabras || []).forEach(p => {
            const id = String(p.id_palabra || '').trim();
            if (!id) return;
            map.set(id, { id, cat: String(p.id_categoria || '').trim(), img: String(p.imagen_referencia || ''), imgId: String(p.imagen_id || ''), credito: String(p.imagen_credito || ''),
                alt: String(p.imagen_alt || ''), galeria: Array.isArray(p.galeria) ? p.galeria.filter(g => g && /^https:\/\//.test(g.url)) : [], imgs: Number(p.imgs) || 0,
                imagenesMax: p.imagenes_max !== undefined && String(p.imagenes_max) !== '' ? Number(p.imagenes_max) : null,
                juegos: p.juegos !== false && String(p.usar_en_juegos || '').toLowerCase() !== 'no', t: {} });
        });
        (d.traducciones || []).forEach(t => {
            const w = map.get(String(t.id_palabra || '').trim());
            const texto = String(t.texto ?? '').trim();
            if (!w || !texto) return;
            // Varias formas ("el rancho, la casa"): la primera es la principal (juegos y tarjetas); todas valen al buscar y en el dictado
            const formas = Array.isArray(t.formas) && t.formas.length ? t.formas.map(String) : [texto];
            w.t[String(t.id_idioma)] = { texto: formas[0], completo: texto, formas, audio: String(t.audio || ''), nota: String(t.nota_variante || t.variante || ''), silabas: String(t.silabas || ''),
                voz: String(t.hablante || ''), lugar: String(t.lugar_hablante || ''), didactico: t.didactico !== false,
                fuente: String(t.id_fuente || ''), ubicacion: String(t.ubicacion_fuente || ''), tipo: String(t.tipo_registro || ''), desc: String(t.descripcion || '').trim() };
        });
        this.words = [...map.values()].filter(w => Object.keys(w.t).length);
        this.byId = new Map(this.words.map(w => [w.id, w]));
        // Relaciones: cada vínculo se guarda una vez; desde la otra palabra se lee con el nombre inverso
        this.tiposRel = new Map((d.tiposRelacion || []).map(t => [String(t.id), t]));
        this.rel = new Map();
        (d.relaciones || []).forEach(r => {
            const a = String(r.id_origen || ''), b = String(r.id_destino || ''), tipo = String(r.tipo || '');
            if (!a || !b || a === b || !this.byId.has(a) || !this.byId.has(b)) return;
            const pon = (x, otra, inv) => { if (!this.rel.has(x)) this.rel.set(x, []); this.rel.get(x).push({ otra, tipo, inv, nota: String(r.nota || '') }); };
            pon(a, b, false); pon(b, a, true);
        });
        // Lengua de referencia de la interfaz: la que eligió la persona, la del catálogo, o la primera con palabras
        const conPalabras = this.langs.filter(l => this.words.some(w => w.t[l.id]));
        this.principal = this.langById.has(String(this.cfg.idiomaPrincipal)) ? String(this.cfg.idiomaPrincipal) : (conPalabras[0]?.id || '');
        const guardada = LS.get('lengua:' + this.base);
        this.lengua = this.langById.has(guardada) ? guardada : this.principal;
        this._conteos.clear(); Texto.limpiarCaches();
        this.cargado = true;
        this.cabecera();
    },
    /** Título, subtítulo y selector de lengua (se vuelve a pintar al cambiar el idioma de los menús). */
    cabecera() {
        if (!this.cargado) return;
        const conPalabras = this.langs.filter(l => this.words.some(w => w.t[l.id]));
        document.title = this.cfg.titulo || _i('visor.glosario');
        $('app-title').textContent = this.cfg.titulo || _i('visor.glosario');
        $('app-sub').textContent = this.cfg.subtitulo || _i('visor.subtitulo', { palabras: this.words.length, lenguas: conPalabras.length });
        const sel = $('sel-lengua');
        sel.innerHTML = conPalabras.map(l => `<option value="${esc(l.id)}">${esc(l.nombre)}${l.otro && l.otro !== l.nombre ? ' · ' + esc(l.otro) : ''}</option>`).join('');
        sel.value = this.lengua;
        UI.ajustarCabecera();
        UI.pie();
    },

    fijarLengua(L) { this.lengua = L; LS.set('lengua:' + this.base, L); UI.refrescar(); },
    cambiarCatalogo() { UI.pantallaConexion(); },

    /** Nombre de un tipo de relación: el que puso el catálogo o el de fábrica en el idioma de la interfaz. */
    etiquetaRel(tipo, inv) {
        const t = this.tiposRel.get(tipo) || { id: tipo, nombre: tipo };
        if (inv && !t.simetrica) return t.inverso || (t.clave ? _i('visor.rel.' + t.clave + '.inv') : _i('visor.rel.inversaDe', { nombre: t.nombre || t.id }));
        return t.nombre || (t.clave ? _i('visor.rel.' + t.clave) : t.id);
    },
    nombreLengua(L) { return this.langById.get(String(L))?.nombre || String(L); },
    /** Texto de una palabra en la lengua pedida; si no existe, en la de referencia; si no, el primero. */
    texto(w, L = this.lengua) { if (!w) return ''; const t = w.t[L] || w.t[this.principal] || Object.values(w.t)[0]; return t ? t.texto : ''; },
    tieneEn(w, L) { return !!(w && w.t[L]); },
    categoria(id, L = this.lengua) {
        const c = this.cats.get(String(id || '').trim());
        if (!c) return _i('visor.general');
        return c.trads[L] || c.trads[this.principal] || c.nombre;
    },
    /** Imagen de una palabra: la guardada en el dispositivo (📥 Llevar sin conexión) o la de internet. `red` = siempre la de internet. */
    img(w, ancho = 400, red = false) {
        const u = w?.img || '';
        if (!u) return '';
        if (!red && typeof Offline !== 'undefined') { const g = Offline.img(w.id); if (g) return g; }
        if (/^https:\/\/lh3\.googleusercontent\.com\/d\/[^=?#]+$/.test(u)) return `${u}=w${ancho}`;
        return u;
    },
    /** Palabras disponibles para juegos en una lengua, con condiciones opcionales. */
    paraJuegos(L, { img = false, audio = false, cats = null, largas = false } = {}) {
        return this.words.filter(w => {
            if (!w.juegos || !w.t[L] || !w.t[L].didactico) return false;   // su fuente puede no permitir juegos
            if (img && !w.img) return false;
            if (audio && !w.t[L].audio) return false;
            if (cats && cats.length && !cats.includes(w.cat)) return false;
            if (largas && Texto.letras(w.t[L].texto, L).length < 3) return false;
            return true;
        });
    },
    conteo(L) {
        if (this._conteos.has(L)) return this._conteos.get(L);
        const c = { img: 0, imgAudio: 0, audio: 0, largas: 0, cats: 0, textos: 0, lenguas: 0, desc: 0, silabas: 0 };
        const cats = new Set();
        this.words.forEach(w => {
            const t = w.t[L];
            if (!t) return;
            c.textos++;
            if (!w.juegos || !t.didactico) return;
            if (t.desc) c.desc++;
            if (t.silabas && t.silabas.split(/[-·.\s|/]+/).filter(Boolean).length >= 2) c.silabas++;
            if (w.img) c.img++;
            if (t.audio) c.audio++;
            if (w.img && t.audio) c.imgAudio++;
            if (Texto.letras(t.texto, L).length >= 3) c.largas++;
            if (w.cat && this.cats.has(w.cat)) cats.add(w.cat);
        });
        c.cats = cats.size;
        c.lenguas = this.langs.filter(l => this.words.filter(w => w.t[l.id]).length >= 10).length;
        this._conteos.set(L, c);
        return c;
    },
    palabrasDe(ids) { return ids.map(id => this.byId.get(id)).filter(Boolean); }
};

/* =============================================
   SONIDO (audios de hablantes) y efectos
   ============================================= */
const Sonido = {
    cache: new Map(), actual: null,
    async tocar(archivo, boton) {
        if (!archivo) return;
        try {
            if (boton) boton.classList.add('sonando');
            let datos = this.cache.get(archivo);
            if (!datos && typeof Offline !== 'undefined' && Offline.tieneAudio(archivo)) datos = await Offline.audio(archivo);
            if (!datos) {
                const r = await Datos.api({ type: 'audio', file: archivo }, 20000);
                if (r.status !== 'ok' || !r.audioData) throw new Error(r.error || 'audio');
                datos = r.audioData;
                if (this.cache.size > 80) this.cache.delete(this.cache.keys().next().value);
                this.cache.set(archivo, datos);
            }
            if (this.actual) this.actual.pause();
            this.actual = new Audio(datos);
            await this.actual.play();
            this.actual.onended = () => boton && boton.classList.remove('sonando');
        } catch (e) {
            if (boton) boton.classList.remove('sonando');
            UI.toast(_i(navigator.onLine ? 'visor.audio.error' : 'visor.audio.conexion'));
        }
    },
    deTexto(w, L) { return w?.t[L]?.audio || ''; },
    boton(w, L, etiqueta = '🔊') {
        const a = this.deTexto(w, L);
        return a ? `<button class="audio-btn" data-audio="${esc(a)}" aria-label="${esc(_i('visor.escucharPalabra', { palabra: Datos.texto(w, L) }))}" title="${esc(_i('visor.escuchar'))}">${etiqueta}</button>` : '';
    }
};
// Cualquier botón con data-audio reproduce su audio (evita código dentro de atributos)
document.addEventListener('click', e => { const b = e.target.closest('[data-audio]'); if (b) { e.stopPropagation(); Sonido.tocar(b.dataset.audio, b); } }, true);

const SoundFX = {
    _ctx: null, _get() { if (!this._ctx) this._ctx = new (window.AudioContext || window.webkitAudioContext)(); return this._ctx; },
    _tone(f, d, type = 'sine', vol = 0.25) { try { const c = this._get(), o = c.createOscillator(), g = c.createGain(); o.connect(g); g.connect(c.destination); o.type = type; o.frequency.value = f; g.gain.setValueAtTime(vol, c.currentTime); g.gain.exponentialRampToValueAtTime(0.001, c.currentTime + d); o.start(); o.stop(c.currentTime + d); } catch (e) {} },
    ok() { this._tone(880, 0.12); }, error() { this._tone(220, 0.25, 'sawtooth', 0.12); },
    win() { [523, 659, 784, 1047].forEach((f, i) => setTimeout(() => this._tone(f, 0.18, 'sine', 0.2), i * 80)); },
    fail() { [440, 349, 294, 220].forEach((f, i) => setTimeout(() => this._tone(f, 0.2, 'sawtooth', 0.12), i * 80)); },
    vibrar(p) { try { navigator.vibrate && navigator.vibrate(p); } catch (e) {} }
};

/* =============================================
   INTERFAZ: secciones, ventanas y avisos
   ============================================= */
const Juego = { activo: null, detener() { if (this.activo && this.activo.detener) this.activo.detener(); this.activo = null; Pantalla.liberar(); } };
const Pantalla = {
    lock: null,
    async mantener() { try { if ('wakeLock' in navigator) this.lock = await navigator.wakeLock.request('screen'); } catch (e) {} },
    liberar() { try { this.lock && this.lock.release(); } catch (e) {} this.lock = null; }
};

const UI = {
    seccion: 'explorar',
    ajustarCabecera() { const h = $('cabecera'); if (h) document.documentElement.style.setProperty('--hdr-h', h.offsetHeight + 'px'); },
    estadoConexion(texto, off = false, icono = '') { const c = $('conn'); c.textContent = (icono || (off ? '📶' : '✓')) + ' ' + texto; c.classList.toggle('off', off); },
    ir(sec) {
        this.seccion = sec;
        document.querySelectorAll('#nav button').forEach(b => b.setAttribute('aria-current', b.dataset.sec === sec ? 'page' : 'false'));
        history.replaceState(null, '', location.pathname + location.search + '#' + sec);
        this.refrescar();
        window.scrollTo({ top: 0 });
    },
    refrescar() {
        if (Estado.cerrado) return Estado.pantalla();
        if (!Datos.cargado) return;
        const v = $('vista');
        ({ explorar: Explorar.pintar, aprender: Hub.aprender, jugar: Hub.jugar, imprimir: Hub.imprimir, colabora: Colabora.pintar }[this.seccion] || Explorar.pintar).call(null, v);
    },
    /** Pie del visor: el contenido es de la comunidad (con su licencia); el programa, de su autor (AGPL-3.0). */
    pie() {
        const p = $('pie'); if (!p) return;
        const cfg = Datos.cfg || {};
        p.innerHTML = `<div>${cfg.titulo ? `<b>${esc(cfg.titulo)}</b> · ` : ''}${_ih('visor.pie.contenido')}${cfg.licencia ? ', ' + _ih('visor.pie.licencia', { licencia: cfg.licencia }) : ''}</div>
            <div>${esc(_i('visor.pie.credito', { autor: AUTOR }))}${APOYO ? ' · ' + esc(_i('visor.pie.apoyo', { x: APOYO })) : ''} · <a href="${esc(REPO_URL || LICENCIA_URL)}" target="_blank" rel="noopener">${_ih('visor.pie.codigo')}</a> · © 2025</div>`;
    },
    abrir(html, { ancho = false, completa = false } = {}) {
        const m = $('modal');
        m.classList.toggle('ancho', ancho); m.classList.toggle('completa', completa);
        $('modal-body').innerHTML = html;
        m.style.display = 'flex';
        m.querySelector('.modal-content').scrollTop = 0;
    },
    cuerpo(html) { $('modal-body').innerHTML = html; $('modal').querySelector('.modal-content').scrollTop = 0; },
    cerrarModal() { $('modal').style.display = 'none'; Juego.detener(); },
    clickFondo(e) { if (e.target.id === 'modal' && !Juego.activo) this.cerrarModal(); },
    toast(texto, ms = 2600) { const d = document.createElement('div'); d.className = 'toast'; d.setAttribute('role', 'status'); d.textContent = texto; document.body.appendChild(d); setTimeout(() => d.remove(), ms); },
    anunciar(texto) {
        const d = document.createElement('div');
        d.className = 'anuncio'; d.setAttribute('role', 'status'); d.setAttribute('aria-live', 'polite'); d.textContent = '🎉 ' + texto;
        document.body.appendChild(d);
        for (let i = 0; i < 14; i++) {
            const c = document.createElement('span'); c.className = 'confeti';
            c.textContent = ['🎉', '✨', '🫘', '⭐', '🎊'][i % 5];
            c.style.left = Math.round(Math.random() * 96) + 'vw'; c.style.animationDelay = (Math.random() * 0.8).toFixed(2) + 's';
            document.body.appendChild(c); setTimeout(() => c.remove(), 4000);
        }
        setTimeout(() => d.remove(), 4500);
        SoundFX.win();
    },
    zoom(src) { $('lightbox-img').src = src; $('lightbox').style.display = 'flex'; },
    /** 🔌 Abrir otro catálogo, cambiar la clave o volver a uno que ya se abrió. pideClave = el catálogo actual rechazó la clave. */
    pantallaConexion(error = '', pideClave = false) {
        if (Juego.activo) Juego.activo.detener?.();
        $('modal').style.display = 'none';
        if (!Datos.cargado) $('app-sub').textContent = _i('visor.cx.sub');
        const rec = Catalogos.lista().filter(c => c.url !== Datos.base);
        const pred = API_URL && Catalogos.limpiar(API_URL) !== Datos.base;
        $('vista').innerHTML = `<div class="caja" style="max-width:560px;margin:20px auto">
            <h2 style="margin-top:0">🔌 ${_ih('visor.cx.titulo')}</h2>
            ${error ? `<div class="error-box">${esc(error)}</div>` : ''}
            <p class="nota">${_ih('visor.cx.d')}</p>
            <label class="nota" for="cx-url">${_ih('visor.cx.direccion')}</label><input id="cx-url" class="input-box" value="${esc(Datos.base)}" autocomplete="off" spellcheck="false">
            <p class="nota">${_ih('visor.cx.enlace.d')}</p>
            ${navigator.clipboard && navigator.clipboard.readText ? `<button class="btn btn-secondary" style="min-height:34px;padding:4px 12px" id="cx-pegar">📋 ${_ih('visor.cx.pegar')}</button>` : ''}
            <label class="nota" for="cx-tok">${_ih('visor.cx.clave')}</label><input id="cx-tok" class="input-box" value="${esc(pideClave ? '' : Datos.token)}" autocomplete="off" spellcheck="false">
            <p class="nota">${_ih('visor.cx.clave.d')}</p>
            <button class="btn btn-primary btn-block" style="margin-top:10px" id="cx-ok">${_ih('visor.cx.conectar')}</button>
            ${Datos.cargado ? `<button class="btn btn-secondary btn-block" style="margin-top:6px" id="cx-volver">← ${_ih('visor.cx.cancelar')}</button>` : ''}
            ${pred ? `<button class="btn btn-secondary btn-block" style="margin-top:6px" id="cx-pred">🏠 ${_ih('visor.cx.predeterminado')}</button>` : ''}
            ${rec.length ? `<h3>${_ih('visor.cx.recientes')}</h3>${rec.map((c, i) => `<div class="fila" style="margin:4px 0"><button class="btn btn-secondary" style="flex:1;text-align:left" data-rec="${i}">${esc(c.titulo || c.url.replace(/^https:\/\/script\.google\.com\/macros\/s\//, '').slice(0, 24) + '…')}</button><button class="btn btn-secondary" data-olvidar="${i}" title="${esc(_i('visor.cx.olvidar'))}" aria-label="${esc(_i('visor.cx.olvidar'))}">✕</button></div>`).join('')}` : ''}
        </div>`;
        const ir = (url, tok) => { url = Catalogos.limpiar(url); if (!url) return; Catalogos.recordar(url, tok); LS.set('glosario_url', url); location.href = location.pathname; };
        // Si se pega un enlace de invitación, la dirección y la clave se toman de ahí
        const leerEnlace = () => { const e = Catalogos.desdeEnlace($('cx-url').value); if (e.url && e.url !== $('cx-url').value.trim()) { $('cx-url').value = e.url; if (e.token) $('cx-tok').value = e.token; UI.toast(_i('visor.cx.enlaceLeido')); } return e; };
        $('cx-url').addEventListener('change', leerEnlace); $('cx-url').addEventListener('paste', () => setTimeout(leerEnlace, 0));
        $('cx-pegar')?.addEventListener('click', async () => { try { $('cx-url').value = await navigator.clipboard.readText(); leerEnlace(); } catch (e) { UI.toast(_i('visor.cx.pegar.error')); } });
        $('cx-ok').onclick = () => {
            const e = leerEnlace(), url = e.url || Catalogos.limpiar($('cx-url').value);
            if (!/^https:\/\/script\.google(usercontent)?\.com\//.test(url)) return UI.toast(_i('visor.cx.error.direccion'));
            ir(url, $('cx-tok').value.trim() || e.token);
        };
        $('cx-volver')?.addEventListener('click', () => UI.refrescar());
        $('cx-pred')?.addEventListener('click', () => { LS.del('glosario_url'); location.href = location.pathname; });
        $('vista').querySelectorAll('[data-rec]').forEach(b => b.onclick = () => { const c = rec[+b.dataset.rec]; ir(c.url, Catalogos.clave(c.url)); });
        $('vista').querySelectorAll('[data-olvidar]').forEach(b => b.onclick = () => { Catalogos.olvidar(rec[+b.dataset.olvidar].url); UI.pantallaConexion(error, pideClave); });
        if (pideClave) setTimeout(() => $('cx-tok').focus(), 50);
    },
    /** Aviso de requisitos de una actividad para la lengua elegida. */
    requisito(req, L = Datos.lengua) {
        if (!req) return { ok: true, texto: '' };
        const c = Datos.conteo(L), tiene = c[req.campo] || 0;
        const ok = tiene >= req.min;
        const que = _i('visor.req.' + req.campo);
        return { ok, texto: ok ? _i('visor.req.disponible', { tiene, que }) : _i('visor.req.necesita', { min: req.min, que, tiene }) };
    }
};

/* =============================================
   EXPLORAR
   ============================================= */
const Explorar = {
    q: '', cat: 'all', limite: 120,
    pintar(v) {
        const cats = [...Datos.cats.values()].filter(c => Datos.words.some(w => w.cat === c.id));
        v.innerHTML = `
            <h2 class="sr">${_ih('visor.sec.explorar')}</h2>
            <div class="buscador">
                <input type="search" id="buscar" class="input-box" placeholder="${esc('🔍 ' + _i('visor.buscar'))}" value="${esc(Explorar.q)}" aria-label="${esc(_i('visor.buscarPalabra'))}" autocomplete="off">
                <select id="filtro-cat" class="modern-select" aria-label="${esc(_i('visor.categoria'))}">
                    <option value="all">${_ih('visor.todasCategorias')}</option>
                    ${cats.map(c => `<option value="${esc(c.id)}" ${Explorar.cat === c.id ? 'selected' : ''}>${esc(Datos.categoria(c.id))}</option>`).join('')}
                </select>
            </div>
            <div id="grid" class="grid" role="list"></div>
            <div style="text-align:center;margin-top:14px"><button id="mas" class="btn btn-secondary oculto">${_ih('visor.mostrarMas')}</button></div>`;
        $('buscar').addEventListener('input', e => { Explorar.q = e.target.value; Explorar.limite = 120; Explorar.lista(); });
        $('filtro-cat').addEventListener('change', e => { Explorar.cat = e.target.value; Explorar.limite = 120; Explorar.lista(); });
        $('mas').addEventListener('click', () => { Explorar.limite += 120; Explorar.lista(); });
        $('grid').addEventListener('click', e => { const c = e.target.closest('.card'); if (c) Ficha.abrir(c.dataset.id); });
        Explorar.lista();
    },
    lista() {
        const L = Datos.lengua;
        const q = Texto.pliegue(Texto.atajos(Explorar.q, L));
        const res = Datos.words.filter(w => (Explorar.cat === 'all' || w.cat === Explorar.cat) &&
            (!q || Object.values(w.t).some(t => t.formas.some(f => Texto.pliegue(f).includes(q)))));
        // También se busca en las descripciones: esas coincidencias van después de las de la palabra misma
        const porDesc = new Set();
        if (q && q.length >= 3) {
            const ya = new Set(res.map(w => w.id));
            Datos.words.forEach(w => { if (!ya.has(w.id) && (Explorar.cat === 'all' || w.cat === Explorar.cat) && Object.values(w.t).some(t => t.desc && Texto.pliegue(t.desc).includes(q))) { porDesc.add(w.id); res.push(w); } });
        }
        res.sort((a, b) => Number(porDesc.has(a.id)) - Number(porDesc.has(b.id)) || Number(!a.t[L]) - Number(!b.t[L]) || Datos.texto(a).localeCompare(Datos.texto(b), 'es'));
        const g = $('grid');
        if (!res.length) { g.innerHTML = `<div class="vacio">${_ih('visor.sinResultados')}${q ? ' ' + _ih('visor.sinResultados.comparte') : ''}</div>`; $('mas').classList.add('oculto'); return; }
        g.innerHTML = res.slice(0, Explorar.limite).map(w => {
            const t = Datos.texto(w), propia = w.t[L];
            const ph = placeholder([...t][0]);
            return `<button class="card" role="listitem" data-id="${esc(w.id)}" aria-label="${esc(t)}">
                <img src="${esc(Datos.img(w, 400) || ph)}" loading="lazy" alt="" onerror="this.onerror=null;this.src='${ph}'">
                <div class="card-body"><b${Texto.lang(propia ? L : '')}>${esc(t)}</b>
                <span class="card-tag">${esc(Datos.categoria(w.cat))}</span>${propia ? '' : `<span class="nota" style="display:block">${_ih('visor.sinRegistro', { lengua: Datos.nombreLengua(L) })}</span>`}</div>
            </button>`;
        }).join('');
        $('mas').classList.toggle('oculto', res.length <= Explorar.limite);
    }
};

const Ficha = {
    abrir(id) {
        const w = Datos.byId.get(id); if (!w) return;
        const L = Datos.lengua, t = Datos.texto(w);
        const orden = [L, ...Datos.langs.map(l => l.id).filter(x => x !== L)].filter(x => w.t[x]);
        const img = Datos.img(w, 1000), ph = placeholder([...t][0]);
        UI.abrir(`
            <h2 class="modal-titulo"${Texto.lang(w.t[L] ? L : '')}>${esc(t)}</h2>
            ${this.imagenes(w, img, ph)}
            <div class="nota" style="margin-bottom:8px">📁 ${esc(Datos.categoria(w.cat))}</div>
            ${orden.map(x => { const tr = w.t[x], f = tr.fuente && Datos.fuentes.get(tr.fuente); return `<div class="trad">
                <div class="l">${esc(Datos.nombreLengua(x))}</div>
                <div class="x"><span${Texto.lang(x)}>${tr.formas.map(esc).join(' · ')}</span>${Sonido.boton(w, x)}</div>
                ${tr.desc ? `<div class="desc"${Texto.lang(x)}>📖 ${esc(tr.desc)}</div>` : ''}
                ${tr.silabas ? `<div class="nota">${_ih('visor.silabas', { s: tr.silabas })}</div>` : ''}
                ${tr.audio && tr.voz ? `<div class="nota">🎙️ ${_ih('visor.voz', { voz: tr.voz + (tr.lugar ? ' · ' + tr.lugar : '') })}</div>` : ''}
                ${tr.nota ? `<div class="n">📌 ${esc(tr.nota)}</div>` : ''}
                ${f ? `<div class="nota">📚 ${_ih('visor.fuente', { credito: [f.credito || f.titulo, tr.ubicacion].filter(Boolean).join(' · ') })}</div>` : ''}</div>`; }).join('')}
            ${this.relacionadas(w)}
            ${Propuestas.activo() ? Propuestas.botonesFicha(w) : Colabora.botonesSugerencia(w)}`);
        $('modal-body').querySelectorAll('[data-zoom]').forEach(i => i.addEventListener('click', () => UI.zoom(i.dataset.zoom === '1' ? (img || ph) : i.dataset.zoom)));
        if (Propuestas.activo()) Propuestas.enlazarFicha(w);
        $('modal-body').querySelectorAll('[data-rel]').forEach(b => b.addEventListener('click', () => Ficha.abrir(b.dataset.rel)));
    },
    /** Imagen principal y, si hay, galería que se desliza (cada imagen con su crédito y su descripción para lectores de pantalla). */
    imagenes(w, img, ph) {
        const cred = c => c ? `<div class="nota credito">🎨 ${_ih('visor.imagenDe', { credito: c })}</div>` : '';
        const principal = `<figure class="gal-it"><img class="ficha-img" src="${esc(img || ph)}" alt="${esc(w.alt || '')}" onerror="this.onerror=null;this.src='${ph}'" data-zoom="1">${img ? cred(w.credito) : ''}</figure>`;
        if (!w.galeria || !w.galeria.length) return principal;
        const ico = { dibujo: '🖍️', foto: '📷', ilustracion: '🎨' };
        return `<div class="galeria" role="region" aria-label="${esc(_i('visor.galeria', { n: w.galeria.length + 1 }))}" tabindex="0">${principal}${w.galeria.map(g =>
            `<figure class="gal-it"><img class="ficha-img" src="${esc(g.url + '=w1000')}" alt="${esc(g.alt || '')}" loading="lazy" data-zoom="${esc(g.url + '=w1600')}">${g.credito ? `<div class="nota credito">${ico[g.tipo] || '🎨'} ${_ih('visor.imagenDe', { credito: g.credito })}</div>` : ''}</figure>`).join('')}</div>
            <div class="nota gal-n">🖼️ ${_ih('visor.galeria', { n: w.galeria.length + 1 })} · ${_ih('visor.galeria.desliza')}</div>`;
    },
    /** Palabras relacionadas, agrupadas por tipo; cada una abre su propia ficha. */
    relacionadas(w) {
        const l = (Datos.rel.get(w.id) || []).filter(r => Datos.byId.has(r.otra));
        if (!l.length) return '';
        const orden = t => (Datos.tiposRel.get(t) || {}).orden ?? 999;
        const grupos = new Map();
        l.slice().sort((a, b) => orden(a.tipo) - orden(b.tipo) || Number(a.inv) - Number(b.inv)).forEach(r => {
            const et = Datos.etiquetaRel(r.tipo, r.inv);
            if (!grupos.has(et)) grupos.set(et, []);
            grupos.get(et).push(r);
        });
        const L = Datos.lengua;
        return `<div class="relacionadas"><h3>🔗 ${_ih('visor.relacionadas')}</h3>${[...grupos].map(([et, rs]) => `<div class="rel-g"><div class="l">${esc(et)}</div><div class="rel-items">${
            rs.map(r => { const o = Datos.byId.get(r.otra); return `<button class="chip-rel" data-rel="${esc(r.otra)}"${Texto.lang(o.t[L] ? L : '')}>${esc(Datos.texto(o))}</button>${r.nota ? `<span class="nota">${esc(r.nota)}</span>` : ''}`; }).join('')
        }</div></div>`).join('')}</div>`;
    }
};

/* =============================================
   COLABORA (contacto fuera del visor) y ACERCA
   ============================================= */
const Colabora = {
    contacto() { return Datos.cfg.contacto || {}; },
    mensajeWA(texto) { const c = this.contacto(); return c.whatsapp ? `https://wa.me/${encodeURIComponent(c.whatsapp.replace(/[^\d]/g, ''))}?text=${encodeURIComponent(texto)}` : ''; },
    mensajeCorreo(asunto, texto) { const c = this.contacto(); return c.correo ? `mailto:${encodeURIComponent(c.correo)}?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(texto)}` : ''; },
    enlaces(asunto, texto) {
        const wa = this.mensajeWA(texto), co = this.mensajeCorreo(asunto, texto);
        return `${wa ? `<a class="btn btn-success" href="${esc(wa)}" target="_blank" rel="noopener">💬 WhatsApp</a>` : ''}${co ? `<a class="btn btn-secondary" href="${esc(co)}">✉️ ${_ih('visor.correo')}</a>` : ''}`;
    },
    botonesSugerencia(w) {
        const enl = this.enlaces(_i('visor.sug.asunto', { palabra: Datos.texto(w) }),
            _i('visor.sug.texto', { palabra: Datos.texto(w), lengua: Datos.nombreLengua(Datos.lengua), catalogo: Datos.cfg.titulo || '', id: w.id }) + '\n');
        return enl ? `<div class="caja"><b>${_ih('visor.sug.titulo')}</b><p class="nota">${_ih('visor.sug.d')}</p><div class="fila">${enl}</div></div>` : '';
    },
    pintar(v) {
        const c = Colabora.contacto(), cfg = Datos.cfg;
        const hay = c.correo || c.telefono || c.whatsapp;
        const general = Colabora.enlaces(_i('visor.col.asunto', { catalogo: cfg.titulo || '' }), _i('visor.col.texto', { catalogo: cfg.titulo || '' }) + '\n');
        const lenguas = Datos.langs.filter(l => Datos.words.some(w => w.t[l.id]));
        v.innerHTML = `
            <h2 class="titulo-seccion">🤝 ${_ih('visor.sec.colabora')}</h2>
            <div class="caja">
                <p class="texto-largo" style="margin-top:0">${c.texto ? esc(c.texto) : _ih('visor.col.intro')}</p>
                ${hay ? `<div class="fila">${general}${c.telefono ? `<a class="btn btn-secondary" href="tel:${esc(c.telefono.replace(/[^\d+]/g, ''))}">📞 ${_ih('visor.col.llamar')}</a>` : ''}</div>
                    <p class="nota">${[c.nombre ? _ih('visor.col.atiende', { x: c.nombre }) : '', c.horario ? _ih('visor.col.horario', { x: c.horario }) : '', c.correo ? _ih('visor.col.correo', { x: c.correo }) : '', c.telefono ? _ih('visor.col.telefono', { x: c.telefono }) : ''].filter(Boolean).join(' ')}</p>`
                : `<p class="nota">${_ih('visor.col.sinContacto')}</p>`}
            </div>
            ${Colabora.participar()}
            <h3>${_ih('visor.col.como')}</h3>
            <ul>
                <li>${_ih('visor.col.c1')}</li>
                <li>${_ih('visor.col.c2')}</li>
                <li>${_ih('visor.col.c3')}</li>
                <li>${_ih('visor.col.c4')}</li>
            </ul>
            <p class="nota">${_ih('visor.col.revision')}</p>
            ${cfg.acerca ? `<h3>📖 ${_ih('visor.acerca')}</h3><p class="texto-largo">${esc(cfg.acerca)}</p>` : ''}
            ${Colabora.gobernanza()}
            ${cfg.creditos ? `<h3>🙏 ${_ih('visor.creditos')}</h3><p class="texto-largo">${esc(cfg.creditos)}</p>` : ''}
            ${Colabora.fuentes()}
            ${cfg.licencia ? `<h3>📜 ${_ih('visor.licencia')}</h3><p>${esc(cfg.licencia)}</p>` : ''}
            ${Colabora.terminos()}
            <h3>🌐 ${_ih('visor.lenguasCatalogo')}</h3>
            <div class="chips">${lenguas.map(l => `<span class="chip" style="cursor:default"><b${Texto.lang(l.id)}>${esc(l.nombre)}</b>${l.otro && l.otro !== l.nombre ? ' · ' + esc(l.otro) : ''} · ${Datos.words.filter(w => w.t[l.id]).length}</span>`).join('')}</div>
            <div class="fila" style="margin-top:16px">
                <button class="btn btn-secondary" onclick="Colabora.compartir()">📤 ${_ih('visor.compartir')}</button>
                <button class="btn btn-secondary" onclick="Interfaz.abrir()">🌐 ${_ih('visor.interfaz.boton')}</button>
            </div>
            <h3>🛠️ ${_ih('visor.sobreVisor')}</h3>
            <p>${esc(_i('visor.pie.credito', { autor: AUTOR }))}${APOYO ? ' · ' + esc(_i('visor.pie.apoyo', { x: APOYO })) : ''} · © 2025.
                ${_ih('visor.sobreVisor.d')} <a href="${esc(REPO_URL || LICENCIA_URL)}" target="_blank" rel="noopener">${_ih('visor.pie.codigo')}</a></p>
            <p class="nota" style="margin-top:16px">${_ih('visor.version', { v: VERSION_VISOR })}${cfg.version ? ' · ' + _ih('visor.versionCatalogo', { v: cfg.version }) : ''}${Datos.version ? ' · ' + _ih('visor.versionPublicacion', { v: Datos.version }) : ''}</p>`;
    },
    /** Proponer desde el visor (si el equipo lo abrió), Mis propuestas y quiénes han aportado. */
    participar() {
        const act = Propuestas.activo(), mias = Propuestas.mias().length;
        let h = '';
        if (act) h += `<div class="caja"><h3 style="margin-top:0">🤝 ${_ih('visor.prop.titulo')}</h3><p class="nota">${_ih('visor.prop.intro')}</p>
            <div class="actividades">
                <button class="actividad" onclick="Propuestas.abrirPalabras()"><span class="ico">➕</span><span><b>${_ih('visor.prop.palabras')}</b><span class="d">${_ih('visor.prop.palabras.d2')}</span></span></button>
                <button class="actividad" onclick="Propuestas.abrirFotos()"><span class="ico">📷</span><span><b>${_ih('visor.prop.fotos')}</b><span class="d">${_ih('visor.prop.fotos.d2')}</span></span></button>
                <button class="actividad" onclick="Propuestas.dondeAyudar()"><span class="ico">🤝</span><span><b>${_ih('visor.prop.donde')}</b><span class="d">${_ih('visor.prop.donde.d2')}</span></span></button>
                <button class="actividad" onclick="Propuestas.misPropuestas()"><span class="ico">📬</span><span><b>${_ih('visor.prop.mias')}</b><span class="d">${_ih('visor.prop.mias.d2')}</span></span></button>
            </div></div>`;
        else if (mias) h += `<div class="fila"><button class="btn btn-secondary" onclick="Propuestas.misPropuestas()">📬 ${_ih('visor.prop.mias')}</button></div>`;
        const ap = Datos.aportes || [];
        if (ap.length) h += `<details class="caja"><summary><b>💛 ${_ih('visor.prop.aportes', { n: ap.length })}</b></summary><p>${ap.map(a => `${esc(a.nombre)} <span class="nota">(${a.n})</span>`).join(' · ')}</p></details>`;
        return h;
    },
    /** Quién decide sobre estos datos: comunidades y acuerdos (sin datos personales). */
    gobernanza() {
        const g = Datos.gobernanza; if (!g) return '';
        const com = (g.comunidades || []);
        return `<div class="caja"><p class="nota" style="margin-top:0">🤝 ${esc(g.aviso || '')}</p>${com.length ? `<b>${_ih('visor.gob.comunidades')}</b><ul>${com.map(c => `<li><b>${esc(c.nombre_propio)}</b>${c.nombre_otro ? ' (' + esc(c.nombre_otro) + ')' : ''}${c.variante || c.region ? ' · ' + esc([c.variante, c.region].filter(Boolean).join(' · ')) : ''}${c.consulta ? '<br><span class="nota">' + _ih('visor.gob.consulta', { x: c.consulta }) + '</span>' : ''}</li>`).join('')}</ul>` : ''}</div>`;
    },
    fuentes() {
        const f = [...(Datos.fuentes || new Map()).values()];
        return f.length ? `<h3>📚 ${_ih('visor.fuentes')}</h3><ul>${f.map(x => `<li>${esc(x.credito || x.titulo)}${x.anio && !(x.credito || '').includes(x.anio) ? ' (' + esc(x.anio) + ')' : ''}${x.licencia ? ' · ' + esc(x.licencia) : ''}${x.enlace ? ` · <a href="${esc(x.enlace)}" target="_blank" rel="noopener">${_ih('visor.verFuente')}</a>` : ''}</li>`).join('')}</ul>` : '';
    },
    terminos() {
        const t = Datos.gobernanza && Datos.gobernanza.terminos; if (!t) return '';
        return `<details class="caja"><summary><b>📄 ${esc(t.titulo)}</b></summary><ol>${t.puntos.map(p => `<li><b>${esc(p.titulo)}</b> ${esc(p.texto)}</li>`).join('')}</ol><p class="nota">${esc(t.acepto)} · ${esc(t.version)}</p></details>`;
    },
    async compartir() {
        const u = new URL(location.origin + location.pathname);
        if (Datos.base !== (API_URL || '').split('?')[0]) u.searchParams.set('api', Datos.base);
        const texto = `${Datos.cfg.titulo || _i('visor.catalogoPalabras')}: ${u}`;
        if (navigator.share) { try { await navigator.share({ title: Datos.cfg.titulo || _i('visor.catalogoPalabras'), text: texto, url: u.toString() }); return; } catch (e) {} }
        try { await navigator.clipboard.writeText(u.toString()); UI.toast(_i('visor.enlaceCopiado')); } catch (e) { prompt(_i('visor.copiaEnlace'), u.toString()); }
    }
};

/* =============================================
   APRENDER · JUGAR · IMPRIMIR (tableros de actividades)
   ============================================= */
const ACTIVIDADES = [
    // Jugar
    { sec: 'jugar', ico: '🎴', id: 'loteria', d: '', req: { campo: 'img', min: 32 }, abrir: () => Loteria.abrir() },
    { sec: 'jugar', ico: '🎲', id: 'unirse', d: '', abrir: () => Loteria.pantallaUnirse() },
    { sec: 'jugar', ico: '🧠', id: 'memorama', d: '', req: { campo: 'img', min: 6 }, abrir: () => Memorama.setup() },
    { sec: 'jugar', ico: '✋', id: 'basta', d: '', req: { campo: 'cats', min: 3 }, abrir: () => Basta.setup() },
    { sec: 'jugar', ico: '🎭', id: 'charadas', d: '', req: { campo: 'img', min: 10 }, abrir: () => Charadas.setup() },
    { sec: 'jugar', ico: '😵', id: 'ahorcado', d: '', req: { campo: 'largas', min: 6 }, abrir: () => Ahorcado.setup() },
    { sec: 'jugar', ico: '🧩', id: 'sopa', d: '', req: { campo: 'largas', min: 6 }, abrir: () => Sopa.setup() },
    // Aprender
    { sec: 'aprender', ico: '📈', id: 'progreso', d: '', abrir: () => Progreso.abrir() },
    { sec: 'aprender', ico: '🃏', id: 'repaso', d: '', req: { campo: 'img', min: 5 }, abrir: () => Repaso.abrir() },
    { sec: 'aprender', ico: '👂', id: 'escucha', d: '', req: { campo: 'imgAudio', min: 4 }, abrir: () => Escucha.abrir() },
    { sec: 'aprender', ico: '🌎', id: 'enquelengua', d: '', req: { campo: 'lenguas', min: 2 }, abrir: () => EnQueLengua.abrir() },
    { sec: 'aprender', ico: '🔤', id: 'arma', d: '', req: { campo: 'largas', min: 6 }, abrir: () => Arma.abrir() },
    { sec: 'aprender', ico: '✍️', id: 'dictado', d: '', req: { campo: 'audio', min: 5 }, abrir: () => Dictado.abrir() },
    { sec: 'aprender', ico: '🔤', id: 'queLetra', d: '', req: { campo: 'img', min: 6 }, abrir: () => QueLetra.abrir() },
    { sec: 'aprender', ico: '📖', id: 'quees', d: '', req: { campo: 'desc', min: 4 }, soloSiHay: true, abrir: () => QueEs.abrir() },
    // Imprimir
    { sec: 'imprimir', ico: '🎴', id: 'loteriaImp', d: '', req: { campo: 'img', min: 32 }, abrir: () => Loteria.abrir() },
    { sec: 'imprimir', ico: '📝', id: 'hojas', d: '', req: { campo: 'textos', min: 6 }, abrir: () => { Hojas.ultimo = null; Hojas.base = null; Hojas.abrir('une'); } },
    { sec: 'imprimir', ico: '🏷️', id: 'etiquetas', d: '', req: { campo: 'img', min: 4 }, abrir: () => { Hojas.ultimo = null; Hojas.base = null; Hojas.abrir('etiquetas'); } },
    { sec: 'imprimir', ico: '🗂️', id: 'tarjetas', d: '', req: { campo: 'img', min: 4 }, abrir: () => { Hojas.ultimo = null; Hojas.base = null; Hojas.abrir('tarjetas'); } },
    { sec: 'imprimir', ico: '🧩', id: 'sopaImp', d: '', req: { campo: 'largas', min: 6 }, abrir: () => Impresos.abrir('sopa') },
    { sec: 'imprimir', ico: '✏️', id: 'crucigrama', d: '', req: { campo: 'largas', min: 6 }, abrir: () => Impresos.abrir('crucigrama') },
    { sec: 'imprimir', ico: '🔀', id: 'ordena', d: '', req: { campo: 'largas', min: 6 }, abrir: () => Impresos.abrir('ordena') },
    { sec: 'imprimir', ico: '✋', id: 'bastaImp', d: '', req: { campo: 'cats', min: 2 }, abrir: () => Impresos.abrir('basta') },
    { sec: 'imprimir', ico: '🧠', id: 'memoramaImp', d: '', req: { campo: 'img', min: 8 }, abrir: () => Impresos.abrir('memorama') },
    { sec: 'imprimir', ico: '🏡', id: 'preguntaCasa', d: '', abrir: () => Impresos.abrir('preguntaCasa') },
    { sec: 'imprimir', ico: '📋', id: 'recoleccion', d: '', abrir: () => Impresos.abrir('recoleccion') },
    { sec: 'imprimir', ico: '🔤', id: 'inicialImp', d: '', req: { campo: 'textos', min: 6 }, abrir: () => Impresos.abrir('inicial') },
    { sec: 'imprimir', ico: '👏', id: 'silabasImp', d: '', req: { campo: 'silabas', min: 4 }, soloSiHay: true, abrir: () => Impresos.abrir('silabas') },
    { sec: 'imprimir', ico: '🧱', id: 'armaSilabas', d: '', req: { campo: 'silabas', min: 3 }, soloSiHay: true, abrir: () => Impresos.abrir('armaSilabas') },
    { sec: 'imprimir', ico: '✍️', id: 'trazo', d: '', abrir: () => Impresos.abrir('trazo') },
    { sec: 'imprimir', ico: '👂', id: 'escuchaImp', d: '', req: { campo: 'img', min: 6 }, abrir: () => Impresos.abrir('escucha') },
    { sec: 'imprimir', ico: '📚', id: 'glosario', d: '', req: { campo: 'textos', min: 4 }, abrir: () => Impresos.abrir('glosario') },
    { sec: 'imprimir', ico: '✅', id: 'cotejo', d: '', abrir: () => Impresos.abrir('cotejo') },
    { sec: 'imprimir', ico: '📂', id: 'misActividades', d: '', abrir: () => Actividades.abrir() },
    { sec: 'imprimir', ico: '📋', id: 'misListas', d: '', abrir: () => Listas.abrir() }
];

const Hub = {
    tarjetas(sec) {
        return `<div class="actividades">${ACTIVIDADES.map((a, i) => ({ a, i })).filter(x => x.a.sec === sec && !(x.a.soloSiHay && !UI.requisito(x.a.req).ok)).map(({ a, i }) => {
            const r = UI.requisito(a.req);
            return `<button class="actividad" data-act="${i}" ${r.ok ? '' : 'aria-disabled="true"'}>
                <span class="ico" aria-hidden="true">${a.ico}</span>
                <span><b>${_ih('visor.act.' + a.id)}</b><span class="d">${_ih('visor.act.' + a.id + '.d')}</span>${a.req ? `<span class="req ${r.ok ? 'si' : 'no'}">${esc(r.texto)}</span>` : ''}</span></button>`;
        }).join('')}</div>`;
    },
    enlazar(v) {
        v.querySelectorAll('[data-act]').forEach(b => b.addEventListener('click', () => {
            const a = ACTIVIDADES[+b.dataset.act], r = UI.requisito(a.req);
            if (!r.ok) return UI.toast(_i('visor.req.toast', { actividad: _i('visor.act.' + a.id), req: r.texto, lengua: Datos.nombreLengua(Datos.lengua) }), 4000);
            a.abrir();
        }));
    },
    aprender(v) {
        const L = Datos.lengua;
        const pool = Datos.words.filter(w => w.t[L] && w.img);
        let dia = '';
        if (pool.length) {
            const hoy = new Date().toISOString().slice(0, 10);
            const conAudio = pool.filter(w => w.t[L].audio);
            const lista = conAudio.length >= 5 ? conAudio : pool;
            const w = lista[hash32(hoy + L) % lista.length];
            dia = `<div class="dia"><img src="${esc(Datos.img(w, 300))}" alt="" onerror="this.onerror=null;this.src='${placeholder([...Datos.texto(w)][0])}'">
                <div><div class="nota">${_ih('visor.dia', { lengua: Datos.nombreLengua(L) })}</div>
                <div class="w"${Texto.lang(L)}>${esc(w.t[L].texto)}</div>
                <div class="fila" style="margin-top:6px">${Sonido.boton(w, L)}<button class="btn btn-secondary" data-ficha="${esc(w.id)}">${_ih('visor.verTodas')}</button></div></div></div>`;
        }
        v.innerHTML = `<h2 class="titulo-seccion">🎓 ${_ih('visor.sec.aprender')}</h2>${dia}${Hub.tarjetas('aprender')}`;
        Hub.enlazar(v);
    },
    jugar(v) {
        v.innerHTML = `<h2 class="titulo-seccion">🎲 ${_ih('visor.sec.jugar')}</h2>
            <p class="intro">${_ih('visor.jugar.intro')}</p>
            ${Hub.tarjetas('jugar')}`;
        Hub.enlazar(v);
    },
    imprimir(v) {
        v.innerHTML = `<h2 class="titulo-seccion">🖨️ ${_ih('visor.sec.imprimir')}</h2>
            <p class="intro">${_ih('visor.imprimir.intro')}</p>
            ${Hub.tarjetas('imprimir')}`;
        Hub.enlazar(v);
    }
};
</script>
