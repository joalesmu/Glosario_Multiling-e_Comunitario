<script>
/* =============================================
   RECURSOS PARA DOCENTES — base común de lo que se imprime
   - Selección de palabras: al azar por categorías o «Mi lista», con filtros.
   - Pistas por palabra: imagen, otra lengua, descripción del catálogo,
     pista propia (editable, se guarda en este dispositivo), espacio o nada.
   - Versiones (A, B, C), hoja del estudiante y hoja del docente con respuestas,
     y un código por hoja para volver a generarla igual.
   - Ajustes de impresión: papel, márgenes seguros, tinta, letra y pie.
   - Mis actividades y Mis listas: guardadas en el dispositivo, con respaldo en archivo.
   - Llevar sin conexión: imágenes y audios guardados en el dispositivo.
   Nada de esto se envía al catálogo: vive en el navegador de quien lo usa.
   ============================================= */
const AUTOR = 'Alejandro Estrada';          // crédito de autoría (AGPL-3.0, sección 7 b): no se traduce ni se quita
const REPO_URL = 'https://github.com/joalesmu/Glosario_Multiling-e_Comunitario';                         // dirección del repositorio original (se muestra en el pie)
const APOYO = '';                            // mención de apoyo institucional que indique el titular (vacío = ninguna)
const LICENCIA_URL = 'https://www.gnu.org/licenses/agpl-3.0.html';

/* ---------- Llevar sin conexión: imágenes y audios en el dispositivo (IndexedDB) ---------- */
const Offline = {
    db: null, imgs: new Map(), audios: new Set(), listo: false,
    abrirDB() {
        if (this.db) return Promise.resolve(this.db);
        return new Promise(res => {
            try {
                const r = indexedDB.open('catalogo-lenguas-medios', 1);
                r.onupgradeneeded = () => { const d = r.result; if (!d.objectStoreNames.contains('m')) d.createObjectStore('m'); };
                r.onsuccess = () => { this.db = r.result; res(this.db); };
                r.onerror = () => res(null);
            } catch (e) { res(null); }
        });
    },
    async tx(modo, fn) {
        const d = await this.abrirDB(); if (!d) return null;
        return new Promise(res => {
            try { const t = d.transaction('m', modo), s = t.objectStore('m'); const r = fn(s); t.oncomplete = () => res(r && 'result' in r ? r.result : true); t.onerror = () => res(null); }
            catch (e) { res(null); }
        });
    },
    clave(tipo, id) { return `${Datos.base}|${tipo}|${id}`; },
    /** Carga en memoria lo guardado de este catálogo (las imágenes se usan al pintar sin esperar). */
    async iniciar() {
        this.imgs.clear(); this.audios.clear();
        const d = await this.abrirDB(); if (!d) { this.listo = true; return; }
        const pref = Datos.base + '|';
        await new Promise(res => {
            try {
                const t = d.transaction('m', 'readonly'), c = t.objectStore('m').openCursor();
                c.onsuccess = () => {
                    const cur = c.result; if (!cur) return;
                    const k = String(cur.key);
                    if (k.startsWith(pref + 'img|')) this.imgs.set(k.slice(pref.length + 4), cur.value);
                    else if (k.startsWith(pref + 'aud|')) this.audios.add(k.slice(pref.length + 4));
                    cur.continue();
                };
                t.oncomplete = res; t.onerror = res;
            } catch (e) { res(); }
        });
        this.listo = true;
    },
    img(id) { return this.imgs.get(String(id)) || ''; },
    tieneAudio(archivo) { return this.audios.has(String(archivo)); },
    async audio(archivo) { return this.audios.has(String(archivo)) ? await this.tx('readonly', s => s.get(this.clave('aud', archivo))) : null; },
    /** Guarda imágenes (por el catálogo, que las entrega ligeras) y, si se pide, audios de las palabras. */
    async llevar(ws, { audios = false, L = '' } = {}, avance = () => {}) {
        try { if (navigator.storage && navigator.storage.persist) await navigator.storage.persist(); } catch (e) {}
        const faltan = ws.filter(w => w.img && !this.imgs.has(w.id)).map(w => w.id);
        let hechas = ws.length - faltan.length, fallos = 0, total = ws.length;
        for (let i = 0; i < faltan.length; i += 25) {
            const lote = faltan.slice(i, i + 25);
            let r = null;
            try { r = await Datos.api({ action: 'paquete', ids: lote.join(','), ancho: 600 }, 240000); } catch (e) { r = null; }
            for (const id of lote) {
                let datos = r && r.status === 200 ? (r.imagenes || {})[id] : '';
                if (!datos) {   // respaldo: copiar la imagen pública desde el navegador
                    try { const b = await (await fetch(Datos.img(Datos.byId.get(id), 600, true))).blob();
                        if (/^image\//.test(b.type)) datos = await new Promise(ok => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.onerror = () => ok(''); fr.readAsDataURL(b); }); } catch (e) { datos = ''; }
                }
                if (datos) { await this.tx('readwrite', s => s.put(datos, this.clave('img', id))); this.imgs.set(id, datos); hechas++; } else fallos++;
                avance(hechas, total);
            }
        }
        let aud = 0, audFallos = 0;
        if (audios) {
            const lista = [...new Set(ws.flatMap(w => (L ? [w.t[L]] : Object.values(w.t)).filter(t => t && t.audio).map(t => t.audio)))].filter(a => !this.audios.has(a));
            total += lista.length;
            for (const a of lista) {
                try {
                    const r = await Datos.api({ type: 'audio', file: a }, 30000);
                    if (r.status === 'ok' && r.audioData) { await this.tx('readwrite', s => s.put(r.audioData, this.clave('aud', a))); this.audios.add(a); aud++; } else audFallos++;
                } catch (e) { audFallos++; }
                avance(hechas + aud, total);
            }
        }
        return { imagenes: hechas, fallos, audios: aud, audFallos };
    },
    /** Cuántas de estas palabras tienen su imagen guardada (para avisar antes de imprimir sin internet). */
    faltantes(ws) { return ws.filter(w => w.img && !this.imgs.has(w.id)).length; },
    async borrarTodo() {
        const d = await this.abrirDB(); if (!d) return;
        const pref = Datos.base + '|';
        await new Promise(res => {
            try {
                const t = d.transaction('m', 'readwrite'), c = t.objectStore('m').openCursor();
                c.onsuccess = () => { const cur = c.result; if (!cur) return; if (String(cur.key).startsWith(pref)) cur.delete(); cur.continue(); };
                t.oncomplete = res; t.onerror = res;
            } catch (e) { res(); }
        });
        this.imgs.clear(); this.audios.clear();
    },
    /** Panel para llevar sin conexión un grupo de palabras. */
    panel(ws, titulo, volver) {
        const conAudio = ws.some(w => Object.values(w.t).some(t => t.audio));
        const guardadas = ws.filter(w => w.img && this.imgs.has(w.id)).length, conImg = ws.filter(w => w.img).length;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📥 ${_ih('visor.off.titulo')}</h2>
            <p>${_ih('visor.off.d', { x: titulo, n: ws.length })}</p>
            <div class="caja">${_ih('visor.off.estado', { g: guardadas, t: conImg })}</div>
            ${conAudio ? `<label class="chip" style="margin-top:8px"><input type="checkbox" id="of-aud" checked> 🔊 ${_ih('visor.off.audios')}</label>` : ''}
            <p class="nota">${_ih('visor.off.nota')}</p>
            <div id="of-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="of-ok">📥 ${_ih('visor.off.boton')}</button>
            ${volver ? `<button class="btn btn-secondary" id="of-volver">↩️ ${_ih('visor.volver')}</button>` : ''}</div></div>`);
        if (volver) $('of-volver').addEventListener('click', volver);
        $('of-ok').addEventListener('click', async () => {
            if (!navigator.onLine) return $('of-msg').innerHTML = `<div class="error-box">${_ih('visor.off.sinRed')}</div>`;
            $('of-ok').disabled = true;
            $('of-msg').innerHTML = `<div class="caja" role="status">⏳ ${_ih('visor.off.bajando')} <span id="of-av">0</span></div>`;
            const r = await this.llevar(ws, { audios: !!$('of-aud')?.checked }, (h, t) => { const a = $('of-av'); if (a) a.textContent = `${h}/${t}`; });
            $('of-msg').innerHTML = `<div class="caja">✅ ${_ih('visor.off.listo', { i: r.imagenes, a: r.audios })}${r.fallos || r.audFallos ? '<br>⚠️ ' + _ih('visor.off.fallos', { n: r.fallos + r.audFallos }) : ''}</div>`;
            $('of-ok').disabled = false;
        });
    }
};

/* ---------- Mis listas de palabras (vocabulario de la unidad, en este dispositivo) ---------- */
const Listas = {
    clave() { return 'listas:' + Datos.base; },
    todas() { return (LS.get(this.clave(), []) || []).filter(l => l && l.id && Array.isArray(l.ids)); },
    guardarTodas(l) { return LS.set(this.clave(), l); },
    de(id) { return this.todas().find(l => l.id === id) || null; },
    palabras(l) { return l ? Datos.palabrasDe(l.ids) : []; },
    guardar(lista) {
        const t = this.todas(), i = t.findIndex(x => x.id === lista.id);
        lista.modificada = new Date().toISOString();
        if (i > -1) t[i] = lista; else t.unshift(lista);
        if (!this.guardarTodas(t)) UI.toast(_i('visor.act.lleno'), 5000);
    },
    borrar(id) { this.guardarTodas(this.todas().filter(l => l.id !== id)); },
    /** Pantalla de listas; `volver` regresa a la ventana desde la que se abrió. */
    abrir(volver) {
        const ls = this.todas();
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📋 ${_ih('visor.listas.titulo')}</h2>
            <p class="nota">${_ih('visor.listas.d')}</p>
            ${ls.length ? ls.map(l => `<div class="caja act-fila"><div><b>${esc(l.nombre)}</b><div class="nota">${_ih('visor.listas.n', { n: Listas.palabras(l).length })}</div></div>
                <div class="fila"><button class="btn btn-secondary btn-mini" data-ed="${esc(l.id)}">✏️ ${_ih('visor.editar')}</button>
                <button class="btn btn-secondary btn-mini" data-off="${esc(l.id)}">📥 ${_ih('visor.off.corto')}</button>
                <button class="btn btn-secondary btn-mini" data-bo="${esc(l.id)}" aria-label="${esc(_i('visor.borrar'))}">🗑️</button></div></div>`).join('')
            : `<div class="caja">${_ih('visor.listas.vacia')}</div>`}
            <div class="acciones"><button class="btn btn-primary" id="li-nueva">＋ ${_ih('visor.listas.nueva')}</button>
            ${volver ? `<button class="btn btn-secondary" id="li-volver">↩️ ${_ih('visor.volver')}</button>` : ''}</div></div>`);
        $('li-nueva').addEventListener('click', () => this.editar(null, volver));
        if (volver) $('li-volver').addEventListener('click', volver);
        $('modal-body').querySelectorAll('[data-ed]').forEach(b => b.addEventListener('click', () => this.editar(this.de(b.dataset.ed), volver)));
        $('modal-body').querySelectorAll('[data-off]').forEach(b => b.addEventListener('click', () => { const l = this.de(b.dataset.off); Offline.panel(this.palabras(l), l.nombre, () => this.abrir(volver)); }));
        $('modal-body').querySelectorAll('[data-bo]').forEach(b => b.addEventListener('click', () => { if (confirm(_i('visor.listas.borrar'))) { this.borrar(b.dataset.bo); this.abrir(volver); } }));
    },
    editar(lista, volver) {
        const l = lista ? JSON.parse(JSON.stringify(lista)) : { id: 'l' + Date.now().toString(36) + azarSeguro(1296).toString(36), nombre: '', ids: [], creada: new Date().toISOString() };
        const sel = new Set(l.ids.filter(id => Datos.byId.has(id)));
        const L = Datos.lengua;
        const cats = [...Datos.cats.values()].filter(c => Datos.words.some(w => w.cat === c.id));
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📋 ${_ih(lista ? 'visor.listas.editar' : 'visor.listas.nueva')}</h2>
            <label class="t" for="le-n">${_ih('visor.listas.nombre')}</label><input id="le-n" class="input-box" maxlength="80" value="${esc(l.nombre)}" placeholder="${esc(_i('visor.listas.nombre.ph'))}">
            <div class="fila" style="margin-top:8px"><input id="le-q" class="input-box" style="flex:2" type="search" placeholder="${esc('🔍 ' + _i('visor.buscar'))}" aria-label="${esc(_i('visor.buscarPalabra'))}">
            <select id="le-c" class="modern-select" style="flex:1" aria-label="${esc(_i('visor.categoria'))}"><option value="">${_ih('visor.todasCategorias')}</option>${cats.map(c => `<option value="${esc(c.id)}">${esc(Datos.categoria(c.id))}</option>`).join('')}</select></div>
            <div class="fila" style="margin:6px 0"><label class="chip"><input type="checkbox" id="le-solo"> ${_ih('visor.listas.soloElegidas')}</label><span class="nota" id="le-cuenta"></span></div>
            <div id="le-lista" class="le-lista" role="group" aria-label="${esc(_i('visor.listas.palabras'))}"></div>
            <div id="le-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="le-ok">💾 ${_ih('visor.guardar')}</button><button class="btn btn-secondary" id="le-no">${_ih('visor.cancelar')}</button></div></div>`);
        const pintar = () => {
            const q = Texto.pliegue(Texto.atajos($('le-q').value, L)), c = $('le-c').value, solo = $('le-solo').checked;
            const res = Datos.words.filter(w => (!c || w.cat === c) && (!solo || sel.has(w.id)) &&
                (!q || Object.values(w.t).some(t => t.formas.some(f => Texto.pliegue(f).includes(q)))))
                .sort((a, b) => Datos.texto(a).localeCompare(Datos.texto(b), 'es')).slice(0, 300);
            $('le-lista').innerHTML = res.map(w => `<label class="le-it"><input type="checkbox" value="${esc(w.id)}" ${sel.has(w.id) ? 'checked' : ''}>
                ${w.img ? `<img src="${esc(Datos.img(w, 120))}" alt="" loading="lazy">` : '<span class="le-sin"></span>'}
                <span><b${Texto.lang(w.t[L] ? L : '')}>${esc(Datos.texto(w))}</b><small>${esc(Datos.categoria(w.cat))}</small></span></label>`).join('') || `<p class="nota">${_ih('visor.sinResultados')}</p>`;
            $('le-cuenta').textContent = _i('visor.listas.elegidas', { n: sel.size });
        };
        $('le-lista').addEventListener('change', e => { const x = e.target; if (x.type !== 'checkbox') return; x.checked ? sel.add(x.value) : sel.delete(x.value); $('le-cuenta').textContent = _i('visor.listas.elegidas', { n: sel.size }); });
        ['le-q', 'le-c', 'le-solo'].forEach(i => $(i).addEventListener(i === 'le-q' ? 'input' : 'change', pintar));
        pintar();
        $('le-no').addEventListener('click', () => this.abrir(volver));
        $('le-ok').addEventListener('click', () => {
            const nombre = $('le-n').value.trim();
            if (!nombre) return $('le-msg').innerHTML = `<div class="error-box">${_ih('visor.listas.faltaNombre')}</div>`;
            if (!sel.size) return $('le-msg').innerHTML = `<div class="error-box">${_ih('visor.listas.faltan')}</div>`;
            l.nombre = nombre.slice(0, 80); l.ids = [...sel];
            this.guardar(l); UI.toast(_i('visor.listas.guardada'));
            this.abrir(volver);
        });
    }
};

/* ---------- Selección de palabras para un material ---------- */
const Seleccion = {
    LARGOS: [[0, 'visor.sel.largo.todos'], [5, 'visor.sel.largo.5'], [8, 'visor.sel.largo.8'], [12, 'visor.sel.largo.12']],
    html(px, L, c = {}) {
        const ls = Listas.todas();
        return `<fieldset class="grupo"><legend>📚 ${_ih('visor.sel.titulo')}</legend>
            <div class="chips"><label class="chip"><input type="radio" name="${px}-f" value="azar" ${c.fuente !== 'lista' ? 'checked' : ''}> 🎲 ${_ih('visor.sel.azar')}</label>
            <label class="chip"><input type="radio" name="${px}-f" value="lista" ${c.fuente === 'lista' ? 'checked' : ''} ${ls.length ? '' : 'disabled'}> 📋 ${_ih('visor.sel.lista')}</label></div>
            <div id="${px}-lis" class="${c.fuente === 'lista' ? '' : 'oculto'}" style="margin-top:6px"><select id="${px}-li" class="modern-select" aria-label="${esc(_i('visor.sel.lista'))}">${ls.map(l => `<option value="${esc(l.id)}" ${l.id === c.lista ? 'selected' : ''}>${esc(l.nombre)} (${Listas.palabras(l).length})</option>`).join('')}</select></div>
            <div id="${px}-cat-b" class="${c.fuente === 'lista' ? 'oculto' : ''}"><details ${c.cats && c.cats.length ? 'open' : ''}><summary class="t" style="cursor:pointer">📁 ${_ih('visor.elegirCategorias')}</summary><div id="${px}-cats">${Comun.categorias(L, px + '-cat', c.cats || [])}</div></details></div>
            <details><summary class="t" style="cursor:pointer">🔎 ${_ih('visor.sel.filtros')}</summary>
                <label class="t" for="${px}-lg">${_ih('visor.sel.largo')}</label><select id="${px}-lg" class="modern-select">${this.LARGOS.map(([n, k]) => `<option value="${n}" ${Number(c.largo || 0) === n ? 'selected' : ''}>${_ih(k)}</option>`).join('')}</select>
                <label class="chip" style="margin-top:8px"><input type="checkbox" id="${px}-au" ${c.audio ? 'checked' : ''}> 🔊 ${_ih('visor.sel.audio')}</label>
                <span class="t">🔡 ${_ih('visor.sel.conocidas')}</span><p class="nota">${_ih('visor.sel.conocidas.d')}</p><div class="chips" id="${px}-con">${this.letrasHtml(px, L, c.conocidas || [])}</div></details>
            <button type="button" class="btn btn-secondary btn-mini" id="${px}-mis" style="margin-top:6px">📋 ${_ih('visor.listas.titulo')}</button></fieldset>`;
    },
    letrasHtml(px, L, marcadas) { return Texto.alfabeto(L).map(a => `<label class="chip"><input type="checkbox" name="${px}-co" value="${esc(a)}" ${marcadas.includes(a) ? 'checked' : ''}> <b${Texto.lang(L)}>${esc(a)}</b></label>`).join(''); },
    enlazar(px, alCambiarLengua, volver) {
        const f = () => document.querySelector(`input[name="${px}-f"]:checked`)?.value || 'azar';
        const vis = () => { $(px + '-lis').classList.toggle('oculto', f() !== 'lista'); $(px + '-cat-b').classList.toggle('oculto', f() === 'lista'); };
        document.querySelectorAll(`input[name="${px}-f"]`).forEach(r => r.addEventListener('change', vis));
        $(px + '-mis').addEventListener('click', () => Listas.abrir(volver));
        return { recategorizar: L => { $(px + '-cats').innerHTML = Comun.categorias(L, px + '-cat', Comun.marcados(px + '-cat')); if ($(px + '-con')) $(px + '-con').innerHTML = this.letrasHtml(px, L, []); } };
    },
    leer(px) {
        return { fuente: document.querySelector(`input[name="${px}-f"]:checked`)?.value || 'azar', lista: $(px + '-li')?.value || '',
            cats: Comun.marcados(px + '-cat'), largo: Number($(px + '-lg')?.value || 0), audio: !!$(px + '-au')?.checked, conocidas: Comun.marcados(px + '-co') };
    },
    /** Palabras que cumplen la selección. Respeta siempre lo que la comunidad excluye de juegos y materiales. */
    palabras(c, L, { img = false, largas = false } = {}) {
        const permitidas = new Set(Datos.paraJuegos(L, { img, largas, audio: !!c.audio }).map(w => w.id));
        let ws;
        if (c.fuente === 'lista') { const l = Listas.de(c.lista); ws = l ? Listas.palabras(l).filter(w => permitidas.has(w.id)) : []; }
        else ws = Datos.paraJuegos(L, { img, largas, audio: !!c.audio, cats: c.cats });
        if (c.largo) ws = ws.filter(w => Texto.letras(w.t[L].texto, L).length <= c.largo);
        // Solo con las letras que el grupo ya conoce (si se marcaron)
        if (c.conocidas && c.conocidas.length) { const k = new Set(c.conocidas.map(x => x.toLowerCase())); ws = ws.filter(w => Texto.letras(w.t[L].texto, L).every(g => k.has(g.toLowerCase()))); }
        return ws;
    }
};

/* ---------- Pistas por palabra ---------- */
const Pistas = {
    MODOS: ['palabra', 'imagen', 'lengua', 'descripcion', 'propia', 'espacio', 'imagenEspacio', 'ninguna'],
    ICO: { palabra: '🔤', imagen: '🖼️', lengua: '🌐', descripcion: '📖', propia: '✍️', espacio: '＿', imagenEspacio: '🖼️＿', ninguna: '∅' },
    clave() { return 'pistas:' + Datos.base; },
    propias() { return LS.get(this.clave(), {}) || {}; },
    propia(id, L) { return this.propias()[id + '|' + L] || ''; },
    guardarPropia(id, L, texto) {
        const p = this.propias(), k = id + '|' + L, v = Texto.atajos(String(texto || ''), L).replace(/\s+/g, ' ').trim().slice(0, 600);
        if (v) p[k] = v; else delete p[k];
        LS.set(this.clave(), p);
        return v;
    },
    descripcion(w, L) { return (w && w.t[L] && w.t[L].desc) || ''; },
    /** La descripción no debe regalar la respuesta: si menciona la palabra (en la misma lengua), se cambia por ____. */
    sinRespuesta(texto, w, L) {
        const t = w && w.t[L]; if (!texto || !t) return texto;
        let r = texto;
        [...new Set([t.completo || t.texto, ...(t.formas || [])])].filter(f => f && f.length > 1).sort((a, b) => b.length - a.length).forEach(f => {
            r = r.replace(new RegExp(f.normalize('NFC').replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'giu'), '____');
        });
        return r;
    },
    /** Controles: qué pista lleva cada palabra, en qué lengua y qué usar si no hay descripción. */
    html(px, L, modos, c = {}) {
        const m = modos.includes(c.modo) ? c.modo : modos[0];
        const otras = Datos.langs.filter(l => Datos.conteo(l.id).textos > 0);
        return `<fieldset class="grupo"><legend>💡 ${_ih('visor.pista.titulo')}</legend>
            <div class="chips" id="${px}-modos">${modos.map(k => `<label class="chip"><input type="radio" name="${px}-m" value="${k}" ${k === m ? 'checked' : ''}> ${this.ICO[k]} ${_ih('visor.pista.' + k)}</label>`).join('')}</div>
            <div id="${px}-lp" class="oculto"><label class="t" for="${px}-pl">${_ih('visor.pista.enLengua')}</label><select id="${px}-pl" class="modern-select">${otras.map(l => `<option value="${esc(l.id)}" ${l.id === (c.lengua || '') ? 'selected' : ''}>${esc(l.nombre)}</option>`).join('')}</select></div>
            <div id="${px}-sin" class="oculto"><label class="t" for="${px}-ps">${_ih('visor.pista.siNoHay')}</label><select id="${px}-ps" class="modern-select">
                <option value="imagen" ${c.siNoHay !== 'espacio' ? 'selected' : ''}>🖼️ ${_ih('visor.pista.imagen')}</option><option value="espacio" ${c.siNoHay === 'espacio' ? 'selected' : ''}>＿ ${_ih('visor.pista.espacio')}</option></select></div>
            <p class="nota" id="${px}-ayuda"></p></fieldset>`;
    },
    enlazar(px, getL) {
        const ajustar = () => {
            const m = document.querySelector(`input[name="${px}-m"]:checked`)?.value || '';
            const conLengua = ['lengua', 'descripcion', 'propia'].includes(m);
            $(px + '-lp').classList.toggle('oculto', !conLengua);
            $(px + '-sin').classList.toggle('oculto', !['descripcion', 'propia'].includes(m));
            const sel = $(px + '-pl'), L = getL();
            if (conLengua && m !== 'lengua' && !sel.dataset.tocado) sel.value = L;                // descripciones: por omisión, en la misma lengua
            if (m === 'lengua' && sel.value === L && !sel.dataset.tocado) { const o = [...sel.options].find(x => x.value !== L); if (o) sel.value = o.value; }
            const n = Datos.conteo(sel.value).desc || 0;
            $(px + '-ayuda').innerHTML = m === 'descripcion' ? _ih(n ? 'visor.pista.hayDesc' : 'visor.pista.noHayDesc', { n, lengua: Datos.nombreLengua(sel.value) })
                : m === 'propia' ? _ih('visor.pista.propia.d') : m === 'espacio' || m === 'imagenEspacio' ? _ih('visor.pista.espacio.d') : '';
        };
        document.querySelectorAll(`input[name="${px}-m"]`).forEach(r => r.addEventListener('change', ajustar));
        $(px + '-pl').addEventListener('change', () => { $(px + '-pl').dataset.tocado = '1'; ajustar(); });
        ajustar();
        return { ajustar };
    },
    leer(px) {
        return { modo: document.querySelector(`input[name="${px}-m"]:checked`)?.value || 'imagen', lengua: $(px + '-pl')?.value || '', siNoHay: $(px + '-ps')?.value || 'imagen' };
    },
    /** ¿Esta palabra puede llevar esta pista? (para no poner en la hoja palabras sin lo que pide) */
    sirve(w, c, L) {
        if (['imagen', 'imagenEspacio'].includes(c.modo)) return !!w.img;
        if (c.modo === 'lengua') return !!(w.t[c.lengua] && c.lengua !== L);
        if (['descripcion', 'propia'].includes(c.modo) && c.siNoHay === 'imagen') return !!(this.descripcion(w, c.lengua) || this.propia(w.id, c.lengua) || w.img);
        return true;
    },
    necesitaImagen(c) { return c.modo === 'imagen' || c.modo === 'imagenEspacio'; },
    /** HTML de la pista en la hoja. `editable`: las pistas propias se pueden escribir en la vista previa. */
    render(w, c, L) {
        const img = () => `<img src="${esc(Datos.img(w, 600))}" alt="">`;
        const linea = '<div class="p-linea"></div>';
        if (c.modo === 'imagen') return img();
        if (c.modo === 'imagenEspacio') return `${img()}${linea}`;
        if (c.modo === 'espacio') return `<div class="p-espacio">${linea}${linea}</div>`;
        if (c.modo === 'ninguna') return '';
        if (c.modo === 'palabra') return `<div class="p-texto"${Texto.lang(L)}>${esc(w.t[L].texto)}</div>`;
        if (c.modo === 'lengua') { const t = w.t[c.lengua]; return t ? `<div class="p-texto"${Texto.lang(c.lengua)}>${esc(t.texto)}</div>` : ''; }
        const desc = c.lengua === L ? this.sinRespuesta(this.descripcion(w, c.lengua), w, L) : this.descripcion(w, c.lengua);
        if (c.modo === 'descripcion' && desc) return `<div class="p-texto p-desc"${Texto.lang(c.lengua)}>${esc(desc)}</div>`;
        if (c.modo === 'propia') {
            const propia = this.propia(w.id, c.lengua), texto = propia || desc;
            if (texto || c.siNoHay === 'espacio' || !w.img)
                return `<div class="p-texto p-desc p-ed${texto ? '' : ' vacia'}" contenteditable="true" spellcheck="false" data-pista="${esc(w.id)}" data-lengua="${esc(c.lengua)}" data-orig="${esc(desc)}" data-ph="${esc(_i('visor.pista.escribeAqui'))}"${Texto.lang(c.lengua)}>${esc(texto)}</div>${desc ? `<button type="button" class="no-imp p-orig" data-orig-de="1">↺ ${esc(_i('visor.pista.volver'))}</button>` : ''}${typeof Propuestas !== 'undefined' && Propuestas.activo() ? `<button type="button" class="no-imp p-orig" data-proponer="1" title="${esc(_i('visor.pista.proponer.d'))}">📤 ${esc(_i('visor.pista.proponer'))}</button>` : ''}`;
        }
        return c.siNoHay === 'espacio' || !w.img ? `<div class="p-espacio">${linea}${linea}</div>` : img();
    }
};

/* ---------- Ajustes de impresión, pie y ventana de vista previa ---------- */
const Impresion = {
    PAPEL: { carta: [216, 279], oficio: [216, 340], a4: [210, 297] },
    pred() { return { papel: 'carta', margen: 'seguro', tinta: 'color', letra: 'normal', mayus: 'como', creditos: true }; },
    ajustes() { return Object.assign(this.pred(), LS.get('imp:ajustes', {}) || {}); },
    html() {
        const a = this.ajustes(), op = (k, vs) => vs.map(v => `<label class="chip"><input type="radio" name="ia-${k}" value="${v}" ${a[k] === v ? 'checked' : ''}> ${_ih(`visor.imp.${k}.${v}`)}</label>`).join('');
        return `<details class="grupo-d"><summary class="t" style="cursor:pointer">⚙️ ${_ih('visor.imp.ajustes')}</summary>
            <span class="t">${_ih('visor.imp.papel')}</span><div class="chips">${op('papel', ['carta', 'oficio', 'a4'])}</div>
            <span class="t">${_ih('visor.imp.margen')}</span><div class="chips">${op('margen', ['seguro', 'minimo'])}</div>
            <span class="t">${_ih('visor.imp.tinta')}</span><div class="chips">${op('tinta', ['color', 'bn'])}</div>
            <span class="t">${_ih('visor.imp.letra')}</span><div class="chips">${op('letra', ['normal', 'grande', 'escolar'])}</div>
            <span class="t">${_ih('visor.imp.mayus')}</span><div class="chips">${op('mayus', ['como', 'minus', 'mayus'])}</div>
            <label class="chip" style="margin-top:8px"><input type="checkbox" id="ia-creditos" ${a.creditos ? 'checked' : ''}> ${_ih('visor.imp.creditos')}</label>
            <p class="nota">${_ih('visor.imp.ajustes.d')}</p></details>`;
    },
    leer() {
        const a = this.ajustes(), v = k => document.querySelector(`input[name="ia-${k}"]:checked`)?.value;
        ['papel', 'margen', 'tinta', 'letra', 'mayus'].forEach(k => { if (v(k)) a[k] = v(k); });
        if ($('ia-creditos')) a.creditos = $('ia-creditos').checked;
        LS.set('imp:ajustes', a);
        return a;
    },
    /** CSS de página según los ajustes. Las medidas del contenido usan --w y --h (área útil de la hoja). */
    css(a, orientacion = 'portrait') {
        let [W, H] = this.PAPEL[a.papel] || this.PAPEL.carta;
        if (orientacion === 'landscape') [W, H] = [H, W];
        const m = a.margen === 'minimo' ? 5 : 10;
        const esc2 = a.letra === 'grande' ? 1.25 : 1;
        return `@page{size:${W}mm ${H}mm;margin:${m}mm}
            :root{--w:${W - 2 * m}mm;--h:${H - 2 * m - 1}mm;--k:${esc2}}
            ${a.letra === 'escolar' ? `@import url('https://fonts.googleapis.com/css2?family=Andika:wght@400;700&display=swap');
            .hoja{font-family:'Andika','Comic Sans MS',system-ui,sans-serif}` : ''}
            ${a.mayus === 'mayus' || a.mayus === 'minus' ? `.hoja>:not(header):not(footer){text-transform:${a.mayus === 'mayus' ? 'uppercase' : 'lowercase'}}` : ''}
            ${a.tinta === 'bn' ? `img{filter:grayscale(1) contrast(1.1)}.hoja *{color:#000!important;border-color:#000!important}` : ''}`;
    },
    /** Pie de cada hoja: catálogo, lengua, licencia, leyenda, código de la hoja y créditos. */
    pie(a, { L, codigo, version } = {}) {
        const cfg = Datos.cfg || {}, leyenda = (Estado.e && Estado.e.leyenda) || '';
        const partes = [cfg.titulo, L ? Datos.nombreLengua(L) : '', cfg.licencia ? _i('visor.imp.licencia', { x: cfg.licencia }) : '', leyenda].filter(Boolean).map(esc);
        const cod = codigo ? `<b class="codigo">${esc(_i('visor.imp.codigo', { c: codigo + (version ? '-' + version : '') }))}</b>` : '';
        const cred = a.creditos ? `<span class="cred">${esc(_i('visor.pie.credito', { autor: AUTOR }))} · AGPL-3.0</span>` : '';
        return `<footer><span>${partes.join(' · ')}</span>${cod}${cred}</footer>`;
    },
    /** Abre la vista previa lista para imprimir. Las pistas propias se pueden escribir ahí mismo. */
    ventana({ titulo, paginas, desc, css = '', ajustes, orientacion = 'portrait', aviso = '' }) {
        const v = VistaImpresion.abrir();
        if (!v) { msg(_ih('visor.imp.bloqueada')); return null; }
        const nHojas = (paginas.match(/<section class="hoja/g) || []).length;
        const hayEditables = /data-pista=/.test(paginas);
        v.document.write(`<!DOCTYPE html><html lang="${esc(document.documentElement.lang || 'es')}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(titulo)}</title>
            <style>${this.css(ajustes, orientacion)}${Impresion.CSS}${css}</style></head><body>
            <div class="barra"><b>${esc(desc)}</b><span>${esc(_i('visor.imp.nHojas', { n: nHojas }))} · ${esc(_i('visor.imp.papel.' + ajustes.papel))}</span>
            <button onclick="print()">🖨️ ${_ih('visor.imp.boton')}</button>
            ${hayEditables ? `<span class="ayuda-ed">✍️ ${_ih('visor.imp.editables')}</span>` : ''}${aviso ? `<span class="aviso">⚠️ ${esc(aviso)}</span>` : ''}</div>${paginas}
            <script>(function(){
              var op=window.opener||(window.parent!==window?window.parent:null);
              document.addEventListener('focusout',function(e){var el=e.target;if(!el||!el.dataset||!el.dataset.pista||!op||!op.Pistas)return;
                var v=op.Pistas.guardarPropia(el.dataset.pista,el.dataset.lengua,el.innerText);el.innerText=v;el.classList.toggle('vacia',!v);
                [].slice.call(document.querySelectorAll('[data-pista="'+el.dataset.pista+'"][data-lengua="'+el.dataset.lengua+'"]')).forEach(function(o){if(o!==el){o.innerText=v;o.classList.toggle('vacia',!v);}});});
              document.addEventListener('keydown',function(e){if(e.key==='Enter'&&e.target&&e.target.dataset&&e.target.dataset.pista){e.preventDefault();e.target.blur();}});
              document.addEventListener('click',function(e){var pp=e.target.closest&&e.target.closest('[data-proponer]');if(pp&&op&&op.Propuestas){var ed=pp.parentNode.querySelector('[data-pista]');if(ed){op.Propuestas.abrirDescripcion(ed.dataset.pista,ed.dataset.lengua,ed.innerText);try{op.focus();}catch(x){}}return;}});
              document.addEventListener('click',function(e){var b=e.target.closest&&e.target.closest('[data-orig-de]');if(!b||!op||!op.Pistas)return;var el=b.previousElementSibling;if(!el)return;
                op.Pistas.guardarPropia(el.dataset.pista,el.dataset.lengua,'');el.innerText=el.dataset.orig;el.classList.toggle('vacia',!el.dataset.orig);});
            })();<\/script>
            ${hayEditables ? '' : `<script>(function(){var h=false;function ir(){if(h)return;h=true;setTimeout(function(){window.print()},300)}var im=[].slice.call(document.images),p=im.length;if(!p)return ir();im.forEach(function(i){if(i.complete){if(--p===0)ir()}else{i.onload=i.onerror=function(){if(--p===0)ir()}}});setTimeout(ir,15000)})();<\/script>`}
            </body></html>`);
        v.document.close();
        msg(_ih('visor.imp.abierta', { x: desc }));
        return v;
    },
    CSS: `html,body{margin:0;background:#fff;color:#111;font-family:system-ui,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
        .barra{position:sticky;top:0;background:#1e3a5f;color:#fff;padding:10px 14px;font-size:14px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;z-index:5}
        .barra button{font-size:15px;font-weight:700;padding:8px 14px;border:0;border-radius:8px;cursor:pointer}
        .barra .ayuda-ed{background:#fef9c3;color:#713f12;border-radius:6px;padding:3px 8px}.barra .aviso{background:#fee2e2;color:#7f1d1d;border-radius:6px;padding:3px 8px}
        @media print{.barra,.no-imp{display:none!important}.p-ed{outline:none!important;background:none!important}.p-ed.vacia{border-bottom:.4mm solid #111}}
        @media screen{.hoja{margin:10px auto;box-shadow:0 0 0 1px #ccc,0 4px 14px rgba(0,0,0,.15)}}
        .hoja{width:var(--w);height:var(--h);box-sizing:border-box;padding:4mm;display:flex;flex-direction:column;gap:3mm;overflow:hidden;break-after:page;page-break-after:always}
        .hoja:last-of-type{break-after:auto;page-break-after:auto}
        header{display:flex;justify-content:space-between;gap:6mm;font-size:calc(13pt*var(--k));border-bottom:.4mm solid #111;padding-bottom:2mm}
        header .datos{white-space:nowrap;color:#444}header .ver{font-weight:800;white-space:nowrap}
        footer{font-size:7pt;color:#555;display:flex;justify-content:space-between;gap:4mm;flex-wrap:wrap;border-top:.2mm solid #bbb;padding-top:1mm}
        footer .codigo{color:#111}footer .cred{color:#777}
        img{max-width:100%;max-height:100%;object-fit:contain}
        .p-texto{font-size:calc(14pt*var(--k));text-align:center;line-height:1.3}.p-desc{font-size:calc(13pt*var(--k));text-align:left}
        .p-ed{outline:1.5px dashed #2563eb;outline-offset:1px;background:#eff6ff;min-height:1.4em;border-radius:1mm;cursor:text}
        .p-ed.vacia::before{content:attr(data-ph);color:#94a3b8}
        .p-linea{width:100%;border-bottom:.4mm solid #111;height:8mm}.p-espacio{width:100%;display:flex;flex-direction:column;gap:1mm}
        .resp-ok{color:#b91c1c;font-weight:800}
        .p-orig{font-size:9pt;border:1px solid #94a3b8;background:#fff;border-radius:4px;padding:1px 6px;cursor:pointer;margin-top:1mm}`
};

/* ---------- Mis actividades: lo preparado, guardado en este dispositivo ---------- */
const FORMATO_ACTIVIDADES = 'catalogo-lenguas-actividades';
const Actividades = {
    MAX_RECIENTES: 25, TIPOS: ['hojas', 'impresos'],
    nombreDe(a) { return a.tipo === 'impresos' ? Impresos.nombre(a.sub) : Hojas.nombre(a.sub); },
    existe(a) { return a.tipo === 'impresos' ? !!Impresos.TIPOS[a.sub] : !!Hojas.TIPOS[a.sub]; },
    clave() { return 'actividades:' + Datos.base; },
    todas() { return (LS.get(this.clave(), []) || []).filter(a => a && a.id && a.tipo); },
    guardarTodas(l) { const ok = LS.set(this.clave(), l); if (!ok) UI.toast(_i('visor.act.lleno'), 6000); return ok; },
    nuevoCodigo() { const A = 'ACDEFGHJKMNPQRTUVWXY3479'; return Array.from({ length: 4 }, () => A[azarSeguro(A.length)]).join(''); },
    /** Cada material generado queda como reciente; «guardar» lo conserva aunque se generen muchos más. */
    registrar({ tipo, sub, config, codigo, nombre }) {
        const t = this.todas(), ahora = new Date().toISOString();
        let a = t.find(x => x.codigo === codigo && x.tipo === tipo && x.sub === sub);
        if (a) { a.config = config; a.usada = ahora; }
        else { a = { id: 'a' + Date.now().toString(36) + azarSeguro(1296).toString(36), tipo, sub, config, codigo, nombre, creada: ahora, usada: ahora, fija: false }; t.unshift(a); }
        const fijas = t.filter(x => x.fija), rec = t.filter(x => !x.fija).sort((x, y) => String(y.usada).localeCompare(String(x.usada))).slice(0, this.MAX_RECIENTES);
        this.guardarTodas([...fijas, ...rec]);
        return a;
    },
    actualizar(id, cambios) { const t = this.todas(), a = t.find(x => x.id === id); if (a) { Object.assign(a, cambios); this.guardarTodas(t); } return a; },
    abrirActividad(a) { if (a.tipo === 'hojas') Hojas.abrir(a.sub, a.config, a.codigo); else if (a.tipo === 'impresos') Impresos.abrir(a.sub, a.config, a.codigo); },
    palabrasDe(a) { const c = a.config || {}; return c.L && c.sel ? Seleccion.palabras(c.sel, c.L) : []; },
    respaldo() { return LS.get('actividades:respaldo:' + Datos.base, ''); },
    /** Pantalla «📂 Mis actividades». */
    abrir() {
        const t = this.todas().sort((x, y) => Number(y.fija) - Number(x.fija) || String(y.usada).localeCompare(String(x.usada)));
        const resp = this.respaldo(), dias = resp ? Math.floor((Date.now() - new Date(resp).getTime()) / 864e5) : null;
        const fila = a => `<div class="caja act-fila"><div><b>${a.fija ? '⭐ ' : ''}${esc(a.nombre)}</b>
            <div class="nota">${esc(this.nombreDe(a))} · ${esc(_i('visor.imp.codigo', { c: a.codigo }))} · ${esc(new Date(a.usada).toLocaleDateString())}</div></div>
            <div class="fila"><button class="btn btn-primary btn-mini" data-ab="${esc(a.id)}">🖨️ ${_ih('visor.act.abrir')}</button>
            <button class="btn btn-secondary btn-mini" data-fi="${esc(a.id)}">${a.fija ? '☆ ' + _ih('visor.act.quitarGuardada') : '⭐ ' + _ih('visor.act.guardar')}</button>
            <button class="btn btn-secondary btn-mini" data-re="${esc(a.id)}">✏️ ${_ih('visor.act.renombrar')}</button>
            <button class="btn btn-secondary btn-mini" data-du="${esc(a.id)}">⧉ ${_ih('visor.act.duplicar')}</button>
            <button class="btn btn-secondary btn-mini" data-ar="${esc(a.id)}">💾</button>
            <button class="btn btn-secondary btn-mini" data-en="${esc(a.id)}">🔗</button>
            <button class="btn btn-secondary btn-mini" data-of="${esc(a.id)}">📥</button>
            <button class="btn btn-secondary btn-mini" data-bo="${esc(a.id)}" aria-label="${esc(_i('visor.borrar'))}">🗑️</button></div></div>`;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📂 ${_ih('visor.act.titulo')}</h2>
            <div class="aviso-local" role="note">${Instalar.instalada() ? `<b>⚠️ ${_ih('visor.act.avisoApp.t')}</b>
                <p>${_ih('visor.act.aviso.1')}</p><ul><li>${_ih('visor.act.avisoApp.c1')}</li><li>${_ih('visor.act.avisoApp.c2')}</li><li>${_ih('visor.act.avisoApp.c3')}</li>${Instalar.esIOS ? `<li>${_ih('visor.act.avisoApp.ios')}</li>` : ''}</ul>` : `<b>⚠️ ${_ih('visor.act.aviso.t')}</b>
                <p>${_ih('visor.act.aviso.1')}</p><ul><li>${_ih('visor.act.aviso.c1')}</li><li>${_ih('visor.act.aviso.c2')}</li><li>${_ih('visor.act.aviso.c3')}</li><li>${_ih('visor.act.aviso.c4')}</li></ul>
                <p class="nota">${_ih('visor.act.aviso.instalar')}</p>`}
                <p>${_ih('visor.act.aviso.resp')}</p><p>${_ih('visor.act.aviso.recup')}</p>
                <p class="nota">${resp ? _ih('visor.act.ultimoRespaldo', { n: dias }) : _ih('visor.act.sinRespaldo')}</p></div>
            ${t.length && (dias === null || dias > 30) && t.some(a => a.fija) ? `<div class="error-box">💾 ${_ih('visor.act.recordar')}</div>` : ''}
            <div class="fila" style="margin:10px 0"><button class="btn btn-primary" id="ma-gtodo">💾 ${_ih('visor.act.guardarArchivo')}</button>
            <button class="btn btn-secondary" id="ma-abrir">📂 ${_ih('visor.act.abrirArchivo')}</button><input type="file" id="ma-file" accept=".json,application/json" class="oculto">
            <button class="btn btn-secondary" id="ma-listas">📋 ${_ih('visor.listas.titulo')}</button></div>
            <div id="ma-msg"></div>
            ${t.length ? `<h3>${_ih('visor.act.lista')}</h3>${t.map(fila).join('')}<p class="nota">${_ih('visor.act.recientes', { n: this.MAX_RECIENTES })}</p>` : `<div class="caja">${_ih('visor.act.vacia')}</div>`}
            <details class="caja"><summary>🗑️ ${_ih('visor.act.borrarDisp')}</summary><p class="nota">${_ih('visor.act.borrarDisp.d')}</p>
                <button class="btn btn-secondary" id="ma-borrar-off">${_ih('visor.act.borrarOff')}</button></details></div>`);
        const B = $('modal-body'), de = id => this.todas().find(a => a.id === id);
        $('ma-gtodo').addEventListener('click', () => this.exportar());
        $('ma-abrir').addEventListener('click', () => $('ma-file').click());
        $('ma-file').addEventListener('change', e => this.importar(e.target.files[0]));
        $('ma-listas').addEventListener('click', () => Listas.abrir(() => this.abrir()));
        $('ma-borrar-off').addEventListener('click', async () => { if (confirm(_i('visor.act.borrarOff.c'))) { await Offline.borrarTodo(); UI.toast(_i('visor.listo')); } });
        B.querySelectorAll('[data-ab]').forEach(b => b.addEventListener('click', () => this.abrirActividad(de(b.dataset.ab))));
        B.querySelectorAll('[data-fi]').forEach(b => b.addEventListener('click', () => { const a = de(b.dataset.fi); this.actualizar(a.id, { fija: !a.fija }); this.abrir(); }));
        B.querySelectorAll('[data-re]').forEach(b => b.addEventListener('click', () => { const a = de(b.dataset.re), n = prompt(_i('visor.act.nuevoNombre'), a.nombre); if (n && n.trim()) { this.actualizar(a.id, { nombre: n.trim().slice(0, 120), fija: true }); this.abrir(); } }));
        B.querySelectorAll('[data-du]').forEach(b => b.addEventListener('click', () => { const a = de(b.dataset.du);
            const t2 = this.todas(); t2.unshift(Object.assign(JSON.parse(JSON.stringify(a)), { id: 'a' + Date.now().toString(36) + azarSeguro(1296).toString(36), codigo: this.nuevoCodigo(), nombre: _i('visor.act.copiaDe', { x: a.nombre }), fija: true, creada: new Date().toISOString(), usada: new Date().toISOString() }));
            this.guardarTodas(t2); this.abrir(); }));
        B.querySelectorAll('[data-ar]').forEach(b => b.addEventListener('click', () => this.exportar([b.dataset.ar])));
        B.querySelectorAll('[data-en]').forEach(b => b.addEventListener('click', () => this.compartirEnlace(de(b.dataset.en))));
        B.querySelectorAll('[data-of]').forEach(b => b.addEventListener('click', () => { const a = de(b.dataset.of); Offline.panel(this.palabrasDe(a), a.nombre, () => this.abrir()); }));
        B.querySelectorAll('[data-bo]').forEach(b => b.addEventListener('click', () => { if (confirm(_i('visor.act.borrar.c'))) { this.guardarTodas(this.todas().filter(a => a.id !== b.dataset.bo)); this.abrir(); } }));
    },
    /** Lo que viaja en un archivo o enlace: actividades, las listas que usan y sus pistas propias. */
    paquete(ids) {
        const acts = this.todas().filter(a => !ids || ids.includes(a.id));
        const listasUsadas = new Set(acts.map(a => a.config && a.config.sel && a.config.sel.lista).filter(Boolean));
        const listas = ids ? Listas.todas().filter(l => listasUsadas.has(l.id)) : Listas.todas();
        const propias = Pistas.propias(), pistas = {};
        if (ids) { const ws = new Set(acts.flatMap(a => this.palabrasDe(a).map(w => w.id))); Object.keys(propias).forEach(k => { if (ws.has(k.split('|')[0])) pistas[k] = propias[k]; }); }
        else Object.assign(pistas, propias);
        return { formato: FORMATO_ACTIVIDADES, version: 1, creado: new Date().toISOString(),
            catalogo: { id: Datos.cfg.idCatalogo || '', titulo: Datos.cfg.titulo || '', base: Datos.base }, actividades: acts, listas, pistas };
    },
    exportar(ids) {
        const p = this.paquete(ids);
        if (!p.actividades.length && !p.listas.length && !Object.keys(p.pistas).length) return UI.toast(_i('visor.act.nadaQueGuardar'));
        const nombre = ids && p.actividades.length === 1 ? p.actividades[0].nombre : (Datos.cfg.titulo || 'catalogo');
        descargar(`actividades-${String(nombre).replace(/[^\p{L}\p{N}]+/gu, '-').toLowerCase().slice(0, 40)}-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(p, null, 1), 'application/json');
        if (!ids) LS.set('actividades:respaldo:' + Datos.base, new Date().toISOString());
        UI.toast(_i('visor.act.descargado'), 4000);
        if (!ids) setTimeout(() => { if ($('ma-msg')) this.abrir(); }, 400);
    },
    /** Revisa un archivo antes de usarlo: formato, tamaño, tipos y largos. Lo escrito siempre se trata como texto. */
    validar(o) {
        const txt = (v, max) => typeof v === 'string' ? v.slice(0, max) : '';
        const limpio = (v, d = 0) => {   // copia solo datos simples y con tope
            if (d > 5) return null;
            if (v == null || typeof v === 'boolean' || (typeof v === 'number' && isFinite(v))) return v;
            if (typeof v === 'string') return v.slice(0, 2000);
            if (Array.isArray(v)) return v.slice(0, 3000).map(x => limpio(x, d + 1));
            if (typeof v === 'object') { const r = {}; Object.keys(v).slice(0, 60).forEach(k => { if (/^[\w-]{1,40}$/.test(k)) r[k] = limpio(v[k], d + 1); }); return r; }
            return null;
        };
        if (!o || o.formato !== FORMATO_ACTIVIDADES || typeof o !== 'object') throw new Error(_i('visor.act.error.formato'));
        const actividades = (Array.isArray(o.actividades) ? o.actividades : []).slice(0, 500).filter(a => a && this.TIPOS.includes(a.tipo) && typeof a.sub === 'string' && this.existe(a))
            .map(a => ({ id: txt(a.id, 40).replace(/[^\w-]/g, '') || 'a' + azarSeguro(1e9).toString(36), tipo: a.tipo, sub: a.sub, nombre: txt(a.nombre, 120) || this.nombreDe(a),
                codigo: /^[A-Z0-9]{4}$/.test(a.codigo) ? a.codigo : this.nuevoCodigo(), config: limpio(a.config) || {}, creada: txt(a.creada, 30), usada: txt(a.usada, 30) || new Date().toISOString(), fija: true }));
        const listas = (Array.isArray(o.listas) ? o.listas : []).slice(0, 200).filter(l => l && Array.isArray(l.ids))
            .map(l => ({ id: txt(l.id, 40).replace(/[^\w-]/g, '') || 'l' + azarSeguro(1e9).toString(36), nombre: txt(l.nombre, 80) || '—', ids: l.ids.filter(x => typeof x === 'string').slice(0, 3000).map(x => x.slice(0, 80)), creada: txt(l.creada, 30), modificada: txt(l.modificada, 30) }));
        const pistas = {};
        if (o.pistas && typeof o.pistas === 'object') Object.keys(o.pistas).slice(0, 5000).forEach(k => { if (k.length <= 200 && typeof o.pistas[k] === 'string') pistas[k] = o.pistas[k].slice(0, 600); });
        const cat = o.catalogo && typeof o.catalogo === 'object' ? { id: txt(o.catalogo.id, 80), titulo: txt(o.catalogo.titulo, 200), base: txt(o.catalogo.base, 300) } : {};
        return { actividades, listas, pistas, catalogo: cat };
    },
    importar(archivo) {
        const m = h => { if ($('ma-msg')) $('ma-msg').innerHTML = h; };
        if (!archivo) return;
        if (archivo.size > 1.5 * 1048576) return m(`<div class="error-box">${_ih('visor.act.error.grande')}</div>`);
        const fr = new FileReader();
        fr.onload = () => {
            let p;
            try { p = this.validar(JSON.parse(String(fr.result))); }
            catch (e) { return m(`<div class="error-box">${esc(e.message && e.message.startsWith('visor') ? e.message : _i('visor.act.error.formato'))}</div>`); }
            const mio = Datos.cfg.idCatalogo || '';
            if (p.catalogo.id && mio && p.catalogo.id !== mio) {
                const ir = /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(p.catalogo.base) ? `<button class="btn btn-secondary" id="ma-ir">🔌 ${_ih('visor.act.irCatalogo')}</button>` : '';
                m(`<div class="error-box">${_ih('visor.act.otroCatalogo', { x: p.catalogo.titulo || '?' })}<div class="fila" style="margin-top:6px">${ir}</div></div>`);
                if (ir) $('ma-ir').addEventListener('click', () => { location.href = location.pathname + '?api=' + encodeURIComponent(p.catalogo.base); });
                return;
            }
            const n = this.fusionar(p);
            this.abrir();
            m(`<div class="caja">✅ ${_ih('visor.act.importado', n)}${n.faltan ? '<br>⚠️ ' + _ih('visor.act.faltanPalabras', { n: n.faltan }) : ''}</div>`);
        };
        fr.onerror = () => m(`<div class="error-box">${_ih('visor.act.error.formato')}</div>`);
        fr.readAsText(archivo);
    },
    /** Junta lo del archivo con lo de este dispositivo sin perder nada: lo repetido se conserva una vez. */
    fusionar(p) {
        const t = this.todas(), ids = new Set(t.map(a => a.id)), codigos = new Set(t.map(a => a.tipo + a.sub + a.codigo));
        let act = 0, lis = 0, pis = 0;
        p.actividades.forEach(a => { if (ids.has(a.id) || codigos.has(a.tipo + a.sub + a.codigo)) return; t.unshift(a); act++; });
        this.guardarTodas(t);
        const ls = Listas.todas(), lids = new Set(ls.map(l => l.id));
        p.listas.forEach(l => { if (!lids.has(l.id)) { ls.push(l); lis++; } });
        Listas.guardarTodas(ls);
        const pr = Pistas.propias();
        Object.keys(p.pistas).forEach(k => { if (!pr[k]) { pr[k] = p.pistas[k]; pis++; } });
        LS.set(Pistas.clave(), pr);
        const faltan = new Set(p.listas.flatMap(l => l.ids).filter(id => !Datos.byId.has(id))).size;
        return { act, lis, pis, faltan };
    },
    /** Enlace para compartir una actividad corta (sin archivo). Con muchas pistas propias, mejor el archivo. */
    compartirEnlace(a) {
        const p = this.paquete([a.id]);
        const b64 = btoa(String.fromCharCode(...new TextEncoder().encode(JSON.stringify(p)))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
        if (b64.length > 1800) return UI.toast(_i('visor.act.enlaceLargo'), 6000);
        const u = new URL(location.origin + location.pathname);
        if (Datos.base !== (API_URL || '').split('?')[0]) u.searchParams.set('api', Datos.base);
        const url = u.toString() + '#act=' + b64;
        (navigator.share ? navigator.share({ title: a.nombre, url }).catch(() => {}) : Promise.reject()).catch(() => {});
        if (!navigator.share) { try { navigator.clipboard.writeText(url).then(() => UI.toast(_i('visor.enlaceCopiado'))); } catch (e) { prompt(_i('visor.copiaEnlace'), url); } }
    },
    /** Al abrir un enlace con #act=…: ofrece agregar la actividad a Mis actividades. */
    revisarEnlace() {
        const h = location.hash || '';
        if (!h.startsWith('#act=')) return;
        history.replaceState(null, '', location.pathname + location.search);
        let p;
        try { const s = h.slice(5).replace(/-/g, '+').replace(/_/g, '/'); p = this.validar(JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(s), c => c.charCodeAt(0))))); }
        catch (e) { return UI.toast(_i('visor.act.error.formato'), 5000); }
        if (p.catalogo.id && Datos.cfg.idCatalogo && p.catalogo.id !== Datos.cfg.idCatalogo) return UI.toast(_i('visor.act.otroCatalogo', { x: p.catalogo.titulo || '?' }), 6000);
        const a = p.actividades[0];
        if (!a) return;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📂 ${_ih('visor.act.recibida')}</h2><div class="caja"><b>${esc(a.nombre)}</b><div class="nota">${esc(this.nombreDe(a))}</div></div>
            <div class="acciones"><button class="btn btn-primary" id="ae-si">💾 ${_ih('visor.act.agregar')}</button><button class="btn btn-secondary" onclick="UI.cerrarModal()">${_ih('visor.cancelar')}</button></div></div>`);
        $('ae-si').addEventListener('click', () => { this.fusionar(p); UI.ir('imprimir'); this.abrir(); });
    }
};

// La vista previa de impresión (otra ventana) guarda las pistas propias a través de este puente
window.Pistas = Pistas;

/* ---------- ¿Qué es? (leer la descripción y elegir) ---------- */
const QueEs = {
    abrir() {
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📖 ${_ih('visor.act.quees')}</h2>
            <p class="nota">${_ih('visor.qe.d')}</p>
            <label class="t" for="q-l">${_ih('visor.lengua')}</label><select id="q-l" class="modern-select">${Comun.lenguas({ campo: 'desc', min: 4 })}</select>
            <div class="acciones"><button class="btn btn-primary" onclick="QueEs.iniciar()">🚀 ${_ih('visor.empezar')}</button></div></div>`);
    },
    iniciar() {
        const L = $('q-l').value;
        const conDesc = Datos.paraJuegos(L).filter(w => w.t[L].desc);
        const conImg = conDesc.filter(w => w.img);
        const imagenes = conImg.length >= 4;
        this.s = { L, imagenes, pool: imagenes ? conImg : conDesc };
        this.s.orden = barajar(this.s.pool).slice(0, 10);
        this.s.r = Rondas.nuevo('📖 ' + _i('visor.act.quees'), L, this.s.orden.length);
        Juego.activo = { detener() {} };
        this.ronda();
    },
    ronda() {
        const s = this.s, w = s.orden[s.r.n];
        if (!w) return Rondas.final(s.r, () => this.abrir());
        const alternativas = s.imagenes ? Datos.paraJuegos(s.L, { img: true }) : Datos.paraJuegos(s.L);
        const otras = barajar(alternativas.filter(x => x !== w && (!s.imagenes || x.img !== w.img))).slice(0, 3);
        const ops = barajar([w, ...otras]);
        UI.cuerpo(`${Rondas.cabeza(s.r)}
            <div class="caja qe-desc"${Texto.lang(s.L)}>${esc(Pistas.sinRespuesta(w.t[s.L].desc, w, s.L))}</div>
            <div class="${s.imagenes ? 'opciones-img' : 'opciones-txt'}" id="q-ops">${ops.map(o => s.imagenes
                ? `<button class="opcion-img" data-id="${esc(o.id)}" aria-label="${esc(_i('visor.opcion'))}"><img src="${esc(Datos.img(o, 400))}" alt=""></button>`
                : `<button class="btn btn-secondary opcion-txt" data-id="${esc(o.id)}"${Texto.lang(s.L)}>${esc(o.t[s.L].texto)}</button>`).join('')}</div>
            <div id="q-res"></div>`);
        $('q-ops').addEventListener('click', e => {
            const b = e.target.closest('[data-id]'); if (!b || s.resp) return;
            s.resp = true;
            const ok = b.dataset.id === w.id;
            b.classList.add(ok ? 'bien' : 'mal');
            $('q-ops').querySelector(`[data-id="${CSS.escape(w.id)}"]`).classList.add('bien');
            ok ? (SoundFX.ok(), s.r.bien++) : SoundFX.error();
            s.r.jugadas.push({ id: w.id, ok }); s.r.n++;
            $('q-res').innerHTML = `<div class="round-result ${ok ? 'win' : 'fail'}">${ok ? '✅' : '❌'} <span${Texto.lang(s.L)}>${esc(w.t[s.L].texto)}</span> ${Sonido.boton(w, s.L)}</div>
                <div class="acciones"><button class="btn btn-primary btn-block" id="q-sig">➡️ ${_ih('visor.siguiente')}</button></div>`;
            $('q-sig').addEventListener('click', () => { s.resp = false; this.ronda(); });
        });
    }
};
</script>
