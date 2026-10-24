<script>
/* =============================================
   COPIA DEL CATÁLOGO, PERFILES, PROGRESO E INSTALACIÓN (versión 1.2)
   Reglas: todo queda en este aparato; nada se envía solo; sin nombres reales (apodos);
   compartir es siempre una decisión de la persona; cada catálogo guarda lo suyo por separado.
   ============================================= */

/** Copia del catálogo para usarlo sin conexión: en IndexedDB (cabe mucho más que en localStorage). */
const Copia = {
    bases: new Set(),
    clave: base => `${base}|cat|datos`,
    hay(base) { return this.bases.has(base) || !!LS.get('cat:' + base); },
    async leer(base) {
        if (!base) return null;
        const d = await Offline.tx('readonly', s => s.get(this.clave(base)));
        if (d && d.data) { this.bases.add(base); return d; }
        // Versiones anteriores la guardaban en localStorage: se pasa a IndexedDB y se libera el espacio
        const vieja = LS.get('cat:' + base);
        if (vieja && vieja.data) { if (await this.guardar(base, vieja)) LS.del('cat:' + base); return vieja; }
        return null;
    },
    async guardar(base, obj) {
        const ok = await Offline.tx('readwrite', s => s.put(obj, this.clave(base)));
        if (ok === null) return LS.set('cat:' + base, obj);   // sin IndexedDB (p. ej. ventana privada): como antes
        this.bases.add(base); LS.del('cat:' + base); Instalar.persistir();
        return true;
    },
    borrar(base) { this.bases.delete(base); LS.del('cat:' + base); Offline.tx('readwrite', s => s.delete(this.clave(base))); }
};

/* ---------- Perfiles locales: apodo y dibujo, nunca el nombre real ---------- */
const EMOJIS_PERFIL = ['🦊', '🐢', '🦉', '🐝', '🌽', '🌻', '🐟', '🦋', '🌙', '⭐', '🍄', '🐇', '🦌', '🐦', '🌵', '🔥', '🌊', '🐸', '🦎', '🌈'];
const Perfiles = {
    lista() { const l = LS.get('perfiles', []); return Array.isArray(l) ? l.filter(p => p && p.id) : []; },
    actual() { const a = LS.get('perfil_actual', 'p0') || 'p0'; return a === 'p0' || this.lista().some(p => p.id === a) ? a : 'p0'; },
    datos(id = this.actual()) { return id === 'p0' ? { id: 'p0', apodo: _i('visor.prog.esteAparato'), emoji: '📱' } : (this.lista().find(p => p.id === id) || this.datos('p0')); },
    sufijo() { const a = this.actual(); return a === 'p0' ? '' : ':' + a; },
    etiqueta(id) { const p = this.datos(id); return `${p.emoji} ${p.apodo}`; },
    elegir(id) { LS.set('perfil_actual', id); Progreso.pintarBoton(); UI.toast(_i('visor.prog.ahoraEres', { perfil: this.etiqueta(id) })); },
    crear(apodo, emoji) {
        const l = this.lista(), p = { id: 'u' + Date.now().toString(36), apodo: apodo.slice(0, 20), emoji, creado: Date.now() };
        l.push(p); LS.set('perfiles', l.slice(0, 40)); return p;
    },
    renombrar(id, apodo, emoji) { LS.set('perfiles', this.lista().map(p => p.id === id ? Object.assign(p, { apodo: apodo.slice(0, 20), emoji }) : p)); },
    borrar(id) {
        // Se quita el perfil y todo su progreso en este aparato (de todos los catálogos)
        LS.set('perfiles', this.lista().filter(p => p.id !== id));
        try { Object.keys(localStorage).filter(k => (k.startsWith('prog:') || k.startsWith('repaso:')) && k.endsWith(':' + id)).forEach(k => LS.del(k)); } catch (e) {}
        if (LS.get('perfil_actual') === id) LS.set('perfil_actual', 'p0');
        Progreso.pintarBoton();
    }
};

/* ---------- Progreso: palabras vistas y escuchadas, favoritas, partidas ---------- */
const Progreso = {
    MAX_PARTIDAS: 300, DIAS: 120,
    clave(perfil = Perfiles.actual(), base = Datos.base) { return `prog:${base}:${perfil}`; },
    leer(perfil) { return Object.assign({ vistas: {}, escuchadas: {}, favoritas: [], juegos: [], dias: {} }, LS.get(this.clave(perfil), null) || {}); },
    guardar(d) {
        if (LS.set(this.clave(), d)) { Instalar.persistir(); return; }
        // Sin espacio: se recortan las partidas más viejas y se intenta de nuevo
        d.juegos = d.juegos.slice(-60); LS.set(this.clave(), d);
    },
    hoy() { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; },
    cambiar(fn) {
        if (!Datos.base) return;
        const d = this.leer(); fn(d);
        const h = this.hoy(); d.dias[h] = (d.dias[h] || 0) + 1;
        const ks = Object.keys(d.dias).sort(); ks.slice(0, Math.max(0, ks.length - this.DIAS)).forEach(k => delete d.dias[k]);
        this.guardar(d);
    },
    vista(id) { if (id) this.cambiar(d => { d.vistas[id] = (d.vistas[id] || 0) + 1; }); },
    _audios: null, _audiosBase: '',
    escuchada(archivo) {
        if (this._audiosBase !== Datos.base) { this._audios = new Map(); this._audiosBase = Datos.base; Datos.words.forEach(w => Object.values(w.t || {}).forEach(t => t && t.audio && this._audios.set(t.audio, w.id))); }
        const id = this._audios.get(archivo); if (id) this.cambiar(d => { d.escuchadas[id] = (d.escuchadas[id] || 0) + 1; });
    },
    esFavorita(id) { return this.leer().favoritas.includes(id); },
    alternarFavorita(id) {
        let ahora = false;
        this.cambiar(d => { const i = d.favoritas.indexOf(id); if (i >= 0) d.favoritas.splice(i, 1); else { d.favoritas.push(id); ahora = true; } });
        return ahora;
    },
    juego(j) {
        if (!j || (!j.total && !j.grupo)) return;
        this.cambiar(d => { d.juegos.push(Object.assign({ f: Date.now() }, j)); if (d.juegos.length > this.MAX_PARTIDAS) d.juegos = d.juegos.slice(-this.MAX_PARTIDAS); });
    },
    /** Palabras «aprendidas» en el repaso con tarjetas (cajas 4 y 5), de todas las lenguas, de este perfil. */
    aprendidas(perfil = Perfiles.actual()) {
        let n = 0; const suf = perfil === 'p0' ? '' : ':' + perfil;
        Datos.langs.forEach(l => { const est = LS.get(`repaso:${Datos.base}:${l.id}${suf}`, {}) || {}; n += Object.values(est).filter(e => e && e.c >= 4).length; });
        return n;
    },
    resumen(perfil = Perfiles.actual()) {
        const d = this.leer(perfil), ahora = Date.now(), sem = ahora - 7 * 864e5, mes = ahora - 30 * 864e5;
        const mesJ = d.juegos.filter(j => j.f >= mes && j.total), bien = mesJ.reduce((a, j) => a + (j.bien || 0), 0), tot = mesJ.reduce((a, j) => a + (j.total || 0), 0);
        const dias7 = Object.keys(d.dias).filter(k => new Date(k + 'T12:00').getTime() >= sem).length;
        const ult = Math.max(0, ...d.juegos.map(j => j.f), ...Object.keys(d.dias).map(k => new Date(k + 'T12:00').getTime()));
        return { vistas: Object.keys(d.vistas).length, escuchadas: Object.keys(d.escuchadas).length, favoritas: d.favoritas.filter(id => Datos.byId.has(id)).length,
            aprendidas: this.aprendidas(perfil), partidas: d.juegos.length, partidas7: d.juegos.filter(j => j.f >= sem).length,
            aciertos: tot ? Math.round(100 * bien / tot) : null, dias7, ultimo: ult || null, d };
    },

    /* ---------- Pantalla «Mi progreso» ---------- */
    pintarBoton() { const b = $('btn-perfil'); if (b) { const p = Perfiles.datos(); b.textContent = p.emoji; b.title = _i('visor.prog.boton', { perfil: Perfiles.etiqueta() }); b.setAttribute('aria-label', b.title); } },
    abrir() {
        if (!Datos.cargado) return UI.toast(_i('visor.prog.sinCatalogo'));
        const yo = Perfiles.actual(), r = this.resumen(), perfiles = [Perfiles.datos('p0'), ...Perfiles.lista()];
        const tarjeta = (ico, n, k) => `<div class="prog-dato"><span class="ico" aria-hidden="true">${ico}</span><b>${n}</b><span>${_ih(k)}</span></div>`;
        const fecha = f => new Date(f).toLocaleDateString(document.documentElement.lang || undefined, { day: 'numeric', month: 'short' });
        const favs = r.d.favoritas.map(id => Datos.byId.get(id)).filter(Boolean);
        UI.abrir(`<div class="form">
            <h2 class="modal-titulo">📈 ${_ih('visor.prog.titulo')}</h2>
            <p class="nota">🔒 ${_ih('visor.prog.privado')}</p>
            <span class="t">${_ih('visor.prog.perfil')}</span>
            <div class="chips" role="radiogroup">${perfiles.map(p => `<label class="chip"><input type="radio" name="pf" value="${esc(p.id)}" ${p.id === yo ? 'checked' : ''}> ${esc(p.emoji)} ${esc(p.apodo)}</label>`).join('')}
                <button class="btn btn-secondary" id="pf-nuevo" style="min-height:36px;padding:4px 12px">＋ ${_ih('visor.prog.nuevo')}</button></div>
            ${yo !== 'p0' ? `<div class="acciones" style="justify-content:flex-start;margin-top:4px"><button class="btn btn-secondary" id="pf-editar" style="min-height:34px;padding:4px 10px">✏️ ${_ih('visor.prog.editarPerfil')}</button></div>` : ''}
            <div class="prog-datos">
                ${tarjeta('👀', r.vistas, 'visor.prog.vistas')}${tarjeta('👂', r.escuchadas, 'visor.prog.escuchadas')}${tarjeta('⭐', r.favoritas, 'visor.prog.favoritas')}
                ${tarjeta('🃏', r.aprendidas, 'visor.prog.aprendidas')}${tarjeta('🎲', r.partidas7, 'visor.prog.partidas7')}${tarjeta('📅', r.dias7, 'visor.prog.dias7')}
            </div>
            ${r.aciertos != null ? `<p class="nota">${_ih('visor.prog.aciertos', { n: r.aciertos })}</p>` : ''}
            ${r.d.juegos.length ? `<h3>${_ih('visor.prog.ultimas')}</h3><div class="prog-lista">${r.d.juegos.slice(-8).reverse().map(j => `<div class="prog-fila"><span>${fecha(j.f)}</span><span>${esc(j.act || '')}${j.L ? ' · ' + esc(Datos.nombreLengua(j.L)) : ''}</span><b>${j.total ? `${j.bien || 0}/${j.total}` : '👥'}</b></div>`).join('')}</div>` : `<p class="nota">${_ih('visor.prog.sinPartidas')}</p>`}
            ${favs.length ? `<h3>⭐ ${_ih('visor.prog.favoritas')}</h3><div class="rel-items">${favs.map(w => `<button class="chip-rel" data-ficha="${esc(w.id)}">${esc(Datos.texto(w))}</button>`).join('')}</div>` : ''}
            ${perfiles.length > 1 ? `<details style="margin-top:12px"><summary><b>👥 ${_ih('visor.prog.todos')}</b></summary><p class="nota">${_ih('visor.prog.todos.d')}</p>
                <div class="prog-lista">${perfiles.map(p => { const x = this.resumen(p.id); return `<div class="prog-fila"><span>${esc(p.emoji)} ${esc(p.apodo)}</span><span>👀 ${x.vistas} · 🃏 ${x.aprendidas} · 🎲 ${x.partidas}</span><b>${x.ultimo ? fecha(x.ultimo) : '—'}</b></div>`; }).join('')}</div></details>` : ''}
            <h3>${_ih('visor.prog.compartir')}</h3>
            <p class="nota">${_ih('visor.prog.compartir.d')}</p>
            <div class="acciones" style="justify-content:flex-start;flex-wrap:wrap">
                <button class="btn btn-primary" id="pf-compartir">📤 ${_ih('visor.prog.compartirResumen')}</button>
                <button class="btn btn-secondary" id="pf-guardar">⬇️ ${_ih('visor.prog.guardarCopia')}</button>
                <button class="btn btn-secondary" id="pf-recuperar">📥 ${_ih('visor.prog.recuperar')}</button>
                <button class="btn btn-secondary" id="pf-borrar">🗑 ${_ih('visor.prog.borrar')}</button></div>
            <input type="file" id="pf-archivo" accept=".json,application/json" class="oculto" hidden></div>`);
        const B = $('modal-body');
        B.querySelectorAll('input[name="pf"]').forEach(x => x.addEventListener('change', () => { Perfiles.elegir(x.value); this.abrir(); }));
        $('pf-nuevo').addEventListener('click', () => this.editarPerfil());
        if ($('pf-editar')) $('pf-editar').addEventListener('click', () => this.editarPerfil(yo));
        $('pf-compartir').addEventListener('click', () => this.compartir());
        $('pf-guardar').addEventListener('click', () => this.guardarCopia());
        $('pf-recuperar').addEventListener('click', () => $('pf-archivo').click());
        $('pf-archivo').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) this.recuperar(f); });
        $('pf-borrar').addEventListener('click', () => {
            if (!confirm(_i('visor.prog.borrar.c', { perfil: Perfiles.etiqueta() }))) return;
            LS.del(this.clave()); Datos.langs.forEach(l => LS.del(`repaso:${Datos.base}:${l.id}${Perfiles.sufijo()}`)); this.abrir();
        });
    },
    editarPerfil(id) {
        const p = id ? Perfiles.datos(id) : { apodo: '', emoji: EMOJIS_PERFIL[Math.floor(Math.random() * EMOJIS_PERFIL.length)] };
        UI.abrir(`<div class="form"><h2 class="modal-titulo">${id ? '✏️ ' + _ih('visor.prog.editarPerfil') : '＋ ' + _ih('visor.prog.nuevo')}</h2>
            <p class="nota">${_ih('visor.prog.apodo.d')}</p>
            <label class="t" for="pf-apodo">${_ih('visor.prog.apodo')}</label>
            <input id="pf-apodo" class="input-box" maxlength="20" value="${esc(p.apodo)}" placeholder="${esc(_i('visor.prog.apodo.ej'))}">
            <span class="t">${_ih('visor.prog.dibujo')}</span>
            <div class="chips">${EMOJIS_PERFIL.map(e => `<label class="chip" style="font-size:1.3em"><input type="radio" name="pf-e" value="${e}" ${e === p.emoji ? 'checked' : ''}> ${e}</label>`).join('')}</div>
            <div id="pf-msg"></div>
            <div class="acciones">${id ? `<button class="btn btn-secondary" id="pf-quitar">🗑 ${_ih('visor.prog.quitarPerfil')}</button>` : ''}<button class="btn btn-secondary" id="pf-no">${_ih('visor.cancelar')}</button><button class="btn btn-primary" id="pf-ok">💾 ${_ih('visor.prog.guardarPerfil')}</button></div></div>`);
        $('pf-no').addEventListener('click', () => this.abrir());
        $('pf-ok').addEventListener('click', () => {
            const apodo = $('pf-apodo').value.trim(), emoji = (document.querySelector('input[name="pf-e"]:checked') || {}).value || '⭐';
            if (!apodo) { $('pf-msg').innerHTML = `<div class="error-box">${_ih('visor.prog.apodo.falta')}</div>`; return; }
            if (id) Perfiles.renombrar(id, apodo, emoji); else Perfiles.elegir(Perfiles.crear(apodo, emoji).id);
            Progreso.pintarBoton(); this.abrir();
        });
        if ($('pf-quitar')) $('pf-quitar').addEventListener('click', () => { if (confirm(_i('visor.prog.quitarPerfil.c', { perfil: Perfiles.etiqueta(id) }))) { Perfiles.borrar(id); this.abrir(); } });
    },
    /** Resumen breve, sin nombres: lo comparte la persona si quiere (menú de compartir del aparato, o se copia). */
    textoResumen() {
        const r = this.resumen(), p = Perfiles.datos();
        return [_i('visor.prog.resumen.titulo', { catalogo: Datos.cfg.titulo || '', perfil: `${p.emoji} ${p.id === 'p0' ? '' : p.apodo}`.trim(), fecha: new Date().toLocaleDateString() }),
            `👀 ${_i('visor.prog.vistas')}: ${r.vistas}`, `👂 ${_i('visor.prog.escuchadas')}: ${r.escuchadas}`, `⭐ ${_i('visor.prog.favoritas')}: ${r.favoritas}`,
            `🃏 ${_i('visor.prog.aprendidas')}: ${r.aprendidas}`, `🎲 ${_i('visor.prog.partidas7')}: ${r.partidas7}`, `📅 ${_i('visor.prog.dias7')}: ${r.dias7}`,
            r.aciertos != null ? _i('visor.prog.aciertos', { n: r.aciertos }).replace(/\*/g, '') : ''].filter(Boolean).join('\n');
    },
    async compartir() {
        const texto = this.textoResumen();
        try { if (navigator.share) { await navigator.share({ title: _i('visor.prog.titulo'), text: texto }); return; } } catch (e) { if (e.name === 'AbortError') return; }
        try { await navigator.clipboard.writeText(texto); UI.toast(_i('visor.prog.copiado')); } catch (e) { prompt(_i('visor.prog.copia'), texto); }
    },
    guardarCopia() {
        const p = Perfiles.datos(), suf = Perfiles.sufijo(), repaso = {};
        Datos.langs.forEach(l => { const v = LS.get(`repaso:${Datos.base}:${l.id}${suf}`, null); if (v) repaso[l.id] = v; });
        const datos = { formato: 'glosario-progreso-1', catalogo: Datos.cfg.idCatalogo || '', titulo: Datos.cfg.titulo || '', perfil: { apodo: p.apodo, emoji: p.emoji }, fecha: new Date().toISOString(), progreso: this.leer(), repaso };
        const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([JSON.stringify(datos)], { type: 'application/json' }));
        a.download = `progreso_${(p.apodo || 'aparato').normalize('NFD').replace(/[^\w-]+/g, '')}_${this.hoy()}.json`; document.body.appendChild(a); a.click(); a.remove();
        UI.toast(_i('visor.prog.copiaGuardada'));
    },
    async recuperar(f) {
        try {
            const j = JSON.parse(await f.text());
            if (j.formato !== 'glosario-progreso-1' || !j.progreso) throw new Error(_i('visor.prog.recuperar.mal'));
            if (j.catalogo && Datos.cfg.idCatalogo && j.catalogo !== Datos.cfg.idCatalogo) throw new Error(_i('visor.prog.recuperar.otro', { titulo: j.titulo || j.catalogo }));
            if (!confirm(_i('visor.prog.recuperar.c', { perfil: Perfiles.etiqueta() }))) return;
            LS.set(this.clave(), j.progreso);
            Object.entries(j.repaso || {}).forEach(([L, v]) => LS.set(`repaso:${Datos.base}:${L}${Perfiles.sufijo()}`, v));
            UI.toast(_i('visor.listo')); this.abrir();
        } catch (e) { UI.toast(e.message, 5000); }
    },
    /** Botón ⭐ en la ficha de cada palabra. */
    botonFavorita(id) {
        const h = $('modal-body') && $('modal-body').querySelector('.modal-titulo'); if (!h || $('f-fav')) return;
        const pinta = b => { const si = this.esFavorita(id); b.textContent = si ? '⭐ ' + _i('visor.prog.esFavorita') : '☆ ' + _i('visor.prog.marcarFavorita'); b.setAttribute('aria-pressed', si ? 'true' : 'false'); };
        h.insertAdjacentHTML('afterend', `<div style="text-align:center;margin:-4px 0 8px"><button class="btn btn-secondary" id="f-fav" style="min-height:34px;padding:4px 12px"></button></div>`);
        const b = $('f-fav'); pinta(b);
        b.addEventListener('click', () => { this.alternarFavorita(id); pinta(b); });
    }
};

/* Se registran las actividades sin tocar cada juego: se envuelven sus finales */
(() => {
    const abrirFicha = Ficha.abrir.bind(Ficha);
    Ficha.abrir = id => { abrirFicha(id); try { Progreso.vista(id); Progreso.botonFavorita(id); } catch (e) {} };
    const tocar = Sonido.tocar.bind(Sonido);
    Sonido.tocar = async (archivo, boton) => { await tocar(archivo, boton); try { if (Sonido.actual && !Sonido.actual.paused) Progreso.escuchada(archivo); } catch (e) {} };
    const finMotor = Motor.prototype.final;
    Motor.prototype.final = function (titulo, L) {
        try { Progreso.juego({ act: String(titulo || '').replace(/^\W+\s*/u, ''), L, bien: this.registro.filter(x => x.ok).length, total: this.registro.length, grupo: this.jugadores.length > 1 }); } catch (e) {}
        return finMotor.call(this, titulo, L);
    };
    const finRondas = Rondas.final.bind(Rondas);
    Rondas.final = (r, otra) => { try { Progreso.juego({ act: r.titulo, L: r.L, bien: r.bien, total: r.total }); } catch (e) {} return finRondas(r, otra); };
    [['Basta', 'basta'], ['Charadas', 'charadas']].forEach(([nombre, id]) => {
        const o = window[nombre] || (nombre === 'Basta' ? Basta : Charadas), fin = o.final.bind(o);
        o.final = (...a) => { try { Progreso.juego({ act: _i('visor.act.' + id), L: (o.cfg && o.cfg.L) || Datos.lengua, grupo: true }); } catch (e) {} return fin(...a); };
    });
})();

/* ---------- Instalar como app y funcionar sin conexión ---------- */
const Instalar = {
    pedido: null,
    esIOS: /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1),
    instalada: () => matchMedia('(display-mode: standalone)').matches || navigator.standalone === true,
    _persistido: false,
    /** Pide al aparato que no borre lo guardado (progreso, copia del catálogo, descargas). */
    persistir() { if (this._persistido) return; this._persistido = true; try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch (e) {} },
    iniciar() {
        const b = $('btn-instalar');
        window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); this.pedido = e; b.hidden = false; });
        window.addEventListener('appinstalled', () => { b.hidden = true; UI.toast(_i('visor.inst.lista')); this.persistir(); });
        if (this.esIOS && !this.instalada()) b.hidden = false;
        b.addEventListener('click', () => this.abrir());
        if (this.instalada()) this.persistir();
        if (!('serviceWorker' in navigator) || !(location.protocol === 'https:' || location.hostname === 'localhost')) return;
        const habia = !!navigator.serviceWorker.controller;
        navigator.serviceWorker.addEventListener('controllerchange', () => { if (habia && !Juego.activo) location.reload(); });
        navigator.serviceWorker.register('sw.js').then(r => r.addEventListener('updatefound', () => {
            const n = r.installing; if (!n) return;
            n.addEventListener('statechange', () => {
                if (n.state !== 'installed' || !navigator.serviceWorker.controller) return;
                const d = document.createElement('div'); d.className = 'toast toast-accion'; d.setAttribute('role', 'status');
                d.innerHTML = `${_ih('visor.inst.nueva')} <button class="btn btn-primary" style="min-height:32px;padding:4px 10px;margin-left:8px">${_ih('visor.inst.actualizar')}</button>`;
                d.querySelector('button').addEventListener('click', () => { if (Juego.activo) return UI.toast(_i('visor.inst.termina')); n.postMessage('saltar'); });
                document.body.appendChild(d);
            });
        })).catch(() => {});
    },
    async abrir() {
        if (this.pedido) { this.pedido.prompt(); await this.pedido.userChoice.catch(() => {}); this.pedido = null; $('btn-instalar').hidden = true; return; }
        UI.abrir(`<div class="form"><h2 class="modal-titulo">⬇️ ${_ih('visor.inst.titulo')}</h2>
            ${this.esIOS ? `<p>${_ih('visor.inst.ios1')}</p><p>${_ih('visor.inst.ios2')}</p><p>${_ih('visor.inst.ios3')}</p><div class="caja">${_ih('visor.inst.ios.d')}</div>`
                : `<p>${_ih('visor.inst.android')}</p><p>${_ih('visor.inst.pc')}</p>`}
            <p class="nota">${_ih('visor.inst.d')}</p>
            <div class="acciones"><button class="btn btn-primary" onclick="UI.cerrarModal()">${_ih('visor.inst.entendido')}</button></div></div>`);
    }
};
</script>
