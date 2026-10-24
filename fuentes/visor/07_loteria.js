<script>
/* =============================================
   LOTERÍA 1.1
   - Las tablas salen de la semilla del juego (fija); el orden de canto, de la
     semilla + la ronda (4 números). Así se pueden jugar muchas rondas con las
     mismas tablas impresas sin repetir el orden, y retomar una ronda exacta.
   - Partidas a distancia: “Nombre + 4 números” (ej. Juguemos 4320).
   - Archivo para jugar sin internet (se comparte por Bluetooth o WhatsApp).
   ============================================= */

/** Núcleo compartido con el archivo sin conexión (se copia tal cual dentro de él). */
function LOT_NUCLEO() {
    const VARIANTES = {
        llena:    { nombre: 'Tabla llena',     patrones: [[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]] },
        linea:    { nombre: 'Línea (chorro)',  patrones: [[0, 1, 2, 3], [4, 5, 6, 7], [8, 9, 10, 11], [12, 13, 14, 15], [0, 4, 8, 12], [1, 5, 9, 13], [2, 6, 10, 14], [3, 7, 11, 15], [0, 5, 10, 15], [3, 6, 9, 12]] },
        esquinas: { nombre: 'Cuatro esquinas', patrones: [[0, 3, 12, 15]] },
        centro:   { nombre: 'Centro (pozo)',   patrones: [[5, 6, 9, 10]] }
    };
    function hash32(str) {
        let h = 1779033703 ^ str.length;
        for (let i = 0; i < str.length; i++) { h = Math.imul(h ^ str.charCodeAt(i), 3432918353); h = h << 13 | h >>> 19; }
        h = Math.imul(h ^ h >>> 16, 2246822507); h = Math.imul(h ^ h >>> 13, 3266489909);
        return (h ^= h >>> 16) >>> 0;
    }
    function azar(semilla) {
        let a = hash32(String(semilla));
        return () => { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
    }
    function barajar(arr, r = Math.random) {
        const a = arr.slice();
        for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(r() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
        return a;
    }
    /** Orden de canto de una ronda. */
    const ordenCanto = (cfg, ronda) => barajar(cfg.ids, azar(cfg.semilla + '|' + ronda));
    /** Tablas 1..hasta. La tabla k nunca cambia aunque se pidan más. */
    const cache = {};
    function tablas(cfg, hasta) {
        const k = cfg.semilla + '|' + cfg.ids.join(',');
        const c = cache[k] = cache[k] || { r: azar(cfg.semilla + '|tablas'), vistas: new Set(), lista: [] };
        let intentos = 0;
        while (c.lista.length < hasta && intentos < hasta * 200) {
            intentos++;
            const t = barajar(cfg.ids, c.r).slice(0, 16), firma = t.slice().sort().join('|');
            if (c.vistas.has(firma)) continue;
            c.vistas.add(firma); c.lista.push(t);
        }
        return c.lista.slice(0, hasta);
    }
    const tabla = (cfg, k) => tablas(cfg, k)[k - 1];
    /** Revisa una tabla contra las cartas que ya salieron. */
    function revisar(cfg, t, salidas, marcas) {
        const s = new Set(salidas);
        return cfg.variantes.map(v => {
            const pat = VARIANTES[v].patrones;
            const completo = pat.find(p => p.every(i => s.has(t[i])));
            const marcado = marcas && pat.find(p => p.every(i => marcas.has(i)));
            return { v, nombre: VARIANTES[v].nombre, completo: !!completo, patron: completo || marcado || null,
                faltan: marcado ? marcado.filter(i => !s.has(t[i])) : [] };
        });
    }
    return { VARIANTES, azar, barajar, ordenCanto, tablas, tabla, revisar };
}
const LN = LOT_NUCLEO();

const Loteria = (() => {
    const MIN = 32, MAX = 120, REC = 54, ALFA = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    const nuevaSemilla = (n = 8) => Array.from(crypto.getRandomValues(new Uint32Array(n)), x => ALFA[x % ALFA.length]).join('');
    const llaveAnf = () => 'lot_anfitrion:' + Datos.base;

    let B = null;   // configuración en preparación
    const S = { rol: '', cfg: null, ronda: '', nombre: '', numero: '', secreto: '', enLinea: false, orden: [], n: 0, fin: false,
        ganadores: [], jugadores: [], tabla: null, tablaN: 0, jugador: '', marcas: new Set(), reclamadas: new Set(), vistos: new Set(),
        lenguaVista: '', poll: null, auto: null, fallos: 0, manualN: false };

    /* ---------- Datos ---------- */
    const w = id => Datos.byId.get(id);
    function nombreCarta(id, L) { const p = w(id); return p ? (p.t[L] || p.t[(S.cfg || B || {}).lengua] || {}).texto || Datos.texto(p) : '?'; }
    const imagen = (id, ancho = 400) => Datos.img(w(id), ancho) || placeholder('?');
    const numeroCarta = (cfg, id) => cfg.ids.indexOf(id) + 1;
    const nombreVariante = v => LN.VARIANTES[v] ? _i('visor.lot.var.' + v) : v;
    function elegibles(c, conCats = true) { return Datos.paraJuegos(c.lengua, { img: true, audio: c.audio, cats: conCats ? c.categorias : null }); }
    const audioDe = id => S.cfg && w(id)?.t[S.cfg.lengua]?.audio || '';
    function msg(html, tipo = 'caja') { const m = $('lot-msg'); if (m) m.innerHTML = `<div class="${tipo === 'error' ? 'error-box' : 'caja'}">${html}</div>`; }
    function detener() { clearInterval(S.poll); S.poll = null; clearInterval(S.auto); S.auto = null; }
    function activo() { Juego.activo = { detener }; }
    function avisarGanadores() {
        (S.ganadores || []).forEach(g => {
            const k = S.ronda + '|' + g.variante + '|' + g.tabla;
            if (S.vistos.has(k)) return; S.vistos.add(k);
            UI.anunciar(S.rol === 'jugador' && g.tabla === S.tablaN ? _i('visor.lot.ganaste', { forma: nombreVariante(g.variante) }) : _i('visor.lot.ganoOtro', { nombre: g.nombre, forma: nombreVariante(g.variante) }));
        });
    }
    function enlace(nombre, numero) {
        const u = new URL(location.origin + location.pathname);
        if (Datos.base !== (API_URL || '').split('?')[0]) u.searchParams.set('api', Datos.base);
        if (Datos.token && !API_TOKEN) u.searchParams.set('token', Datos.token);
        u.searchParams.set('loteria', nombre); u.searchParams.set('n', numero);
        return u.toString();
    }

    /* =============================================
       1. CONFIGURACIÓN
       ============================================= */
    function predeterminada() {
        return { tipo: 'loteria-configuracion', v: 2, semilla: nuevaSemilla(), nombre: _i('visor.act.loteria'), titulo: Datos.cfg.titulo || '',
            lengua: Datos.lengua, categorias: [], cantidad: REC, seleccion: 'auto', audio: false, lenguasJugador: true,
            variantes: ['llena', 'linea'], cantor: { auto: false, seg: 8 }, ids: [], impresas: 0 };
    }
    function leerAnfitrion() { const g = LS.get(llaveAnf()); return g && Date.now() - g.t < 6 * 3600e3 ? g : null; }
    function guardarAnfitrion() { if (S.rol === 'cantor' && S.enLinea) LS.set(llaveAnf(), { nombre: S.nombre, numero: S.numero, secreto: S.secreto, cfg: S.cfg, ronda: S.ronda, n: S.n, t: Date.now() }); }

    function abrir() { detener(); activo(); config(B || predeterminada()); }
    function config(c) {
        B = c;
        const g = leerAnfitrion();
        const vars = Object.entries(LN.VARIANTES).map(([k, v]) => `<label class="chip"><input type="checkbox" name="lot-var" value="${k}" ${c.variantes.includes(k) ? 'checked' : ''}> ${esc(nombreVariante(k))}</label>`).join('');
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🎴 ${_ih('visor.act.loteria')}</h2>
            ${g ? `<div class="caja">${_ih('visor.lot.enCurso', { x: g.nombre + ' ' + g.numero })}
                <button class="btn btn-secondary btn-block" style="margin-top:6px" id="lot-volver">↩️ ${_ih('visor.lot.volverPartida')}</button></div>` : ''}
            <label class="t" for="lot-nombre">${_ih('visor.lot.nombreJuego')} <span class="nota">${_ih('visor.lot.nombreJuego.d')}</span></label>
            <input id="lot-nombre" class="input-box" maxlength="20" value="${esc(c.nombre)}">
            <label class="t" for="lot-lengua">${_ih('visor.lot.lengua')}</label>
            <select id="lot-lengua" class="modern-select">${Comun.lenguas({ campo: 'img', min: MIN }, c.lengua)}</select>
            <p class="nota">${_ih('visor.lot.lengua.d')}</p>
            <span class="t">${_ih('visor.lot.categorias')} <span class="nota">${_ih('visor.lot.sinMarcar')}</span></span><div id="lot-cats"></div>
            <label class="t" for="lot-cantidad">${_ih('visor.lot.cantidad', { min: MIN, max: MAX, rec: REC })}</label>
            <input id="lot-cantidad" type="number" class="input-box" min="${MIN}" max="${MAX}" value="${c.cantidad}">
            <span class="t">${_ih('visor.lot.comoEligen')}</span>
            <div class="chips"><label class="chip"><input type="radio" name="lot-sel" value="auto" ${c.seleccion === 'auto' ? 'checked' : ''}> 🎲 ${_ih('visor.alAzar')}</label>
                <label class="chip"><input type="radio" name="lot-sel" value="manual" ${c.seleccion === 'manual' ? 'checked' : ''}> ✋ ${_ih('visor.lot.lasElijo')}</label></div>
            <span class="t">${_ih('visor.opciones')}</span>
            <div class="chips" style="flex-direction:column;align-items:stretch">
                <label class="chip"><input type="checkbox" id="lot-audio" ${c.audio ? 'checked' : ''}> 🔊 ${_ih('visor.lot.audio')}</label>
                <label class="chip"><input type="checkbox" id="lot-lj" ${c.lenguasJugador ? 'checked' : ''}> 🌐 ${_ih('visor.lot.otraLengua')}</label></div>
            <span class="t">${_ih('visor.lot.formasGanar')}</span><div class="chips">${vars}</div>
            <span class="t">${_ih('visor.lot.cantor')}</span>
            <div class="chips"><label class="chip"><input type="radio" name="lot-cantor" value="manual" ${!c.cantor.auto ? 'checked' : ''}> ${_ih('visor.lot.avanzoYo')}</label>
                <label class="chip"><input type="radio" name="lot-cantor" value="auto" ${c.cantor.auto ? 'checked' : ''}> ${_ih('visor.lot.automatico')}
                <select id="lot-seg" class="modern-select" style="width:auto;min-height:32px;padding:4px 8px">${[5, 8, 10, 15, 20, 30].map(s => `<option ${c.cantor.seg === s ? 'selected' : ''}>${s}</option>`).join('')}</select> s</label></div>
            <div id="lot-msg"></div>
            <div class="acciones" style="flex-direction:column;align-items:stretch">
                <button class="btn btn-primary" id="lot-cont">${_ih('visor.continuar')} ➔</button>
                <label class="btn btn-secondary">📂 ${_ih('visor.lot.cargar')}<input type="file" accept=".json,application/json" style="display:none" id="lot-file"></label>
                <button class="btn btn-secondary" onclick="Loteria.pantallaUnirse()">🎲 ${_ih('visor.lot.unirmePartida')}</button></div></div>`);
        const cats = () => { $('lot-cats').innerHTML = Comun.categorias($('lot-lengua').value, 'lot-cat', B.categorias); };
        $('lot-lengua').addEventListener('change', cats); cats();
        $('lot-cont').addEventListener('click', continuar);
        $('lot-file').addEventListener('change', e => importar(e.target.files[0]));
        $('lot-volver')?.addEventListener('click', () => { const x = leerAnfitrion(); Object.assign(S, { cfg: x.cfg, nombre: x.nombre, numero: x.numero, secreto: x.secreto, ronda: x.ronda, n: x.n, enLinea: true }); B = x.cfg; cantor(true); });
    }
    function leer() {
        const c = Object.assign({}, B);
        c.nombre = ($('lot-nombre').value.trim() || _i('visor.act.loteria')).slice(0, 20);
        c.lengua = $('lot-lengua').value;
        c.categorias = Comun.marcados('lot-cat');
        c.cantidad = parseInt($('lot-cantidad').value, 10) || REC;
        c.seleccion = document.querySelector('input[name="lot-sel"]:checked').value;
        c.audio = $('lot-audio').checked; c.lenguasJugador = $('lot-lj').checked;
        c.variantes = Comun.marcados('lot-var');
        c.cantor = { auto: document.querySelector('input[name="lot-cantor"]:checked').value === 'auto', seg: parseInt($('lot-seg').value, 10) || 8 };
        // Si cambian lengua o categorías, las tablas ya no serían las mismas: nueva semilla
        if (B.ids.length && (c.lengua !== B.lengua || c.audio !== B.audio)) { c.ids = []; c.semilla = nuevaSemilla(); c.impresas = 0; }
        return c;
    }
    function continuar() {
        const c = leer();
        if (!/[\p{L}\p{N}]/u.test(c.nombre)) return msg(_ih('visor.lot.error.nombre'), 'error');
        if (c.cantidad < MIN || c.cantidad > MAX) return msg(_ih('visor.lot.error.cantidad', { min: MIN, max: MAX }), 'error');
        if (!c.variantes.length) return msg(_ih('visor.lot.error.formas'), 'error');
        B = c;
        const pool = elegibles(c), validas = c.ids.filter(id => pool.some(x => x.id === id));
        if (c.seleccion === 'manual') return seleccionar(validas);
        if (validas.length === c.cantidad && c.ids.length === c.cantidad) return resumen();
        if (pool.length < c.cantidad) return insuficiente(c, pool);
        cambiarIds(barajar(pool.map(x => x.id)).slice(0, c.cantidad));
        resumen();
    }
    function cambiarIds(ids) { if (B.ids.join() !== ids.join()) { B.ids = ids; B.semilla = nuevaSemilla(); B.impresas = 0; } B.cantidad = ids.length; }
    function insuficiente(c, pool) {
        const todas = elegibles(c, false), puede = c.categorias.length && todas.length >= c.cantidad;
        UI.cuerpo(`<div class="form"><h2 class="modal-titulo">🎴 ${_ih('visor.lot.faltan')}</h2>
            <div class="caja">${_ih('visor.lot.faltan.d', { hay: pool.length, pediste: c.cantidad })}${c.audio ? '<br>' + _ih('visor.lot.faltan.audio') : ''}</div>
            <div class="acciones" style="flex-direction:column;align-items:stretch">
                <button class="btn btn-secondary" id="l-mas">➕ ${_ih('visor.lot.masCats')}</button>
                ${puede ? '<button class="btn btn-primary" id="l-comp">🎲 ' + _ih('visor.lot.completarOtras') + '</button>' : ''}
                ${pool.length >= MIN ? `<button class="btn btn-secondary" id="l-con">${_ih('visor.lot.jugarCon', { n: pool.length })}</button>` : ''}</div>
            ${!puede && pool.length < MIN ? `<p class="nota">${c.audio ? _ih('visor.lot.hayConAudio', { n: todas.length }) : _ih('visor.lot.hay', { n: todas.length })}</p>` : ''}</div>`);
        $('l-mas').addEventListener('click', () => config(B));
        $('l-comp')?.addEventListener('click', () => { const p = barajar(pool.map(x => x.id)); cambiarIds(p.concat(barajar(todas.map(x => x.id).filter(id => !p.includes(id)))).slice(0, c.cantidad)); resumen(); });
        $('l-con')?.addEventListener('click', () => { cambiarIds(barajar(pool.map(x => x.id))); resumen(); });
    }

    /* ---------- Selección manual ---------- */
    let sel = new Set();
    function seleccionar(ids) {
        sel = new Set(ids);
        UI.cuerpo(`<div class="form"><h2 class="modal-titulo">✋ ${_ih('visor.lot.eligeCartas')}</h2>
            <div class="fila"><input id="ls-q" class="input-box" style="flex:2;min-width:140px" placeholder="🔍 ${esc(_i('visor.buscarCorto'))}" aria-label="${esc(_i('visor.buscarCorto'))}">
                <select id="ls-cat" class="modern-select" style="flex:1;min-width:130px" aria-label="${esc(_i('visor.categoria'))}"><option value="">${_ih('visor.todas')}</option>${[...Datos.cats.values()].map(c => `<option value="${esc(c.id)}">${esc(Datos.categoria(c.id, B.lengua))}</option>`).join('')}</select></div>
            <div class="lot-contador" id="ls-n"></div>
            <div class="lot-sel-grid" id="ls-grid"></div>
            <div class="fila" style="margin-top:8px"><button class="btn btn-secondary" style="flex:1" id="ls-comp">🎲 ${_ih('visor.lot.completarAzar')}</button><button class="btn btn-secondary" style="flex:1" id="ls-0">${_ih('visor.lot.quitarTodas')}</button></div>
            <div id="lot-msg"></div>
            <div class="acciones"><button class="btn btn-secondary" id="ls-atras">← ${_ih('visor.volver')}</button><button class="btn btn-primary" id="ls-ok">${_ih('visor.continuar')} ➔</button></div></div>`);
        const pool = elegibles(B, false);
        const pinta = () => {
            const q = Texto.pliegue($('ls-q').value), cat = $('ls-cat').value;
            $('ls-grid').innerHTML = pool.filter(x => (!cat || x.cat === cat) && (!q || Texto.pliegue(x.t[B.lengua].texto).includes(q))).map(x =>
                `<button class="lot-sel ${sel.has(x.id) ? 'on' : ''}" data-id="${esc(x.id)}" aria-pressed="${sel.has(x.id)}"><img src="${esc(Datos.img(x, 200))}" alt="" loading="lazy"><span${Texto.lang(B.lengua)}>${esc(x.t[B.lengua].texto)}</span></button>`).join('') || '<p class="nota">' + _ih('visor.nadaFiltros') + '</p>';
            cuenta();
        };
        const cuenta = () => { const n = sel.size, ok = n >= MIN && n <= MAX; $('ls-n').innerHTML = `${_ih('visor.lot.elegidas')} <span style="color:${ok ? 'var(--success)' : 'var(--error)'}">${n}</span> <span class="nota">${_ih('visor.lot.deA', { min: MIN, max: MAX })}</span>`; };
        $('ls-q').addEventListener('input', pinta); $('ls-cat').addEventListener('change', pinta);
        $('ls-grid').addEventListener('click', e => { const b = e.target.closest('.lot-sel'); if (!b) return; const id = b.dataset.id; sel.has(id) ? sel.delete(id) : sel.add(id); b.classList.toggle('on', sel.has(id)); b.setAttribute('aria-pressed', sel.has(id)); cuenta(); });
        $('ls-comp').addEventListener('click', () => {
            const meta = Math.max(MIN, Math.min(MAX, B.cantidad));
            const resto = barajar(elegibles(B).map(x => x.id).filter(id => !sel.has(id))), otras = barajar(pool.map(x => x.id).filter(id => !sel.has(id) && !resto.includes(id)));
            while (sel.size < meta && resto.length) sel.add(resto.pop());
            while (sel.size < meta && otras.length) sel.add(otras.pop());
            pinta();
        });
        $('ls-0').addEventListener('click', () => { sel.clear(); pinta(); });
        $('ls-atras').addEventListener('click', () => config(B));
        $('ls-ok').addEventListener('click', () => { if (sel.size < MIN || sel.size > MAX) return msg(_ih('visor.lot.error.elige', { min: MIN, max: MAX, n: sel.size }), 'error'); cambiarIds([...sel]); resumen(); });
        pinta();
    }

    /* ---------- Resumen ---------- */
    function resumen() {
        const c = B;
        UI.cuerpo(`<div class="form"><h2 class="modal-titulo">🎴 ${esc(c.nombre)}</h2>
            <div class="caja">${_ih('visor.lot.nCartas', { n: c.ids.length })} · ${esc(Datos.nombreLengua(c.lengua))} · ${_ih('visor.lot.ganar', { formas: c.variantes.map(nombreVariante).join(', ') })}
                <div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:8px">${c.ids.slice(0, 12).map(id => `<img src="${esc(imagen(id, 100))}" alt="" style="width:44px;height:44px;object-fit:cover;border-radius:8px;border:2px solid var(--border);background:#fff">`).join('')}${c.ids.length > 12 ? `<span class="nota" style="align-self:center">+${c.ids.length - 12}</span>` : ''}</div>
                <p class="nota">${_ih('visor.lot.juegoSemilla', { x: c.semilla })}</p></div>
            <span class="t">🌐 ${_ih('visor.lot.distancia')}</span>
            <button class="btn btn-primary btn-block" id="l-linea">${_ih('visor.lot.crear', { nombre: c.nombre })}</button>
            <span class="t">📴 ${_ih('visor.lot.sinInternet')}</span>
            <div class="acciones" style="flex-direction:column;align-items:stretch;margin-top:0">
                <button class="btn btn-secondary" id="l-local">📣 ${_ih('visor.lot.cantarAqui')}</button>
                <button class="btn btn-secondary" id="l-arch">📲 ${_ih('visor.lot.archivo')}</button></div>
            <span class="t">🖨️ ${_ih('visor.lot.imprimir')}</span>
            <div class="caja" style="margin-top:0">
                <button class="btn btn-secondary btn-block" id="l-cartas">${_ih('visor.lot.cartasCantar')}</button>
                <label class="chip" style="margin-top:6px"><input type="checkbox" id="l-verso"> ✍️ ${_ih('visor.lot.verso')}</label>
                <div class="fila" style="margin-top:10px"><label for="l-nt" class="nota"><b>${_ih('visor.lot.tablasDos')}</b></label>
                    <input id="l-nt" type="number" min="1" max="200" value="${Math.max(10, c.impresas || 0)}" class="input-box" style="width:90px">
                    <select id="l-fmt" class="modern-select" style="flex:1;min-width:170px" aria-label="${esc(_i('visor.lot.formato'))}"><option value="media">${_ih('visor.lot.media')}</option><option value="carta">${_ih('visor.lot.carta')}</option></select></div>
                <button class="btn btn-secondary btn-block" style="margin-top:8px" id="l-tablas">${_ih('visor.lot.tablas4')}</button>${Impresion.html()}</div>
            <div id="lot-msg"></div>
            <div class="acciones"><button class="btn btn-secondary" id="l-guardar">💾 ${_ih('visor.lot.guardar')}</button><button class="btn btn-secondary" id="l-editar">✏️ ${_ih('visor.lot.cambiarCartas')}</button><button class="btn btn-secondary" id="l-opc">← ${_ih('visor.opciones')}</button></div></div>`);
        $('l-linea').addEventListener('click', crearEnLinea);
        $('l-local').addEventListener('click', () => { Object.assign(S, { cfg: JSON.parse(JSON.stringify(B)), enLinea: false, ronda: cuatroDigitos(), n: 0, nombre: B.nombre, numero: '' }); cantor(); });
        $('l-arch').addEventListener('click', archivoSinConexion);
        $('l-cartas').addEventListener('click', () => imprimir('cartas'));
        $('l-tablas').addEventListener('click', () => imprimir('tablas'));
        $('l-guardar').addEventListener('click', exportar);
        $('l-editar').addEventListener('click', () => seleccionar(B.ids));
        $('l-opc').addEventListener('click', () => config(B));
    }
    function exportar() {
        descargar(`loteria-${B.nombre.replace(/[^\p{L}\p{N}]+/gu, '-').toLowerCase()}-${B.semilla}.json`, JSON.stringify(B, null, 1), 'application/json');
        msg(_ih('visor.lot.guardada'));
    }
    function importar(f) {
        if (!f) return;
        const r = new FileReader();
        r.onload = () => {
            try {
                const c = JSON.parse(r.result);
                if (c.tipo !== 'loteria-configuracion' || !Array.isArray(c.ids)) throw new Error(_i('visor.lot.error.archivo'));
                const base = predeterminada();
                B = Object.assign(base, c, { cantor: Object.assign(base.cantor, c.cantor || {}), nombre: c.nombre || c.titulo || _i('visor.act.loteria'), v: 2 });
                const pool = new Set(elegibles(B, false).map(x => x.id)), validas = B.ids.filter(id => pool.has(id)), perdidas = B.ids.length - validas.length;
                if (perdidas) { B.ids = validas; B.semilla = nuevaSemilla(); B.impresas = 0; }
                if (validas.length >= MIN) { B.cantidad = validas.length; resumen(); if (perdidas) msg(_ih('visor.lot.perdidas', { n: perdidas })); }
                else { seleccionar(validas); msg(_ih('visor.lot.quedan', { n: validas.length })); }
            } catch (e) { config(B || predeterminada()); msg(_ih('visor.lot.error.cargar', { x: e.message }), 'error'); }
        };
        r.readAsText(f);
    }

    /* =============================================
       2. IMPRESIÓN (tamaño carta, 5 mm de margen)
       ============================================= */
    const CSS_IMPRESION = `
        html,body{margin:0;padding:0;background:#fff;color:#111;font-family:system-ui,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
        .barra{position:sticky;top:0;background:#1e3a5f;color:#fff;padding:10px 14px;font-size:14px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;z-index:5}
        .barra button{font-size:15px;font-weight:700;padding:8px 14px;border:0;border-radius:8px;cursor:pointer}
        @media print{.barra{display:none}}
        @media screen{.hoja{margin:10px auto;box-shadow:0 0 0 1px #ccc,0 4px 14px rgba(0,0,0,.15)}}
        .hoja{box-sizing:border-box;overflow:hidden;break-after:page;page-break-after:always;background:#fff}
        .hoja:last-of-type{break-after:auto;page-break-after:auto}
        .num{font-weight:900;line-height:1}
        .img{flex:1;min-height:0;display:flex;align-items:center;justify-content:center;overflow:hidden}
        .img img{max-width:100%;max-height:100%;object-fit:contain}
        .nombre{border:.4mm solid #111;border-radius:1.5mm;text-align:center;font-weight:800;line-height:1.15;word-break:break-word;background:#fff}
        .cartas{width:var(--w);height:var(--h);display:grid;grid-template-columns:repeat(4,1fr);grid-template-rows:repeat(2,1fr);gap:3mm}
        .carta{border:.7mm solid #111;border-radius:2.5mm;padding:2.5mm;display:flex;flex-direction:column;gap:1.5mm;min-height:0;overflow:hidden}
        .carta .num{font-size:15pt}.carta .nombre{font-size:12pt;padding:1.2mm 1.5mm}
        .tabla{border:.8mm solid #111;border-radius:3mm;padding:3mm;display:flex;flex-direction:column;box-sizing:border-box;height:100%;min-height:0}
        .tabla header{display:flex;justify-content:space-between;align-items:baseline;gap:3mm;font-weight:900;margin-bottom:2mm}
        .tabla header .tit{font-weight:600;flex:1;text-align:center;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
        .tabla footer{font-size:7pt;color:#555;text-align:right;margin-top:1mm}
        .rejilla{flex:1;min-height:0;display:grid;grid-template-columns:repeat(4,1fr);grid-template-rows:repeat(4,1fr);gap:1.5mm}
        .celda{border:.5mm solid #111;border-radius:1.5mm;padding:1mm;display:flex;flex-direction:column;gap:.8mm;min-height:0;overflow:hidden;position:relative}
        .celda .num{position:absolute;top:.8mm;left:1mm;background:#fff;border-radius:1mm;padding:0 .8mm}
        .medias{width:var(--w);height:var(--h);display:grid;grid-template-columns:1fr 1fr;column-gap:8mm;background:repeating-linear-gradient(to bottom,#999 0 2mm,transparent 2mm 4mm) center/0.3mm 100% no-repeat}
        .medias header{font-size:11pt}.medias .celda .num{font-size:7pt}.medias .celda .nombre{font-size:7.5pt;padding:.5mm .8mm}
        .completa{width:var(--w);height:var(--h)}
        .verso{display:flex;flex-direction:column;gap:1mm}.verso .p-linea{border-bottom:.3mm solid #111;height:5mm}
        .pie-lot{font-size:7pt;color:#555;text-align:center}
        .completa header{font-size:15pt}.completa .celda .num{font-size:10pt}.completa .celda .nombre{font-size:10.5pt;padding:.8mm 1mm}`;
    let VERSO = false, PIE = '';
    function cartaHTML(cfg, id, clase, L = cfg.lengua) {
        return `<div class="${clase}"><div class="num">${numeroCarta(cfg, id)}</div><div class="img"><img src="${esc(imagen(id, clase === 'carta' ? 800 : 500))}" alt=""></div><div class="nombre"${Texto.lang(L)}>${esc(nombreCarta(id, L))}</div>${clase === 'carta' && VERSO ? '<div class="verso"><div class="p-linea"></div><div class="p-linea"></div></div>' : ''}</div>`;
    }
    function tablaHTML(cfg, t, k) {
        return `<div class="tabla"><header><span>${esc(_i('visor.act.loteria').toLocaleUpperCase())}</span><span class="tit">${esc(cfg.nombre)}</span><span>${_ih('visor.lot.tablaN', { n: k })}</span></header>
            <div class="rejilla">${t.map(id => cartaHTML(cfg, id, 'celda')).join('')}</div><footer>${_ih('visor.lot.juegoN', { x: cfg.semilla })}${PIE ? ' · ' + PIE : ''}</footer></div>`;
    }
    /** Abre una ventana con hojas listas para imprimir; espera a que carguen las imágenes. */
    function ventanaImpresion(titulo, paginas, orientacion, descripcion, css, aj) {
        const v = window.open('', '_blank');
        if (!v) return msg(_ih('visor.imp.bloqueada'), 'error');
        v.document.write(`<!DOCTYPE html><html lang="${esc(document.documentElement.lang || 'es')}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
            <title>${esc(titulo)}</title><style>${Impresion.css(aj, orientacion)}${css}</style></head><body>
            <div class="barra"><b>${esc(descripcion)}</b><span>${esc(_i('visor.imp.papel.' + aj.papel))} · ${orientacion === 'landscape' ? _ih('visor.imp.horizontalCorto') : _ih('visor.imp.verticalCorto')}</span><button onclick="print()">🖨️ ${_ih('visor.imp.boton')}</button></div>
            ${paginas}
            <script>(function(){var h=false;function ir(){if(h)return;h=true;setTimeout(function(){window.print()},300)}var im=[].slice.call(document.images),p=im.length;if(!p)return ir();im.forEach(function(i){if(i.complete){if(--p===0)ir()}else{i.onload=i.onerror=function(){if(--p===0)ir()}}});setTimeout(ir,15000)})();<\/script></body></html>`);
        v.document.close();
        msg(_ih('visor.imp.abierta', { x: descripcion }));
    }
    function imprimir(tipo) {
        const c = B, aj = Impresion.leer(); let pag = '', o = 'landscape', d = '';
        VERSO = !!($('l-verso') && $('l-verso').checked);
        const cfg = Datos.cfg || {}, ley = (Estado.e && Estado.e.leyenda) || '';
        PIE = [cfg.titulo, cfg.licencia ? _i('visor.imp.licencia', { x: cfg.licencia }) : '', ley, aj.creditos ? _i('visor.pie.credito', { autor: AUTOR }) : ''].filter(Boolean).map(esc).join(' · ');
        if (tipo === 'cartas') {
            for (let p = 0; p < c.ids.length; p += 8) pag += `<section class="hoja cartas">${c.ids.slice(p, p + 8).map(id => cartaHTML(c, id, 'carta')).join('')}</section>`;
            d = _i('visor.lot.d.cartas', { n: c.ids.length, h: Math.ceil(c.ids.length / 8) });
        } else {
            const n = Math.max(1, Math.min(200, parseInt($('l-nt').value, 10) || 1)), fmt = $('l-fmt').value, ts = LN.tablas(c, n);
            c.impresas = Math.max(c.impresas || 0, ts.length);
            if (fmt === 'media') { for (let p = 0; p < ts.length; p += 2) pag += `<section class="hoja medias">${ts.slice(p, p + 2).map((t, i) => tablaHTML(c, t, p + i + 1)).join('')}</section>`; d = _i('visor.lot.d.medias', { n: ts.length, h: Math.ceil(ts.length / 2) }); }
            else { o = 'portrait'; pag = ts.map((t, i) => `<section class="hoja completa">${tablaHTML(c, t, i + 1)}</section>`).join(''); d = _i('visor.lot.d.carta', { n: ts.length, h: ts.length }); }
        }
        ventanaImpresion(`${_i('visor.act.loteria')} · ${c.nombre} · ${c.semilla}`, pag, o, d, CSS_IMPRESION, aj);
    }

    /* =============================================
       3. CANTOR
       ============================================= */
    async function crearEnLinea() {
        msg('⏳ ' + _ih('visor.lot.creando'));
        try {
            const secreto = nuevaSemilla(20), cfg = JSON.parse(JSON.stringify(B)), ronda = cuatroDigitos();
            const r = await Datos.post({ accion: 'lot_crear', config: cfg, secreto, ronda });
            if (r.status !== 200 || !r.numero) throw new Error(r.error || _i('visor.lot.error.script'));
            Object.assign(S, { cfg, nombre: r.nombre, numero: r.numero, secreto, ronda: r.ronda || ronda, enLinea: true, n: 0 });
            cantor();
        } catch (e) { msg(_ih('visor.lot.error.crear', { x: e.message || _i('visor.sinConexion') }), 'error'); }
    }
    function cantor(retomar = false) {
        detener(); activo();
        Object.assign(S, { rol: 'cantor', orden: LN.ordenCanto(S.cfg, S.ronda), fin: retomar ? S.fin : false, ganadores: [], jugadores: [], vistos: new Set() });
        UI.abrir('<div></div>');
        pintarCantor();
        if (S.enLinea) { guardarAnfitrion(); sincronizarCantor(); S.poll = setInterval(sincronizarCantor, 4000); }
        if (S.cfg.cantor.auto) auto(true);
        Pantalla.mantener();
    }
    function cartaGrande(id) {
        const p = w(id), L = S.cfg.lengua;
        const otras = S.cfg.lenguasJugador ? Object.keys(p.t).filter(x => x !== L).map(x => `${esc(Datos.nombreLengua(x))}: <b${Texto.lang(x)}>${esc(p.t[x].texto)}</b>`) : [];
        return `<div class="lot-carta" aria-live="polite"><div class="num">${numeroCarta(S.cfg, id)}</div><div class="img"><img src="${esc(imagen(id, 600))}" alt="${esc(nombreCarta(id, L))}"></div>
            <div class="lot-nombre"${Texto.lang(L)}>${esc(nombreCarta(id, L))}</div></div>
            ${otras.length ? `<div class="nota" style="text-align:center;margin-top:6px">${otras.join(' · ')}</div>` : ''}
            ${audioDe(id) ? `<div style="text-align:center;margin-top:6px">${Sonido.boton(p, L, '🔊 ' + esc(_i('visor.escuchar')))}</div>` : ''}`;
    }
    function variantesHTML() {
        return S.cfg.variantes.map(v => { const g = S.ganadores.find(x => x.variante === v); return `<span class="lot-var ${g ? 'ganada' : ''}">${g ? '🏆' : '⏳'} ${esc(nombreVariante(v))}${g ? ': ' + esc(g.nombre) : ''}</span>`; }).join('');
    }
    function pintarCantor() {
        const c = S.cfg, total = S.orden.length, id = S.n ? S.orden[S.n - 1] : null;
        const hist = S.orden.slice(0, Math.max(0, S.n - 1)).reverse().map(i => `<figure><img src="${esc(imagen(i, 120))}" alt=""><figcaption${Texto.lang(c.lengua)}>${esc(nombreCarta(i, c.lengua))}</figcaption></figure>`).join('');
        UI.cuerpo(`${S.enLinea ? `<div class="lot-codigo">${_ih('visor.lot.paraUnirse', { x: S.nombre + ' ' + S.numero })}
                <div class="fila" style="justify-content:center;margin-top:8px"><button class="btn btn-secondary" id="lc-comp">📤 ${_ih('visor.compartirEnlace')}</button><button class="btn btn-secondary" id="lc-copia">📋 ${_ih('visor.copiar')}</button></div>
                <div class="nota" id="lc-jug">👥 ${_ih('visor.lot.esperando')}</div></div>`
            : `<div class="caja nota" style="margin-top:0">${_ih('visor.lot.local', { x: c.semilla })}</div>`}
            <div class="lot-contador">${_ih('visor.lot.contador', { ronda: S.ronda, n: S.n, total })}</div>
            ${id ? cartaGrande(id) : '<div class="caja" style="text-align:center">' + _ih('visor.lot.toca') + '</div>'}
            <div class="lot-variantes" id="lc-vars">${variantesHTML()}</div>
            ${S.fin ? '<div class="round-result win">🏁 ' + _ih('visor.lot.terminada') + '</div>' : `<div class="acciones" style="flex-direction:column;align-items:stretch">
                <button class="btn btn-primary" id="lc-sig" style="min-height:64px;font-size:1.2em" ${S.n >= total ? 'disabled' : ''}>▶ ${_ih('visor.lot.siguiente')}</button>
                <label class="chip" style="justify-content:center"><input type="checkbox" id="lc-auto" ${S.auto ? 'checked' : ''}> ${_ih('visor.lot.avanzarSolo')}
                    <select id="lc-seg" class="modern-select" style="width:auto;min-height:32px;padding:4px 8px">${[5, 8, 10, 15, 20, 30].map(s => `<option ${c.cantor.seg === s ? 'selected' : ''}>${s}</option>`).join('')}</select> s</label></div>`}
            ${hist ? `<details style="margin-top:12px"><summary>${_ih('visor.lot.yaSalieron', { n: S.n - 1 })}</summary><div class="lot-hist">${hist}</div></details>` : ''}
            <div class="acciones">
                <button class="btn btn-secondary" id="lc-ver">🔍 ${_ih('visor.lot.verificar')}</button>
                <button class="btn btn-secondary" id="lc-ronda">🔁 ${_ih('visor.lot.nuevaRonda')}</button>
                ${!S.fin ? '<button class="btn btn-danger" id="lc-fin">🏁 ' + _ih('visor.lot.terminar') + '</button>' : ''}
                <button class="btn btn-secondary" id="lc-nueva">🎴 ${_ih('visor.lot.otra')}</button></div>`);
        const on = (i, f, ev = 'click') => { const e = $(i); if (e) e.addEventListener(ev, f); };
        on('lc-sig', siguiente); on('lc-auto', e => auto(e.target.checked), 'change');
        on('lc-seg', e => { S.cfg.cantor.seg = +e.target.value; if (S.auto) auto(true); }, 'change');
        on('lc-ver', () => verificar()); on('lc-ronda', nuevaRonda); on('lc-fin', terminar); on('lc-nueva', () => { if (confirm(_i('visor.lot.salir'))) { detener(); LS.del(llaveAnf()); config(B || predeterminada()); } });
        on('lc-comp', compartir); on('lc-copia', copiar);
        pintarJugadores();
    }
    function pintarJugadores() { const el = $('lc-jug'); if (el) el.textContent = S.jugadores.length ? `👥 ${S.jugadores.length}: ${S.jugadores.map(j => _i('visor.lot.jugadorTabla', { nombre: j.nombre, n: j.tabla })).join(', ')}` : '👥 ' + _i('visor.lot.esperando'); }
    function siguiente() {
        if (S.fin || S.n >= S.orden.length) return auto(false);
        S.n++;
        if (S.cfg.audio) { const a = audioDe(S.orden[S.n - 1]); if (a) Sonido.tocar(a); }
        pintarCantor(); guardarAnfitrion();
        if (S.enLinea) Datos.post({ accion: 'lot_actualizar', nombre: S.nombre, numero: S.numero, secreto: S.secreto, n: S.n, ronda: S.ronda }).catch(() => {});
        if (S.n >= S.orden.length) { auto(false); UI.toast(_i('visor.lot.todas')); }
    }
    function auto(si) {
        clearInterval(S.auto); S.auto = null;
        if (si && !S.fin && S.n < S.orden.length) S.auto = setInterval(siguiente, (S.cfg.cantor.seg || 8) * 1000);
        const c = $('lc-auto'); if (c) c.checked = !!S.auto;
    }
    async function sincronizarCantor() {
        try {
            const r = await Datos.api({ action: 'lot_estado', nombre: S.nombre, numero: S.numero }, 10000);
            if (r.status !== 200) return;
            S.jugadores = r.estado.jugadores || [];
            if (r.estado.ronda === S.ronda) { S.ganadores = r.estado.ganadores || []; avisarGanadores(); }
            const v = $('lc-vars'); if (v) v.innerHTML = variantesHTML();
            pintarJugadores();
            if (r.estado.fin && r.estado.ronda === S.ronda && !S.fin) { S.fin = true; auto(false); pintarCantor(); }
        } catch (e) { /* sin señal: se reintenta */ }
    }
    function terminar() {
        if (!confirm(_i('visor.lot.terminarTodos'))) return;
        S.fin = true; auto(false);
        if (S.enLinea) Datos.post({ accion: 'lot_actualizar', nombre: S.nombre, numero: S.numero, secreto: S.secreto, fin: true }).catch(() => {});
        pintarCantor();
    }
    function nuevaRonda() {
        if (S.n && !S.fin && !confirm(_i('visor.lot.confirmaRonda'))) return;
        let r; do { r = cuatroDigitos(); } while (r === S.ronda);
        Object.assign(S, { ronda: r, n: 0, fin: false, ganadores: [], orden: LN.ordenCanto(S.cfg, r) });
        auto(false);
        if (S.enLinea) Datos.post({ accion: 'lot_actualizar', nombre: S.nombre, numero: S.numero, secreto: S.secreto, ronda: r, n: 0, fin: false }).catch(() => {});
        guardarAnfitrion(); pintarCantor();
        if (S.cfg.cantor.auto) auto(true);
    }
    async function compartir() {
        const u = enlace(S.nombre, S.numero), t = _i('visor.lot.invitacion', { partida: S.nombre + ' ' + S.numero, enlace: u });
        if (navigator.share) { try { await navigator.share({ title: _i('visor.act.loteria'), text: t, url: u }); return; } catch (e) { return; } }
        window.open('https://wa.me/?text=' + encodeURIComponent(t), '_blank');
    }
    async function copiar() { const u = enlace(S.nombre, S.numero); try { await navigator.clipboard.writeText(u); UI.toast(_i('visor.enlaceCopiado')); } catch (e) { prompt(_i('visor.copiaEnlace'), u); } }

    /** Verificación de cualquier tabla (impresa, descargada o en pantalla) contra las cartas que ya salieron. */
    function verificar(k = '') {
        const c = S.cfg, salidas = S.orden.slice(0, S.n);
        let res = '';
        const n = parseInt(k, 10);
        if (n > 0) {
            const t = LN.tabla(c, n);
            if (!t) res = '<div class="error-box">' + _ih('visor.lot.error.tabla') + '</div>';
            else {
                const rev = LN.revisar(c, t, salidas), s = new Set(salidas);
                const pat = new Set((rev.find(x => x.completo) || {}).patron || []);
                res = `<div class="lot-variantes">${rev.map(x => `<span class="lot-var ${x.completo ? 'ganada' : ''}">${x.completo ? '✅' : '—'} ${esc(nombreVariante(x.v))}</span>`).join('')}</div>
                    <div class="lot-tabla">${t.map((id, i) => `<div class="lot-celda ${s.has(id) ? 'salio' : 'nosalio'} ${pat.has(i) ? 'marcada' : ''}"><span class="n">${numeroCarta(c, id)}</span><img src="${esc(imagen(id, 200))}" alt=""><span class="nm"${Texto.lang(c.lengua)}>${esc(nombreCarta(id, c.lengua))}</span></div>`).join('')}</div>
                    <p class="nota">${_ih('visor.lot.colores')} ${rev.some(x => x.completo) ? '<b>' + _ih('visor.lot.gana') + '</b>' : _ih('visor.lot.noGana')}</p>
                    ${rev.filter(x => !x.completo).length ? `<p class="nota">${_ih('visor.lot.revisen', { x: t.filter(id => !s.has(id)).slice(0, 8).map(id => nombreCarta(id, c.lengua)).join(', ') })}</p>` : ''}`;
            }
        }
        UI.cuerpo(`<h2 class="modal-titulo">🔍 ${_ih('visor.lot.verificar')}</h2>
            <p class="nota">${_ih('visor.lot.verificar.d', { ronda: S.ronda, n: S.n })}</p>
            <div class="fila"><input id="lv-n" type="number" min="1" inputmode="numeric" class="input-box" style="width:120px" value="${esc(k)}" aria-label="${esc(_i('visor.lot.numTabla'))}"><button class="btn btn-primary" id="lv-ok">${_ih('visor.revisar')}</button></div>
            ${res}
            <div class="acciones"><button class="btn btn-secondary" id="lv-volver">← ${_ih('visor.lot.volverCantar')}</button></div>`);
        $('lv-ok').addEventListener('click', () => verificar($('lv-n').value));
        $('lv-n').addEventListener('keydown', e => { if (e.key === 'Enter') verificar($('lv-n').value); });
        $('lv-volver').addEventListener('click', pintarCantor);
    }

    /* =============================================
       4. JUGADOR EN LÍNEA
       ============================================= */
    function pantallaUnirse(nombre = '', numero = '') {
        detener(); activo();
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🎲 ${_ih('visor.act.unirse')}</h2>
            <p class="nota">${_ih('visor.lot.pide')}</p>
            <div class="fila"><input id="lu-nom" class="input-box" style="flex:2;min-width:140px" maxlength="20" placeholder="${esc(_i('visor.nombrePartida'))}" value="${esc(nombre)}" aria-label="${esc(_i('visor.nombrePartida'))}">
                <input id="lu-num" class="input-box" style="flex:1;min-width:90px;font-size:1.3em;letter-spacing:3px;text-align:center" inputmode="numeric" maxlength="4" placeholder="0000" value="${esc(numero)}" aria-label="${esc(_i('visor.cuatroNumeros'))}"></div>
            <label class="t" for="lu-yo">${_ih('visor.lot.tuNombre')}</label>
            <input id="lu-yo" class="input-box" maxlength="40" value="${esc(LS.get('lot_nombre', ''))}" placeholder="${esc(_i('visor.lot.tuNombre.d'))}">
            <div id="lot-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="lu-ok">${_ih('visor.lot.entrar')} ➔</button></div></div>`);
        $('lu-ok').addEventListener('click', () => unirse());
    }
    async function unirse(nombreD, numeroD) {
        const nombre = String(nombreD ?? $('lu-nom')?.value ?? '').trim(), numero = String(numeroD ?? $('lu-num')?.value ?? '').replace(/\D/g, '');
        if (!nombre || numero.length !== 4) return msg(_ih('visor.error.partida'), 'error');
        const yo = ($('lu-yo')?.value ?? LS.get('lot_nombre', '')).trim();
        if ($('lu-yo')) LS.set('lot_nombre', yo);
        msg('⏳ ' + _ih('visor.lot.entrando'));
        const llave = `lot_jugador:${Texto.pliegue(nombre)}:${numero}`;
        try {
            const previo = LS.get(llave);
            let r;
            if (previo) {
                r = await Datos.api({ action: 'lot_info', nombre, numero });
                if (r.status !== 200) throw new Error(r.error);
                Object.assign(S, { tablaN: previo.tabla, jugador: previo.jugador, marcas: new Set(previo.ronda === r.estado.ronda ? previo.marcas : []), reclamadas: new Set(previo.ronda === r.estado.ronda ? previo.reclamadas : []) });
            } else {
                r = await Datos.post({ accion: 'lot_unirse', nombre, numero, jugador: yo });
                if (r.status !== 200) throw new Error(r.error);
                Object.assign(S, { tablaN: r.tabla, jugador: r.jugador, marcas: new Set(), reclamadas: new Set() });
            }
            if (r.config.ids.some(id => !w(id))) throw new Error(_i('visor.lot.error.otroCatalogo'));
            Object.assign(S, { rol: 'jugador', llave, cfg: r.config, nombre: r.config.nombre, numero, enLinea: true, ronda: r.estado.ronda, orden: LN.ordenCanto(r.config, r.estado.ronda),
                n: r.estado.n, fin: r.estado.fin, ganadores: r.estado.ganadores || [], vistos: new Set((r.estado.ganadores || []).map(g => r.estado.ronda + '|' + g.variante + '|' + g.tabla)),
                fallos: 0, manualN: false, lenguaVista: LS.get('lot_lengua', '') && w(r.config.ids[0])?.t[LS.get('lot_lengua', '')] ? LS.get('lot_lengua', '') : r.config.lengua });
            S.tabla = LN.tabla(S.cfg, S.tablaN);
            guardarJugador(); pintarJugador();
            S.poll = setInterval(sincronizarJugador, 3000);
            Pantalla.mantener();
        } catch (e) {
            if (!$('lu-nom')) pantallaUnirse(nombre, numero);
            msg(_ih('visor.lot.error.entrar', { x: e.message || _i('visor.sinConexion') }), 'error');
        }
    }
    function guardarJugador() { LS.set(S.llave, { tabla: S.tablaN, jugador: S.jugador, ronda: S.ronda, marcas: [...S.marcas], reclamadas: [...S.reclamadas] }); }
    /** Lenguas en que existen todas las cartas de la tabla. */
    function lenguasTabla() { return Datos.langs.filter(l => S.tabla.every(id => w(id)?.t[l.id])); }
    function pintarJugador(aviso = '') {
        const c = S.cfg, id = S.n ? S.orden[S.n - 1] : null, L = S.lenguaVista;
        const lenguas = c.lenguasJugador && lenguasTabla().length > 1 ? `<div class="fila" style="justify-content:center;margin:6px 0"><label for="lj-l" class="nota"><b>${_ih('visor.lot.verEn')}</b></label>
            <select id="lj-l" class="modern-select" style="width:auto">${lenguasTabla().map(l => `<option value="${esc(l.id)}" ${l.id === L ? 'selected' : ''}>${esc(l.nombre)}</option>`).join('')}</select></div>` : '';
        UI.cuerpo(`<div style="text-align:center"><b>${_ih('visor.lot.tablaN', { n: S.tablaN })}</b> · ${esc(S.jugador)} · <span class="nota">${esc(S.nombre)} ${esc(S.numero)} · ${_ih('visor.lot.rondaX', { x: S.ronda })}</span></div>
            <div class="lot-variantes" id="lj-vars">${variantesHTML()}</div>
            <div class="lot-salio" id="lj-salio" aria-live="polite">${salioHTML(id)}</div>
            <div id="lj-con"></div>${lenguas}
            <div class="lot-tabla" id="lj-tabla" role="group" aria-label="${esc(_i('visor.lot.miTabla'))}">${S.tabla.map((cid, i) => { const m = S.marcas.has(i), nom = nombreCarta(cid, L);
                return `<button class="lot-celda ${m ? 'marcada' : ''}" data-i="${i}" aria-pressed="${m}" aria-label="${esc(m ? _i('visor.lot.marcada', { x: nom }) : nom)}"><span class="n">${numeroCarta(c, cid)}</span><img src="${esc(imagen(cid, 250))}" alt=""><span class="nm"${Texto.lang(L)}>${esc(nom)}</span></button>`; }).join('')}</div>
            <div id="lj-aviso">${aviso}</div>
            ${S.fin ? '<div class="round-result win">🏁 ' + _ih('visor.lot.esperaSiguiente') + '</div>' : ''}
            <p class="nota" style="text-align:center">${_ih('visor.lot.frijol')}</p>`);
        $('lj-tabla').addEventListener('click', e => {
            const b = e.target.closest('.lot-celda'); if (!b || S.fin) return;
            const i = +b.dataset.i; S.marcas.has(i) ? S.marcas.delete(i) : S.marcas.add(i);
            b.classList.toggle('marcada', S.marcas.has(i)); b.setAttribute('aria-pressed', S.marcas.has(i));
            guardarJugador(); revisarMia();
        });
        $('lj-l')?.addEventListener('change', e => { S.lenguaVista = e.target.value; LS.set('lot_lengua', S.lenguaVista); pintarJugador(); });
        pintarConexion();
    }
    function salioHTML(id) {
        if (!id) return '<span class="nota">' + _ih('visor.lot.primera') + '</span>';
        const p = w(id);
        return `<img src="${esc(imagen(id, 150))}" alt=""><div><div class="nota">${_ih('visor.lot.salio', { n: S.n, total: S.orden.length })}</div>
            <b style="font-size:1.15em"${Texto.lang(S.lenguaVista)}>${esc(nombreCarta(id, S.lenguaVista))}</b> ${Sonido.boton(p, S.cfg.lengua)}</div>`;
    }
    function pintarConexion() {
        const el = $('lj-con'); if (!el) return;
        el.innerHTML = S.fallos >= 3 ? `<div class="caja">📶 ${_ih('visor.lot.sinSenal')}
            <input type="number" min="0" max="${S.orden.length}" value="${S.n}" class="input-box" style="width:100px;margin-top:6px" id="lj-n" aria-label="${esc(_i('visor.lot.cartaCantor'))}"></div>` : '';
        $('lj-n')?.addEventListener('change', e => { S.n = Math.max(0, Math.min(S.orden.length, parseInt(e.target.value, 10) || 0)); S.manualN = true; $('lj-salio').innerHTML = salioHTML(S.n ? S.orden[S.n - 1] : null); });
    }
    async function sincronizarJugador() {
        try {
            const r = await Datos.api({ action: 'lot_estado', nombre: S.nombre, numero: S.numero }, 8000);
            if (r.status !== 200) throw new Error(r.error);
            const antes = S.fallos; S.fallos = 0; if (antes >= 3) pintarConexion();
            const e = r.estado;
            if (e.ronda !== S.ronda) {   // nueva ronda: mismas tablas, otro orden, tabla limpia
                Object.assign(S, { ronda: e.ronda, orden: LN.ordenCanto(S.cfg, e.ronda), n: e.n, fin: e.fin, ganadores: e.ganadores || [], marcas: new Set(), reclamadas: new Set(), manualN: false });
                guardarJugador(); pintarJugador('<div class="caja">🔁 ' + _ih('visor.lot.rondaNueva') + '</div>'); UI.toast(_i('visor.lot.nuevaRondaCorto'));
                return true;
            }
            const cambio = e.n !== S.n;
            if (!S.manualN || e.n >= S.n) { S.n = e.n; S.manualN = false; }
            S.ganadores = e.ganadores || []; avisarGanadores();
            const v = $('lj-vars'); if (v) v.innerHTML = variantesHTML();
            if (cambio) { const s = $('lj-salio'); if (s) s.innerHTML = salioHTML(S.n ? S.orden[S.n - 1] : null); }
            if (e.fin !== S.fin) { S.fin = e.fin; pintarJugador(); }
            return true;
        } catch (e) { S.fallos++; if (S.fallos === 3) pintarConexion(); return false; }
    }
    let revisando = false;
    async function revisarMia() {
        if (revisando || S.fin) return;
        revisando = true;
        try {
            for (const v of S.cfg.variantes) {
                if (S.reclamadas.has(v) || S.ganadores.some(g => g.variante === v)) continue;
                const pat = LN.VARIANTES[v].patrones.find(p => p.every(i => S.marcas.has(i)));
                if (!pat) continue;
                if (!S.manualN) await sincronizarJugador();
                const salidas = new Set(S.orden.slice(0, S.n)), faltan = pat.filter(i => !salidas.has(S.tabla[i]));
                if (faltan.length) {
                    faltan.forEach(i => S.marcas.delete(i)); guardarJugador(); SoundFX.error();
                    return pintarJugador(`<div class="caja">${_ih('visor.lot.casi', { x: faltan.map(i => nombreCarta(S.tabla[i], S.lenguaVista)).join(', ') })}</div>`);
                }
                S.reclamadas.add(v); guardarJugador();
                try {
                    const r = await Datos.post({ accion: 'lot_ganar', nombre: S.nombre, numero: S.numero, tabla: S.tablaN, jugador: S.jugador, variante: v, ronda: S.ronda });
                    if (r.status === 200) {
                        S.ganadores = r.estado.ganadores || S.ganadores; S.fin = !!r.estado.fin; avisarGanadores();
                        if (!r.nuevo) { const g = S.ganadores.find(x => x.variante === v); pintarJugador(`<div class="caja">${_ih('visor.lot.primero', { forma: nombreVariante(v), nombre: g ? g.nombre : _i('visor.alguien') })}</div>`); }
                        else pintarJugador();
                    } else if (r.status === 409) sincronizarJugador();
                } catch (e) { UI.anunciar(_i('visor.lot.grita', { forma: nombreVariante(v) })); }
            }
        } finally { revisando = false; }
    }

    function revisarEnlace() {
        const p = new URLSearchParams(location.search), nombre = p.get('loteria') || '', numero = (p.get('n') || '').replace(/\D/g, '');
        if (!nombre) return;
        if (LS.get(`lot_jugador:${Texto.pliegue(nombre)}:${numero}`)) { activo(); UI.abrir('<div id="lot-msg"></div>'); unirse(nombre, numero); }
        else pantallaUnirse(nombre, numero);
    }

    /* =============================================
       5. ARCHIVO PARA JUGAR SIN INTERNET
       ============================================= */
    /** Imágenes ligeras incrustadas: primero las pide al catálogo; si no puede, intenta desde el navegador. */
    async function imagenesLigeras(ids, avance) {
        const out = {};
        let pend = ids.slice(), vueltas = 0;
        while (pend.length && vueltas < 6) {
            vueltas++;
            const lote = pend.slice(0, 30);
            try {
                const r = await Datos.api({ action: 'paquete', ids: lote.join(','), ancho: 320 }, 240000);
                if (r.status !== 200) throw new Error(r.error);
                Object.assign(out, r.imagenes || {});
                const faltan = new Set(r.faltan || []);
                pend = pend.filter(id => !out[id] && (faltan.has(id) || !lote.includes(id)));
                lote.filter(id => !out[id] && !faltan.has(id)).forEach(id => out[id] = '');   // sin imagen en Drive
            } catch (e) { break; }
            avance(Object.values(out).filter(Boolean).length);
        }
        for (const id of ids) {
            if (out[id]) continue;
            try {   // respaldo: copiar la imagen pública desde el navegador
                const b = await (await fetch(Datos.img(w(id), 320))).blob();
                out[id] = await new Promise(res => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = () => res(''); fr.readAsDataURL(b); });
            } catch (e) { out[id] = ''; }
            avance(Object.values(out).filter(Boolean).length);
        }
        return out;
    }
    async function archivoSinConexion() {
        const c = JSON.parse(JSON.stringify(B));
        msg(`⏳ ${_ih('visor.lot.preparando')} <span id="l-av">0</span>/${c.ids.length}`);
        const imgs = await imagenesLigeras(c.ids, n => { const a = $('l-av'); if (a) a.textContent = n; });
        const sin = c.ids.filter(id => !imgs[id]).length;
        const lenguas = {}, palabras = {};
        c.ids.forEach(id => {
            const p = w(id); palabras[id] = { t: {}, img: imgs[id] || '' };
            Object.entries(p.t).forEach(([L, t]) => { if (L === c.lengua || c.lenguasJugador) { palabras[id].t[L] = t.texto; lenguas[L] = Datos.nombreLengua(L); } });
        });
        // Textos del archivo en el idioma de menús elegido (el archivo no carga traducciones)
        const T = {}; Object.keys(V_ES).filter(x => x.startsWith('visor.lot.') || x === 'visor.act.loteria').forEach(x => T[x] = (I18n.dic || V_ES)[x]);
        const D = { cfg: c, palabras, lenguas, titulo: Datos.cfg.titulo || '', creado: new Date().toISOString().slice(0, 10), T, lang: document.documentElement.lang || 'es' };
        const html = `<!DOCTYPE html><html lang="${esc(D.lang)}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(_i('visor.act.loteria'))} ${esc(c.nombre)} · ${esc(c.semilla)}</title><style>${CSS_OFFLINE}</style></head><body><div id="app"></div>
<script>
const D = ${JSON.stringify(D).replace(/</g, '\\u003c')};
${LOT_NUCLEO.toString()}
(${LOT_APP_OFFLINE.toString()})(D, LOT_NUCLEO());
<\/script></body></html>`;
        descargar(`loteria-${c.nombre.replace(/[^\p{L}\p{N}]+/gu, '-').toLowerCase()}-${c.semilla}.html`, html);
        msg('✅ ' + _ih('visor.lot.listo', { mb: (html.length / 1048576).toFixed(1) }) + (sin ? '<br>⚠️ ' + _ih('visor.lot.sinImagen', { n: sin }) : ''));
    }

    return { abrir, pantallaUnirse, unirse, revisarEnlace, _S: S, _B: () => B, _setB: c => { B = c; } };
})();

const CSS_OFFLINE = `*{box-sizing:border-box}body{margin:0;font-family:system-ui,Arial,sans-serif;background:#f8fafc;color:#0f172a}
header{background:#1e3a5f;color:#fff;padding:10px 14px}header h1{margin:0;font-size:1.15rem}header p{margin:2px 0 0;font-size:.8rem;opacity:.85}
nav{display:flex;gap:4px;padding:6px;background:#e2e8f0;position:sticky;top:0;z-index:3}nav button{flex:1;border:0;border-radius:10px;padding:10px 4px;font-weight:800;font-size:.95rem;background:#fff;cursor:pointer}
nav button[aria-current="true"]{background:#2563eb;color:#fff}main{padding:12px;max-width:720px;margin:0 auto}
.btn{border:0;border-radius:12px;padding:12px 16px;font-weight:800;font-size:1rem;cursor:pointer;background:#2563eb;color:#fff;min-height:48px}.btn.s{background:#fff;color:#0f172a;border:2px solid #cbd5e1}
.fila{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:8px 0}input,select{font-size:1.1rem;padding:10px;border-radius:10px;border:2px solid #cbd5e1;min-height:46px}
.carta{background:#fff;border:3px solid #111;border-radius:14px;padding:10px;width:min(260px,70vw);aspect-ratio:65/101;margin:10px auto;display:flex;flex-direction:column;gap:6px}
.carta .im{flex:1;min-height:0;display:flex;align-items:center;justify-content:center}.carta img{max-width:100%;max-height:100%;object-fit:contain}
.nom{border:2px solid #111;border-radius:8px;padding:4px;text-align:center;font-weight:800}.num{font-weight:900;font-size:1.3rem}
.tabla{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}.celda{position:relative;background:#fff;border:2px solid #555;border-radius:10px;padding:3px;aspect-ratio:3/4;display:flex;flex-direction:column;cursor:pointer;min-width:0}
.celda img{width:100%;flex:1;min-height:0;object-fit:contain}.celda .n{position:absolute;top:2px;left:4px;font-size:.7rem;font-weight:900;background:#fff;border-radius:4px;padding:0 3px}
.celda .t{font-size:clamp(.55rem,2.6vw,.8rem);font-weight:800;text-align:center;word-break:break-word;line-height:1.1}
.celda.m::after{content:'🫘';position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:2.4rem;background:rgba(37,99,235,.22);border-radius:8px}
.celda.si{box-shadow:0 0 0 3px #16a34a}.celda.no{opacity:.4}.nota{color:#475569;font-size:.88rem}.ok{color:#16a34a;font-weight:800}
.hist{display:flex;flex-wrap:wrap;gap:6px}.hist figure{margin:0;width:62px;font-size:.66rem;text-align:center}.hist img{width:62px;height:62px;object-fit:cover;border-radius:8px;border:2px solid #cbd5e1;background:#fff}
.ph{display:flex;align-items:center;justify-content:center;font-size:2rem;font-weight:900;color:#94a3b8;flex:1}`;

/** Aplicación del archivo sin conexión (se copia tal cual dentro del archivo descargado). */
function LOT_APP_OFFLINE(D, LN) {
    const c = D.cfg, $ = id => document.getElementById(id);
    const e = s => String(s ?? '').replace(/[&<>"']/g, x => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[x]));
    /** Texto de la interfaz guardado en el archivo: escapado, *negritas*, {valores}. */
    const T = (k, v) => e((D.T || {})[k] || k).replace(/\{(\w+)\}/g, (m, x) => v && x in v ? e(v[x]) : m).replace(/\*([^*\n]+)\*/g, '<b>$1</b>');
    const Tp = (k, v) => String((D.T || {})[k] || k).replace(/\{(\w+)\}/g, (m, x) => v && x in v ? String(v[x]) : m).replace(/\*([^*\n]+)\*/g, '$1');
    const variante = v => T('visor.lot.var.' + v);
    const ls = { get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (x) { return d; } }, set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (x) {} } };
    const K = 'lotoff:' + c.semilla;
    const st = Object.assign({ vista: 'cantar', ronda: String(Math.floor(Math.random() * 10000)).padStart(4, '0'), n: 0, tabla: 0, marcas: [], L: c.lengua, seg: 0 }, ls.get(K, {}));
    const guardar = () => ls.set(K, st);
    const nombre = (id, L = st.L) => D.palabras[id].t[L] || D.palabras[id].t[c.lengua] || '?';
    const img = (id, cls = '') => D.palabras[id].img ? `<img src="${D.palabras[id].img}" alt="">` : `<div class="ph">${c.ids.indexOf(id) + 1}</div>`;
    const num = id => c.ids.indexOf(id) + 1;
    let orden = LN.ordenCanto(c, st.ronda), reloj = null;
    document.body.insertAdjacentHTML('afterbegin', `<header><h1>🎴 ${T('visor.act.loteria')} · ${e(c.nombre)}</h1><p>${e(D.titulo)} · ${T('visor.lot.juegoN', { x: c.semilla })} · ${T('visor.lot.off.cartas', { n: c.ids.length })}</p></header>
        <nav><button data-v="cantar">📣 ${T('visor.lot.off.cantar')}</button><button data-v="jugar">🎴 ${T('visor.lot.miTabla')}</button><button data-v="verificar">🔍 ${T('visor.lot.off.verificar')}</button></nav>`);
    document.querySelector('nav').addEventListener('click', ev => { const b = ev.target.closest('button'); if (b) { st.vista = b.dataset.v; guardar(); pinta(); } });
    function pinta() {
        clearInterval(reloj); reloj = null;
        document.querySelectorAll('nav button').forEach(b => b.setAttribute('aria-current', b.dataset.v === st.vista));
        ({ cantar, jugar, verificar })[st.vista]();
    }
    function cantar() {
        const id = st.n ? orden[st.n - 1] : null;
        $('app').innerHTML = `<main><div class="nota" style="text-align:center">${T('visor.lot.contador', { ronda: st.ronda, n: st.n, total: orden.length })}</div>
            ${id ? `<div class="carta"><div class="num">${num(id)}</div><div class="im">${img(id)}</div><div class="nom">${e(nombre(id, c.lengua))}</div></div>` : '<p style="text-align:center">' + T('visor.lot.toca') + '</p>'}
            <div class="fila" style="justify-content:center"><button class="btn" id="sig" style="flex:1;min-height:64px" ${st.n >= orden.length ? 'disabled' : ''}>▶ ${T('visor.lot.siguiente')}</button></div>
            <div class="fila" style="justify-content:center"><label>${T('visor.lot.avanzarSolo')} <select id="seg"><option value="0">—</option>${[5, 8, 10, 15, 20].map(s => `<option ${st.seg == s ? 'selected' : ''}>${s}</option>`).join('')}</select> s</label></div>
            <details><summary>${T('visor.lot.yaSalieron', { n: Math.max(0, st.n - 1) })}</summary><div class="hist">${orden.slice(0, Math.max(0, st.n - 1)).reverse().map(i => `<figure>${D.palabras[i].img ? `<img src="${D.palabras[i].img}" alt="">` : num(i)}<figcaption>${e(nombre(i, c.lengua))}</figcaption></figure>`).join('')}</div></details>
            <div class="fila"><button class="btn s" id="nueva">🔁 ${T('visor.lot.nuevaRondaCorto')}</button><label class="nota">${T('visor.lot.off.retomar')} <input id="rr" inputmode="numeric" maxlength="4" style="width:90px" value="${e(st.ronda)}"></label></div></main>`;
        $('sig').onclick = () => { if (st.n < orden.length) { st.n++; guardar(); cantar(); } };
        $('seg').onchange = ev => { st.seg = +ev.target.value; guardar(); cantar(); };
        $('nueva').onclick = () => { if (st.n && !confirm(Tp('visor.lot.off.confirma'))) return; st.ronda = String(Math.floor(Math.random() * 10000)).padStart(4, '0'); st.n = 0; orden = LN.ordenCanto(c, st.ronda); guardar(); cantar(); };
        $('rr').onchange = ev => { const r = ev.target.value.replace(/\D/g, ''); if (r.length === 4) { st.ronda = r; orden = LN.ordenCanto(c, r); st.n = 0; guardar(); cantar(); } };
        if (st.seg && st.n < orden.length) reloj = setTimeout(() => { st.n++; guardar(); cantar(); }, st.seg * 1000);
    }
    function jugar() {
        if (!st.tabla) {
            $('app').innerHTML = `<main><p>${T('visor.lot.off.numero')}</p>
                <div class="fila"><input id="tn" type="number" min="1" inputmode="numeric" style="width:120px" aria-label="${T('visor.lot.numTabla')}"><button class="btn" id="ok">${T('visor.lot.off.usar')}</button><button class="btn s" id="az">🎲 ${T('visor.lot.off.azar')}</button></div></main>`;
            const usar = n => { if (n > 0) { st.tabla = n; st.marcas = []; guardar(); jugar(); } };
            $('ok').onclick = () => usar(parseInt($('tn').value, 10));
            $('az').onclick = () => usar(1 + (c.impresas || 0) + Math.floor(Math.random() * 500));
            return;
        }
        const t = LN.tabla(c, st.tabla), m = new Set(st.marcas);
        const lenguas = Object.keys(D.lenguas).filter(L => t.every(id => D.palabras[id].t[L]));
        $('app').innerHTML = `<main><div class="fila" style="justify-content:space-between"><b>${T('visor.lot.tablaN', { n: st.tabla })}</b>
                ${lenguas.length > 1 ? `<select id="lj">${lenguas.map(L => `<option value="${e(L)}" ${L === st.L ? 'selected' : ''}>${e(D.lenguas[L])}</option>`).join('')}</select>` : ''}</div>
            <div class="tabla" id="tb">${t.map((id, i) => `<div class="celda ${m.has(i) ? 'm' : ''}" data-i="${i}"><span class="n">${num(id)}</span>${img(id)}<span class="t">${e(nombre(id))}</span></div>`).join('')}</div>
            <p class="nota">${T('visor.lot.off.frijol')}</p>
            <div class="fila"><button class="btn s" id="limpia">${T('visor.lot.off.quitar')}</button><button class="btn s" id="otra">${T('visor.lot.off.cambiar')}</button></div></main>`;
        $('tb').onclick = ev => { const d = ev.target.closest('.celda'); if (!d) return; const i = +d.dataset.i; m.has(i) ? m.delete(i) : m.add(i); st.marcas = [...m]; guardar(); d.classList.toggle('m', m.has(i)); };
        if ($('lj')) $('lj').onchange = ev => { st.L = ev.target.value; guardar(); jugar(); };
        $('limpia').onclick = () => { st.marcas = []; guardar(); jugar(); };
        $('otra').onclick = () => { st.tabla = 0; guardar(); jugar(); };
    }
    function verificar(k) {
        let res = '';
        if (k > 0) {
            const t = LN.tabla(c, k), sal = orden.slice(0, st.n), s = new Set(sal), rev = LN.revisar(c, t, sal);
            res = `<p>${rev.map(x => `<span class="${x.completo ? 'ok' : ''}">${x.completo ? '✅' : '—'} ${variante(x.v)}</span>`).join(' · ')}</p>
                <div class="tabla">${t.map(id => `<div class="celda ${s.has(id) ? 'si' : 'no'}"><span class="n">${num(id)}</span>${img(id)}<span class="t">${e(nombre(id, c.lengua))}</span></div>`).join('')}</div>
                <p class="nota">${T('visor.lot.off.colores', { ronda: st.ronda, n: st.n })} ${rev.some(x => x.completo) ? '<b class="ok">' + T('visor.lot.gana') + '</b>' : T('visor.lot.noGana')}</p>`;
        }
        $('app').innerHTML = `<main><p>${T('visor.lot.off.revisa')}</p>
            <div class="fila"><input id="vn" type="number" min="1" inputmode="numeric" placeholder="${T('visor.lot.numTabla')}" aria-label="${T('visor.lot.numTabla')}" style="width:140px" value="${k || ''}"><button class="btn" id="vk">${T('visor.lot.off.revisar')}</button></div>${res}
            <p class="nota">${T('visor.lot.formasGanar')}: ${c.variantes.map(variante).join(', ')}.</p></main>`;
        $('vk').onclick = () => verificar(parseInt($('vn').value, 10));
    }
    pinta();
}
</script>
