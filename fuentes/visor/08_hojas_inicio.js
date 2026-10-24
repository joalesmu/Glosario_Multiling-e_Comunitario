<script>
/* =============================================
   HOJAS PARA IMPRIMIR (carta vertical, 5 mm de margen)
   Las instrucciones se pueden escribir en la lengua de la comunidad.
   ============================================= */
const Hojas = {
    TIPOS: {
        une:       { ico: '🔗', por: 6, inst: true, pistas: ['imagen', 'lengua', 'descripcion', 'propia'], resp: true },
        completa:  { ico: '🔡', por: 6, inst: true, pistas: ['imagen', 'lengua', 'descripcion', 'propia', 'ninguna'], resp: true, largas: true },
        escribe:   { ico: '✏️', por: 6, inst: true, pistas: ['imagen', 'lengua', 'descripcion', 'propia', 'espacio'], resp: true },
        dibuja:    { ico: '🎨', por: 4, inst: true, pistas: ['ninguna', 'lengua', 'descripcion', 'propia'] },
        etiquetas: { ico: '🏷️', por: 6, act: true, img: true, pistas: ['ninguna', 'lengua', 'descripcion', 'propia', 'espacio'], segunda: true },
        tarjetas:  { ico: '🗂️', por: 8, act: true, img: true, pistas: ['ninguna', 'lengua', 'descripcion', 'propia', 'espacio'], segunda: true }
    },
    nombre(k) { return this.TIPOS[k] ? _i((this.TIPOS[k].act ? 'visor.act.' : 'visor.hoja.') + k) : String(k); },
    /** Instrucción por omisión: la que el docente dejó la última vez, o una que corresponde a la pista elegida. */
    instruccion(k, modo = 'imagen') {
        if (!this.TIPOS[k].inst) return '';
        const propia = LS.get('hoja:inst:' + k, '');
        if (propia) return propia;
        const conPista = !['imagen', 'imagenEspacio', 'ninguna'].includes(modo) && V_ES['visor.hoja.' + k + '.inst.pista'];
        return _i('visor.hoja.' + k + '.inst' + (conPista ? '.pista' : ''));
    },
    instruccionesPred(k) { return [_i('visor.hoja.' + k + '.inst'), V_ES['visor.hoja.' + k + '.inst.pista'] ? _i('visor.hoja.' + k + '.inst.pista') : null].filter(Boolean); },
    /** Palabras posibles de una actividad guardada (para llevarlas sin conexión). */
    /** Ventana de opciones. `c` y `codigo` llegan al reabrir desde Mis actividades: con las mismas opciones sale la misma hoja. */
    abrir(tipo = 'une', c = null, codigo = '') {
        c = c || this.ultimo || {};
        if (codigo) this.base = { codigo, config: JSON.stringify(c) }; else if (!c.tipo) this.base = null;
        tipo = c.tipo || tipo;
        const L0 = c.L && Datos.langById.has(c.L) ? c.L : Datos.lengua;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🖨️ ${_ih('visor.hoja.titulo')}</h2>
            ${this.base ? `<div class="caja">🔁 ${_ih('visor.hoja.reabierta', { c: this.base.codigo })}</div>` : ''}
            <span class="t">${_ih('visor.hoja.cual')}</span>
            <div class="chips">${Object.entries(this.TIPOS).map(([k, t]) => `<label class="chip"><input type="radio" name="h-t" value="${k}" ${k === tipo ? 'checked' : ''}> ${t.ico} ${esc(this.nombre(k))}</label>`).join('')}</div>
            <label class="t" for="h-l">${_ih('visor.lengua')}</label><select id="h-l" class="modern-select">${Comun.lenguas({ campo: 'textos', min: 4 }, L0)}</select>
            <div id="h-sel">${Seleccion.html('hs', L0, c.sel || {})}</div>
            <div id="h-pis"></div>
            <div id="h-instb"><label class="t" for="h-inst">${_ih('visor.hoja.instruccion')} <span class="nota">${_ih('visor.hoja.instruccion.d')}</span></label>
            <input id="h-inst" class="input-box" maxlength="200"></div>
            <div id="h-calcar" class="oculto"><label class="chip"><input type="checkbox" id="h-calca" ${c.calca ? 'checked' : ''}> ${_ih('visor.hoja.calcar')}</label></div>
            <fieldset class="grupo"><legend>📄 ${_ih('visor.hoja.cuantas')}</legend>
                <span class="t">${_ih('visor.hoja.versiones')}</span><div class="chips">${[1, 2, 3].map(n => `<label class="chip"><input type="radio" name="h-v" value="${n}" ${(c.versiones || 1) === n ? 'checked' : ''}> ${n === 1 ? _ih('visor.hoja.unaVersion') : ['A', 'B', 'C'].slice(0, n).join(' · ')}</label>`).join('')}</div>
                <p class="nota">${_ih('visor.hoja.versiones.d')}</p>
                <span class="t">${_ih('visor.hoja.hojas')}</span><div class="chips">${[1, 2, 3, 5, 10].map(n => `<label class="chip"><input type="radio" name="h-n" value="${n}" ${(c.n || 1) === n ? 'checked' : ''}> ${n}</label>`).join('')}</div>
                <div id="h-parab"><span class="t">${_ih('visor.hoja.para')}</span><div class="chips">${['estudiante', 'docente', 'ambas'].map(k => `<label class="chip"><input type="radio" name="h-p" value="${k}" ${(c.para || 'ambas') === k ? 'checked' : ''}> ${_ih('visor.hoja.para.' + k)}</label>`).join('')}</div>
                <p class="nota">${_ih('visor.hoja.para.d')}</p></div></fieldset>
            ${Impresion.html()}
            <div id="lot-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="h-ok">🖨️ ${_ih('visor.hoja.preparar')}</button>
            <button class="btn btn-secondary" id="h-off">📥 ${_ih('visor.off.boton')}</button></div></div>`);
        const volver = () => this.abrir(this.leerOpciones().tipo, this.leerOpciones(), this.base ? this.base.codigo : '');
        const sel = Seleccion.enlazar('hs', null, volver);
        const pintarPistas = () => {
            const t = document.querySelector('input[name="h-t"]:checked').value, T = this.TIPOS[t];
            const prev = $('hp-pl') ? Pistas.leer('hp') : (c.pista || {});
            $('h-pis').innerHTML = Pistas.html('hp', $('h-l').value, T.pistas, prev).replace('💡 ' + _ih('visor.pista.titulo'), '💡 ' + _ih(T.segunda ? 'visor.pista.segunda' : 'visor.pista.titulo'));
            this._pis = Pistas.enlazar('hp', () => $('h-l').value);
            // Si la instrucción no la escribió el docente, se ajusta a la pista elegida
            document.querySelectorAll('input[name="hp-m"]').forEach(r => r.addEventListener('change', () => {
                const t2 = document.querySelector('input[name="h-t"]:checked').value;
                if (this.instruccionesPred(t2).includes($('h-inst').value.trim())) $('h-inst').value = this.instruccion(t2, Pistas.leer('hp').modo);
            }));
        };
        const ajustar = (inicial) => {
            const t = document.querySelector('input[name="h-t"]:checked').value, T = this.TIPOS[t];
            $('h-instb').classList.toggle('oculto', !T.inst);
            $('h-calcar').classList.toggle('oculto', t !== 'escribe');
            $('h-parab').classList.toggle('oculto', !T.resp);
            pintarPistas();
            if (!inicial || !c.inst) $('h-inst').value = this.instruccion(t, Pistas.leer('hp').modo); else $('h-inst').value = c.inst;
        };
        document.querySelectorAll('input[name="h-t"]').forEach(r => r.addEventListener('change', () => ajustar(false))); ajustar(true);
        $('h-l').addEventListener('change', () => { sel.recategorizar($('h-l').value); if (this._pis) this._pis.ajustar(); });
        $('h-ok').addEventListener('click', () => this.generar());
        $('h-off').addEventListener('click', () => { const o = this.leerOpciones(); Offline.panel(Seleccion.palabras(o.sel, o.L), this.nombre(o.tipo), volver); });
    },
    leerOpciones() {
        const tipo = document.querySelector('input[name="h-t"]:checked').value;
        return { tipo, L: $('h-l').value, sel: Seleccion.leer('hs'), pista: Pistas.leer('hp'), inst: $('h-inst').value.trim(),
            calca: $('h-calca').checked, versiones: +(document.querySelector('input[name="h-v"]:checked')?.value || 1),
            n: +(document.querySelector('input[name="h-n"]:checked')?.value || 1), para: this.TIPOS[tipo].resp ? (document.querySelector('input[name="h-p"]:checked')?.value || 'estudiante') : 'estudiante' };
    },
    /**
     * Arma el contenido con azar reproducible (código + versión + lengua) y lo pinta dos veces:
     * para estudiantes y, si se pide, para el docente con las respuestas.
     */
    generar() {
        const c = this.leerOpciones(), T = this.TIPOS[c.tipo], L = c.L, aj = Impresion.leer();
        if (T.inst && !this.instruccionesPred(c.tipo).includes(c.inst)) LS.set('hoja:inst:' + c.tipo, c.inst); else LS.del('hoja:inst:' + c.tipo);
        this.ultimo = c;
        const pis = c.pista;
        let pool = Seleccion.palabras(c.sel, L, { img: !!T.img || Pistas.necesitaImagen(pis), largas: !!T.largas }).filter(w => Pistas.sirve(w, pis, L));
        if (pool.length < Math.min(4, T.por)) return msg(_ih('visor.hoja.pocas'));
        const codigo = this.base && this.base.config === JSON.stringify(c) ? this.base.codigo : Actividades.nuevoCodigo();
        const versiones = ['A', 'B', 'C'].slice(0, c.versiones);
        const modelo = versiones.map(v => {
            const r = azar(`${codigo}|${v}|${L}|${c.tipo}`), orden = barajar(pool, r);
            return { v, paginas: Array.from({ length: c.n }, (_, i) => {
                const ws = []; for (let k = 0; k < T.por && ws.length < orden.length; k++) { const w = orden[(i * T.por + k) % orden.length]; if (!ws.includes(w)) ws.push(w); }
                const p = { ws };
                if (c.tipo === 'une') p.der = barajar(ws, r);
                if (c.tipo === 'completa') p.quitar = ws.map(w => { const g = Texto.grafemas(w.t[L].texto, L), letras = g.map((x, j) => x.letra ? j : -1).filter(j => j >= 0);
                    return barajar(letras, r).slice(0, Math.max(1, Math.round(letras.length / 3))); });
                return p;
            }) };
        });
        const usadas = [...new Set(modelo.flatMap(m => m.paginas.flatMap(p => p.ws)))];
        let pags = '';
        const ver = v => c.versiones > 1 ? v : '';
        if (c.para !== 'docente') modelo.forEach(m => m.paginas.forEach(p => { pags += this.pagina(c, p, ver(m.v), codigo, aj, false); }));
        if (T.resp && c.para !== 'estudiante') modelo.forEach(m => m.paginas.forEach(p => { pags += this.pagina(c, p, ver(m.v), codigo, aj, true); }));
        const nombre = `${this.nombre(c.tipo)} · ${Datos.nombreLengua(L)}`;
        Actividades.registrar({ tipo: 'hojas', sub: c.tipo, config: c, codigo, nombre: `${nombre} · ${new Date().toLocaleDateString()}` });
        this.base = { codigo, config: JSON.stringify(c) };
        const sinRed = !navigator.onLine ? Offline.faltantes(usadas) : 0;
        Impresion.ventana({ titulo: `${nombre} · ${codigo}`, paginas: pags, css: this.CSS, ajustes: aj,
            desc: _i('visor.hoja.desc2', { tipo: this.nombre(c.tipo), c: codigo, v: c.versiones, n: c.n }),
            aviso: sinRed ? _i('visor.off.faltanImg', { n: sinRed }) : '' });
    },
    /** Una hoja. `docente` = con las respuestas a la vista, para calificar. */
    pagina(c, p, v, codigo, aj, docente) {
        const L = c.L, T = this.TIPOS[c.tipo], pis = c.pista;
        const pal = (w, X = L) => `<span${Texto.lang(X)}>${esc(w.t[X].texto)}</span>`;
        const pista = w => Pistas.render(w, pis, L);
        const cab = T.act ? '' : `<header><div><b>${esc(T.ico)} </b>${esc(c.inst)}</div>${docente ? `<div class="ver">✔️ ${esc(_i('visor.hoja.docente'))}${v ? ' · ' + esc(_i('visor.hoja.version', { v })) : ''}</div>`
            : `<div class="datos">${v ? `<span class="ver">${esc(_i('visor.hoja.version', { v }))}</span> &nbsp; ` : ''}______________________ &nbsp; ____/____/______</div>`}</header>`;
        let cuerpo = '';
        const ws = p.ws;
        if (c.tipo === 'une') {
            cuerpo = `<div class="une">${ws.map((w, k) => `<div class="ui">${pista(w)}</div><div class="punto">${docente ? `<span class="resp-ok">${p.der.indexOf(w) + 1}</span>` : '●'}</div><div></div><div class="punto">●</div><div class="up"><span class="num">${k + 1}.</span> ${pal(p.der[k])}</div>`).join('')}</div>`;
        } else if (c.tipo === 'completa') {
            cuerpo = `<div class="rej2">${ws.map((w, k) => {
                const g = Texto.grafemas(w.t[L].texto, L), quitar = new Set(p.quitar[k]), ps = pista(w);
                return `<div class="item">${ps ? `<div class="ps">${ps}</div>` : ''}<div class="pal"${Texto.lang(L)}>${g.map((x, j) => quitar.has(j) ? (docente ? `<span class="hue resp-ok">${esc(x.g)}</span>` : '<span class="hue"></span>') : esc(x.g)).join('')}</div></div>`; }).join('')}</div>`;
        } else if (c.tipo === 'escribe') {
            cuerpo = `<div class="rej2">${ws.map(w => `<div class="item"><div class="ps">${pista(w)}</div><div class="lin">${docente ? `<span class="resp-ok"${Texto.lang(L)}>${esc(w.t[L].texto)}</span>` : c.calca ? `<span class="calca"${Texto.lang(L)}>${esc(w.t[L].texto)}</span>` : ''}</div><div class="lin"></div></div>`).join('')}</div>`;
        } else if (c.tipo === 'dibuja') {
            cuerpo = `<div class="rej2 dib">${ws.map(w => { const ps = pista(w); return `<div class="item"><div class="pal">${pal(w)}</div>${ps ? `<div class="ps-txt">${ps}</div>` : ''}<div class="caja-d"></div></div>`; }).join('')}</div>`;
        } else {
            const cls = c.tipo === 'etiquetas' ? 'etq' : 'tar';
            cuerpo = `<div class="rej2 ${cls}">${ws.map(w => { const ps = pista(w); return `<div class="item"><img src="${esc(Datos.img(w, 600))}" alt=""><div class="pal">${pal(w)}</div>${ps ? `<div class="sub">${ps}</div>` : ''}</div>`; }).join('')}</div>`;
        }
        return `<section class="hoja${docente ? ' docente' : ''}">${cab}${cuerpo}${Impresion.pie(aj, { L, codigo, version: v })}</section>`;
    },
    CSS: `.une{flex:1;display:grid;grid-template-columns:48mm 8mm 1fr 8mm 62mm;grid-auto-rows:1fr;align-items:center;row-gap:3mm;min-height:0}
        .une .ui{height:100%;max-height:36mm;display:flex;flex-direction:column;align-items:center;justify-content:center;border:.4mm solid #111;border-radius:2mm;padding:1mm;box-sizing:border-box;overflow:hidden}
        .une .punto{font-size:16pt;text-align:center}.une .up{font-size:calc(20pt*var(--k));font-weight:800}.une .num{font-size:10pt;color:#666;font-weight:600}
        .rej2{flex:1;display:grid;grid-template-columns:1fr 1fr;grid-auto-rows:1fr;gap:4mm;min-height:0}
        .item{border:.5mm solid #111;border-radius:3mm;padding:3mm;display:flex;flex-direction:column;gap:2mm;align-items:center;justify-content:center;min-height:0;overflow:hidden}
        .item img{flex:1;min-height:0}.item .ps{flex:1;min-height:0;width:100%;display:flex;flex-direction:column;align-items:center;justify-content:center}
        .item .ps img{flex:1;min-height:0}
        .pal{font-size:calc(22pt*var(--k));font-weight:800;letter-spacing:.5mm;text-align:center;word-break:break-word}
        .hue{display:inline-block;min-width:9mm;height:11mm;border-bottom:.6mm solid #111;margin:0 .8mm;vertical-align:bottom;text-align:center}
        .lin{width:100%;height:13mm;border-bottom:.4mm solid #111;position:relative;display:flex;align-items:flex-end;justify-content:center;font-size:calc(22pt*var(--k))}
        .calca{font-size:calc(24pt*var(--k));color:transparent;-webkit-text-stroke:.3mm #999;letter-spacing:1mm;font-weight:700}
        .dib .caja-d{flex:1;width:100%;border:.4mm dashed #777;border-radius:2mm}.dib .ps-txt{width:100%}
        .etq .pal{font-size:calc(30pt*var(--k))}.etq .sub,.tar .sub{font-size:calc(14pt*var(--k));color:#333;width:100%;text-align:center}
        .tar{grid-template-rows:repeat(4,1fr)}.tar .pal{font-size:calc(18pt*var(--k))}
        .docente .hue{border-bottom-color:#b91c1c}`
};
function msg(html) { const m = $('lot-msg'); if (m) m.innerHTML = `<div class="caja">${html}</div>`; else UI.toast(String(html).replace(/<[^>]+>/g, '')); }

/* =============================================
   INICIO
   ============================================= */
(function () {
    I18n.init();   // idioma de los menús: lo elegido en este dispositivo, antes de pintar nada
    A11y.init();
    const secs = ['explorar', 'aprender', 'jugar', 'imprimir', 'colabora'];
    const h = (location.hash || '').slice(1);
    UI.seccion = secs.includes(h) ? h : 'explorar';
    document.querySelectorAll('#nav button').forEach(b => b.setAttribute('aria-current', b.dataset.sec === UI.seccion ? 'page' : 'false'));
    window.addEventListener('resize', () => UI.ajustarCabecera());
    document.addEventListener('keydown', e => {
        if (e.key !== 'Escape' || $('modal').style.display !== 'flex') return;
        if ($('lightbox').style.display === 'flex') { $('lightbox').style.display = 'none'; return; }
        if (!Juego.activo || confirm(_i('visor.salirActividad'))) UI.cerrarModal();
    });
    window.addEventListener('online', () => UI.estadoConexion(_i('visor.con.conectado')));
    window.addEventListener('offline', () => UI.estadoConexion(_i('visor.con.sinConexion'), true));
    $('avisos').addEventListener('click', e => Estado.clic(e));
    document.addEventListener('visibilitychange', () => Estado.alVolver());
    window.addEventListener('online', () => Estado.alVolver());
    window.addEventListener('online', () => Propuestas.enviarCola());
    document.addEventListener('visibilitychange', () => { if (!document.hidden) Propuestas.enviarCola(); });
    Instalar.iniciar(); Progreso.pintarBoton();
    Datos.iniciar().then(async () => { if (Datos.cargado) { await Offline.iniciar(); UI.refrescar(); Loteria.revisarEnlace(); Actividades.revisarEnlace(); Propuestas.revisarEnlace(); revisarEnlaceOir(); Propuestas.enviarCola(); } });
})();
</script>
</body>
</html>
