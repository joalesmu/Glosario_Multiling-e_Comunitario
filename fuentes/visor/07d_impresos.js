<script>
/* =============================================
   MATERIALES PARA IMPRIMIR (juegos y recursos en papel)
   Todos usan la misma base: palabras (al azar o de mi lista), pistas configurables,
   versiones A·B·C con hoja del docente, niveles para grupos multigrado, ayudas graduadas,
   ajustes de impresión, código de hoja y Mis actividades.
   Nada decide por el docente: las «Ideas de uso» solo sugieren.
   ============================================= */
const Impresos = {
    /* Cada material: ico, req (para la tarjeta), pistas (modos), resp (hay hoja del docente), niveles, ayudas,
       opciones() → HTML extra, leer() → objeto, modelo(ws, c, rng, nivel) → datos de una hoja, pagina(m, c, docente) → HTML. */
    TIPOS: {},
    nombre(k) { return _i('visor.imp2.' + k); },
    registrar(k, def) { this.TIPOS[k] = def; },
    /** La lista desde la palabra i·k (cada hoja con otras palabras, hasta dar la vuelta). */
    rotar(orden, i, k) { const s = orden.length ? (i * k) % orden.length : 0; return orden.slice(s).concat(orden.slice(0, s)); },
    /** Palabras cuyas letras son todas letras de su alfabeto (crucigramas, sopas y ordenar letras). */
    deLetras(ws, L, min = 3, max = 12) {
        return ws.map(w => ({ w, g: Texto.grafemas(w.t[L].texto, L) })).filter(x => x.g.length >= min && x.g.length <= max && x.g.every(y => y.letra)).map(x => ({ w: x.w, g: x.g.map(y => y.g) }));
    },
    abrir(tipo, c = null, codigo = '') {
        const T = this.TIPOS[tipo]; if (!T) return;
        c = c || (this.ultimo && this.ultimo.tipo === tipo ? this.ultimo : {});
        this.base = codigo ? { codigo, config: JSON.stringify(c) } : (c.tipo ? this.base : null);
        const L0 = c.L && Datos.langById.has(c.L) ? c.L : Datos.lengua;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">${T.ico} ${esc(this.nombre(tipo))}</h2>
            <p class="nota">${_ih('visor.imp2.' + tipo + '.d')}</p>
            ${this.base ? `<div class="caja">🔁 ${_ih('visor.hoja.reabierta', { c: this.base.codigo })}</div>` : ''}
            ${T.sinPalabras ? '' : `<label class="t" for="im-l">${_ih('visor.lengua')}</label><select id="im-l" class="modern-select">${Comun.lenguas({ campo: 'textos', min: 3 }, L0)}</select>
            ${T.sinSeleccion ? '' : `<div id="im-sel">${Seleccion.html('is', L0, c.sel || {})}</div>`}`}
            ${T.pistas ? `<div id="im-pis">${Pistas.html('ip', L0, T.pistas, c.pista || {})}</div>` : ''}
            ${T.opciones ? `<fieldset class="grupo"><legend>⚙️ ${_ih('visor.imp2.opciones')}</legend>${T.opciones(c.op || {}, L0)}</fieldset>` : ''}
            ${T.ayudas ? `<fieldset class="grupo"><legend>🪜 ${_ih('visor.imp2.ayudas')}</legend><div class="chips">${T.ayudas.map(a => `<label class="chip"><input type="checkbox" name="im-ay" value="${a}" ${(c.ayudas || T.ayudasPred || []).includes(a) ? 'checked' : ''}> ${_ih('visor.imp2.ayuda.' + a)}</label>`).join('')}</div>
                ${T.niveles ? `<label class="chip" style="margin-top:8px"><input type="checkbox" id="im-niv" ${c.niveles ? 'checked' : ''}> 🏫 ${_ih('visor.imp2.niveles')}</label><p class="nota">${_ih('visor.imp2.niveles.d')}</p>` : ''}</fieldset>` : ''}
            ${T.inst !== false ? `<label class="t" for="im-inst">${_ih('visor.hoja.instruccion')} <span class="nota">${_ih('visor.hoja.instruccion.d')}</span></label><input id="im-inst" class="input-box" maxlength="200" value="${esc(c.inst != null ? c.inst : (LS.get('imp2:inst:' + tipo, '') || _i('visor.imp2.' + tipo + '.inst')))}">` : ''}
            <fieldset class="grupo"><legend>📄 ${_ih('visor.hoja.cuantas')}</legend>
                ${T.resp || T.versiones ? `<span class="t">${_ih('visor.hoja.versiones')}</span><div class="chips">${[1, 2, 3].map(n => `<label class="chip"><input type="radio" name="im-v" value="${n}" ${(c.versiones || 1) === n ? 'checked' : ''}> ${n === 1 ? _ih('visor.hoja.unaVersion') : ['A', 'B', 'C'].slice(0, n).join(' · ')}</label>`).join('')}</div>` : ''}
                ${T.hojas === false ? '' : `<span class="t">${_ih('visor.hoja.hojas')}</span>`}<div class="chips">${(T.hojas === false ? [] : T.hojas || [1, 2, 3, 5, 10]).map(n => `<label class="chip"><input type="radio" name="im-n" value="${n}" ${(c.n || 1) === n ? 'checked' : ''}> ${n}</label>`).join('')}</div>
                ${T.resp ? `<span class="t">${_ih('visor.hoja.para')}</span><div class="chips">${['estudiante', 'docente', 'ambas'].map(k => `<label class="chip"><input type="radio" name="im-p" value="${k}" ${(c.para || 'ambas') === k ? 'checked' : ''}> ${_ih('visor.hoja.para.' + k)}</label>`).join('')}</div>` : ''}</fieldset>
            ${Impresion.html()}
            <details class="caja"><summary>💡 ${_ih('visor.imp2.ideas')}</summary><ul>${[1, 2, 3].map(i => `<li>${_ih('visor.imp2.' + tipo + '.idea' + i)}</li>`).join('')}</ul><p class="nota">${_ih('visor.imp2.ideas.d')}</p></details>
            <div id="lot-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="im-ok">🖨️ ${_ih('visor.hoja.preparar')}</button>
            ${T.sinPalabras || T.sinSeleccion ? '' : `<button class="btn btn-secondary" id="im-off">📥 ${_ih('visor.off.boton')}</button>`}</div></div>`);
        const volver = () => { const o = this.leer(tipo); this.abrir(tipo, o, this.base ? this.base.codigo : ''); };
        let sel = null, pis = null;
        if (!T.sinPalabras) {
            if (!T.sinSeleccion) sel = Seleccion.enlazar('is', null, volver);
            $('im-l').addEventListener('change', () => { if (sel) sel.recategorizar($('im-l').value); if (pis) pis.ajustar(); if (T.alCambiarLengua) T.alCambiarLengua($('im-l').value); });
            if ($('im-off')) $('im-off').addEventListener('click', () => { const o = this.leer(tipo); Offline.panel(Seleccion.palabras(o.sel, o.L), this.nombre(tipo), volver); });
        }
        if (T.pistas) pis = Pistas.enlazar('ip', () => ($('im-l') || {}).value || Datos.lengua);
        if (T.enlazar) T.enlazar(c.op || {});
        $('im-ok').addEventListener('click', () => this.generar(tipo));
    },
    leer(tipo) {
        const T = this.TIPOS[tipo], v = (n, d) => document.querySelector(`input[name="${n}"]:checked`)?.value || d;
        return { tipo, L: $('im-l') ? $('im-l').value : Datos.lengua, sel: T.sinPalabras || T.sinSeleccion ? { fuente: 'azar', cats: [] } : Seleccion.leer('is'), pista: T.pistas ? Pistas.leer('ip') : null,
            op: T.leer ? T.leer() : {}, ayudas: T.ayudas ? [...document.querySelectorAll('input[name="im-ay"]:checked')].map(x => x.value) : [], niveles: !!($('im-niv') && $('im-niv').checked),
            inst: $('im-inst') ? $('im-inst').value.trim() : '', versiones: +(v('im-v', 1)), n: +(v('im-n', 1)), para: T.resp ? v('im-p', 'estudiante') : 'estudiante' };
    },
    NIVELES: { facil: ['inicial', 'largo', 'una', 'lista'], medio: ['largo'], dificil: [] },
    generar(tipo) {
        const T = this.TIPOS[tipo], c = this.leer(tipo), aj = Impresion.leer(), L = c.L;
        if ($('im-inst')) { if (c.inst !== _i('visor.imp2.' + tipo + '.inst')) LS.set('imp2:inst:' + tipo, c.inst); else LS.del('imp2:inst:' + tipo); }
        this.ultimo = c;
        let pool = [];
        if (!T.sinPalabras) {
            pool = Seleccion.palabras(c.sel, L, { img: !!T.img || (c.pista && Pistas.necesitaImagen(c.pista)) });
            if (c.pista) pool = pool.filter(w => Pistas.sirve(w, c.pista, L));
            if (T.filtrar) pool = T.filtrar(pool, c);
            if (pool.length < (T.min ?? 4)) return msg(_ih('visor.imp2.pocas', { n: T.min ?? 4 }));
        }
        const codigo = this.base && this.base.config === JSON.stringify(c) ? this.base.codigo : Actividades.nuevoCodigo();
        // Niveles: la misma selección en varias dificultades (cada una con sus ayudas); si no, las ayudas elegidas
        const niveles = T.niveles && c.niveles ? Object.keys(this.NIVELES) : [''];
        const versiones = ['A', 'B', 'C'].slice(0, c.versiones || 1);
        const hojas = [];
        versiones.forEach(v => niveles.forEach(nv => {
            const r = azar(`${codigo}|${v}|${nv}|${L}|${tipo}`), orden = barajar(pool, r);
            const cc = Object.assign({}, c, { ayudas: nv ? this.NIVELES[nv] : c.ayudas, nivel: nv, codigo });
            for (let i = 0, nh = T.cuantas ? T.cuantas(c, pool) : c.n; i < nh; i++) hojas.push({ v: versiones.length > 1 ? v : '', nivel: nv, m: T.modelo(orden, cc, r, i), cc });
        }));
        let pags = '';
        const marco = (h, docente) => `<section class="hoja ${tipo}${docente ? ' docente' : ''}">${T.inst === false ? '' : this.cabecera(h, c, docente)}${T.pagina(h.m, h.cc, docente)}${Impresion.pie(aj, { L: T.sinPalabras ? '' : L, codigo, version: [h.v, h.nivel ? _i('visor.imp2.nivel.' + h.nivel) : ''].filter(Boolean).join(' · ') })}</section>`;
        if (c.para !== 'docente') hojas.forEach(h => { pags += marco(h, false); if (T.reverso && c.op && c.op.reverso) pags += `<section class="hoja ${tipo} reverso">${T.reverso(h.m, c)}</section>`; });
        if (T.resp && c.para !== 'estudiante') hojas.forEach(h => { pags += marco(h, true); });
        if (T.extra) pags += T.extra(c, aj, codigo);
        const nombre = `${this.nombre(tipo)}${T.sinPalabras ? '' : ' · ' + Datos.nombreLengua(L)}`;
        Actividades.registrar({ tipo: 'impresos', sub: tipo, config: c, codigo, nombre: `${nombre} · ${new Date().toLocaleDateString()}` });
        this.base = { codigo, config: JSON.stringify(c) };
        const usadas = [...new Set(hojas.flatMap(h => (h.m.ws || []).map(x => x.w || x)))];
        const sinRed = !navigator.onLine ? Offline.faltantes(usadas) : 0;
        Impresion.ventana({ titulo: `${nombre} · ${codigo}`, paginas: pags, css: this.CSS + (T.css || ''), ajustes: aj, orientacion: T.orientacion || 'portrait',
            desc: _i('visor.hoja.desc2', { tipo: this.nombre(tipo), c: codigo, v: c.versiones || 1, n: c.n }), aviso: sinRed ? _i('visor.off.faltanImg', { n: sinRed }) : '' });
    },
    cabecera(h, c, docente) {
        const tag = [h.v ? _i('visor.hoja.version', { v: h.v }) : '', h.nivel ? _i('visor.imp2.nivel.' + h.nivel) : ''].filter(Boolean).join(' · ');
        return `<header><div><b>${esc(this.TIPOS[c.tipo].ico)} </b>${esc(c.inst)}</div>${docente ? `<div class="ver">✔️ ${esc(_i('visor.hoja.docente'))}${tag ? ' · ' + esc(tag) : ''}</div>`
            : `<div class="datos">${tag ? `<span class="ver">${esc(tag)}</span> &nbsp; ` : ''}______________________ &nbsp; ____/____/______</div>`}</header>`;
    },
    /** Pista de una palabra para las listas de pistas (en crucigramas y sopas), con ayudas. */
    pista(x, c, n) {
        const L = c.L, w = x.w, ay = c.ayudas || [];
        const extra = [ay.includes('largo') ? `(${x.g.length})` : '', ay.includes('inicial') ? _i('visor.imp2.empiezaCon', { l: x.g[0] }) : ''].filter(Boolean).join(' ');
        if (!c.pista || c.pista.modo === 'palabra') return `<span${Texto.lang(L)}>${esc(w.t[L].texto)}</span>`;
        return `${Pistas.render(w, c.pista, L)}${extra ? `<span class="ay"> ${esc(extra)}</span>` : ''}`;
    },
    CSS: `.grid{display:grid;margin:0 auto}.cl{box-sizing:border-box;display:flex;align-items:center;justify-content:center;position:relative;font-weight:700}
        .pistas-l{columns:2;column-gap:6mm;font-size:calc(11pt*var(--k));min-height:0}.pistas-l .it{break-inside:avoid;margin:0 0 2mm;display:flex;gap:2mm;align-items:flex-start}
        .pistas-l .it img{width:16mm;height:16mm;object-fit:contain;flex:none}.pistas-l .n{font-weight:800;min-width:6mm}.pistas-l .p-texto{text-align:left;font-size:inherit}
        .pistas-l h4{margin:0 0 1.5mm;font-size:calc(12pt*var(--k));column-span:all}.ay{color:#555;font-size:.9em}
        .rej2{flex:1;display:grid;grid-template-columns:1fr 1fr;grid-auto-rows:1fr;gap:4mm;min-height:0}
        .rej2 .item{border:.5mm solid #111;border-radius:3mm;padding:3mm;display:flex;flex-direction:column;gap:2mm;align-items:center;justify-content:center;min-height:0;overflow:hidden}
        .rej2 .ps{flex:1;min-height:0;width:100%;display:flex;flex-direction:column;align-items:center;justify-content:center}.rej2 .ps img{flex:1;min-height:0}
        .letras{display:flex;gap:1.5mm;flex-wrap:wrap;justify-content:center}.ficha{border:.4mm solid #111;border-radius:1.5mm;min-width:8mm;height:9mm;display:inline-flex;align-items:center;justify-content:center;font-size:calc(15pt*var(--k));font-weight:800;padding:0 1mm}
        .ficha.fija{background:#e5e7eb}.ordena .lin{width:100%;height:11mm;border-bottom:.4mm solid #111;display:flex;align-items:flex-end;justify-content:center;gap:1.5mm;font-size:calc(18pt*var(--k))}
        .hueco{display:inline-block;width:7mm;border-bottom:.5mm solid #111;height:8mm}
        .corte{border:.3mm dashed #888}.p-linea{height:7mm}`
};

/* ---------- 🧩 Sopa de letras ---------- */
Impresos.registrar('sopa', {
    ico: '🧩', resp: true, niveles: true, min: 4,
    pistas: ['palabra', 'imagen', 'lengua', 'descripcion', 'propia', 'espacio'],
    ayudas: ['largo', 'inicial'], ayudasPred: [],
    opciones: op => `<span class="t">${_ih('visor.sopa.tamano')}</span><div class="chips">${[10, 12, 14, 16].map(n => `<label class="chip"><input type="radio" name="op-n" value="${n}" ${(op.tam || 12) === n ? 'checked' : ''}> ${n}×${n}</label>`).join('')}</div>
        <span class="t">${_ih('visor.sopa.nivel')}</span><div class="chips">${['facil', 'normal', 'dificil'].map(k => `<label class="chip"><input type="radio" name="op-d" value="${k}" ${(op.dir || 'normal') === k ? 'checked' : ''}> ${_ih('visor.sopa.' + k)}</label>`).join('')}</div>
        <span class="t">${_ih('visor.imp2.cuantas')}</span><div class="chips">${[6, 8, 10, 12].map(n => `<label class="chip"><input type="radio" name="op-c" value="${n}" ${(op.cuantas || 8) === n ? 'checked' : ''}> ${n}</label>`).join('')}</div>`,
    leer: () => ({ tam: +(document.querySelector('input[name="op-n"]:checked')?.value || 12), dir: document.querySelector('input[name="op-d"]:checked')?.value || 'normal', cuantas: +(document.querySelector('input[name="op-c"]:checked')?.value || 8) }),
    filtrar(pool, c) { return Impresos.deLetras(pool, c.L, 3, c.op.tam || 12).map(x => x.w); },
    modelo(orden, c, r, i) {
        const n = c.op.tam || 12, dir = c.nivel === 'facil' ? 'facil' : c.nivel === 'dificil' ? 'dificil' : c.nivel === 'medio' ? 'normal' : (c.op.dir || 'normal');
        const lista = Impresos.deLetras(Impresos.rotar(orden, i, c.op.cuantas || 8), c.L, 3, n);
        const g = Sopa.generar(lista, n, c.op.cuantas || 8, dir, r);
        const alfa = Texto.alfabeto(c.L).filter(a => a.length <= 3), relleno = alfa.length ? alfa : 'abcdefghijklmnopqrstuvwxyz'.split('');
        const grid = g.grid.map(f => f.map(x => x || relleno[Math.floor(r() * relleno.length)]));
        return { grid, ws: g.palabras, n };
    },
    pagina(m, c, docente) {
        const cs = Math.min(13, Math.floor(175 / m.n)), en = new Map();
        m.ws.forEach((p, i) => p.celdas.forEach(k => en.set(k, i)));
        const celdas = m.grid.map((f, y) => f.map((l, x) => `<div class="cl${docente && en.has(x + ',' + y) ? ' hall' : ''}" style="width:${cs}mm;height:${cs}mm;font-size:${(cs * 1.5).toFixed(1)}pt">${esc(l)}</div>`).join('')).join('');
        const lista = m.ws.map((p, i) => `<div class="it"><span class="n">${docente ? '✓' : '☐'}</span><span>${Impresos.pista(p, c, i)}</span></div>`).join('');
        return `<div class="grid sopa-g" style="grid-template-columns:repeat(${m.n},${cs}mm)"${Texto.lang(c.L)}>${celdas}</div><div class="pistas-l">${lista}</div>`;
    },
    css: `.sopa-g .cl{border:.2mm solid #bbb}.sopa-g .hall{background:#fde68a;border-color:#b45309}.sopa .pistas-l .it img{width:14mm;height:14mm}`
});

/* ---------- ✏️ Crucigrama ---------- */
const Cruci = {
    /** Acomoda palabras (listas de letras) cruzándolas; devuelve { celdas: Map 'x,y'→letra, palabras: [{x,y,dir,g,w}] } normalizado a 0,0. */
    armar(lista, cuantas, rng, intentos = 12, maxLado = 17) {
        let mejor = null;
        for (let t = 0; t < intentos; t++) {
            const orden = t === 0 ? lista.slice().sort((a, b) => b.g.length - a.g.length) : barajar(lista, rng).sort((a, b) => b.g.length - a.g.length + (rng() - 0.5) * 3);
            const r = this.intento(orden, cuantas, rng, maxLado);
            if (!mejor || r.palabras.length > mejor.palabras.length || (r.palabras.length === mejor.palabras.length && r.area < mejor.area)) mejor = r;
            if (mejor.palabras.length >= cuantas && t > 3) break;
        }
        return this.normalizar(mejor);
    },
    intento(orden, cuantas, rng, maxLado) {
        const celdas = new Map(), dirs = new Map(), palabras = [];
        const k = (x, y) => x + ',' + y;
        const caja = () => { let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity; celdas.forEach((v, key) => { const [x, y] = key.split(',').map(Number); x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); }); return { x0, y0, x1, y1 }; };
        const valida = (g, x, y, d) => {
            const [dx, dy] = d === 'h' ? [1, 0] : [0, 1];
            if (celdas.has(k(x - dx, y - dy)) || celdas.has(k(x + dx * g.length, y + dy * g.length))) return -1;   // nada pegado antes ni después
            let cruces = 0;
            for (let i = 0; i < g.length; i++) {
                const cx = x + dx * i, cy = y + dy * i, v = celdas.get(k(cx, cy));
                if (v != null) { if (v !== g[i] || (dirs.get(k(cx, cy)) || '').includes(d)) return -1; cruces++; continue; }
                // celda nueva: sus vecinas a los lados deben estar vacías (no se forman palabras de más)
                if (d === 'h' ? (celdas.has(k(cx, cy - 1)) || celdas.has(k(cx, cy + 1))) : (celdas.has(k(cx - 1, cy)) || celdas.has(k(cx + 1, cy)))) return -1;
            }
            return cruces === g.length ? -1 : cruces;
        };
        const poner = (p, x, y, d) => { const [dx, dy] = d === 'h' ? [1, 0] : [0, 1]; p.g.forEach((l, i) => { const key = k(x + dx * i, y + dy * i); celdas.set(key, l); dirs.set(key, (dirs.get(key) || '') + d); }); palabras.push({ x, y, dir: d, g: p.g, w: p.w }); };
        for (const p of orden) {
            if (palabras.length >= cuantas) break;
            if (!palabras.length) { poner(p, 0, 0, rng() < 0.5 ? 'h' : 'v'); continue; }
            let op = null;
            celdas.forEach((v, key) => {
                const [cx, cy] = key.split(',').map(Number);
                p.g.forEach((l, i) => {
                    if (l !== v) return;
                    ['h', 'v'].forEach(d => {
                        const x = d === 'h' ? cx - i : cx, y = d === 'v' ? cy - i : cy, cr = valida(p.g, x, y, d);
                        if (cr < 1) return;
                        const b = caja(), nx0 = Math.min(b.x0, x), ny0 = Math.min(b.y0, y), nx1 = Math.max(b.x1, d === 'h' ? x + p.g.length - 1 : x), ny1 = Math.max(b.y1, d === 'v' ? y + p.g.length - 1 : y);
                        if (nx1 - nx0 + 1 > maxLado || ny1 - ny0 + 1 > maxLado) return;
                        const pts = cr * 10 - (nx1 - nx0 + 1) * (ny1 - ny0 + 1) / 10 + rng() * 2;
                        if (!op || pts > op.pts) op = { x, y, d, pts };
                    });
                });
            });
            if (op) poner(p, op.x, op.y, op.d);
        }
        const b = caja();
        return { celdas, palabras, area: (b.x1 - b.x0 + 1) * (b.y1 - b.y0 + 1), b };
    },
    normalizar(r) {
        const { x0, y0, x1, y1 } = r.b, celdas = new Map();
        r.celdas.forEach((v, key) => { const [x, y] = key.split(',').map(Number); celdas.set((x - x0) + ',' + (y - y0), v); });
        const palabras = r.palabras.map(p => Object.assign({}, p, { x: p.x - x0, y: p.y - y0 }));
        // Números en orden de lectura
        const inicio = new Map(); let n = 0;
        for (let y = 0; y <= y1 - y0; y++) for (let x = 0; x <= x1 - x0; x++) {
            const ps = palabras.filter(p => p.x === x && p.y === y);
            if (ps.length) { n++; inicio.set(x + ',' + y, n); ps.forEach(p => { p.num = n; }); }
        }
        return { celdas, palabras, ancho: x1 - x0 + 1, alto: y1 - y0 + 1, inicio };
    }
};
Impresos.registrar('crucigrama', {
    ico: '✏️', resp: true, niveles: true, min: 5,
    pistas: ['imagen', 'lengua', 'descripcion', 'propia', 'imagenEspacio', 'espacio'],
    ayudas: ['inicial', 'largo', 'una', 'dobles', 'lista'], ayudasPred: ['largo'],
    /* modelo(orden, c, r, i) */
    opciones: op => `<span class="t">${_ih('visor.imp2.cuantas')}</span><div class="chips">${[6, 8, 10, 12].map(n => `<label class="chip"><input type="radio" name="op-c" value="${n}" ${(op.cuantas || 8) === n ? 'checked' : ''}> ${n}</label>`).join('')}</div>`,
    leer: () => ({ cuantas: +(document.querySelector('input[name="op-c"]:checked')?.value || 8) }),
    filtrar(pool, c) { return Impresos.deLetras(pool, c.L, 3, 12).map(x => x.w); },
    modelo(orden, c, r, i) {
        const R = Cruci.armar(Impresos.deLetras(Impresos.rotar(orden, i, c.op.cuantas || 8), c.L, 3, 12), c.op.cuantas || 8, r);
        // «una letra puesta»: una letra al azar de cada palabra (no la primera)
        const reveladas = new Set();
        R.palabras.forEach(p => { if (p.g.length > 2) { const i = 1 + Math.floor(r() * (p.g.length - 1)); reveladas.add(p.dir === 'h' ? (p.x + i) + ',' + p.y : p.x + ',' + (p.y + i)); } });
        return Object.assign(R, { ws: R.palabras, reveladas });
    },
    pagina(m, c, docente) {
        const ay = c.ayudas || [], cs = Math.max(6, Math.min(11, Math.floor(Math.min(180 / m.ancho, 125 / m.alto))));
        const iniciales = new Set(m.palabras.map(p => p.x + ',' + p.y));
        let h = '';
        for (let y = 0; y < m.alto; y++) for (let x = 0; x < m.ancho; x++) {
            const key = x + ',' + y, l = m.celdas.get(key);
            if (l == null) { h += `<div class="cl vacia" style="width:${cs}mm;height:${cs}mm"></div>`; continue; }
            const ver = docente || (ay.includes('inicial') && iniciales.has(key)) || (ay.includes('una') && m.reveladas.has(key));
            h += `<div class="cl letra${ay.includes('dobles') && l.length > 1 ? ' doble' : ''}" style="width:${cs}mm;height:${cs}mm;font-size:${(cs * 1.6).toFixed(1)}pt">${m.inicio.has(key) ? `<span class="nn">${m.inicio.get(key)}</span>` : ''}${ver ? `<span class="${docente ? 'resp-ok' : ''}">${esc(l)}</span>` : ''}</div>`;
        }
        const lista = d => m.palabras.filter(p => p.dir === d).sort((a, b) => a.num - b.num).map(p => `<div class="it"><span class="n">${p.num}.</span><span>${Impresos.pista(p, c)}</span></div>`).join('');
        const banco = ay.includes('lista') && !docente ? `<div class="banco"${Texto.lang(c.L)}>${_ih('visor.imp2.banco')} ${barajar(m.palabras.map(p => p.w.t[c.L].texto), azar(m.palabras.length + c.L)).map(esc).join(' · ')}</div>` : '';
        return `<div class="grid cruci-g" style="grid-template-columns:repeat(${m.ancho},${cs}mm)"${Texto.lang(c.L)}>${h}</div>${banco}
            <div class="pistas-l"><h4>➡️ ${_ih('visor.imp2.horizontales')}</h4>${lista('h')}<h4>⬇️ ${_ih('visor.imp2.verticales')}</h4>${lista('v')}</div>`;
    },
    css: `.cruci-g .letra{border:.35mm solid #111;background:#fff}.cruci-g .vacia{background:transparent}.cruci-g .nn{position:absolute;top:.2mm;left:.5mm;font-size:6pt;font-weight:600}
        .cruci-g .doble::after{content:'';position:absolute;right:.6mm;bottom:.6mm;width:1.2mm;height:1.2mm;border-radius:50%;background:#111;box-shadow:-1.6mm 0 0 #111}
        .banco{font-size:calc(11pt*var(--k));border:.3mm dashed #888;border-radius:2mm;padding:2mm;text-align:center}.crucigrama .pistas-l{flex:1;overflow:hidden}`
});

/* ---------- 🔀 Ordena las letras ---------- */
Impresos.registrar('ordena', {
    ico: '🔀', resp: true, niveles: true, min: 4,
    pistas: ['imagen', 'lengua', 'descripcion', 'propia', 'ninguna'],
    ayudas: ['inicial', 'largo'], ayudasPred: [],
    filtrar(pool, c) { return Impresos.deLetras(pool, c.L, 3, 10).map(x => x.w); },
    modelo(orden, c, r, i) {
        let lista = Impresos.deLetras(orden, c.L, 3, 10);
        if (c.nivel === 'facil') lista = lista.filter(x => x.g.length <= 5).concat(lista.filter(x => x.g.length > 5));
        if (c.nivel === 'dificil') lista = lista.filter(x => x.g.length >= 5).concat(lista.filter(x => x.g.length < 5));
        const ws = []; for (let k = 0; k < 6 && k < lista.length; k++) ws.push(lista[(i * 6 + k) % lista.length]);
        return { ws: [...new Set(ws)].map(x => {
            const fija = (c.ayudas || []).includes('inicial');
            let rev = x.g.slice(fija ? 1 : 0); for (let t = 0; t < 8; t++) { rev = barajar(rev, r); if (rev.join('') !== x.g.slice(fija ? 1 : 0).join('')) break; }
            return Object.assign({}, x, { rev: fija ? [x.g[0], ...rev] : rev, fija });
        }) };
    },
    pagina(m, c, docente) {
        return `<div class="rej2">${m.ws.map(x => { const ps = Pistas.render(x.w, c.pista, c.L);
            return `<div class="item"><div class="ps">${ps}</div><div class="letras"${Texto.lang(c.L)}>${x.rev.map((l, i) => `<span class="ficha${x.fija && i === 0 ? ' fija' : ''}">${esc(l)}</span>`).join('')}</div>
            <div class="lin">${docente ? `<span class="resp-ok"${Texto.lang(c.L)}>${esc(x.w.t[c.L].texto)}</span>` : (c.ayudas || []).includes('largo') ? x.g.map(() => '<span class="hueco"></span>').join('') : ''}</div></div>`; }).join('')}</div>`;
    },
    css: ''
});

/* ---------- ✋ Basta en papel ---------- */
Impresos.registrar('basta', {
    ico: '✋', resp: true, min: 0, sinSeleccion: true, hojas: [1, 2, 3, 5, 10, 20, 30],
    opciones: (op, L) => {
        const usadas = new Set(Datos.words.map(w => w.cat)), cats = [...Datos.cats.values()].filter(c => usadas.has(c.id));
        return `<span class="t">${_ih('visor.imp2.basta.cats')}</span><div class="chips">${cats.map(c => `<label class="chip"><input type="checkbox" name="op-cat" value="${esc(c.id)}" ${(op.cats || cats.slice(0, 4).map(x => x.id)).includes(c.id) ? 'checked' : ''}> ${esc(Datos.categoria(c.id, L))}</label>`).join('')}</div>
            <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-blanco" ${op.blanco ? 'checked' : ''}> ${_ih('visor.imp2.basta.blanco')}</label>
            <span class="t">${_ih('visor.imp2.basta.filas')}</span><div class="chips">${[6, 8, 10, 12].map(n => `<label class="chip"><input type="radio" name="op-f" value="${n}" ${(op.filas || 8) === n ? 'checked' : ''}> ${n}</label>`).join('')}</div>`;
    },
    leer: () => ({ cats: Comun.marcados('op-cat').slice(0, 6), blanco: $('op-blanco').checked, filas: +(document.querySelector('input[name="op-f"]:checked')?.value || 8) }),
    filtrar: pool => pool,
    modelo(orden, c) { return { ws: [] }; },
    pagina(m, c, docente) {
        const cats = c.op.blanco ? Array.from({ length: Math.max(4, c.op.cats.length) }, () => '') : c.op.cats.map(id => Datos.categoria(id, c.L));
        if (docente) {
            // Hoja del docente: letras para cantar (las que tienen palabras) y ejemplos del catálogo por categoría
            const alfa = Texto.alfabeto(c.L), ej = {};
            Datos.paraJuegos(c.L).forEach(w => { const g = Texto.grafemas(w.t[c.L].texto, c.L).find(x => x.letra); if (!g) return; (ej[g.g] = ej[g.g] || {}); (ej[g.g][w.cat] = ej[g.g][w.cat] || []).push(w.t[c.L].texto); });
            const letras = alfa.filter(l => ej[l]);
            return `<table class="basta"${Texto.lang(c.L)}><thead><tr><th>${_ih('visor.imp2.basta.letra')}</th>${c.op.cats.map(id => `<th>${esc(Datos.categoria(id, c.L))}</th>`).join('')}</tr></thead>
                <tbody>${letras.map(l => `<tr><td class="le">${esc(l)}</td>${c.op.cats.map(id => `<td>${(ej[l][id] || []).slice(0, 3).map(esc).join(', ')}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
        }
        return `<table class="basta"><thead><tr><th class="le">${_ih('visor.imp2.basta.letra')}</th>${cats.map(n => `<th>${esc(n)}</th>`).join('')}<th class="pt">${_ih('visor.imp2.basta.puntos')}</th></tr></thead>
            <tbody>${Array.from({ length: c.op.filas || 8 }, () => `<tr><td class="le"></td>${cats.map(() => '<td></td>').join('')}<td class="pt"></td></tr>`).join('')}</tbody>
            <tfoot><tr><td colspan="${cats.length + 1}" style="text-align:right">${_ih('visor.imp2.basta.total')}</td><td class="pt"></td></tr></tfoot></table>`;
    },
    css: `.basta{width:100%;border-collapse:collapse;flex:1;font-size:calc(11pt*var(--k))}.basta th,.basta td{border:.35mm solid #111;padding:1.5mm;vertical-align:middle}
        .basta th{background:#f1f5f9;height:10mm}.basta td{height:13mm}.basta .le{width:13mm;text-align:center;font-weight:800;font-size:1.3em}.basta .pt{width:16mm}
        .docente .basta td{height:auto;font-size:.9em}`
});

/* ---------- 🧠 Memorama recortable ---------- */
Impresos.registrar('memorama', {
    ico: '🧠', min: 4, versiones: false, inst: false,
    opciones: op => `<span class="t">${_ih('visor.imp2.mem.pares')}</span><div class="chips">${['imagenPalabra', 'dosLenguas', 'dosImagenes', 'imagenEspacio'].map(k => `<label class="chip"><input type="radio" name="op-m" value="${k}" ${(op.modo || 'imagenPalabra') === k ? 'checked' : ''}> ${_ih('visor.imp2.mem.' + k)}</label>`).join('')}</div>
        <div id="op-l2b"><label class="t" for="op-l2">${_ih('visor.otraLengua')}</label><select id="op-l2" class="modern-select">${Comun.lenguas(null, op.L2 || '')}</select></div>
        <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-rev" ${op.reverso ? 'checked' : ''}> ${_ih('visor.imp2.mem.reverso')}</label>`,
    enlazar() { const f = () => $('op-l2b').classList.toggle('oculto', document.querySelector('input[name="op-m"]:checked').value !== 'dosLenguas'); document.querySelectorAll('input[name="op-m"]').forEach(r => r.addEventListener('change', f)); f(); },
    leer: () => ({ modo: document.querySelector('input[name="op-m"]:checked')?.value || 'imagenPalabra', L2: $('op-l2').value, reverso: $('op-rev').checked }),
    filtrar(pool, c) {
        const m = c.op.modo;
        if (m === 'dosLenguas') return pool.filter(w => w.t[c.op.L2] && c.op.L2 !== c.L);
        if (m === 'dosImagenes') return pool.filter(w => w.img && w.galeria && w.galeria.length);
        return pool.filter(w => w.img);
    },
    modelo(orden, c, r, i) {
        const ws = []; for (let k = 0; k < 8 && k < orden.length; k++) ws.push(orden[(i * 8 + k) % orden.length]);
        const u = [...new Set(ws)], cartas = [];
        u.forEach(w => { cartas.push({ w, lado: 'a' }); cartas.push({ w, lado: 'b' }); });
        return { ws: u, cartas: barajar(cartas, r) };
    },
    pagina(m, c) {
        const L = c.L, mo = c.op.modo;
        const img = (u, alt) => `<img src="${esc(u)}" alt="${esc(alt || '')}">`;
        const cara = x => {
            const w = x.w;
            if (x.lado === 'a') return mo === 'dosLenguas' ? `<div class="mp"${Texto.lang(L)}>${esc(w.t[L].texto)}</div>` : img(Datos.img(w, 500), w.alt);
            if (mo === 'imagenPalabra') return `<div class="mp"${Texto.lang(L)}>${esc(w.t[L].texto)}</div>`;
            if (mo === 'dosLenguas') return `<div class="mp"${Texto.lang(c.op.L2)}>${esc(w.t[c.op.L2].texto)}</div>`;
            if (mo === 'dosImagenes') return img(w.galeria[0].url + '=w500', w.galeria[0].alt);
            return '<div class="p-linea"></div><div class="p-linea"></div>';
        };
        return `<div class="mem">${m.cartas.map(x => `<div class="carta-m corte">${cara(x)}</div>`).join('')}</div>`;
    },
    reverso(m, c) { return `<div class="mem">${m.cartas.map(() => `<div class="carta-m corte dorso"><span>${esc(Datos.cfg.titulo || '')}</span></div>`).join('')}</div>`; },
    css: `.mem{flex:1;display:grid;grid-template-columns:repeat(4,1fr);grid-template-rows:repeat(4,1fr);gap:0;min-height:0}
        .carta-m{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:3mm;min-height:0;overflow:hidden}.carta-m img{max-width:100%;max-height:100%;flex:1;min-height:0}
        .mp{font-size:calc(17pt*var(--k));font-weight:800;text-align:center;word-break:break-word}
        .dorso{background:repeating-linear-gradient(45deg,#e5e7eb 0 3mm,#fff 3mm 6mm)}.dorso span{background:#fff;padding:1mm 2mm;border-radius:1mm;font-weight:700;font-size:9pt;text-align:center}
        .reverso{padding:4mm}`
});

/* ---------- 🏡 Pregunta en casa ---------- */
const QR = {
    /** Código QR como SVG (para que un teléfono abra el visor justo donde se necesita). */
    svg(texto, tam = 30) {
        try { const q = qrcode(0, 'M'); q.addData(texto); q.make(); return q.createSvgTag({ cellSize: 2, margin: 0, scalable: true }).replace('<svg ', `<svg style="width:${tam}mm;height:${tam}mm" `); }
        catch (e) { return ''; }
    },
    enlace(hash) {
        const u = new URL(location.origin + location.pathname);
        if (Datos.base !== (API_URL || '').split('?')[0]) u.searchParams.set('api', Datos.base);
        if (Datos.token) u.searchParams.set('token', Datos.token);
        return u.toString() + hash;
    }
};
Impresos.registrar('preguntaCasa', {
    ico: '🏡', min: 0, versiones: false, hojas: [1, 5, 10, 20, 30, 40],
    opciones: (op, L) => {
        const usadas = new Set(Datos.words.map(w => w.cat)), cats = [...Datos.cats.values()].filter(c => usadas.has(c.id));
        return `<label class="t" for="op-tema">${_ih('visor.imp2.pc.tema')}</label><select id="op-tema" class="modern-select"><option value="">${_ih('visor.imp2.pc.temaLibre')}</option>${cats.map(c => `<option value="${esc(c.id)}" ${op.tema === c.id ? 'selected' : ''}>${esc(Datos.categoria(c.id, L))}</option>`).join('')}</select>
            <span class="t">${_ih('visor.imp2.pc.tipo')}</span><div class="chips">${['nuevas', 'investiga', 'ambas'].map(k => `<label class="chip"><input type="radio" name="op-t" value="${k}" ${(op.tipo || 'ambas') === k ? 'checked' : ''}> ${_ih('visor.imp2.pc.' + k)}</label>`).join('')}</div>
            <label class="t" for="op-inv">${_ih('visor.imp2.pc.invitacion')}</label><textarea id="op-inv" class="input-box" rows="4" maxlength="600">${esc(op.inv != null ? op.inv : (LS.get('pc:inv:' + Datos.base, '') || _i('visor.imp2.pc.invitacion.pred', { catalogo: Datos.cfg.titulo || '' })))}</textarea>
            <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-dib" ${op.dibuja !== false ? 'checked' : ''}> 🖍️ ${_ih('visor.imp2.pc.dibuja')}</label>
            <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-qr" ${op.qr !== false ? 'checked' : ''} ${Propuestas.activo() ? '' : 'disabled'}> 📱 ${_ih('visor.imp2.pc.qr')}</label>
            <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-guia" ${op.guia !== false ? 'checked' : ''}> 🧭 ${_ih('visor.imp2.pc.guia')}</label>
            ${Propuestas.activo() ? '' : `<p class="nota">${_ih('visor.imp2.pc.sinPropuestas')}</p>`}`;
    },
    leer: () => { const inv = $('op-inv').value.trim(); LS.set('pc:inv:' + Datos.base, inv); return { tema: $('op-tema').value, tipo: document.querySelector('input[name="op-t"]:checked')?.value || 'ambas', inv, dibuja: $('op-dib').checked, qr: $('op-qr').checked && !$('op-qr').disabled, guia: $('op-guia').checked }; },
    filtrar: pool => pool,
    modelo(orden, c, r) {
        const L = c.L, tema = c.op.tema;
        const sinDesc = Datos.paraJuegos(L).filter(w => (!tema || w.cat === tema) && !w.t[L].desc);
        const ya = Datos.paraJuegos(L).filter(w => !tema || w.cat === tema);
        return { ws: [], investiga: barajar(sinDesc, r).slice(0, 4), ya: barajar(ya, r).slice(0, 14) };
    },
    pagina(m, c) {
        const L = c.L, op = c.op, tema = op.tema ? Datos.categoria(op.tema, L) : '';
        const act = `${_i('visor.imp2.preguntaCasa')}${tema ? ' · ' + tema : ''}`;
        const qr = op.qr ? QR.svg(QR.enlace('#proponer=' + encodeURIComponent(act + ' · ' + (c.codigo || ''))), 24) : '';
        const ren = n => Array.from({ length: n }, () => '<tr><td></td><td></td><td></td></tr>').join('');
        const nuevas = op.tipo !== 'investiga' ? `<h3>🗣️ ${_ih('visor.imp2.pc.nuevas.t', { tema: tema || _i('visor.imp2.pc.loQueQuieras') })}</h3>
            ${m.ya.length ? `<p class="ya">${_ih('visor.imp2.pc.yaEstan')} <span${Texto.lang(L)}>${m.ya.map(w => esc(w.t[L].texto)).join(' · ')}</span></p>` : ''}
            <table class="pc"><thead><tr><th>${_ih('visor.imp2.pc.palabra')}</th><th>${_ih('visor.imp2.pc.queEs')}</th><th>${_ih('visor.imp2.pc.donde')}</th></tr></thead><tbody>${ren(op.tipo === 'nuevas' ? 6 : 4)}</tbody></table>` : '';
        const inv = op.tipo !== 'nuevas' && m.investiga.length ? `<h3>🔎 ${_ih('visor.imp2.pc.investiga.t')}</h3>${m.investiga.map(w => `<div class="inv"><b${Texto.lang(L)}>${esc(w.t[L].texto)}</b>: <span class="linea"></span></div>`).join('')}` : '';
        return `<div class="pc-inv">${esc(op.inv)}</div>${nuevas}${inv}
            ${op.dibuja ? `<div class="dib-pc">🖍️ ${_ih('visor.imp2.pc.dibujaAqui')}</div>` : ''}
            <div class="pie-pc"><div class="permiso">☐ ${_ih('visor.imp2.pc.permiso')} <span class="linea corta"></span> ${_ih('visor.imp2.pc.quienEnseno')} <span class="linea corta"></span></div>
            ${qr ? `<div class="qr-pc">${qr}<span>${_ih('visor.imp2.pc.qr.d')}</span></div>` : ''}</div>
            <div class="talon">✂ ${_ih('visor.imp2.pc.talon')} ____________________________ · ${_ih('visor.imp2.pc.grupo')} ________</div>`;
    },
    extra(c, aj) {
        if (!c.op.guia) return '';
        return `<section class="hoja guia"><header><div><b>🧭 ${_ih('visor.imp2.pc.guia.t')}</b></div></header>
            <div class="guia-t" contenteditable="true" spellcheck="false">${[1, 2, 3, 4, 5, 6].map(i => `<p>${_ih('visor.imp2.pc.guia.' + i)}</p>`).join('')}</div>${Impresion.pie(aj, {})}</section>`;
    },
    css: `.pc-inv{font-size:calc(11pt*var(--k));border:.3mm solid #111;border-radius:2mm;padding:2mm 3mm;white-space:pre-wrap}.preguntaCasa h3{margin:1mm 0;font-size:calc(12pt*var(--k))}
        .ya{font-size:9pt;color:#444;margin:0}.pc{width:100%;border-collapse:collapse}.pc th,.pc td{border:.3mm solid #111;padding:1mm 2mm;font-size:10pt}.pc td{height:9mm}
        .inv{font-size:calc(11pt*var(--k));display:flex;flex-direction:column;gap:1mm;margin-bottom:1.5mm}.linea{display:block;border-bottom:.3mm solid #111;height:6mm;flex:1}
        .linea.corta{display:inline-block;width:45mm;height:4mm}.dib-pc{flex:1;min-height:14mm;border:.4mm dashed #777;border-radius:2mm;padding:2mm;font-size:9pt;color:#555}
        .pie-pc{display:flex;gap:4mm;align-items:center}.permiso{font-size:9.5pt;flex:1}.qr-pc span{max-width:40mm}.qr-pc{display:flex;gap:3mm;align-items:center;font-size:9pt}.talon{border-top:.4mm dashed #555;padding-top:2mm;font-size:9pt}
        .guia-t{font-size:calc(12pt*var(--k));line-height:1.5;flex:1}.guia-t p{margin:0 0 3mm}`
});

/* ---------- 📋 Hoja de recolección (asambleas y talleres) ---------- */
Impresos.registrar('recoleccion', {
    ico: '📋', sinPalabras: true, min: 0, versiones: false, hojas: [1, 2, 3, 5, 10],
    opciones: op => `<label class="t" for="op-ev">${_ih('visor.imp2.rec.evento')}</label><input id="op-ev" class="input-box" maxlength="120" value="${esc(op.evento || '')}" placeholder="${esc(_i('visor.imp2.rec.evento.ph'))}">
        <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-qr" ${op.qr !== false ? 'checked' : ''} ${Propuestas.activo() ? '' : 'disabled'}> 📱 ${_ih('visor.imp2.pc.qr')}</label>`,
    leer: () => ({ evento: $('op-ev').value.trim(), qr: $('op-qr').checked && !$('op-qr').disabled }),
    modelo: () => ({ ws: [] }),
    pagina(m, c) {
        const qr = c.op.qr ? QR.svg(QR.enlace('#proponer=' + encodeURIComponent(c.op.evento || _i('visor.imp2.recoleccion'))), 22) : '';
        return `<div class="rec-cab"><div>${_ih('visor.imp2.rec.lugar')} ______________________ ${_ih('visor.imp2.rec.fecha')} ____/____/______<br>${esc(c.op.evento || '')}</div>${qr ? `<div class="qr-pc">${qr}</div>` : ''}</div>
            <table class="pc rec"><thead><tr><th>${_ih('visor.imp2.pc.palabra')}</th><th>${_ih('visor.imp2.pc.queEs')}</th><th>${_ih('visor.imp2.pc.donde')}</th><th>${_ih('visor.imp2.rec.quien')}</th><th>${_ih('visor.imp2.rec.permiso')}</th></tr></thead>
            <tbody>${Array.from({ length: 11 }, () => `<tr><td></td><td></td><td></td><td></td><td class="si">☐ ${_ih('visor.si')} ☐ ${_ih('visor.no')}</td></tr>`).join('')}</tbody></table>
            <p class="nota-rec">${_ih('visor.imp2.rec.nota')}</p>`;
    },
    css: `.rec-cab{display:flex;justify-content:space-between;gap:4mm;font-size:10pt}.rec td{height:12.5mm}.rec .si{font-size:9pt;white-space:nowrap;width:24mm}.nota-rec{font-size:9pt;color:#444;margin:0}`
});
window.Impresos = Impresos;
</script>
