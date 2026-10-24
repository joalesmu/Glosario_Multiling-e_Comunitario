<script>
/* =============================================
   LETRAS, SÍLABAS Y ESCUCHA (en pantalla y en papel)
   Sirven tanto para aprender a leer y escribir en la lengua como para aprenderla
   como segunda lengua: el docente decide cómo usarlos.
   ============================================= */
const Silabas = {
    /** Sílabas registradas en el catálogo («ki-i-tɨ», «kii·tɨ», «kii tɨ»). Sin registro → []. */
    de(w, L) { const s = String((w.t[L] || {}).silabas || '').trim(); return s ? s.split(/[-·.\s/|]+/).map(x => x.trim()).filter(Boolean) : []; }
};
const _letrasHoja = (w, L) => Texto.grafemas(w.t[L].texto, L).filter(x => x.letra).map(x => x.g);

/* ---------- 🔤 ¿Con qué letra empieza? ---------- */
Impresos.registrar('inicial', {
    ico: '🔤', resp: true, min: 4,
    pistas: ['imagen', 'lengua', 'descripcion', 'propia'],
    opciones: op => `<span class="t">${_ih('visor.imp2.ini.modo')}</span><div class="chips">${['elegir', 'escribir'].map(k => `<label class="chip"><input type="radio" name="op-m" value="${k}" ${(op.modo || 'elegir') === k ? 'checked' : ''}> ${_ih('visor.imp2.ini.' + k)}</label>`).join('')}</div>`,
    leer: () => ({ modo: document.querySelector('input[name="op-m"]:checked')?.value || 'elegir' }),
    filtrar(pool, c) { return pool.filter(w => _letrasHoja(w, c.L).length); },
    modelo(orden, c, r, i) {
        const alfa = Texto.alfabeto(c.L), ws = Impresos.rotar(orden, i, 8).slice(0, 8);
        return { ws: ws.map(w => { const ini = _letrasHoja(w, c.L)[0], otras = barajar(alfa.filter(a => a !== ini), r).slice(0, 2); return { w, ini, ops: barajar([ini, ...otras], r) }; }) };
    },
    pagina(m, c, docente) {
        return `<div class="rej2 ini">${m.ws.map(x => `<div class="item"><div class="ps">${Pistas.render(x.w, c.pista, c.L)}</div>
            ${c.op.modo === 'escribir' ? `<div class="caja-l">${docente ? `<span class="resp-ok">${esc(x.ini)}</span>` : ''}</div>`
            : `<div class="letras"${Texto.lang(c.L)}>${x.ops.map(l => `<span class="ficha${docente && l === x.ini ? ' circ' : ''}">${esc(l)}</span>`).join('')}</div>`}</div>`).join('')}</div>`;
    },
    css: `.ini .item .ps{flex:1;min-height:0}.caja-l{width:16mm;height:16mm;border:.5mm solid #111;border-radius:2mm;display:flex;align-items:center;justify-content:center;font-size:24pt}
        .circ{border-radius:50%;border:.8mm solid #b91c1c!important;color:#b91c1c}`
});

/* ---------- 👏 Cuenta las sílabas ---------- */
Impresos.registrar('silabas', {
    ico: '👏', resp: true, min: 4, req: { campo: 'silabas', min: 4 },
    pistas: ['imagen', 'palabra', 'lengua'],
    filtrar(pool, c) { return pool.filter(w => Silabas.de(w, c.L).length); },
    modelo(orden, c, r, i) { return { ws: Impresos.rotar(orden, i, 8).slice(0, 8).map(w => ({ w, s: Silabas.de(w, c.L) })) }; },
    pagina(m, c, docente) {
        return `<div class="rej2 sil">${m.ws.map(x => `<div class="item"><div class="ps">${Pistas.render(x.w, c.pista, c.L)}</div>
            <div class="palmas">${[1, 2, 3, 4, 5].map(n => `<span class="palma${docente && n <= x.s.length ? ' llena' : ''}">👏</span>`).join('')}</div>
            ${docente ? `<div class="nota-d"${Texto.lang(c.L)}>${x.s.map(esc).join(' · ')}</div>` : ''}</div>`).join('')}</div>`;
    },
    css: `.palmas{display:flex;gap:2mm}.palma{width:10mm;height:10mm;border:.4mm solid #111;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:12pt;filter:grayscale(1);opacity:.45}
        .palma.llena{opacity:1;filter:none;border-color:#b91c1c}.nota-d{font-size:12pt;color:#b91c1c;font-weight:700}`
});

/* ---------- 🧱 Arma con sílabas (recortable) ---------- */
Impresos.registrar('armaSilabas', {
    ico: '🧱', min: 3, versiones: false, req: { campo: 'silabas', min: 3 },
    pistas: ['imagen', 'lengua', 'ninguna'],
    filtrar(pool, c) { return pool.filter(w => Silabas.de(w, c.L).length >= 2); },
    modelo(orden, c, r, i) { return { ws: Impresos.rotar(orden, i, 4).slice(0, 4).map(w => ({ w, s: barajar(Silabas.de(w, c.L), r) })) }; },
    pagina(m, c) {
        return `<div class="arma">${m.ws.map(x => `<div class="fila-a"><div class="ps-a">${Pistas.render(x.w, c.pista, c.L)}</div>
            <div class="base">${x.s.map(() => '<span class="hueco-s"></span>').join('')}</div>
            <div class="piezas"${Texto.lang(c.L)}>${x.s.map(s => `<span class="pieza corte">${esc(s)}</span>`).join('')}</div></div>`).join('')}</div>`;
    },
    css: `.arma{flex:1;display:grid;grid-template-rows:repeat(4,1fr);gap:3mm;min-height:0}.fila-a{display:grid;grid-template-columns:40mm 1fr;grid-template-rows:1fr 1fr;gap:2mm;border-bottom:.3mm solid #ccc;min-height:0}
        .ps-a{grid-row:span 2;display:flex;align-items:center;justify-content:center;min-height:0}.ps-a img{max-height:100%}
        .base{display:flex;gap:2mm;align-items:flex-end}.hueco-s{width:24mm;height:12mm;border:.4mm solid #111;border-radius:2mm}
        .piezas{display:flex;gap:3mm;flex-wrap:wrap;align-items:center}.pieza{width:24mm;height:12mm;display:inline-flex;align-items:center;justify-content:center;font-size:calc(16pt*var(--k));font-weight:800}`
});

/* ---------- ✍️ Hoja de trazo por letra ---------- */
Impresos.registrar('trazo', {
    ico: '✍️', min: 0, versiones: false, sinSeleccion: true, hojas: [1, 2, 3, 5],
    cuantas: c => Math.max(1, c.op.letras.length) * c.n,
    opciones: (op, L) => `<span class="t">${_ih('visor.imp2.trazo.letras')}</span><div class="chips" id="op-letras">${Texto.alfabeto(L).map(a => `<label class="chip"><input type="checkbox" name="op-le" value="${esc(a)}" ${(op.letras || []).includes(a) ? 'checked' : ''}> <b>${esc(a)}</b></label>`).join('')}</div>
        <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-pal" ${op.palabras !== false ? 'checked' : ''}> ${_ih('visor.imp2.trazo.palabras')}</label>`,
    alCambiarLengua(L) { const d = $('op-letras'); if (d) d.innerHTML = Texto.alfabeto(L).map(a => `<label class="chip"><input type="checkbox" name="op-le" value="${esc(a)}"> <b>${esc(a)}</b></label>`).join(''); },
    leer: () => ({ letras: Comun.marcados('op-le').slice(0, 12), palabras: $('op-pal').checked }),
    filtrar: pool => pool,
    modelo(orden, c, r, i) {
        const letras = c.op.letras.length ? c.op.letras : Texto.alfabeto(c.L).slice(0, 1), le = letras[Math.floor(i / c.n) % letras.length];
        const ws = c.op.palabras ? barajar(Datos.paraJuegos(c.L).filter(w => _letrasHoja(w, c.L).includes(le)), r).slice(0, 3) : [];
        return { le, ws };
    },
    pagina(m, c) {
        const L = c.L, fila = t => `<div class="tr-f"${Texto.lang(L)}>${Array.from({ length: 6 }, () => `<span class="tr-l">${esc(t)}</span>`).join('')}</div>`;
        return `<div class="trazo-g"><div class="tr-gran"${Texto.lang(L)}>${esc(m.le)}</div>${fila(m.le)}${fila(m.le)}<div class="tr-f vacia"></div>
            ${m.ws.map(w => `<div class="tr-p">${w.img ? `<img src="${esc(Datos.img(w, 300))}" alt="">` : ''}<span class="tr-l"${Texto.lang(L)}>${esc(w.t[L].texto)}</span></div><div class="tr-f vacia"></div>`).join('')}</div>`;
    },
    css: `.trazo-g{flex:1;display:flex;flex-direction:column;gap:2mm;min-height:0}.tr-gran{font-size:90pt;font-weight:700;color:transparent;-webkit-text-stroke:.6mm #555;text-align:center;line-height:1}
        .tr-f{display:flex;justify-content:space-around;border-bottom:.4mm solid #111;border-top:.2mm dashed #999;min-height:16mm;align-items:flex-end}
        .tr-l{font-size:calc(34pt*var(--k));color:transparent;-webkit-text-stroke:.35mm #888;letter-spacing:1mm;font-weight:600}
        .tr-p{display:flex;align-items:center;gap:4mm}.tr-p img{width:18mm;height:18mm;object-fit:contain}`
});

/* ---------- 👂 Escucha y señala ---------- */
Impresos.registrar('escucha', {
    ico: '👂', resp: true, min: 4, img: true,
    opciones: op => `<span class="t">${_ih('visor.imp2.cuantas')}</span><div class="chips">${[6, 9].map(n => `<label class="chip"><input type="radio" name="op-c" value="${n}" ${(op.cuantas || 6) === n ? 'checked' : ''}> ${n}</label>`).join('')}</div>
        <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-qr" ${op.qr !== false ? 'checked' : ''}> 📱 ${_ih('visor.imp2.esc.qr')}</label>`,
    leer: () => ({ cuantas: +(document.querySelector('input[name="op-c"]:checked')?.value || 6), qr: $('op-qr').checked }),
    filtrar: pool => pool.filter(w => w.img),
    modelo(orden, c, r, i) {
        const ws = Impresos.rotar(orden, i, c.op.cuantas).slice(0, c.op.cuantas);
        return { ws, lectura: barajar(ws, r) };
    },
    pagina(m, c, docente) {
        const L = c.L;
        if (docente) return `<ol class="lectura">${m.lectura.map(w => `<li><b${Texto.lang(L)}>${esc(w.t[L].texto)}</b>${c.op.qr && w.t[L].audio ? `<span class="qr-mini">${QR.svg(QR.enlace('#oir=' + encodeURIComponent(w.id) + '&l=' + encodeURIComponent(L)), 16)}</span>` : ''}
            <span class="nota-d">→ ${m.ws.indexOf(w) + 1}</span></li>`).join('')}</ol><p class="nota-l">${_ih('visor.imp2.esc.docente')}</p>`;
        return `<div class="esc-g" style="grid-template-columns:repeat(${m.ws.length > 6 ? 3 : 2},1fr)">${m.ws.map((w, i) => `<div class="item"><span class="num-e">${i + 1}</span><img src="${esc(Datos.img(w, 500))}" alt=""><span class="caja-n"></span></div>`).join('')}</div>`;
    },
    css: `.esc-g{flex:1;display:grid;grid-auto-rows:1fr;gap:4mm;min-height:0}.esc-g .item{position:relative;border:.5mm solid #111;border-radius:3mm;padding:3mm;display:flex;flex-direction:column;align-items:center;min-height:0}
        .esc-g img{flex:1;min-height:0}.num-e{position:absolute;top:1.5mm;left:2mm;font-size:9pt;color:#666}.caja-n{width:14mm;height:12mm;border:.5mm solid #111;border-radius:2mm}
        .lectura{font-size:calc(16pt*var(--k));line-height:2}.lectura li{display:flex;align-items:center;gap:4mm}.qr-mini svg{vertical-align:middle}.nota-l{font-size:10pt;color:#444}`
});

/* ---------- 📚 Glosario de la unidad ---------- */
Impresos.registrar('glosario', {
    ico: '📚', min: 2, versiones: false, hojas: false,
    cuantas: (c, pool) => Math.ceil(pool.length / 12),
    opciones: op => `<label class="t" for="op-l2">${_ih('visor.imp2.glo.otra')}</label><select id="op-l2" class="modern-select"><option value="">— ${_ih('visor.imp2.glo.solo')} —</option>${Datos.langs.filter(l => Datos.conteo(l.id).textos > 0).map(l => `<option value="${esc(l.id)}" ${l.id === op.L2 ? 'selected' : ''}>${esc(l.nombre)}</option>`).join('')}</select>
        <label class="chip" style="margin-top:6px"><input type="checkbox" id="op-desc" ${op.desc ? 'checked' : ''}> 📖 ${_ih('visor.imp2.glo.desc')}</label>`,
    leer: () => ({ L2: $('op-l2').value, desc: $('op-desc').checked }),
    filtrar: pool => pool,
    modelo(orden, c, r, i) { const ws = orden.slice().sort((a, b) => a.t[c.L].texto.localeCompare(b.t[c.L].texto)); return { ws: ws.slice(i * 12, i * 12 + 12) }; },
    pagina(m, c) {
        const L = c.L, L2 = c.op.L2 && c.op.L2 !== L ? c.op.L2 : '';
        return `<div class="glo">${m.ws.map(w => `<div class="g-it">${w.img ? `<img src="${esc(Datos.img(w, 300))}" alt="">` : '<span></span>'}<div><b${Texto.lang(L)}>${esc(w.t[L].texto)}</b>${L2 && w.t[L2] ? `<div class="g-2"${Texto.lang(L2)}>${esc(w.t[L2].texto)}</div>` : ''}${c.op.desc && w.t[L].desc ? `<div class="g-d"${Texto.lang(L)}>${esc(w.t[L].desc)}</div>` : ''}</div></div>`).join('')}</div>`;
    },
    css: `.glo{flex:1;display:grid;grid-template-columns:1fr 1fr;grid-auto-rows:min-content;gap:3mm}.g-it{display:grid;grid-template-columns:22mm 1fr;gap:3mm;align-items:center;border-bottom:.3mm solid #ccc;padding-bottom:2mm;break-inside:avoid}
        .g-it img{width:22mm;height:22mm;object-fit:contain}.g-it b{font-size:calc(15pt*var(--k))}.g-2{font-size:calc(12pt*var(--k));color:#333}.g-d{font-size:calc(10pt*var(--k));color:#444}`
});

/* ---------- ✅ Lista de cotejo (el docente la llena a mano; el visor no registra a nadie) ---------- */
Impresos.registrar('cotejo', {
    ico: '✅', min: 0, versiones: false,
    opciones: op => `<span class="t">${_ih('visor.imp2.cot.columnas')}</span><div class="chips">${['palabras', 'habilidades'].map(k => `<label class="chip"><input type="radio" name="op-k" value="${k}" ${(op.cols || 'palabras') === k ? 'checked' : ''}> ${_ih('visor.imp2.cot.' + k)}</label>`).join('')}</div>
        <input id="op-hab" class="input-box" style="margin-top:6px" maxlength="300" value="${esc(op.hab || '')}" placeholder="${esc(_i('visor.imp2.cot.hab.ph'))}">
        <span class="t">${_ih('visor.imp2.cot.filas')}</span><div class="chips">${[15, 20, 25, 30].map(n => `<label class="chip"><input type="radio" name="op-f" value="${n}" ${(op.filas || 25) === n ? 'checked' : ''}> ${n}</label>`).join('')}</div>`,
    leer: () => ({ cols: document.querySelector('input[name="op-k"]:checked')?.value || 'palabras', hab: $('op-hab').value.trim(), filas: +(document.querySelector('input[name="op-f"]:checked')?.value || 25) }),
    filtrar: pool => pool,
    modelo(orden, c) { return { ws: c.op.cols === 'palabras' ? orden.slice(0, 10) : [] }; },
    pagina(m, c) {
        const cols = c.op.cols === 'palabras' ? m.ws.map(w => w.t[c.L].texto) : c.op.hab.split(/[,;]+/).map(x => x.trim()).filter(Boolean).slice(0, 10);
        return `<table class="cot" style="--fh:${c.op.filas}"><thead><tr><th class="nom">${_ih('visor.imp2.cot.nombre')}</th>${cols.map(x => `<th class="v"><span${Texto.lang(c.L)}>${esc(x)}</span></th>`).join('')}</tr></thead>
            <tbody>${Array.from({ length: c.op.filas }, (_, i) => `<tr><td class="nom">${i + 1}.</td>${cols.map(() => '<td></td>').join('')}</tr>`).join('')}</tbody></table>
            <p class="nota-l">${_ih('visor.imp2.cot.nota')}</p>`;
    },
    css: `.cot{width:100%;border-collapse:collapse;font-size:9pt}.cot th,.cot td{border:.3mm solid #111;padding:.8mm}.cot td{box-sizing:border-box;height:calc((var(--h) - 90mm) / var(--fh))}.cot .nom{width:55mm;text-align:left}
        .cot th.v{height:26mm;vertical-align:bottom}.cot th.v span{writing-mode:vertical-rl;transform:rotate(180deg);font-size:10pt}`
});

/* ---------- En pantalla: ¿Con qué letra empieza? ---------- */
const QueLetra = {
    abrir() {
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🔤 ${_ih('visor.act.queLetra')}</h2><p class="nota">${_ih('visor.ql.d')}</p>
            <label class="t" for="ql-l">${_ih('visor.lengua')}</label><select id="ql-l" class="modern-select">${Comun.lenguas({ campo: 'img', min: 6 })}</select>
            <div class="acciones"><button class="btn btn-primary" onclick="QueLetra.iniciar()">🚀 ${_ih('visor.empezar')}</button></div></div>`);
    },
    iniciar() {
        const L = $('ql-l').value;
        this.s = { L, alfa: Texto.alfabeto(L), pool: Datos.paraJuegos(L, { img: true }).filter(w => _letrasHoja(w, L).length) };
        this.s.orden = barajar(this.s.pool).slice(0, 10);
        this.s.r = Rondas.nuevo('🔤 ' + _i('visor.act.queLetra'), L, this.s.orden.length);
        Juego.activo = { detener() {} };
        this.ronda();
    },
    ronda() {
        const s = this.s, w = s.orden[s.r.n];
        if (!w) return Rondas.final(s.r, () => this.abrir());
        const ini = _letrasHoja(w, s.L)[0], ops = barajar([ini, ...barajar(s.alfa.filter(a => a !== ini)).slice(0, 3)]);
        UI.cuerpo(`${Rondas.cabeza(s.r)}<img class="ficha-img" src="${esc(Datos.img(w, 600))}" alt="${esc(w.alt || '')}">${Sonido.boton(w, s.L, '🔊 ' + esc(_i('visor.escuchar')))}
            <div class="opciones-txt ql" id="ql-ops"${Texto.lang(s.L)}>${ops.map(l => `<button class="btn btn-secondary opcion-txt" data-l="${esc(l)}">${esc(l)}</button>`).join('')}</div><div id="ql-res"></div>`);
        $('ql-ops').addEventListener('click', e => {
            const b = e.target.closest('[data-l]'); if (!b || s.resp) return; s.resp = true;
            const ok = b.dataset.l === ini; b.classList.add(ok ? 'bien' : 'mal');
            $('ql-ops').querySelectorAll('[data-l]').forEach(x => { if (x.dataset.l === ini) x.classList.add('bien'); });
            ok ? (SoundFX.ok(), s.r.bien++) : SoundFX.error(); s.r.jugadas.push({ id: w.id, ok }); s.r.n++;
            if (w.t[s.L].audio) Sonido.tocar(w.t[s.L].audio);
            $('ql-res').innerHTML = `<div class="round-result ${ok ? 'win' : 'fail'}">${ok ? '✅' : '❌'} <span${Texto.lang(s.L)}>${esc(w.t[s.L].texto)}</span></div><div class="acciones"><button class="btn btn-primary btn-block" id="ql-sig">➡️ ${_ih('visor.siguiente')}</button></div>`;
            $('ql-sig').addEventListener('click', () => { s.resp = false; this.ronda(); });
        });
    }
};

/** Enlace de un QR de audio (#oir=<palabra>&l=<lengua>): abre la ficha lista para escuchar. */
function revisarEnlaceOir() {
    const h = location.hash || '';
    if (!h.startsWith('#oir=')) return;
    const p = new URLSearchParams(h.slice(1)), id = p.get('oir'), L = p.get('l');
    history.replaceState(null, '', location.pathname + location.search);
    if (L && Datos.langById.has(L)) Datos.fijarLengua(L);
    if (Datos.byId.has(id)) { Ficha.abrir(id); const w = Datos.byId.get(id), a = w.t[Datos.lengua]?.audio; if (a) Promise.resolve(Sonido.tocar(a)).catch(() => {}); }
}
</script>
