<script>
/* =============================================
   APRENDER: actividades tranquilas, una persona a su ritmo
   ============================================= */

/** Rondas con aciertos para las actividades de aprender. */
const Rondas = {
    nuevo(titulo, L, total) { return { titulo, L, total, n: 0, bien: 0, jugadas: [] }; },
    cabeza(r) { return `<div class="score-board"><span class="player-pill" style="border-color:var(--primary)">${r.n}/${r.total} · ✓ ${r.bien}</span></div>`; },
    final(r, reintentar) {
        const pct = r.total ? Math.round(100 * r.bien / r.total) : 0;
        const vistos = new Set(), lista = r.jugadas.filter(x => !vistos.has(x.id) && vistos.add(x.id));
        UI.cuerpo(`<h2 class="modal-titulo">${esc(r.titulo)}</h2>
            <div class="round-result ${pct >= 60 ? 'win' : 'fail'}" style="font-size:1.3em">${pct >= 90 ? '🌟' : pct >= 60 ? '👏' : '💪'} ${_ih('visor.rondas.resultado', { bien: r.bien, total: r.total })}</div>
            ${Comun.palabrasJugadas(lista, r.L)}
            <div class="acciones"><button class="btn btn-secondary" id="r-otra">🔁 ${_ih('visor.otraVez')}</button><button class="btn btn-primary" onclick="UI.cerrarModal()">${_ih('visor.cerrar')}</button></div>`);
        $('r-otra').addEventListener('click', reintentar);
        Juego.activo = null;
    }
};

/** Distancia de edición entre dos listas de letras (para decir “casi”). */
function distancia(a, b) {
    const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
    for (let j = 1; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++)
        d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    return d[a.length][b.length];
}

/* ---------- Repaso con tarjetas (cajas de Leitner, se guarda en el dispositivo) ---------- */
const Repaso = {
    DIAS: [0, 0, 1, 3, 7, 16],
    clave(L) { return `repaso:${Datos.base}:${L}`; },
    abrir() {
        const L = Datos.lengua;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🃏 ${_ih('visor.act.repaso')}</h2>
            <p class="nota">${_ih('visor.rep.d')}</p>
            <label class="t" for="r-l">${_ih('visor.rep.lengua')}</label><select id="r-l" class="modern-select">${Comun.lenguas({ campo: 'img', min: 5 }, L)}</select>
            <span class="t">${_ih('visor.rep.frente')}</span>
            <div class="chips"><label class="chip"><input type="radio" name="r-f" value="img" checked> 🖼️ ${_ih('visor.mem.cara.img')}</label>
                <label class="chip"><input type="radio" name="r-f" value="audio"> 🔊 ${_ih('visor.mem.cara.audio')}</label>
                <label class="chip"><input type="radio" name="r-f" value="otra"> 🌐 ${_ih('visor.otraLengua')}</label></div>
            <select id="r-l2" class="modern-select oculto" aria-label="${esc(_i('visor.otraLengua'))}">${Comun.lenguas(null, '')}</select>
            <div id="r-prog"></div><div id="r-msg"></div>
            <div class="acciones"><button class="btn btn-primary" onclick="Repaso.iniciar()">🚀 ${_ih('visor.rep.empezar', { n: 10 })}</button></div></div>`);
        document.querySelectorAll('input[name="r-f"]').forEach(x => x.addEventListener('change', () => $('r-l2').classList.toggle('oculto', x.value !== 'otra' || !x.checked)));
        const prog = () => {
            const est = LS.get(this.clave($('r-l').value), {}), cajas = [0, 0, 0, 0, 0, 0];
            Object.values(est).forEach(e => cajas[e.c]++);
            const total = Object.keys(est).length;
            $('r-prog').innerHTML = total ? `<div class="caja">${_ih('visor.rep.avance', { vistas: total, aprendidas: cajas[5] + cajas[4] })}
                <button class="btn btn-secondary" style="min-height:32px;padding:4px 10px;margin-left:6px" id="r-borra">${_ih('visor.rep.deCero')}</button></div>` : '';
            $('r-borra')?.addEventListener('click', () => { if (confirm(_i('visor.rep.borrar'))) { LS.del(this.clave($('r-l').value)); prog(); } });
        };
        $('r-l').addEventListener('change', prog); prog();
    },
    iniciar() {
        const L = $('r-l').value, frente = document.querySelector('input[name="r-f"]:checked').value, L2 = $('r-l2').value;
        const pool = Datos.paraJuegos(L, { img: frente === 'img', audio: frente === 'audio' }).filter(w => frente !== 'otra' || (w.t[L2] && L2 !== L));
        if (pool.length < 3) return $('r-msg').innerHTML = '<div class="error-box">' + _ih('visor.rep.error') + '</div>';
        const est = LS.get(this.clave(L), {}), hoy = Date.now();
        const tocan = barajar(pool.filter(w => est[w.id] && est[w.id].p <= hoy)).sort((a, b) => est[a.id].c - est[b.id].c);
        const nuevas = barajar(pool.filter(w => !est[w.id]));
        let sesion = [...tocan, ...nuevas].slice(0, 10);
        if (sesion.length < 10) sesion = [...sesion, ...barajar(pool.filter(w => !sesion.includes(w))).slice(0, 10 - sesion.length)];
        this.s = { L, L2, frente, sesion, i: 0, est, r: Rondas.nuevo('🃏 ' + _i('visor.act.repaso'), L, sesion.length) };
        Juego.activo = { detener() {} };
        this.tarjeta();
    },
    tarjeta(volteada = false) {
        const s = this.s, w = s.sesion[s.i];
        if (!w) { LS.set(this.clave(s.L), s.est); return Rondas.final(s.r, () => this.abrir()); }
        const t = w.t[s.L];
        let frente = '';
        if (s.frente === 'img') frente = `<img src="${esc(Datos.img(w, 600))}" alt="">`;
        else if (s.frente === 'audio') frente = `<div style="font-size:3em">👂</div>${Sonido.boton(w, s.L, '🔊 ' + esc(_i('visor.escuchar')))}`;
        else frente = `<div class="nota">${esc(Datos.nombreLengua(s.L2))}</div><div class="gran-palabra"${Texto.lang(s.L2)}>${esc(w.t[s.L2].texto)}</div>`;
        const caja = s.est[w.id]?.c || 0;
        UI.cuerpo(`${Rondas.cabeza({ ...s.r, n: s.i + 1 })}
            <div class="tarjeta-repaso">${frente}
                <div class="nota" style="margin-top:6px">${'●'.repeat(caja)}${'○'.repeat(5 - caja)}</div>
                ${volteada ? `<div class="gran-palabra" style="margin-top:10px"${Texto.lang(s.L)}>${esc(t.texto)}</div>${s.frente !== 'audio' ? Sonido.boton(w, s.L) : ''}${t.silabas ? `<div class="nota">${esc(t.silabas)}</div>` : ''}` : ''}</div>
            <div class="acciones">${volteada
                ? `<button class="btn btn-secondary" id="r-no" style="flex:1;min-height:60px">😕 ${_ih('visor.rep.no')}</button><button class="btn btn-success" id="r-si" style="flex:1;min-height:60px">🙂 ${_ih('visor.rep.si')}</button>`
                : `<button class="btn btn-primary btn-block" id="r-ver" style="min-height:60px">👀 ${_ih('visor.rep.ver')}</button>`}</div>`);
        if (s.frente === 'audio' && !volteada && t.audio) Sonido.tocar(t.audio);
        if (volteada && s.frente !== 'audio' && t.audio) Sonido.tocar(t.audio);
        $('r-ver')?.addEventListener('click', () => this.tarjeta(true));
        const responder = ok => {
            const c = ok ? Math.min(5, caja + 1) : 1;
            s.est[w.id] = { c, p: Date.now() + this.DIAS[c] * 864e5 };
            if (ok) s.r.bien++;
            s.r.jugadas.push({ id: w.id, ok });
            LS.set(this.clave(s.L), s.est);
            if (!ok && !s.sesion.slice(s.i + 1).includes(w)) s.sesion.splice(Math.min(s.sesion.length, s.i + 4), 0, w); // vuelve en un rato
            s.r.total = s.sesion.length;
            s.i++; this.tarjeta();
        };
        $('r-si')?.addEventListener('click', () => responder(true));
        $('r-no')?.addEventListener('click', () => responder(false));
    }
};

/* ---------- ¿Cuál escuchaste? ---------- */
const Escucha = {
    abrir() {
        UI.abrir(`<div class="form"><h2 class="modal-titulo">👂 ${_ih('visor.act.escucha')}</h2>
            <p class="nota">${_ih('visor.esc.d')}</p>
            <label class="t" for="e-l">${_ih('visor.lengua')}</label><select id="e-l" class="modern-select">${Comun.lenguas({ campo: 'imgAudio', min: 4 })}</select>
            <label class="chip"><input type="checkbox" id="e-txt"> 📝 ${_ih('visor.esc.texto')}</label>
            <div class="acciones"><button class="btn btn-primary" onclick="Escucha.iniciar()">🚀 ${_ih('visor.empezar')}</button></div></div>`);
    },
    iniciar() {
        const L = $('e-l').value;
        this.s = { L, txt: $('e-txt').checked, pool: Datos.paraJuegos(L, { img: true, audio: true }) };
        this.s.orden = barajar(this.s.pool).slice(0, 10);
        this.s.r = Rondas.nuevo('👂 ' + _i('visor.act.escucha'), L, this.s.orden.length);
        Juego.activo = { detener() {} };
        this.ronda();
    },
    ronda() {
        const s = this.s, w = s.orden[s.r.n];
        if (!w) return Rondas.final(s.r, () => this.abrir());
        const otras = barajar(s.pool.filter(x => x !== w && x.img !== w.img)).slice(0, 3);
        const ops = barajar([w, ...otras]);
        UI.cuerpo(`${Rondas.cabeza(s.r)}
            <div style="text-align:center;margin-bottom:10px">${Sonido.boton(w, s.L, '🔊 ' + esc(_i('visor.escucharOtraVez')))}
            ${s.txt ? `<div class="gran-palabra"${Texto.lang(s.L)}>${esc(w.t[s.L].texto)}</div>` : ''}</div>
            <div class="opciones-img" id="e-ops">${ops.map(o => `<button class="opcion-img" data-id="${esc(o.id)}" aria-label="${esc(_i('visor.opcion'))}"><img src="${esc(Datos.img(o, 400))}" alt=""></button>`).join('')}</div>
            <div id="e-res"></div>`);
        Sonido.tocar(w.t[s.L].audio);
        $('e-ops').addEventListener('click', e => {
            const b = e.target.closest('.opcion-img'); if (!b || s.resp) return;
            s.resp = true;
            const ok = b.dataset.id === w.id;
            b.classList.add(ok ? 'bien' : 'mal');
            $('e-ops').querySelector(`[data-id="${CSS.escape(w.id)}"]`).classList.add('bien');
            ok ? (SoundFX.ok(), s.r.bien++) : SoundFX.error();
            s.r.jugadas.push({ id: w.id, ok }); s.r.n++;
            $('e-res').innerHTML = `<div class="round-result ${ok ? 'win' : 'fail'}">${ok ? '✅' : '❌'} <span${Texto.lang(s.L)}>${esc(w.t[s.L].texto)}</span></div>
                <div class="acciones"><button class="btn btn-primary btn-block" id="e-sig">➡️ ${_ih('visor.siguiente')}</button></div>`;
            $('e-sig').addEventListener('click', () => { s.resp = false; this.ronda(); });
        });
    }
};

/* ---------- ¿En qué lengua lo dicen? ---------- */
const EnQueLengua = {
    abrir() {
        this.lenguas = Datos.langs.filter(l => Datos.words.filter(w => w.t[l.id]).length >= 10);
        const s = { pool: Datos.words.filter(w => this.lenguas.filter(l => w.t[l.id]).length >= 2) };
        s.orden = barajar(s.pool).slice(0, 10);
        s.r = Rondas.nuevo('🌎 ' + _i('visor.act.enquelengua'), Datos.lengua, s.orden.length);
        this.s = s; Juego.activo = { detener() {} };
        UI.abrir('<div></div>');
        this.ronda();
    },
    ronda() {
        const s = this.s, w = s.orden[s.r.n];
        if (!w) return Rondas.final(s.r, () => this.abrir());
        const tiene = this.lenguas.filter(l => w.t[l.id]);
        const L = elegir(tiene).id;
        const ops = barajar([L, ...barajar(this.lenguas.map(l => l.id).filter(x => x !== L)).slice(0, 3)]);
        UI.cuerpo(`${Rondas.cabeza(s.r)}
            <div class="tarjeta-repaso">${w.img ? `<img src="${esc(Datos.img(w, 500))}" alt="">` : ''}
                <div class="gran-palabra"${Texto.lang(L)}>${esc(w.t[L].texto)}</div>${Sonido.boton(w, L)}</div>
            <p style="text-align:center;font-weight:800">${_ih('visor.eql.pregunta')}</p>
            <div class="chips" id="q-ops" style="justify-content:center">${ops.map(x => `<button class="btn btn-secondary" data-l="${esc(x)}" style="flex:1;min-width:140px">${esc(Datos.nombreLengua(x))}</button>`).join('')}</div>
            <div id="q-res"></div>`);
        if (w.t[L].audio) Sonido.tocar(w.t[L].audio);
        $('q-ops').addEventListener('click', e => {
            const b = e.target.closest('[data-l]'); if (!b || s.resp) return;
            s.resp = true;
            const ok = b.dataset.l === L;
            ok ? (SoundFX.ok(), s.r.bien++) : SoundFX.error();
            $('q-ops').querySelectorAll('[data-l]').forEach(x => { if (x.dataset.l === L) x.className = 'btn btn-success'; else if (x === b) x.style.borderColor = 'var(--error)'; });
            s.r.jugadas.push({ id: w.id, ok }); s.r.n++;
            $('q-res').innerHTML = `<div class="round-result ${ok ? 'win' : 'fail'}">${ok ? '✅ ' + _ih('visor.eql.si') : '❌'} ${_ih('visor.eql.es', { lengua: Datos.nombreLengua(L) })}</div>
                <div class="nota" style="text-align:center">${_ih('visor.eql.otras')}</div>
                ${tiene.filter(l => l.id !== L).map(l => `<div class="trad"><div class="l">${esc(l.nombre)}</div><div class="x"><span${Texto.lang(l.id)}>${esc(w.t[l.id].texto)}</span>${Sonido.boton(w, l.id)}</div></div>`).join('')}
                <div class="acciones"><button class="btn btn-primary btn-block" id="q-sig">➡️ ${_ih('visor.siguiente')}</button></div>`;
            $('q-sig').addEventListener('click', () => { s.resp = false; this.ronda(); });
        });
    }
};

/* ---------- Arma la palabra (letras reales o sílabas) ---------- */
const Arma = {
    abrir() {
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🔤 ${_ih('visor.act.arma')}</h2>
            <label class="t" for="a-l">${_ih('visor.lengua')}</label><select id="a-l" class="modern-select">${Comun.lenguas({ campo: 'largas', min: 6 })}</select>
            <span class="t">${_ih('visor.arma.piezas')}</span>
            <div class="chips"><label class="chip"><input type="radio" name="a-p" value="sil" checked> ${_ih('visor.arma.silabas')}</label>
                <label class="chip"><input type="radio" name="a-p" value="letra"> ${_ih('visor.arma.letras')}</label></div>
            <div class="acciones"><button class="btn btn-primary" onclick="Arma.iniciar()">🚀 ${_ih('visor.empezar')}</button></div></div>`);
    },
    piezas(w, L, sil) {
        const t = w.t[L];
        if (sil && t.silabas) { const p = t.silabas.split(/[-·.\s|]+/).map(x => x.trim().toLowerCase()).filter(Boolean); if (p.length >= 2) return p; }
        return Texto.letras(t.texto, L);
    },
    iniciar() {
        const L = $('a-l').value, sil = document.querySelector('input[name="a-p"]:checked').value === 'sil';
        const pool = Datos.paraJuegos(L, { largas: true }).filter(w => { const n = this.piezas(w, L, sil).length; return n >= 2 && n <= 12 && !/\s/.test(w.t[L].texto.trim()); });
        const orden = barajar([...barajar(pool.filter(w => w.img)), ...pool.filter(w => !w.img)].slice(0, 30)).sort((a, b) => !!b.img - !!a.img).slice(0, 8);
        this.s = { L, sil, orden, r: Rondas.nuevo('🔤 ' + _i('visor.act.arma'), L, orden.length) };
        Juego.activo = { detener() {} };
        this.ronda();
    },
    ronda() {
        const s = this.s, w = s.orden[s.r.n];
        if (!w) return Rondas.final(s.r, () => this.abrir());
        s.meta = this.piezas(w, s.L, s.sil);
        let mezcla = barajar(s.meta.map((p, i) => ({ p, i })));
        for (let k = 0; k < 5 && mezcla.every((x, j) => x.p === s.meta[j]); k++) mezcla = barajar(mezcla);
        s.fichas = mezcla; s.puestas = []; s.fallos = 0;
        this.pintar();
        if (w.t[s.L].audio) Sonido.tocar(w.t[s.L].audio);
    },
    pintar(final = '') {
        const s = this.s, w = s.orden[s.r.n];
        UI.cuerpo(`${Rondas.cabeza(s.r)}
            <div class="tarjeta-repaso">${w.img ? `<img src="${esc(Datos.img(w, 500))}" alt="" style="max-height:200px">` : `<div class="nota">${esc(Datos.categoria(w.cat, s.L))}</div>`}
                <div>${Sonido.boton(w, s.L, '🔊 ' + esc(_i('visor.escuchar')))}</div></div>
            <div class="fichas" id="a-huecos"${Texto.lang(s.L)}>${s.meta.map((_, j) => { const f = s.puestas[j]; return `<button class="hueco" data-h="${j}" aria-label="${esc(_i('visor.arma.espacio', { n: j + 1 }))}">${f ? esc(f.p) : ''}</button>`; }).join('')}</div>
            <div class="fichas" id="a-fichas"${Texto.lang(s.L)}>${s.fichas.map((f, k) => `<button class="ficha-letra ${s.puestas.includes(f) ? 'usada' : ''}" data-k="${k}" ${s.puestas.includes(f) ? 'disabled' : ''}>${esc(f.p)}</button>`).join('')}</div>
            ${final}`);
        $('a-fichas').addEventListener('click', e => {
            const b = e.target.closest('.ficha-letra'); if (!b || b.disabled || final) return;
            const libre = s.meta.findIndex((_, j) => !s.puestas[j]); if (libre < 0) return;
            s.puestas[libre] = s.fichas[+b.dataset.k];
            this.revisar();
        });
        $('a-huecos').addEventListener('click', e => { const h = e.target.closest('.hueco'); if (!h || final) return; s.puestas[+h.dataset.h] = undefined; this.pintar(); });
        $('a-sig')?.addEventListener('click', () => this.ronda());
    },
    revisar() {
        const s = this.s, w = s.orden[s.r.n];
        if (s.puestas.filter(Boolean).length < s.meta.length) return this.pintar();
        const ok = s.puestas.every((f, j) => f.p === s.meta[j]);
        if (!ok) {
            s.fallos++; SoundFX.error();
            if (s.fallos < 2) { s.puestas = s.puestas.map((f, j) => f.p === s.meta[j] ? f : undefined); UI.toast(_i('visor.arma.casi')); return this.pintar(); }
        } else SoundFX.ok();
        const gano = ok;
        if (gano) s.r.bien++;
        s.r.jugadas.push({ id: w.id, ok: gano }); s.r.n++;
        if (!gano) s.puestas = s.meta.map(p => ({ p }));
        this.s.r.n--; // se muestra la misma tarjeta con el resultado
        this.pintar(`<div class="round-result ${gano ? 'win' : 'fail'}">${gano ? '✅ ' + _ih('visor.muyBien') : '❌ ' + _ih('visor.asiSeEscribe')} <span${Texto.lang(s.L)}>${esc(w.t[s.L].texto)}</span></div>
            <div class="acciones"><button class="btn btn-primary btn-block" id="a-sig">➡️ ${_ih('visor.siguiente')}</button></div>`);
        this.s.r.n++;
        if (w.t[s.L].audio) Sonido.tocar(w.t[s.L].audio);
    }
};

/* ---------- Dictado ---------- */
const Dictado = {
    abrir() {
        UI.abrir(`<div class="form"><h2 class="modal-titulo">✍️ ${_ih('visor.act.dictado')}</h2>
            <p class="nota">${_ih('visor.dic.d')}</p>
            <label class="t" for="d-l">${_ih('visor.lengua')}</label><select id="d-l" class="modern-select">${Comun.lenguas({ campo: 'audio', min: 5 })}</select>
            <label class="chip"><input type="checkbox" id="d-img" checked> 🖼️ ${_ih('visor.dic.img')}</label>
            <div class="acciones"><button class="btn btn-primary" onclick="Dictado.iniciar()">🚀 ${_ih('visor.empezar')}</button></div></div>`);
    },
    iniciar() {
        const L = $('d-l').value;
        const orden = barajar(Datos.paraJuegos(L, { audio: true })).slice(0, 10);
        this.s = { L, img: $('d-img').checked, orden, r: Rondas.nuevo('✍️ ' + _i('visor.act.dictado'), L, orden.length) };
        Juego.activo = { detener() {} };
        this.ronda();
    },
    ronda() {
        const s = this.s, w = s.orden[s.r.n];
        if (!w) return Rondas.final(s.r, () => this.abrir());
        const alfa = Texto.alfabeto(s.L).filter(g => !/^[a-z]$/.test(g)).slice(0, 24);
        UI.cuerpo(`${Rondas.cabeza(s.r)}
            <div class="tarjeta-repaso">${s.img && w.img ? `<img src="${esc(Datos.img(w, 500))}" alt="" style="max-height:200px">` : ''}
                <div>${Sonido.boton(w, s.L, '🔊 ' + esc(_i('visor.escuchar')))}</div></div>
            <input id="d-in" class="input-box" style="font-size:1.4em;text-align:center;margin-top:10px" autocomplete="off" autocapitalize="none" spellcheck="false"${Texto.lang(s.L)} aria-label="${esc(_i('visor.dic.escribe'))}">
            ${alfa.length ? `<div class="keyboard" id="d-kb">${alfa.map(g => `<button class="key-btn" data-g="${esc(g)}">${esc(g)}</button>`).join('')}</div>` : ''}
            <div id="d-res"><div class="acciones"><button class="btn btn-primary btn-block" id="d-ok">✔️ ${_ih('visor.revisar')}</button></div></div>`);
        Sonido.tocar(w.t[s.L].audio);
        const inp = $('d-in');
        inp.addEventListener('input', () => { const p = inp.selectionStart, v = Texto.atajos(inp.value, s.L); if (v !== inp.value) { const d = v.length - inp.value.length; inp.value = v; try { inp.setSelectionRange(p + d, p + d); } catch (e) {} } });
        inp.addEventListener('keydown', e => { if (e.key === 'Enter') $('d-ok')?.click(); });
        $('d-kb')?.addEventListener('click', e => { const b = e.target.closest('[data-g]'); if (!b) return; const p = inp.selectionStart ?? inp.value.length; inp.value = inp.value.slice(0, p) + b.dataset.g + inp.value.slice(p); inp.focus(); try { inp.setSelectionRange(p + b.dataset.g.length, p + b.dataset.g.length); } catch (x) {} });
        $('d-ok').addEventListener('click', () => {
            const dicho = Texto.pliegue(Texto.atajos(inp.value, s.L));
            const reales = (w.t[s.L].formas || [w.t[s.L].texto]).map(f => Texto.pliegue(f));
            if (!dicho) return inp.focus();
            const ok = reales.includes(dicho);
            const casi = !ok && reales.some(real => distancia(Texto.letras(dicho, s.L), Texto.letras(real, s.L)) <= 1);
            ok ? (SoundFX.ok(), s.r.bien++) : SoundFX.error();
            s.r.jugadas.push({ id: w.id, ok }); s.r.n++;
            inp.disabled = true;
            $('d-res').innerHTML = `<div class="round-result ${ok ? 'win' : 'fail'}">${ok ? '✅ ' + _ih('visor.dic.exacto') : casi ? '🟡 ' + _ih('visor.dic.casi') : '❌ ' + _ih('visor.dic.seEscribe')} <span${Texto.lang(s.L)}>${(w.t[s.L].formas || [w.t[s.L].texto]).map(esc).join(' · ')}</span></div>
                ${w.t[s.L].nota ? `<div class="nota">📌 ${esc(w.t[s.L].nota)}</div>` : ''}
                <div class="acciones"><button class="btn btn-primary btn-block" id="d-sig">➡️ ${_ih('visor.siguiente')}</button></div>`;
            $('d-sig').addEventListener('click', () => this.ronda());
        });
        setTimeout(() => inp.focus(), 60);
    }
};
</script>
