<script>
/* =============================================
   BASTA (Stop / Tutti frutti) con las categorías del catálogo
   ============================================= */
const Basta = {
    cfg: null, ronda: 0, total: 0, reloj: null,
    detener() { clearInterval(this.reloj); this.reloj = null; },
    setup() {
        const L = Datos.lengua;
        const lenguas = Comun.lenguas({ campo: 'cats', min: 3 }, L);
        UI.abrir(`<div class="form"><h2 class="modal-titulo">✋ ${_ih('visor.act.basta')}</h2>
            <p class="nota">${_ih('visor.basta.d')}</p>
            <label class="t" for="b-l">${_ih('visor.lengua')}</label><select id="b-l" class="modern-select">${lenguas}</select>
            <span class="t">${_ih('visor.basta.columnas')}</span><div id="b-cats"></div>
            <label class="chip" style="margin-top:6px"><input type="checkbox" id="b-libre"> ➕ ${_ih('visor.basta.libre')}</label>
            <span class="t">${_ih('visor.basta.como')}</span>
            <div class="chips">
                <label class="chip"><input type="radio" name="b-m" value="solo" checked> 📱 ${_ih('visor.basta.celular')}</label>
                <label class="chip"><input type="radio" name="b-m" value="arbitro"> 📋 ${_ih('visor.basta.papel')}</label>
            </div>
            <div id="b-comp" class="caja" style="margin-top:8px">
                <label class="chip"><input type="checkbox" id="b-compartida"> 🔗 ${_ih('visor.basta.compartida')}</label>
                <div id="b-comp-datos" class="oculto" style="margin-top:8px">
                    <p class="nota">${_ih('visor.basta.compartida.d')}</p>
                    <div class="fila"><input id="b-nom" class="input-box" style="flex:2;min-width:140px" placeholder="${esc(_i('visor.nombrePartida'))}" aria-label="${esc(_i('visor.nombrePartida'))}">
                    <input id="b-num" class="input-box" style="flex:1;min-width:90px" inputmode="numeric" maxlength="4" placeholder="0000" value="${cuatroDigitos()}" aria-label="${esc(_i('visor.cuatroNumeros'))}"></div>
                </div>
            </div>
            <span class="t">${_ih('visor.basta.tiempo')}</span><div class="chips">${[[60, '1 min'], [90, '1½ min'], [120, '2 min'], [0, _i('visor.sinLimite')]].map(([n, t], i) => `<label class="chip"><input type="radio" name="b-t" value="${n}" ${i === 1 ? 'checked' : ''}> ${esc(t)}</label>`).join('')}</div>
            <span class="t">${_ih('visor.rondas')}</span><div class="chips">${[3, 5, 8].map((n, i) => `<label class="chip"><input type="radio" name="b-r" value="${n}" ${i === 1 ? 'checked' : ''}> ${n}</label>`).join('')}</div>
            <div id="b-msg"></div>
            <div class="acciones"><button class="btn btn-primary" onclick="Basta.iniciar()">🚀 ${_ih('visor.empezar')}</button></div></div>`);
        const pintarCats = () => {
            const Lx = $('b-l').value;
            const usadas = new Map();
            Datos.paraJuegos(Lx).forEach(w => usadas.set(w.cat, (usadas.get(w.cat) || 0) + 1));
            const top = [...usadas.entries()].filter(([c]) => Datos.cats.has(c)).sort((a, b) => b[1] - a[1]).slice(0, 4).map(x => x[0]);
            $('b-cats').innerHTML = Comun.categorias(Lx, 'b-cat', top);
        };
        $('b-l').addEventListener('change', pintarCats); pintarCats();
        $('b-compartida').addEventListener('change', e => $('b-comp-datos').classList.toggle('oculto', !e.target.checked));
    },
    /** Letras con las que empieza al menos una palabra del catálogo en las columnas elegidas. */
    letrasPosibles(L, cats) {
        const cuenta = new Map();
        Datos.paraJuegos(L, { cats }).forEach(w => {
            const g = Texto.letras(w.t[L].texto, L)[0]; if (!g) return;
            const k = Texto.pliegue(g);
            if (!cuenta.has(k)) cuenta.set(k, { g, cats: new Set() });
            cuenta.get(k).cats.add(w.cat);
        });
        const minimo = Math.min(2, cats.length);
        let letras = [...cuenta.values()].filter(x => x.cats.size >= minimo).map(x => x.g);
        if (letras.length < 3) letras = [...cuenta.values()].map(x => x.g);
        const alfa = Texto.alfabeto(L);
        return letras.sort((a, b) => (alfa.indexOf(a) + 1 || 999) - (alfa.indexOf(b) + 1 || 999));
    },
    iniciar() {
        const L = $('b-l').value, cats = Comun.marcados('b-cat'), libre = $('b-libre').checked;
        const ncols = cats.length + (libre ? 1 : 0);
        if (ncols < 3 || ncols > 6) return $('b-msg').innerHTML = `<div class="error-box">${_ih('visor.basta.error.cols', { n: ncols })}</div>`;
        const letras = this.letrasPosibles(L, cats);
        if (!letras.length) return $('b-msg').innerHTML = '<div class="error-box">' + _ih('visor.basta.error.vacias') + '</div>';
        const compartida = $('b-compartida').checked;
        let semilla = String(Math.random()), etiqueta = '';
        if (compartida) {
            const nom = $('b-nom').value.trim(), num = $('b-num').value.replace(/\D/g, '');
            if (!nom || num.length !== 4) return $('b-msg').innerHTML = '<div class="error-box">' + _ih('visor.error.partida') + '</div>';
            semilla = `basta|${Texto.pliegue(nom)}|${num}|${L}|${cats.slice().sort().join(',')}|${libre}`;
            etiqueta = `${nom} · ${num}`;
        }
        this.cfg = { L, cats, libre, letras: barajar(letras, azar(semilla)), modo: document.querySelector('input[name="b-m"]:checked').value,
            tiempo: +document.querySelector('input[name="b-t"]:checked').value, etiqueta };
        this.total = Math.min(+document.querySelector('input[name="b-r"]:checked').value, this.cfg.letras.length);
        this.ronda = 0; this.puntos = 0; this.nuevas = []; this.jugadas = [];
        Juego.activo = this;
        this.siguiente();
    },
    columnas() {
        return [...this.cfg.cats.map(id => ({ id, nombre: Datos.categoria(id, this.cfg.L) })), ...(this.cfg.libre ? [{ id: '*', nombre: '✨ ' + _i('visor.basta.libreCorto') }] : [])];
    },
    siguiente() {
        if (this.ronda >= this.total) return this.final();
        this.letra = this.cfg.letras[this.ronda]; this.ronda++;
        this.resto = this.cfg.tiempo;
        const cols = this.columnas(), L = this.cfg.L;
        const cab = `<div class="nota" style="text-align:center">${_ih('visor.rondaN', { n: this.ronda, total: this.total })}${this.cfg.etiqueta ? ' · ' + _ih('visor.basta.partida', { x: this.cfg.etiqueta }) : ''}</div>
            <div class="basta-letra"${Texto.lang(L)} aria-live="polite">${esc(this.letra.toLocaleUpperCase())}</div>
            ${this.cfg.tiempo ? `<div style="text-align:center"><span class="temporizador" id="b-reloj">${fmtTiempo(this.resto)}</span></div>` : ''}`;
        if (this.cfg.modo === 'arbitro') {
            UI.cuerpo(`${cab}<div class="caja"><b>${_ih('visor.basta.columnasCorto')}</b><ol>${cols.map(c => `<li style="font-size:1.2em;font-weight:800">${esc(c.nombre)}</li>`).join('')}</ol></div>
                <div class="acciones"><button class="btn btn-danger btn-block" style="background:var(--error);color:#fff;min-height:70px;font-size:1.4em" onclick="Basta.basta()">✋ ${_ih('visor.basta.grito')}</button></div>`);
        } else {
            UI.cuerpo(`${cab}<div id="b-form">${cols.map((c, i) => `<div class="basta-col"><label for="b-c${i}">${esc(c.nombre)}</label>
                <input id="b-c${i}" class="input-box" autocomplete="off" autocapitalize="none"${Texto.lang(L)}></div>`).join('')}</div>
                <div class="acciones"><button class="btn btn-danger btn-block" style="background:var(--error);color:#fff;min-height:64px;font-size:1.3em" onclick="Basta.basta()">✋ ${_ih('visor.basta.grito')}</button></div>`);
            // Los atajos de la lengua (por ejemplo + → ɨ) funcionan al escribir
            $('b-form').addEventListener('input', e => { const t = e.target, p = t.selectionStart, v = Texto.atajos(t.value, L); if (v !== t.value) { const d = v.length - t.value.length; t.value = v; try { t.setSelectionRange(p + d, p + d); } catch (x) {} } });
            setTimeout(() => $('b-c0')?.focus(), 50);
        }
        this.detener();
        if (this.cfg.tiempo) this.reloj = setInterval(() => {
            this.resto--;
            const r = $('b-reloj'); if (r) { r.textContent = fmtTiempo(Math.max(0, this.resto)); r.classList.toggle('poco', this.resto <= 10); }
            if (this.resto <= 5 && this.resto > 0) SoundFX.ok();
            if (this.resto <= 0) { SoundFX.vibrar([200, 100, 200]); this.basta(); }
        }, 1000);
    },
    /** Palabras del catálogo que empiezan con la letra en una columna. */
    delCatalogo(cat) {
        const L = this.cfg.L, k = Texto.pliegue(this.letra);
        return Datos.paraJuegos(L, cat === '*' ? {} : { cats: [cat] }).filter(w => Texto.pliegue(Texto.letras(w.t[L].texto, L)[0] || '') === k);
    },
    empiezaBien(texto) {
        const g = Texto.letras(texto, this.cfg.L)[0] || '';
        return Texto.pliegue(g) === Texto.pliegue(this.letra) || Texto.pliegue(texto).startsWith(Texto.pliegue(this.letra));
    },
    basta() {
        this.detener(); SoundFX.win();
        const cols = this.columnas(), L = this.cfg.L;
        const resp = cols.map((c, i) => Texto.atajos(($('b-c' + i)?.value || '').trim(), L));
        this.filas = cols.map((c, i) => {
            const r = resp[i], lista = this.delCatalogo(c.id);
            const hallada = r && lista.find(w => w.t[L].formas.some(f => Texto.pliegue(f) === Texto.pliegue(r)));
            const estado = !r ? 0 : hallada ? 100 : this.empiezaBien(r) ? null : 0; // null = la decide el grupo
            return { c, r, lista, hallada, estado };
        });
        this.pintarRevision();
    },
    pintarRevision() {
        const L = this.cfg.L, arbitro = this.cfg.modo === 'arbitro';
        const lista = f => f.lista.length ? `<div style="margin-top:6px">${f.lista.slice(0, 8).map(w => `<button class="chip-res" data-ficha="${esc(w.id)}"${Texto.lang(L)}>${esc(w.t[L].texto)}</button>`).join('')}${f.lista.length > 8 ? `<span class="nota"> ${_ih('visor.yMas', { n: f.lista.length - 8 })}</span>` : ''}</div>` : '<div class="nota">' + _ih('visor.basta.sinLetra') + '</div>';
        const filas = this.filas.map((f, i) => `<div class="basta-rev"><div class="r"><b>${esc(f.c.name || f.c.nombre)}</b>
            ${arbitro ? '' : `<span${Texto.lang(L)} style="font-size:1.15em;font-weight:800">${f.r ? esc(f.r) : '<span class="nota">' + _ih('visor.vacio') + '</span>'}</span>
            ${f.r ? `<span class="seg" role="group" aria-label="${esc(_i('visor.puntos'))}">${[[100, '✅ 100'], [50, '🟡 50'], [0, '❌ 0']].map(([p, t]) => `<button data-f="${i}" data-p="${p}" aria-pressed="${f.estado === p}">${t}</button>`).join('')}</span>` : ''}`}</div>
            ${!arbitro && f.hallada ? '<div class="nota" style="color:var(--success)">✓ ' + _ih('visor.basta.esta') + '</div>' : ''}
            ${!arbitro && f.r && !f.hallada && f.estado === null ? '<div class="nota">' + _ih('visor.basta.noEsta') + '</div>' : ''}
            ${!arbitro && f.r && !this.empiezaBien(f.r) ? '<div class="nota" style="color:var(--error)">' + _ih('visor.basta.noEmpieza') + '</div>' : ''}
            <details><summary class="nota">${_ih('visor.basta.delCatalogo')}</summary>${lista(f)}</details></div>`).join('');
        const pend = this.filas.some(f => f.r && f.estado === null);
        const suma = this.filas.reduce((s, f) => s + (f.estado || 0), 0);
        UI.cuerpo(`<div class="basta-letra" style="font-size:3rem"${Texto.lang(L)}>${esc(this.letra.toLocaleUpperCase())}</div>
            ${filas}
            ${arbitro ? '' : `<div class="round-result win">${_ih('visor.basta.suma', { ronda: suma, total: this.puntos + suma })}</div>`}
            <div class="acciones">${pend ? '<span class="nota">' + _ih('visor.basta.pendientes') + '</span>' : ''}
                <button class="btn btn-primary" id="b-sig" ${pend ? 'disabled' : ''}>${this.ronda >= this.total ? '📊 ' + _ih('visor.terminar') : '➡️ ' + _ih('visor.basta.siguiente')}</button></div>`);
        $('modal-body').querySelectorAll('[data-f]').forEach(b => b.addEventListener('click', () => { this.filas[+b.dataset.f].estado = +b.dataset.p; this.pintarRevision(); }));
        $('b-sig').addEventListener('click', () => {
            this.puntos += suma;
            this.filas.forEach(f => {
                if (f.hallada) this.jugadas.push({ id: f.hallada.id, ok: true });
                else if (f.r && f.estado > 0) this.nuevas.push({ palabra: f.r, columna: f.c.nombre, letra: this.letra });
            });
            this.siguiente();
        });
    },
    final() {
        this.detener(); Juego.activo = null;
        const L = this.cfg.L, arbitro = this.cfg.modo === 'arbitro';
        const texto = _i('visor.basta.msg', { lengua: Datos.nombreLengua(L), catalogo: Datos.cfg.titulo || '' }) + '\n' +
            this.nuevas.map(n => `• ${n.palabra} (${n.columna})`).join('\n') + '\n';
        const enl = this.nuevas.length ? Colabora.enlaces(_i('visor.basta.msgAsunto'), texto) : '';
        UI.cuerpo(`<h2 class="modal-titulo">✋ ${_ih('visor.basta.fin')}</h2>
            ${arbitro ? '<p>' + _ih('visor.basta.papelPuntos') + '</p>' : `<div class="round-result win" style="font-size:1.4em">⭐ ${_ih('visor.puntosN', { n: this.puntos })}</div>`}
            ${this.nuevas.length ? `<div class="caja"><b>✨ ${_ih('visor.basta.nuevas', { n: this.nuevas.length })}</b>
                <p${Texto.lang(L)}>${this.nuevas.map(n => esc(n.palabra)).join(' · ')}</p>
                ${enl ? `<p class="nota">${_ih('visor.basta.compartelas')}</p><div class="fila">${enl}</div>` : '<p class="nota">' + _ih('visor.basta.anotalas') + '</p>'}</div>` : ''}
            ${Comun.palabrasJugadas(this.jugadas, L)}
            <div class="acciones"><button class="btn btn-secondary" onclick="Basta.setup()">🔁 ${_ih('visor.otraPartida')}</button><button class="btn btn-primary" onclick="UI.cerrarModal()">${_ih('visor.cerrar')}</button></div>`);
    }
};

/* =============================================
   CHARADAS: mímica, dibujo, sonidos o descripción, por equipos
   ============================================= */
const Charadas = {
    MODOS: {
        mimica: { ico: '🙌' }, dibujo: { ico: '✏️' }, sonido: { ico: '🔉' }, describe: { ico: '💬' }
    },
    nombreModo(k) { return _i('visor.char.' + k); }, descModo(k) { return _i('visor.char.' + k + '.d'); },
    detener() { clearInterval(this.reloj); this.reloj = null; this.quitarSensores(); Pantalla.liberar(); document.removeEventListener('keydown', this._teclas); try { document.fullscreenElement && document.exitFullscreen(); } catch (e) {} },
    setup() {
        const L = Datos.lengua;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🎭 ${_ih('visor.act.charadas')}</h2>
            <label class="t" for="c-l">${_ih('visor.char.lengua')}</label><select id="c-l" class="modern-select">${Comun.lenguas({ campo: 'img', min: 10 }, L)}</select>
            <label class="chip"><input type="checkbox" id="c-decir" checked> 🗣️ ${_ih('visor.char.decir')}</label>
            <label class="t" for="c-l2">${_ih('visor.char.significado')}</label>
            <select id="c-l2" class="modern-select"><option value="">— ${_ih('visor.ninguna')} —</option>${Comun.lenguas(null, '')}</select>
            <span class="t">${_ih('visor.char.formas')}</span>
            <div class="chips">${Object.entries(this.MODOS).map(([k, m], i) => `<label class="chip" title="${esc(this.descModo(k))}"><input type="checkbox" name="c-m" value="${k}" ${i < 2 ? 'checked' : ''}> ${m.ico} ${esc(this.nombreModo(k))}</label>`).join('')}</div>
            <p class="nota">${_ih('visor.char.ruleta')}</p>
            <span class="t">${_ih('visor.char.celular')}</span>
            <div class="chips" style="flex-direction:column;align-items:stretch">
                <label class="chip"><input type="radio" name="c-c" value="presenta" checked> 🧑‍🏫 ${_ih('visor.char.presenta')}</label>
                <label class="chip"><input type="radio" name="c-c" value="frente"> 🤳 ${_ih('visor.char.frente')}</label>
                <label class="chip"><input type="radio" name="c-c" value="tv"> 📺 ${_ih('visor.char.tv')}</label>
            </div>
            <div id="c-incl" class="caja oculto" style="margin-top:8px">
                <label class="chip"><input type="checkbox" id="c-tilt"> 📐 ${_ih('visor.char.tilt')}</label>
                <p class="nota">${_ih('visor.char.tilt.d')}</p>
            </div>
            <span class="t">${_ih('visor.equipos')}</span>
            <div class="chips">${[2, 3, 4, 5, 6].map((n, i) => `<label class="chip"><input type="radio" name="nj" value="${n}" ${i === 0 ? 'checked' : ''}> ${n}</label>`).join('')}</div>
            <div id="nombres" class="fila" style="margin-top:8px"></div>
            <span class="t">${_ih('visor.char.tiempo')}</span><div class="chips">${[30, 60, 90].map((n, i) => `<label class="chip"><input type="radio" name="c-t" value="${n}" ${i === 1 ? 'checked' : ''}> ${n} s</label>`).join('')}</div>
            <span class="t">${_ih('visor.char.turnos')}</span><div class="chips">${[1, 2, 3, 5].map((n, i) => `<label class="chip"><input type="radio" name="c-r" value="${n}" ${i === 1 ? 'checked' : ''}> ${n}</label>`).join('')}</div>
            <label class="chip"><input type="checkbox" id="c-audio" checked> 🔊 ${_ih('visor.char.audio')}</label>
            <details style="margin-top:8px"><summary class="t" style="cursor:pointer">📁 ${_ih('visor.elegirCategorias')}</summary><div id="c-cats"></div></details>
            <div id="c-msg"></div>
            <div class="acciones"><button class="btn btn-primary" onclick="Charadas.iniciar()">🚀 ${_ih('visor.empezar')}</button></div></div>`);
        const pintarCats = () => { $('c-cats').innerHTML = Comun.categorias($('c-l').value, 'c-cat'); };
        $('c-l').addEventListener('change', pintarCats); pintarCats();
        document.querySelectorAll('input[name="c-c"]').forEach(r => r.addEventListener('change', () => $('c-incl').classList.toggle('oculto', r.value !== 'frente' || !r.checked)));
        Comun.enlazarJugadores(_i('visor.equipo'));
    },
    async iniciar() {
        const L = $('c-l').value, modos = Comun.marcados('c-m'), cats = Comun.marcados('c-cat');
        if (!modos.length) return $('c-msg').innerHTML = '<div class="error-box">' + _ih('visor.char.error.formas') + '</div>';
        const pool = Datos.paraJuegos(L, { cats });
        if (pool.length < 10) return $('c-msg').innerHTML = `<div class="error-box">${_ih('visor.char.error.pocas', { n: pool.length })}</div>`;
        this.cfg = { L, L2: $('c-l2').value === L ? '' : $('c-l2').value, decir: $('c-decir').checked, modos, control: document.querySelector('input[name="c-c"]:checked').value,
            tilt: $('c-tilt').checked, tiempo: +document.querySelector('input[name="c-t"]:checked').value, turnos: +document.querySelector('input[name="c-r"]:checked').value, audio: $('c-audio').checked };
        // Primero las palabras con imagen: ayudan a quien no lee
        this.mazo = [...barajar(pool.filter(w => w.img)), ...barajar(pool.filter(w => !w.img))];
        this.equipos = Comun.leerJugadores(_i('visor.equipo'));
        this.turno = 0; this.turnosHechos = 0; this.jugadas = [];
        if (this.cfg.control === 'frente' && this.cfg.tilt) await this.pedirSensores();
        Juego.activo = this;
        this.preparar();
    },
    preparar() {
        if (this.turnosHechos >= this.equipos.length * this.cfg.turnos) return this.final();
        const e = this.equipos[this.turno];
        const instr = _i('visor.char.instr.' + this.cfg.control);
        this.detenerReloj();
        UI.abrir(`<div style="text-align:center">
            <div class="nota">${_ih('visor.char.turnoN', { n: Math.floor(this.turnosHechos / this.equipos.length) + 1, total: this.cfg.turnos })}</div>
            <h2 style="color:${e.color};font-size:2em;margin:10px 0">${esc(e.nombre)}</h2>
            <div class="score-board">${this.equipos.map(x => `<span class="player-pill" style="border-color:${x.color}">${esc(x.nombre)} · ⭐ ${x.puntos}</span>`).join('')}</div>
            <p class="texto-largo">${esc(instr)}</p>
            ${this.cfg.decir ? `<p class="nota">🗣️ ${_ih('visor.char.dicen', { lengua: Datos.nombreLengua(this.cfg.L) })}</p>` : ''}
            <button class="btn btn-primary btn-block" style="min-height:80px;font-size:1.5em" id="c-go">▶️ ${_ih('visor.char.listos')}</button>
            <div class="acciones"><button class="btn btn-secondary" onclick="Charadas.final()">${_ih('visor.terminarJuego')}</button></div></div>`, { completa: this.cfg.control === 'tv' });
        $('c-go').addEventListener('click', () => this.cuentaAtras());
    },
    cuentaAtras() {
        if (this.cfg.control === 'frente') { try { document.documentElement.requestFullscreen?.().then(() => screen.orientation?.lock?.('landscape').catch(() => {})).catch(() => {}); } catch (e) {} }
        Pantalla.mantener();
        let n = 3;
        const pinta = () => UI.cuerpo(`<div class="basta-letra" style="font-size:8rem;margin-top:20vh" aria-live="assertive">${n}</div>`);
        pinta(); SoundFX.ok();
        const t = setInterval(() => { n--; if (n > 0) { pinta(); SoundFX.ok(); } else { clearInterval(t); this.empezarTurno(); } }, 800);
    },
    empezarTurno() {
        this.resto = this.cfg.tiempo; this.aciertos = []; this.pasadas = [];
        this.sacar();
        this.activarSensores();
        this.reloj = setInterval(() => {
            this.resto--;
            document.querySelectorAll('.c-reloj').forEach(r => { r.textContent = fmtTiempo(Math.max(0, this.resto)); r.classList.toggle('poco', this.resto <= 10); });
            if (this.resto <= 5 && this.resto > 0) SoundFX.ok();
            if (this.resto <= 0) { SoundFX.fail(); SoundFX.vibrar([300, 100, 300]); this.finTurno(); }
        }, 1000);
        this._teclas = e => {
            if (!this.actual) return;
            if (e.key === 'ArrowRight' || e.key === 'Enter') { e.preventDefault(); this.marcar(true); }
            else if (e.key === 'ArrowLeft' || e.key === 'Backspace') { e.preventDefault(); this.marcar(false); }
            else if (e.key === ' ') { e.preventDefault(); this.ver(e.type === 'keydown'); }
        };
        document.addEventListener('keydown', this._teclas);
        if (this.cfg.control === 'tv') { this._suelta = e => { if (e.key === ' ') this.ver(false); }; document.addEventListener('keyup', this._suelta); }
    },
    detenerReloj() { clearInterval(this.reloj); this.reloj = null; document.removeEventListener('keydown', this._teclas); if (this._suelta) document.removeEventListener('keyup', this._suelta); },
    sacar() {
        if (!this.mazo.length) this.mazo = barajar(Datos.paraJuegos(this.cfg.L, {}));
        this.actual = this.mazo.shift();
        this.modo = elegir(this.cfg.modos);
        this.pintarTarjeta();
    },
    pintarTarjeta(flash = '') {
        const w = this.actual, L = this.cfg.L, m = this.MODOS[this.modo], e = this.equipos[this.turno];
        const sig = this.cfg.L2 && w.t[this.cfg.L2] ? `<div class="nota" style="font-size:1em">${esc(Datos.nombreLengua(this.cfg.L2))}: <b${Texto.lang(this.cfg.L2)}>${esc(w.t[this.cfg.L2].texto)}</b></div>` : '';
        const img = w.img ? `<img src="${esc(Datos.img(w, 700))}" alt="">` : '';
        const palabra = `<div class="gran-palabra"${Texto.lang(L)}>${esc(w.t[L].texto)}</div>`;
        const mn = esc(this.nombreModo(this.modo)), md = esc(this.descModo(this.modo));
        const cab = `<div class="cabeza"><span class="modo" title="${md}">${m.ico} ${mn}</span>
            <span class="temporizador c-reloj">${fmtTiempo(Math.max(0, this.resto))}</span><span style="color:${e.color};font-weight:800">${esc(e.nombre)} · ✓ ${this.aciertos.length}</span></div>`;
        if (this.cfg.control === 'frente') {
            UI.cuerpo(`<div class="frente ${flash}">
                <button class="lado pasar" id="c-pasa" aria-label="${esc(_i('visor.char.pasar'))}">⏭ ${esc(_i('visor.char.pasar').toLocaleUpperCase())}</button>
                <div class="centro"><div class="cabeza" style="display:flex;gap:14px"><span class="temporizador c-reloj" style="font-size:1.6em">${fmtTiempo(Math.max(0, this.resto))}</span><span class="modo">${m.ico} ${mn}</span></div>
                    ${img}<div class="w"${Texto.lang(L)}>${esc(w.t[L].texto)}</div>${sig}</div>
                <button class="lado acierto" id="c-ok" aria-label="${esc(_i('visor.char.acierto'))}">✓ ${esc(_i('visor.char.acierto').toLocaleUpperCase())}</button></div>`);
        } else if (this.cfg.control === 'tv') {
            UI.cuerpo(`<div class="charada tv ${flash}">${cab}
                <div class="tarjeta"><div class="${this.visto ? '' : 'oculta-palabra'}" id="c-oculta" style="display:flex;flex-direction:column;align-items:center;gap:10px">${img}${palabra}${sig}</div>
                    <p class="nota">${md}</p></div>
                <button class="btn btn-secondary btn-block" id="c-ver" style="min-height:60px">👁 ${_ih('visor.char.ver')}</button>
                <div class="zonas"><button class="zona pasar" id="c-pasa">⏭ ${_ih('visor.char.pasar')}<small>${_ih('visor.tecla', { t: '←' })}</small></button><button class="zona acierto" id="c-ok">✓ ${_ih('visor.char.acierto')}<small>${_ih('visor.tecla', { t: '→' })}</small></button></div></div>`);
            const v = $('c-ver');
            ['pointerdown'].forEach(ev => v.addEventListener(ev, () => this.ver(true)));
            ['pointerup', 'pointerleave', 'pointercancel'].forEach(ev => v.addEventListener(ev, () => this.ver(false)));
        } else {
            UI.cuerpo(`<div class="charada ${flash}">${cab}
                <div class="tarjeta">${img}${palabra}${sig}<p class="nota">${md}</p></div>
                <div class="zonas"><button class="zona pasar" id="c-pasa">⏭<span>${_ih('visor.char.pasar')}</span></button><button class="zona acierto" id="c-ok">✓<span>${_ih('visor.char.acierto')}</span></button></div></div>`);
        }
        $('c-ok').addEventListener('click', () => this.marcar(true));
        $('c-pasa').addEventListener('click', () => this.marcar(false));
    },
    ver(si) { this.visto = si; const o = $('c-oculta'); if (o) o.classList.toggle('oculta-palabra', !si); },
    marcar(ok) {
        if (!this.actual || this.resto <= 0) return;
        const w = this.actual;
        (ok ? this.aciertos : this.pasadas).push(w);
        this.jugadas.push({ id: w.id, ok });
        if (ok) { SoundFX.ok(); SoundFX.vibrar(80); if (this.cfg.audio && w.t[this.cfg.L].audio) Sonido.tocar(w.t[this.cfg.L].audio); }
        else { SoundFX.error(); SoundFX.vibrar([40, 40, 40]); }
        this.visto = false;
        this.actual = null;
        this._bloqueo = Date.now();
        setTimeout(() => { if (this.reloj) { this.sacar(); this.pintarTarjeta(ok ? 'flash-ok' : 'flash-pasa'); } }, 250);
    },
    finTurno() {
        this.detenerReloj(); this.quitarSensores(); Pantalla.liberar();
        try { document.fullscreenElement && document.exitFullscreen(); } catch (e) {}
        const e = this.equipos[this.turno], L = this.cfg.L;
        e.puntos += this.aciertos.length;
        this.actual = null;
        UI.abrir(`<h2 class="modal-titulo" style="color:${e.color}">⏰ ${_ih('visor.char.tiempoFin', { equipo: e.nombre, n: this.aciertos.length })}</h2>
            ${this.aciertos.length ? `<div>${this.aciertos.map(w => `<button class="chip-res ok" data-ficha="${esc(w.id)}"${Texto.lang(L)}>✓ ${esc(w.t[L].texto)}</button>`).join('')}</div>` : ''}
            ${this.pasadas.length ? `<div style="margin-top:6px">${this.pasadas.map(w => `<button class="chip-res fail" data-ficha="${esc(w.id)}"${Texto.lang(L)}>⏭ ${esc(w.t[L].texto)}</button>`).join('')}</div>` : ''}
            <p class="nota">${_ih('visor.char.corrige')}</p>
            <div class="fila" style="justify-content:center"><button class="btn btn-secondary" id="c-menos">−1</button><b style="font-size:1.4em" id="c-pts">${e.puntos}</b><button class="btn btn-secondary" id="c-mas">+1</button></div>
            <div class="acciones"><button class="btn btn-primary" id="c-sig">➡️ ${_ih('visor.char.siguienteEquipo')}</button></div>`);
        const p = $('c-pts');
        $('c-menos').addEventListener('click', () => { e.puntos = Math.max(0, e.puntos - 1); p.textContent = e.puntos; });
        $('c-mas').addEventListener('click', () => { e.puntos++; p.textContent = e.puntos; });
        $('c-sig').addEventListener('click', () => { this.turnosHechos++; this.turno = (this.turno + 1) % this.equipos.length; this.preparar(); });
    },
    final() {
        this.detenerReloj(); this.detener(); Juego.activo = null;
        const orden = [...this.equipos].sort((a, b) => b.puntos - a.puntos), med = ['🥇', '🥈', '🥉'];
        const vistos = new Set(), lista = this.jugadas.filter(x => !vistos.has(x.id) && vistos.add(x.id));
        UI.abrir(`<h2 class="modal-titulo">🎭 ${_ih('visor.char.fin')}</h2>
            ${orden.map((x, i) => `<div class="ranking-row" style="border-left:5px solid ${x.color}"><span style="font-size:1.5em">${med[i] || (i + 1) + '.'}</span><b style="flex:1">${esc(x.nombre)}</b><span>⭐ ${x.puntos}</span></div>`).join('')}
            ${Comun.palabrasJugadas(lista, this.cfg.L)}
            <div class="acciones"><button class="btn btn-secondary" onclick="Charadas.setup()">🔁 ${_ih('visor.otraVez')}</button><button class="btn btn-primary" onclick="UI.cerrarModal()">${_ih('visor.cerrar')}</button></div>`);
    },

    /* --- Inclinación (opcional): usa la gravedad perpendicular a la pantalla --- */
    async pedirSensores() {
        try { if (typeof DeviceMotionEvent !== 'undefined' && DeviceMotionEvent.requestPermission) { const r = await DeviceMotionEvent.requestPermission(); if (r !== 'granted') { this.cfg.tilt = false; UI.toast(_i('visor.char.sinSensor')); return; } } } catch (e) { this.cfg.tilt = false; return; }
    },
    activarSensores() {
        if (!this.cfg.tilt || this.cfg.control !== 'frente') return;
        this.quitarSensores();
        this._armado = true; this._base = null;
        this._mov = e => {
            if (!this.actual || this.cfg.control !== 'frente') return;
            const z = e.accelerationIncludingGravity?.z; if (z == null) return;
            if (this._base == null) this._base = z;     // calibración: la posición en la frente
            const d = z - this._base;
            if (this._armado && Date.now() - (this._bloqueo || 0) > 700) {
                if (d < -6) { this._armado = false; this.marcar(true); }       // pantalla hacia el suelo: acierto
                else if (d > 6) { this._armado = false; this.marcar(false); }  // pantalla hacia el techo: pasar
            } else if (Math.abs(d) < 3) this._armado = true;
        };
        window.addEventListener('devicemotion', this._mov);
    },
    quitarSensores() { if (this._mov) window.removeEventListener('devicemotion', this._mov); this._mov = null; }
};
</script>
