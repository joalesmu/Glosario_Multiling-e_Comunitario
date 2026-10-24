<script>
/* =============================================
   PIEZAS COMUNES DE LOS JUEGOS
   ============================================= */
const Comun = {
    /** Opciones de lengua: solo las que tienen suficientes palabras para la actividad. */
    lenguas(req, sel = Datos.lengua) {
        const ok = Datos.langs.filter(l => req ? (Datos.conteo(l.id)[req.campo] || 0) >= req.min : Datos.conteo(l.id).textos > 0);
        const actual = ok.some(l => l.id === sel) ? sel : ok[0]?.id;
        return ok.map(l => `<option value="${esc(l.id)}" ${l.id === actual ? 'selected' : ''}>${esc(l.nombre)}</option>`).join('');
    },
    categorias(L, nombre = 'cat', marcadas = []) {
        const usadas = new Set(Datos.paraJuegos(L).map(w => w.cat));
        const cats = [...Datos.cats.values()].filter(c => usadas.has(c.id));
        return cats.length ? `<div class="chips">${cats.map(c => `<label class="chip"><input type="checkbox" name="${nombre}" value="${esc(c.id)}" ${marcadas.includes(c.id) ? 'checked' : ''}> ${esc(Datos.categoria(c.id, L))}</label>`).join('')}</div>` : '<p class="nota">' + _ih('visor.sinCategorias') + '</p>';
    },
    marcados(nombre) { return [...document.querySelectorAll(`input[name="${nombre}"]:checked`)].map(x => x.value); },
    jugadores(max = 4, etiqueta = _i('visor.jugadores')) {
        return `<span class="t">${esc(etiqueta)}</span>
            <div class="chips" id="nj">${Array.from({ length: max }, (_, i) => `<label class="chip"><input type="radio" name="nj" value="${i + 1}" ${i === 0 ? 'checked' : ''}> ${i + 1}</label>`).join('')}</div>
            <div id="nombres" class="fila" style="margin-top:8px"></div>`;
    },
    enlazarJugadores(prefijo = _i('visor.jugador')) {
        const pintar = () => {
            const n = +document.querySelector('input[name="nj"]:checked').value;
            $('nombres').innerHTML = n > 1 ? Array.from({ length: n }, (_, i) => `<input class="input-box" style="flex:1;min-width:120px" id="nom${i}" placeholder="${esc(prefijo)} ${i + 1}" aria-label="${esc(_i('visor.nombreN', { n: i + 1 }))}">`).join('') : '';
        };
        document.querySelectorAll('input[name="nj"]').forEach(r => r.addEventListener('change', pintar));
        pintar();
    },
    leerJugadores(prefijo = _i('visor.jugador')) {
        const n = +(document.querySelector('input[name="nj"]:checked')?.value || 1);
        const colores = ['#2563eb', '#e11d48', '#16a34a', '#d97706', '#7c3aed', '#0891b2'];
        return Array.from({ length: n }, (_, i) => ({ nombre: n === 1 ? _i('visor.tu') : ($('nom' + i)?.value.trim() || `${prefijo} ${i + 1}`), puntos: 0, tiempo: 0, color: colores[i % colores.length] }));
    },
    /** Tarjeta de aprendizaje: imagen, la palabra en todas sus lenguas y el audio. */
    tarjeta(w, L) {
        if (!w) return '';
        const orden = [L, ...Datos.langs.map(l => l.id).filter(x => x !== L)].filter(x => w.t[x]);
        return `<div class="caja">
            ${w.img ? `<img src="${esc(Datos.img(w, 600))}" alt="" style="width:100%;max-height:180px;object-fit:contain;border-radius:10px;background:#fff">` : ''}
            ${orden.map(x => `<div class="trad" style="margin:6px 0 0"><div class="l">${esc(Datos.nombreLengua(x))}</div><div class="x"><span${Texto.lang(x)}>${esc(w.t[x].texto)}</span>${Sonido.boton(w, x)}</div></div>`).join('')}
        </div>`;
    },
    /** Lista final de palabras jugadas: tocar una abre su ficha completa. */
    palabrasJugadas(lista, L) {
        if (!lista.length) return '';
        return `<h3 style="font-size:1em;margin:14px 0 6px">📚 ${_ih('visor.palabrasJuego')}</h3><div>${lista.map(x =>
            `<button class="chip-res ${x.ok === true ? 'ok' : x.ok === false ? 'fail' : ''}" data-ficha="${esc(x.id)}">${x.ok === true ? '✓' : x.ok === false ? '✗' : '·'} <span${Texto.lang(L)}>${esc(Datos.byId.get(x.id) ? Datos.byId.get(x.id).t[L]?.texto || Datos.texto(Datos.byId.get(x.id)) : '')}</span></button>`).join('')}</div>`;
    }
};
document.addEventListener('click', e => { const b = e.target.closest('[data-ficha]'); if (b) { Juego.detener(); Ficha.abrir(b.dataset.ficha); } });

/** Motor de turnos, tiempo y puntos para juegos de 1 a 4 jugadores en un mismo dispositivo. */
class Motor {
    constructor() { this.jugadores = []; this.turno = 0; this.reloj = null; this.registro = []; }
    detener() { clearInterval(this.reloj); this.reloj = null; }
    empezar(jugadores) { this.jugadores = jugadores; this.turno = 0; this.registro = []; Juego.activo = this; this.correr(); }
    correr() {
        clearInterval(this.reloj);
        this.reloj = setInterval(() => {
            const p = this.jugadores[this.turno]; if (!p) return;
            p.tiempo++;
            const el = document.querySelector(`.player-pill[data-i="${this.turno}"] .p-time`);
            if (el) el.textContent = fmtTiempo(p.tiempo);
        }, 1000);
    }
    siguiente() { this.turno = (this.turno + 1) % this.jugadores.length; }
    marcador() {
        if (this.jugadores.length < 2) {
            const p = this.jugadores[0];
            return p ? `<div class="score-board"><span class="player-pill" style="border-color:${p.color}">⭐ ${p.puntos} · <span class="p-time" data-i="0">${fmtTiempo(p.tiempo)}</span></span></div>` : '';
        }
        return `<div class="score-board">${this.jugadores.map((p, i) => `<div class="player-pill ${i === this.turno ? 'activo' : ''}" data-i="${i}"
            style="background:${i === this.turno ? p.color : 'transparent'};color:${i === this.turno ? '#fff' : 'var(--text-main)'};border-color:${p.color}">
            <span>${esc(p.nombre)}</span><span>⭐ ${p.puntos}</span><span class="p-time">${fmtTiempo(p.tiempo)}</span></div>`).join('')}</div>`;
    }
    anotar(id, ok) { this.registro.push({ id, ok, jugador: this.jugadores[this.turno]?.nombre }); }
    final(titulo, L) {
        this.detener();
        const orden = [...this.jugadores].sort((a, b) => b.puntos - a.puntos || a.tiempo - b.tiempo);
        const medallas = ['🥇', '🥈', '🥉'];
        const vistos = new Set(), lista = this.registro.filter(x => !vistos.has(x.id) && vistos.add(x.id));
        UI.cuerpo(`<h2 class="modal-titulo">${esc(titulo)}</h2>
            ${orden.map((p, i) => `<div class="ranking-row" style="border-left:5px solid ${p.color}"><span style="font-size:1.5em">${medallas[i] || i + 1 + '.'}</span>
                <b style="flex:1">${esc(p.nombre)}</b><span>⭐ ${p.puntos} · ⏱ ${fmtTiempo(p.tiempo)}</span></div>`).join('')}
            ${Comun.palabrasJugadas(lista, L)}
            <div class="acciones"><button class="btn btn-primary" onclick="UI.cerrarModal()">${_ih('visor.cerrar')}</button></div>`);
        Juego.activo = null;
    }
}

/* =============================================
   MEMORAMA
   ============================================= */
const Memorama = new class extends Motor {
    setup() {
        const tipos = sel => ['img', 'txt', 'audio'].map(t => `<option value="${t}" ${t === sel ? 'selected' : ''}>${{ img: '🖼️ ', txt: '📝 ', audio: '🔊 ' }[t]}${_ih('visor.mem.cara.' + t)}</option>`).join('');
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🧠 ${_ih('visor.act.memorama')}</h2>
            <p class="nota">${_ih('visor.mem.d')}</p>
            <span class="t">${_ih('visor.mem.caraA')}</span><div class="fila"><select id="m-ta" class="modern-select" style="flex:1">${tipos('img')}</select><select id="m-la" class="modern-select" style="flex:1">${Comun.lenguas(null)}</select></div>
            <span class="t">${_ih('visor.mem.caraB')}</span><div class="fila"><select id="m-tb" class="modern-select" style="flex:1">${tipos('txt')}</select><select id="m-lb" class="modern-select" style="flex:1">${Comun.lenguas(null)}</select></div>
            <span class="t">${_ih('visor.mem.pares')}</span><div class="chips">${[6, 8, 10, 12].map((n, i) => `<label class="chip"><input type="radio" name="m-p" value="${n}" ${i === 1 ? 'checked' : ''}> ${n}</label>`).join('')}</div>
            ${Comun.jugadores()}
            <p class="nota">${_ih('visor.mem.turnos')}</p>
            <div id="m-msg"></div>
            <div class="acciones"><button class="btn btn-primary" onclick="Memorama.iniciar()">🚀 ${_ih('visor.jugarBoton')}</button></div></div>`);
        Comun.enlazarJugadores();
    }
    cara(w, tipo, L) {
        if (tipo === 'img') return w.img ? `<img src="${esc(Datos.img(w, 300))}" alt="">` : '';
        const t = w.t[L]; if (!t) return '';
        if (tipo === 'audio') return t.audio ? `<span style="font-size:2.2em" aria-hidden="true">🔊</span><span class="nota">${_ih('visor.escucha')}</span>` : '';
        return `<span${Texto.lang(L)}>${esc(t.texto)}</span>`;
    }
    iniciar() {
        const ta = $('m-ta').value, la = $('m-la').value, tb = $('m-tb').value, lb = $('m-lb').value;
        const pares = +document.querySelector('input[name="m-p"]:checked').value;
        if (ta === tb && (ta === 'img' || la === lb)) return $('m-msg').innerHTML = '<div class="error-box">' + _ih('visor.mem.iguales') + '</div>';
        const pool = Datos.words.filter(w => w.juegos && this.cara(w, ta, la) && this.cara(w, tb, lb));
        if (pool.length < 4) return $('m-msg').innerHTML = `<div class="error-box">${_ih('visor.mem.pocas', { n: pool.length })}</div>`;
        const elegidas = barajar(pool).slice(0, pares);
        this.L = tb === 'txt' || tb === 'audio' ? lb : la;
        this.cartas = barajar(elegidas.flatMap(w => [
            { id: w.id, html: this.cara(w, ta, la), audio: ta === 'audio' ? w.t[la].audio : '', cap: w.t[la]?.texto || Datos.texto(w), L: la },
            { id: w.id, html: this.cara(w, tb, lb), audio: tb === 'audio' ? w.t[lb].audio : '', cap: w.t[lb]?.texto || Datos.texto(w), L: lb }
        ]));
        this.total = elegidas.length; this.hallados = 0; this.vueltas = []; this.bloqueo = false;
        this.empezar(Comun.leerJugadores());
        this.pintar();
    }
    pintar() {
        const cols = this.cartas.length <= 12 ? 3 : 4;
        UI.cuerpo(`${this.marcador()}
            <div class="nota" style="text-align:center">${_ih('visor.mem.contador', { n: this.hallados, total: this.total })}${this.jugadores.length > 1 ? ' · ' + _ih('visor.turnoDe', { nombre: this.jugadores[this.turno].nombre }) : ''}</div>
            <div class="memory-grid" style="grid-template-columns:repeat(${cols},1fr)" id="m-grid">${this.cartas.map((c, i) => `
                <button class="mem-card ${c.ok ? 'matched' : ''} ${this.vueltas.includes(i) ? 'flipped' : ''}" data-i="${i}" aria-label="${c.ok || this.vueltas.includes(i) ? esc(c.cap) : esc(_i('visor.cartaN', { n: i + 1 }))}">
                    <span class="mem-face mem-front" aria-hidden="true">?</span>
                    <span class="mem-face mem-back">${c.html}${c.ok ? `<span class="cap"${Texto.lang(c.L)}>${esc(c.cap)}</span>` : ''}</span></button>`).join('')}</div>
            <div class="acciones"><button class="btn btn-secondary" onclick="Memorama.final('🧠 ' + _i('visor.act.memorama'), Memorama.L)">${_ih('visor.terminar')}</button></div>`);
        $('m-grid').addEventListener('click', e => { const b = e.target.closest('.mem-card'); if (b) this.voltear(+b.dataset.i); });
    }
    voltear(i) {
        if (this.bloqueo || this.vueltas.includes(i) || this.cartas[i].ok) return;
        this.vueltas.push(i);
        if (this.cartas[i].audio) Sonido.tocar(this.cartas[i].audio);
        this.pintar();
        if (this.vueltas.length === 2) { this.bloqueo = true; setTimeout(() => this.revisar(), 1000); }
    }
    revisar() {
        const [a, b] = this.vueltas, A = this.cartas[a], B = this.cartas[b];
        if (A.id === B.id) {
            A.ok = B.ok = true; this.hallados++; this.jugadores[this.turno].puntos++; this.anotar(A.id, true); SoundFX.ok();
        } else { SoundFX.error(); this.siguiente(); }
        this.vueltas = []; this.bloqueo = false;
        if (this.hallados === this.total) { SoundFX.win(); return setTimeout(() => this.final('🧠 ' + _i('visor.mem.completado'), this.L), 500); }
        this.pintar();
    }
};

/* =============================================
   AHORCADO (con el alfabeto real de cada lengua)
   ============================================= */
const Ahorcado = new class extends Motor {
    setup() {
        const req = { campo: 'largas', min: 6 };
        UI.abrir(`<div class="form"><h2 class="modal-titulo">😵 ${_ih('visor.act.ahorcado')}</h2>
            <label class="t" for="h-l">${_ih('visor.lengua')}</label><select id="h-l" class="modern-select">${Comun.lenguas(req)}</select>
            <p class="nota">${_ih('visor.ah.teclado')}</p>
            <span class="t">${_ih('visor.ah.dificultad')}</span>
            <div class="chips">
                <label class="chip"><input type="radio" name="h-d" value="facil"> 🟢 ${_ih('visor.ah.facil')}</label>
                <label class="chip"><input type="radio" name="h-d" value="medio" checked> 🟡 ${_ih('visor.ah.medio')}</label>
                <label class="chip"><input type="radio" name="h-d" value="dificil"> 🔴 ${_ih('visor.ah.dificil')}</label>
            </div>
            <span class="t">${_ih('visor.rondas')}</span><div class="chips">${[3, 5, 10].map((n, i) => `<label class="chip"><input type="radio" name="h-r" value="${n}" ${i === 1 ? 'checked' : ''}> ${n}</label>`).join('')}</div>
            ${Comun.jugadores()}
            <div class="acciones"><button class="btn btn-primary" onclick="Ahorcado.iniciar()">🚀 ${_ih('visor.jugarBoton')}</button></div></div>`);
        Comun.enlazarJugadores();
    }
    iniciar() {
        this.L = $('h-l').value; this.dif = document.querySelector('input[name="h-d"]:checked').value;
        this.rondas = +document.querySelector('input[name="h-r"]:checked').value;
        this.pool = barajar(Datos.paraJuegos(this.L, { largas: true }));
        this.n = 0;
        this.empezar(Comun.leerJugadores());
        this.ronda();
    }
    puntosPor() { return { facil: [3, 3, 3, 3, 3], medio: [3, 2, 2, 1, 1], dificil: [3] }[this.dif]; }
    ronda() {
        if (this.n >= this.rondas || !this.pool.length) return this.final('🏆 ' + _i('visor.act.ahorcado'), this.L);
        this.n++;
        this.w = this.pool.pop();
        this.graf = Texto.grafemas(this.w.t[this.L].texto, this.L);
        this.letras = [...new Set(this.graf.filter(g => g.letra).map(g => g.g))];
        const alfa = Texto.alfabeto(this.L);
        this.teclas = [...new Set([...alfa, ...this.letras])];
        this.dichas = []; this.pistas = 0; this.cerrada = false;
        this.pintar();
    }
    pintar() {
        const errores = this.dichas.filter(l => !this.letras.includes(l)).length;
        const gano = this.letras.every(l => this.dichas.includes(l)), perdio = errores >= 6;
        const arte = ['', '😶', '😶\n |', '😶\n/|', '😶\n/|\\', '😵\n/|\\\n/', '😵\n/|\\\n/ \\'][Math.min(errores, 6)];
        const pal = this.graf.map(g => !g.letra ? `<span class="sep">${esc(g.g === ' ' ? '' : g.g)}</span>` : `<span>${this.dichas.includes(g.g) || gano || perdio ? esc(g.g) : '&nbsp;'}</span>`).join('');
        let cont;
        if ((gano || perdio) && !this.cerrada) {
            this.cerrada = true;
            const pts = gano ? (this.puntosPor()[this.pistas] ?? 1) : 0;
            this.jugadores[this.turno].puntos += pts;
            this.anotar(this.w.id, gano);
            gano ? SoundFX.win() : SoundFX.fail();
            if (this.w.t[this.L].audio) Sonido.tocar(this.w.t[this.L].audio);
            this.resultado = `<div class="round-result ${gano ? 'win' : 'fail'}">${gano ? '✅ ' + _ih('visor.correcto', { pts }) : '❌ ' + _ih('visor.ah.era')} <span${Texto.lang(this.L)}>${esc(this.w.t[this.L].texto)}</span></div>`;
            this.siguiente();
        }
        if (this.cerrada) {
            cont = `${this.resultado}${Comun.tarjeta(this.w, this.L)}
                <div class="acciones"><button class="btn btn-primary" onclick="Ahorcado.ronda()">${this.n >= this.rondas ? '📊 ' + _ih('visor.resultados') : (this.jugadores.length > 1 ? '➡️ ' + _ih('visor.turnoDe', { nombre: this.jugadores[this.turno].nombre }) : '➡️ ' + _ih('visor.siguientePalabra'))}</button></div>`;
        } else {
            const pistas = this.dif === 'dificil' ? '' : this.htmlPistas();
            cont = `<div class="nota" style="text-align:center">${_ih('visor.rondaN', { n: this.n, total: this.rondas })}${this.jugadores.length > 1 ? ' · ' + _ih('visor.turnoDe', { nombre: this.jugadores[this.turno].nombre }) : ''}</div>
                <div class="hangman-wrap" aria-hidden="true">┌──┐\n│  ${arte.split('\n')[0] || ''}\n│  ${arte.split('\n')[1] || ''}\n│  ${arte.split('\n')[2] || ''}\n┴────</div>
                <div class="word-display" aria-label="${esc(_i('visor.palabra'))}">${pal}</div>
                <div class="nota" style="text-align:center;color:var(--error)">${_ih('visor.ah.errores', { n: errores })}</div>
                ${pistas}
                <div class="keyboard" id="h-kb">${this.teclas.map(l => {
                    const d = this.dichas.includes(l);
                    return `<button class="key-btn ${d ? (this.letras.includes(l) ? 'correct' : 'wrong') : ''}" data-l="${esc(l)}" ${d ? 'disabled' : ''}${Texto.lang(this.L)}>${esc(l)}</button>`; }).join('')}</div>`;
        }
        UI.cuerpo(`${this.marcador()}${cont}<div class="acciones"><button class="btn btn-secondary" onclick="Ahorcado.final('🏆 ' + _i('visor.act.ahorcado'), Ahorcado.L)">${_ih('visor.terminar')}</button></div>`);
        const kb = $('h-kb'); if (kb) kb.addEventListener('click', e => { const b = e.target.closest('.key-btn'); if (b && !b.disabled) this.decir(b.dataset.l); });
        const bp = $('h-pista'); if (bp) bp.addEventListener('click', () => { this.pistas++; this.pintar(); });
    }
    htmlPistas() {
        const t = this.w.t[this.L], items = [];
        if (this.pistas >= 1) items.push(`📁 ${esc(Datos.categoria(this.w.cat, this.L))}`);
        if (this.pistas >= 2 && this.w.img) items.push(`<img src="${esc(Datos.img(this.w, 300))}" alt="${esc(_i('visor.pista'))}" style="height:100px;border-radius:8px;${this.pistas === 2 ? 'filter:blur(5px)' : ''}">`);
        if (this.pistas >= 3) items.push('🔤 ' + _ih('visor.ah.empieza', { letra: this.letras.length ? this.graf.find(g => g.letra).g : '' }));
        if (this.pistas >= 4 && t.audio) items.push(Sonido.boton(this.w, this.L, '🔊 ' + _ih('visor.escuchar')));
        const max = t.audio ? 4 : 3, pts = this.puntosPor();
        const boton = this.pistas < max ? `<button class="btn btn-secondary" id="h-pista">💡 ${_ih('visor.ah.pistaN', { n: this.pistas + 1 })}${this.dif === 'medio' ? ' · ' + _ih('visor.ah.valdra', { pts: pts[this.pistas + 1] ?? 1 }) : ''}</button>` : '';
        return `${items.length ? `<div class="hint-box">${items.join('<br>')}</div>` : ''}<div style="text-align:center">${boton}</div>`;
    }
    decir(l) {
        if (this.dichas.includes(l) || this.cerrada) return;
        this.dichas.push(l);
        this.letras.includes(l) ? SoundFX.ok() : SoundFX.error();
        this.pintar();
    }
};

/* =============================================
   SOPA DE LETRAS (cada casilla es una letra real de la lengua)
   ============================================= */
const Sopa = new class extends Motor {
    setup() {
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🧩 ${_ih('visor.act.sopa')}</h2>
            <label class="t" for="s-l">${_ih('visor.lengua')}</label><select id="s-l" class="modern-select">${Comun.lenguas({ campo: 'largas', min: 6 })}</select>
            <span class="t">${_ih('visor.sopa.tamano')}</span><div class="chips">${[[8, 'chica'], [10, 'mediana'], [12, 'grande']].map(([n, t], i) => `<label class="chip"><input type="radio" name="s-t" value="${n}" ${i === 1 ? 'checked' : ''}> ${_ih('visor.sopa.' + t)} (${n}×${n})</label>`).join('')}</div>
            <span class="t">${_ih('visor.sopa.nivel')}</span><div class="chips">${['facil', 'normal', 'dificil'].map(k => `<label class="chip" title="${_ih('visor.sopa.' + k + '.d')}"><input type="radio" name="s-n" value="${k}" ${k === (LS.get('sopa:nivel') || 'normal') ? 'checked' : ''}> ${_ih('visor.sopa.' + k)}</label>`).join('')}</div>
            ${Comun.jugadores()}
            <p class="nota">${_ih('visor.sopa.d')}</p>
            <div id="s-msg"></div>
            <div class="acciones"><button class="btn btn-primary" onclick="Sopa.iniciar()">🚀 ${_ih('visor.jugarBoton')}</button></div></div>`);
        Comun.enlazarJugadores();
    }
    celdaPx(n) {
        const ancho = Math.min(window.innerWidth * 0.96, 680) - 48;
        return Math.floor((ancho - (n - 1) * 2) / n);
    }
    iniciar() {
        this.L = $('s-l').value;
        let n = +document.querySelector('input[name="s-t"]:checked').value;
        while (n > 8 && this.celdaPx(n) < 26) n -= 2;   // en pantallas pequeñas se ajusta solo
        this.n = n;
        const cuantas = { 8: 6, 10: 8, 12: 10 }[n];
        const candidatas = barajar(Datos.paraJuegos(this.L, { largas: true }))
            .map(w => ({ w, g: Texto.letras(w.t[this.L].texto, this.L) })).filter(x => x.g.length >= 3 && x.g.length <= n);
        const vistos = new Set(), unicas = candidatas.filter(x => { const k = x.g.join(''); return !vistos.has(k) && vistos.add(k); });
        if (unicas.length < 4) return $('s-msg').innerHTML = '<div class="error-box">' + _ih('visor.sopa.pocas') + '</div>';
        const nivel = (document.querySelector('input[name="s-n"]:checked') || {}).value || 'normal';
        LS.set('sopa:nivel', nivel);
        const r = this.generar(unicas, n, cuantas, nivel);
        this.grid = r.grid;
        this.palabras = r.palabras.map((x, i) => ({ w: x.w, g: x.g, clave: x.g.join('|'), celdas: x.celdas, hallada: false,
            color: ['#2563eb', '#e11d48', '#16a34a', '#d97706', '#7c3aed', '#0891b2', '#be185d', '#065f46', '#9333ea', '#ca8a04'][i % 10] }));
        const alfa = Texto.alfabeto(this.L).filter(a => a.length <= 3);
        const relleno = alfa.length ? alfa : 'abcdefghijklmnopqrstuvwxyz'.split('');
        for (let y = 0; y < n; y++) for (let x = 0; x < n; x++) if (!this.grid[y][x]) this.grid[y][x] = elegir(relleno);
        this.encontradas = {}; this.sel = [];
        this.empezar(Comun.leerJugadores());
        this.pintar();
    }
    /**
     * Acomoda las palabras en la cuadrícula. Las palabras se cruzan cuando comparten una letra (nunca en la misma
     * línea: una no puede quedar escondida dentro de otra) y las direcciones se reparten para que no se adivinen
     * por los huecos. nivel: facil (→ ↓), normal (+ diagonales), dificil (también al revés).
     * Devuelve { grid, palabras: [{ w, g, celdas, dir }] } con las celdas vacías en ''.
     */
    generar(lista, n, cuantas, nivel = 'normal', rnd = Math.random) {
        const base = { facil: [[1, 0], [0, 1]], normal: [[1, 0], [0, 1], [1, 1], [1, -1]] }[nivel];
        const dirs = base || [[1, 0], [0, 1], [1, 1], [1, -1], [-1, 0], [0, -1], [-1, -1], [-1, 1]];
        const eje = ([dx, dy]) => (dx < 0 || (dx === 0 && dy < 0)) ? `${-dx},${-dy}` : `${dx},${dy}`;
        const grid = Array.from({ length: n }, () => Array(n).fill(''));
        const ejes = Array.from({ length: n }, () => Array.from({ length: n }, () => new Set()));
        const usos = new Map(dirs.map(d => [d.join(','), 0]));
        // Las elegidas, de la más larga a la más corta (las largas son las difíciles de acomodar);
        // si alguna no cabe, entra otra de reserva. El orden de la lista se conserva en el resultado.
        const elegidas = lista.slice(0, Math.max(cuantas * 2, cuantas + 6));
        const orden = elegidas.slice(0, cuantas).sort((a, b) => b.g.length - a.g.length).concat(elegidas.slice(cuantas));
        const puestas = [];
        for (const x of orden) {
            if (puestas.length >= cuantas) break;
            const len = x.g.length, opciones = [];
            const menos = Math.min(...usos.values());
            for (const d of dirs) {
                const [dx, dy] = d, e = eje(d);
                for (let y0 = 0; y0 < n; y0++) for (let x0 = 0; x0 < n; x0++) {
                    const x1 = x0 + dx * (len - 1), y1 = y0 + dy * (len - 1);
                    if (x1 < 0 || x1 >= n || y1 < 0 || y1 >= n) continue;
                    let cruces = 0, ok = true;
                    for (let k = 0; k < len && ok; k++) {
                        const cx = x0 + dx * k, cy = y0 + dy * k, c = grid[cy][cx];
                        if (!c) continue;
                        if (c !== x.g[k] || ejes[cy][cx].has(e)) ok = false; else cruces++;
                    }
                    if (!ok || cruces === len) continue;
                    // Puntaje: cruzarse suma, repartir direcciones suma, y un poco de azar para que cada sopa sea distinta
                    const pts = (cruces ? 3 + Math.min(cruces, 2) : 0) + (usos.get(d.join(',')) === menos ? 2 : 0) + rnd() * 3;
                    opciones.push({ d, x0, y0, pts });
                }
            }
            if (!opciones.length) continue;
            const m = opciones.reduce((a, b) => b.pts > a.pts ? b : a);
            const celdas = [];
            for (let k = 0; k < len; k++) {
                const cx = m.x0 + m.d[0] * k, cy = m.y0 + m.d[1] * k;
                grid[cy][cx] = x.g[k]; ejes[cy][cx].add(eje(m.d)); celdas.push(`${cx},${cy}`);
            }
            usos.set(m.d.join(','), usos.get(m.d.join(',')) + 1);
            puestas.push({ w: x.w, g: x.g, celdas, dir: m.d, i: elegidas.indexOf(x) });
        }
        puestas.sort((a, b) => a.i - b.i);
        return { grid, palabras: puestas };
    }
    pintar() {
        const n = this.n, cs = Math.max(22, Math.min(46, this.celdaPx(n)));
        const max = Math.max(...this.grid.flat().map(g => g.length));
        const fs = Math.max(10, Math.floor(cs * (max > 1 ? 0.42 : 0.55)));
        UI.cuerpo(`${this.marcador()}
            <div class="nota" style="text-align:center">${_ih('visor.sopa.encontradas', { n: this.palabras.filter(p => p.hallada).length, total: this.palabras.length })}${this.jugadores.length > 1 ? ' · ' + _ih('visor.turnoDe', { nombre: this.jugadores[this.turno].nombre }) : ''}</div>
            <div class="ws-wrapper"><div class="ws-grid" id="s-grid" style="grid-template-columns:repeat(${n},${cs}px)"${Texto.lang(this.L)}>
                ${this.grid.map((fila, y) => fila.map((g, x) => { const c = this.encontradas[`${x},${y}`];
                    return `<div class="ws-cell" data-x="${x}" data-y="${y}" style="width:${cs}px;height:${cs}px;font-size:${fs}px;${c ? `background:${c};color:#fff;border-color:${c}` : ''}">${esc(g)}</div>`; }).join('')).join('')}
            </div></div>
            <div class="word-list">${this.palabras.map(p => `<span class="word-chip" style="${p.hallada ? `background:${p.color};color:#fff;border-color:${p.color};text-decoration:line-through` : ''}"${Texto.lang(this.L)}>${esc(p.w.t[this.L].texto)}</span>`).join('')}</div>
            <div class="acciones"><button class="btn btn-secondary" onclick="Sopa.final('🧩 ' + _i('visor.act.sopa'), Sopa.L)">${_ih('visor.terminar')}</button></div>`);
        this.enlazar();
    }
    enlazar() {
        const grid = $('s-grid'); if (!grid) return;
        let ini = null;
        const celda = e => { const el = document.elementFromPoint(e.clientX, e.clientY); return el && el.dataset && el.dataset.x !== undefined ? { x: +el.dataset.x, y: +el.dataset.y } : null; };
        const marcar = () => grid.querySelectorAll('.ws-cell').forEach(el => el.classList.toggle('selected', this.sel.some(p => p.x === +el.dataset.x && p.y === +el.dataset.y)));
        grid.addEventListener('pointerdown', e => { const c = celda(e); if (!c) return; grid.setPointerCapture(e.pointerId); ini = c; this.sel = [c]; marcar(); });
        grid.addEventListener('pointermove', e => {
            if (!ini) return; const c = celda(e); if (!c) return;
            const dx = c.x - ini.x, dy = c.y - ini.y;
            if (dx && dy && Math.abs(dx) !== Math.abs(dy)) return;
            const pasos = Math.max(Math.abs(dx), Math.abs(dy)), sx = Math.sign(dx), sy = Math.sign(dy);
            this.sel = Array.from({ length: pasos + 1 }, (_, i) => ({ x: ini.x + sx * i, y: ini.y + sy * i }));
            marcar();
        });
        const fin = () => { if (!ini) return; ini = null; this.revisar(); };
        grid.addEventListener('pointerup', fin); grid.addEventListener('pointercancel', fin);
    }
    revisar() {
        if (this.sel.length < 2) { this.sel = []; return; }
        const g = this.sel.map(p => this.grid[p.y][p.x]);
        const ida = g.join('|'), vuelta = [...g].reverse().join('|');
        const p = this.palabras.find(x => !x.hallada && (x.clave === ida || x.clave === vuelta));
        if (p) {
            p.hallada = true; this.sel.forEach(c => { this.encontradas[`${c.x},${c.y}`] = p.color; });
            this.jugadores[this.turno].puntos++; this.anotar(p.w.id, true); SoundFX.ok();
            if (p.w.t[this.L].audio) Sonido.tocar(p.w.t[this.L].audio);
            UI.toast(`✓ ${p.w.t[this.L].texto}`);
            if (this.palabras.every(x => x.hallada)) { SoundFX.win(); this.sel = []; return setTimeout(() => this.final('🧩 ' + _i('visor.sopa.completada'), this.L), 600); }
            this.siguiente();
        } else if (g.length > 2) { SoundFX.error(); this.siguiente(); }
        this.sel = [];
        this.pintar();
    }
    final(t, L) { this.palabras.filter(p => !p.hallada).forEach(p => this.anotar(p.w.id, false)); super.final(t, L); }
};
</script>
