<script>
/* =============================================
   PROPONER AL CATÁLOGO (docentes, hablantes, familias)
   Todo lo que se envía llega como pendiente al equipo, que decide. Sin cuentas:
   cada envío recibe un código para ver su estado en «Mis propuestas».
   Sin internet, lo enviado espera en el dispositivo y sale solo al volver la conexión.
   El correo del catálogo siempre aparece como alternativa.
   ============================================= */
const Propuestas = {
    conf() { const p = Estado.e && Estado.e.propuestas; return p && p.activo ? p : null; },
    activo() { return !!this.conf(); },
    dispositivo() { let d = LS.get('dispositivo', ''); if (!d) { d = 'd' + Date.now().toString(36) + azarSeguro(1e9).toString(36); LS.set('dispositivo', d); } return d; },
    quienGuardado() { return LS.get('prop:quien:' + Datos.base, {}) || {}; },
    colab() { return LS.get('prop:colab:' + Datos.base, '') || ''; },
    mias() { return LS.get('prop:mias:' + Datos.base, []) || []; },
    /* ---------- cola en el dispositivo (IndexedDB: las fotos no caben en el almacenamiento simple) ---------- */
    async cola() {
        const d = await Offline.abrirDB(); if (!d) return [];
        const pref = Datos.base + '|cola|', out = [];
        await new Promise(res => { try { const t = d.transaction('m', 'readonly'), c = t.objectStore('m').openCursor();
            c.onsuccess = () => { const cur = c.result; if (!cur) return; if (String(cur.key).startsWith(pref)) out.push(cur.value); cur.continue(); }; t.oncomplete = res; t.onerror = res; } catch (e) { res(); } });
        return out.sort((a, b) => String(a.creado).localeCompare(String(b.creado)));
    },
    async encolar(envio) { envio.id = envio.id || 'e' + Date.now().toString(36) + azarSeguro(1e6).toString(36); envio.creado = envio.creado || new Date().toISOString(); await Offline.tx('readwrite', s => s.put(envio, Datos.base + '|cola|' + envio.id)); return envio; },
    async quitar(id) { await Offline.tx('readwrite', s => s.delete(Datos.base + '|cola|' + id)); },
    /** Envía lo que espera en la cola. Devuelve el resultado del último envío (para mostrarlo). */
    async enviarCola() {
        if (this._enviando || !Datos.base || !navigator.onLine) return null;
        this._enviando = true;
        let ultimo = null;
        try {
            for (const e of await this.cola()) {
                let r;
                try { r = await Datos.post({ accion: 'proponer', dispositivo: this.dispositivo(), colab: this.colab(), web: e.web || '', quien: e.quien, items: e.items }); }
                catch (err) { e.error = 'red'; await this.encolar(e); ultimo = { espera: true }; break; }
                if (r.status === 200) {
                    if (r.codigo) { const m = this.mias(); m.unshift({ codigo: r.codigo, fecha: new Date().toISOString().slice(0, 10), resumen: e.resumen, n: e.items.length }); LS.set('prop:mias:' + Datos.base, m.slice(0, 200)); }
                    await this.quitar(e.id);
                    ultimo = { codigo: r.codigo, resultados: r.resultados || [], items: e.items };
                } else if (r.status === 503 || r.status === 429) { e.error = r.status === 503 ? 'mant' : 'limite'; await this.encolar(e); ultimo = { espera: true, error: e.error }; break; }
                else { e.error = r.error || String(r.status); await this.encolar(e); ultimo = { error: e.error }; }
            }
        } finally { this._enviando = false; }
        return ultimo;
    },
    /* ---------- partes comunes de los formularios ---------- */
    /** Quién propone, consentimiento, criterios del equipo y la alternativa por correo. */
    comun(asunto) {
        const q = this.quienGuardado(), c = this.conf() || {};
        return `<fieldset class="grupo"><legend>🙋 ${_ih('visor.prop.quien')}</legend>
            <label class="t" for="pq-n">${_ih('visor.prop.nombre')}</label><input id="pq-n" class="input-box" maxlength="120" value="${esc(q.nombre || '')}" placeholder="${esc(_i('visor.prop.nombre.ph'))}" autocomplete="name">
            <label class="chip" style="margin-top:6px"><input type="checkbox" id="pq-pub" ${q.publico !== false ? 'checked' : ''}> ${_ih('visor.prop.publico')}</label>
            <label class="t" for="pq-rol">${_ih('visor.prop.rol')}</label><select id="pq-rol" class="modern-select">${['docente', 'hablante', 'familia', 'otro'].map(r => `<option value="${r}" ${q.rol === r ? 'selected' : ''}>${_ih('visor.prop.rol.' + r)}</option>`).join('')}</select>
            <details><summary class="t" style="cursor:pointer">${_ih('visor.prop.mas')}</summary>
                <label class="t" for="pq-rec">${_ih('visor.prop.recopilo')}</label><input id="pq-rec" class="input-box" maxlength="120" value="${esc(q.recopilo || '')}" placeholder="${esc(_i('visor.prop.recopilo.ph'))}">
                <label class="t" for="pq-con">${_ih('visor.prop.contacto')}</label><input id="pq-con" class="input-box" maxlength="160" value="${esc(q.contacto || '')}" placeholder="${esc(_i('visor.prop.contacto.ph'))}"></details>
            ${c.quien === 'clave' ? `<label class="t" for="pq-col">🔑 ${_ih('visor.prop.clave')}</label><input id="pq-col" class="input-box" value="${esc(this.colab())}" autocomplete="off">` : ''}
            <input id="pq-web" class="trampa" tabindex="-1" autocomplete="off" aria-hidden="true">
            <label class="chip" style="margin-top:8px"><input type="checkbox" id="pq-ok"> ${_ih('visor.prop.acepto')}</label>
            <label class="chip" style="margin-top:6px"><input type="checkbox" id="pq-comp" ${q.compartir ? 'checked' : ''}> ${_ih('visor.prop.compartir')}</label>
            <p class="nota">🧒 ${_ih('visor.prop.menores')}</p>
            ${c.criterios ? `<details class="caja"><summary>📜 ${_ih('visor.prop.criterios')}</summary><p class="texto-largo">${esc(c.criterios)}</p></details>` : ''}</fieldset>
            ${this.correoAlt(asunto)}`;
    },
    correoAlt(asunto, cuerpo = '') {
        const enl = Colabora.enlaces(asunto || _i('visor.prop.asunto', { catalogo: Datos.cfg.titulo || '' }), cuerpo || _i('visor.prop.cuerpoCorreo', { catalogo: Datos.cfg.titulo || '' }));
        return enl ? `<details class="caja"><summary>✉️ ${_ih('visor.prop.alternativa')}</summary><p class="nota">${_ih('visor.prop.alternativa.d')}</p><div class="fila">${enl}</div></details>` : '';
    },
    leerQuien() {
        const q = { nombre: $('pq-n').value.trim(), publico: $('pq-pub').checked, rol: $('pq-rol').value, recopilo: $('pq-rec').value.trim(), contacto: $('pq-con').value.trim(),
            compartir: $('pq-comp').checked, acepto: $('pq-ok').checked };
        LS.set('prop:quien:' + Datos.base, { nombre: q.nombre, publico: q.publico, rol: q.rol, recopilo: q.recopilo, contacto: q.contacto, compartir: q.compartir });
        if ($('pq-col')) LS.set('prop:colab:' + Datos.base, $('pq-col').value.trim());
        return q;
    },
    /** Pone en cola y envía; muestra el código o que quedó esperando. */
    async mandar(items, resumen, msgId) {
        const m = h => { const x = $(msgId); if (x) x.innerHTML = h; };
        const q = this.leerQuien();
        if (!q.acepto) return m(`<div class="error-box">${_ih('visor.prop.faltaAcepto')}</div>`);
        if (!items.length) return m(`<div class="error-box">${_ih('visor.prop.vacio')}</div>`);
        await this.encolar({ quien: q, items, resumen: String(resumen).slice(0, 80), web: $('pq-web') ? $('pq-web').value : '' });
        m(`<div class="caja" role="status">⏳ ${_ih('visor.prop.enviando')}</div>`);
        const r = await this.enviarCola();
        if (r && r.codigo) {
            const malas = (r.resultados || []).filter(x => !x.ok);
            m(`<div class="caja exito"><b>✅ ${_ih('visor.prop.recibida')}</b><div class="codigo-grande">${esc(r.codigo)}</div><p class="nota">${_ih('visor.prop.anota')}</p>
                ${malas.length ? `<p>⚠️ ${malas.map(x => esc(_i('visor.prop.err.' + x.error, { n: x.i + 1 }))).join('<br>')}</p>` : ''}
                <button class="btn btn-secondary" onclick="Propuestas.misPropuestas()">📬 ${_ih('visor.prop.mias')}</button></div>`);
        } else if (r && r.espera) m(`<div class="caja">📶 ${_ih(r.error === 'mant' ? 'visor.prop.esperaMant' : r.error === 'limite' ? 'visor.prop.esperaLimite' : 'visor.prop.espera')}</div>`);
        else if (r && r.error) m(`<div class="error-box">${esc(_i('visor.prop.err.' + r.error, { n: '' }))}</div>${this.correoAlt()}`);
        else m(`<div class="caja">📶 ${_ih('visor.prop.espera')}</div>`);
    },
    /* ---------- formularios ---------- */
    selLenguas(id, sel, conVacia) {
        return `<select id="${id}" class="modern-select">${conVacia ? `<option value="">${_ih('visor.prop.lenguaOpc')}</option>` : ''}${Datos.langs.map(l => `<option value="${esc(l.id)}" ${l.id === sel ? 'selected' : ''}>${esc(l.nombre)}</option>`).join('')}</select>`;
    },
    /** ➕ Proponer palabras (una o varias). `actividad` llega desde «Pregunta en casa» o un código de hoja. */
    abrirPalabras(actividad = '') {
        if (!this.activo()) return UI.toast(_i('visor.prop.cerradas'), 4000);
        const L = Datos.lengua, otra = (Datos.langs.find(l => l.id !== L) || {}).id || '';
        const cats = [...Datos.cats.values()].map(c => Datos.categoria(c.id));
        const fila = i => `<div class="caja pp-fila" data-i="${i}"><div class="fila"><input class="input-box pp-t" style="flex:2" maxlength="300" placeholder="${esc(_i('visor.prop.palabra'))}" aria-label="${esc(_i('visor.prop.palabra'))}">${this.selLenguas('pp-l' + i, L)}</div>
            <div class="fila" style="margin-top:6px"><input class="input-box pp-s" style="flex:2" maxlength="300" placeholder="${esc(_i('visor.prop.significa'))}" aria-label="${esc(_i('visor.prop.significa'))}">${this.selLenguas('pp-sl' + i, otra, true)}</div>
            <details><summary class="nota" style="cursor:pointer">${_ih('visor.prop.masDatos')}</summary>
                <textarea class="input-box pp-d" rows="2" maxlength="600" placeholder="${esc(_i('visor.prop.descripcion.ph'))}"></textarea>
                <input class="input-box pp-v" maxlength="200" placeholder="${esc(_i('visor.prop.dondeSeDice'))}" style="margin-top:6px">
                <input class="input-box pp-c" maxlength="120" list="pp-cats" placeholder="${esc(_i('visor.prop.categoria'))}" style="margin-top:6px"></details></div>`;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">➕ ${_ih('visor.prop.palabras')}</h2>
            <p class="nota">${_ih('visor.prop.palabras.d')}</p>
            <datalist id="pp-cats">${cats.map(c => `<option value="${esc(c)}">`).join('')}</datalist>
            <div id="pp-filas">${fila(0)}</div>
            <div class="fila"><button class="btn btn-secondary" id="pp-mas">＋ ${_ih('visor.prop.otraPalabra')}</button><button class="btn btn-secondary" id="pp-pegar">📋 ${_ih('visor.prop.pegar')}</button></div>
            <div id="pp-pegarb" class="oculto"><textarea id="pp-lista" class="input-box" rows="5" placeholder="${esc(_i('visor.prop.pegar.ph'))}"></textarea><button class="btn btn-secondary" id="pp-usar">${_ih('visor.prop.pegar.usar')}</button></div>
            <label class="t" for="pp-act">🧩 ${_ih('visor.prop.actividad')}</label><input id="pp-act" class="input-box" maxlength="120" value="${esc(actividad)}" placeholder="${esc(_i('visor.prop.actividad.ph'))}">
            ${this.comun(_i('visor.prop.asunto.palabras', { act: actividad }))}
            <div id="pp-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="pp-ok">📤 ${_ih('visor.prop.enviar')}</button></div></div>`);
        let n = 1;
        $('pp-mas').addEventListener('click', () => { $('pp-filas').insertAdjacentHTML('beforeend', fila(n++)); });
        $('pp-pegar').addEventListener('click', () => $('pp-pegarb').classList.toggle('oculto'));
        $('pp-usar').addEventListener('click', () => {
            const lineas = $('pp-lista').value.split(/\n+/).map(x => x.trim()).filter(Boolean).slice(0, 40);
            lineas.forEach(lin => {
                const [p, s] = lin.split(/\s*[=:–—-]\s+|\s*=\s*/);
                const vacia = [...document.querySelectorAll('.pp-fila')].find(f => !f.querySelector('.pp-t').value);
                if (!vacia) $('pp-filas').insertAdjacentHTML('beforeend', fila(n++));
                const f = vacia || [...document.querySelectorAll('.pp-fila')].pop();
                f.querySelector('.pp-t').value = Texto.atajos(p || '', f.querySelector('select').value); f.querySelector('.pp-s').value = s || '';
            });
            $('pp-lista').value = ''; $('pp-pegarb').classList.add('oculto');
        });
        $('pp-ok').addEventListener('click', () => {
            const act = $('pp-act').value.trim();
            const items = [...document.querySelectorAll('.pp-fila')].map(f => {
                const i = f.dataset.i, L1 = $('pp-l' + i).value;
                return { tipo: 'palabra', texto: Texto.atajos(f.querySelector('.pp-t').value.trim(), L1), id_idioma: L1, significado: f.querySelector('.pp-s').value.trim(), id_idioma_significado: $('pp-sl' + i).value,
                    descripcion: Texto.atajos(f.querySelector('.pp-d').value.trim(), L1), variante: f.querySelector('.pp-v').value.trim(), categoria: f.querySelector('.pp-c').value.trim(), actividad: act };
            }).filter(x => x.texto);
            this.mandar(items, items.map(x => x.texto).join(', '), 'pp-msg');
        });
    },
    /** 📖 Proponer una descripción (también desde una pista propia de un docente). */
    abrirDescripcion(id, L, texto = '') {
        const w = Datos.byId.get(id); if (!w || !this.activo()) return;
        L = L || Datos.lengua;
        const actual = w.t[L]?.desc || '';
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📖 ${_ih('visor.prop.desc')}</h2>
            <div class="caja"><b${Texto.lang(L)}>${esc(Datos.texto(w, L))}</b> · ${esc(Datos.nombreLengua(L))}${actual ? `<p class="nota">${_ih('visor.prop.descActual')} ${esc(actual)}</p>` : ''}</div>
            <label class="t" for="pd-l">${_ih('visor.prop.enLengua')}</label>${this.selLenguas('pd-l', L)}
            <label class="t" for="pd-t">${_ih('visor.prop.descripcion')}</label><textarea id="pd-t" class="input-box" rows="4" maxlength="600" placeholder="${esc(_i('visor.prop.descripcion.ph'))}">${esc(texto)}</textarea>
            ${this.comun(_i('visor.prop.asunto.desc', { palabra: Datos.texto(w) }))}<div id="pd-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="pd-ok">📤 ${_ih('visor.prop.enviar')}</button></div></div>`);
        $('pd-ok').addEventListener('click', () => { const L2 = $('pd-l').value; this.mandar([{ tipo: 'descripcion', id_palabra: id, id_idioma: L2, descripcion: Texto.atajos($('pd-t').value.trim(), L2) }].filter(x => x.descripcion), Datos.texto(w), 'pd-msg'); });
    },
    /** ✏️ Proponer una corrección de cómo se escribe. */
    abrirCorreccion(id, L) {
        const w = Datos.byId.get(id); if (!w || !this.activo()) return;
        L = L || Datos.lengua;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">✏️ ${_ih('visor.prop.corr')}</h2>
            <label class="t" for="pc-l">${_ih('visor.lengua')}</label>${this.selLenguas('pc-l', w.t[L] ? L : Object.keys(w.t)[0])}
            <div class="caja" id="pc-act"></div>
            <label class="t" for="pc-t">${_ih('visor.prop.comoSeEscribe')}</label><input id="pc-t" class="input-box" maxlength="300">
            <label class="t" for="pc-n">${_ih('visor.prop.porque')}</label><textarea id="pc-n" class="input-box" rows="2" maxlength="500"></textarea>
            ${this.comun(_i('visor.prop.asunto.corr', { palabra: Datos.texto(w) }))}<div id="pc-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="pc-ok">📤 ${_ih('visor.prop.enviar')}</button></div></div>`);
        const act = () => { const t = w.t[$('pc-l').value]; $('pc-act').innerHTML = t ? `${_ih('visor.prop.ahoraDice')} <b${Texto.lang($('pc-l').value)}>${esc(t.completo || t.texto)}</b>` : _ih('visor.prop.noHay'); };
        $('pc-l').addEventListener('change', act); act();
        $('pc-ok').addEventListener('click', () => { const L2 = $('pc-l').value; this.mandar([{ tipo: 'correccion', id_palabra: id, id_idioma: L2, texto: Texto.atajos($('pc-t').value.trim(), L2), nota: $('pc-n').value.trim() }].filter(x => x.texto || x.nota), Datos.texto(w), 'pc-msg'); });
    },
    /** 📍 ¿Así se dice en tu comunidad? (documenta la variación; no vota cuál es «la correcta»). */
    abrirVariante(id, L) {
        const w = Datos.byId.get(id); if (!w || !this.activo()) return;
        L = w.t[L] ? L : (w.t[Datos.lengua] ? Datos.lengua : Object.keys(w.t)[0]);
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📍 ${_ih('visor.prop.var')}</h2>
            <div class="caja"><span class="gran-palabra"${Texto.lang(L)}>${esc(w.t[L].completo || w.t[L].texto)}</span> · ${esc(Datos.nombreLengua(L))}</div>
            <p class="nota">${_ih('visor.prop.var.d')}</p>
            <label class="t" for="pv-c">${_ih('visor.prop.tuComunidad')}</label><input id="pv-c" class="input-box" maxlength="200">
            <div class="chips" style="margin-top:8px"><label class="chip"><input type="radio" name="pv-r" value="igual" checked> ✅ ${_ih('visor.prop.var.igual')}</label><label class="chip"><input type="radio" name="pv-r" value="otra"> 🔀 ${_ih('visor.prop.var.otra')}</label></div>
            <div id="pv-ob" class="oculto"><label class="t" for="pv-t">${_ih('visor.prop.var.como')}</label><input id="pv-t" class="input-box" maxlength="300"></div>
            ${this.comun(_i('visor.prop.asunto.var', { palabra: Datos.texto(w) }))}<div id="pv-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="pv-ok">📤 ${_ih('visor.prop.enviar')}</button></div></div>`);
        document.querySelectorAll('input[name="pv-r"]').forEach(r => r.addEventListener('change', () => $('pv-ob').classList.toggle('oculto', document.querySelector('input[name="pv-r"]:checked').value !== 'otra')));
        $('pv-ok').addEventListener('click', () => {
            const otra = document.querySelector('input[name="pv-r"]:checked').value === 'otra', com = $('pv-c').value.trim();
            if (!com) return $('pv-msg').innerHTML = `<div class="error-box">${_ih('visor.prop.faltaComunidad')}</div>`;
            this.mandar([{ tipo: 'variante', id_palabra: id, id_idioma: L, texto: otra ? Texto.atajos($('pv-t').value.trim(), L) : '', variante: com }], Datos.texto(w), 'pv-msg');
        });
    },
    /** Imagen → JPEG ligero (los datos móviles y el espacio de Drive importan). */
    comprimir(file, max = 1200, objetivo = 300 * 1024) {
        return new Promise((res, rej) => {
            if (!/^image\//.test(file.type)) return rej(new Error('tipo'));
            const url = URL.createObjectURL(file), img = new Image();
            img.onload = () => {
                const k = Math.min(1, max / Math.max(img.width, img.height)), c = document.createElement('canvas');
                c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
                const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
                URL.revokeObjectURL(url);
                let q = 0.85, d = c.toDataURL('image/jpeg', q);
                while (d.length * 0.75 > objetivo && q > 0.4) { q -= 0.1; d = c.toDataURL('image/jpeg', q); }
                res(d);
            };
            img.onerror = () => { URL.revokeObjectURL(url); rej(new Error('img')); };
            img.src = url;
        });
    },
    cupo(w) { const max = w.imagenesMax != null ? w.imagenesMax : ((this.conf() || {}).maxImagenes || 7); return { n: w.imgs || (w.img ? 1 : 0), max }; },
    /** 🖼️ Proponer una imagen (dibujo, foto, ilustración) para la galería de una palabra. */
    abrirImagen(id) {
        const w = Datos.byId.get(id); if (!w || !this.activo()) return;
        const c = this.cupo(w);
        if (c.n >= c.max) return this.completa(w);
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🖼️ ${_ih('visor.prop.img')}</h2>
            <div class="caja"><b>${esc(Datos.texto(w))}</b> · 🖼️ ${_ih('visor.prop.cupo', { n: c.n, max: c.max })}</div>
            <label class="btn btn-secondary" style="margin-top:8px">📷 ${_ih('visor.prop.elegirImg')}<input type="file" id="pi-f" accept="image/*" capture="environment" class="oculto"></label>
            <div id="pi-prev"></div>
            <span class="t">${_ih('visor.prop.tipoImg')}</span><div class="chips">${['dibujo', 'foto', 'ilustracion', 'otro'].map((t, i) => `<label class="chip"><input type="radio" name="pi-t" value="${t}" ${i === 0 ? 'checked' : ''}> ${_ih('visor.prop.tipo.' + t)}</label>`).join('')}</div>
            <span class="t">${_ih('visor.prop.licencia')}</span><div class="chips">${['CC BY 4.0', 'CC BY-NC 4.0', 'Solo este catálogo'].map((t, i) => `<label class="chip"><input type="radio" name="pi-li" value="${t}" ${i === 2 ? 'checked' : ''}> ${_ih('visor.prop.lic.' + i)}</label>`).join('')}</div>
            <label class="chip" style="margin-top:8px"><input type="checkbox" id="pi-mia"> ${_ih('visor.prop.esMia')}</label>
            <p class="nota">⚠️ ${_ih('visor.prop.rostros')}</p>
            ${this.comun(_i('visor.prop.asunto.img', { palabra: Datos.texto(w) }))}<div id="pi-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="pi-ok" disabled>📤 ${_ih('visor.prop.enviar')}</button></div></div>`);
        let datos = '';
        $('pi-f').addEventListener('change', async e => {
            const f = e.target.files[0]; if (!f) return;
            try { datos = await this.comprimir(f); $('pi-prev').innerHTML = `<img src="${datos}" alt="" class="pi-prev"><p class="nota">${_ih('visor.prop.tamano', { kb: Math.round(datos.length * 0.75 / 1024) })}</p>`; $('pi-ok').disabled = false; }
            catch (er) { $('pi-msg').innerHTML = `<div class="error-box">${_ih('visor.prop.err.imagen', { n: '' })}</div>`; }
        });
        $('pi-ok').addEventListener('click', () => {
            if (!$('pi-mia').checked) return $('pi-msg').innerHTML = `<div class="error-box">${_ih('visor.prop.faltaMia')}</div>`;
            this.mandar([{ tipo: 'imagen', id_palabra: id, imagen: datos, tipo_imagen: document.querySelector('input[name="pi-t"]:checked').value, licencia: document.querySelector('input[name="pi-li"]:checked').value }], Datos.texto(w), 'pi-msg');
        });
    },
    /** Palabra con su cupo lleno: se invita a ayudar con otras de la misma categoría. */
    completa(w) {
        const otras = this.necesitan('img').filter(x => x.cat === w.cat && x.id !== w.id).slice(0, 12);
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🖼️ ${_ih('visor.prop.completa.t')}</h2><p>${_ih('visor.prop.completa', { palabra: Datos.texto(w) })}</p>
            ${otras.length ? `<div class="chips">${otras.map(x => `<button class="chip-rel" data-pimg="${esc(x.id)}">${esc(Datos.texto(x))}</button>`).join('')}</div>` : ''}
            <div class="acciones"><button class="btn btn-secondary" onclick="Propuestas.dondeAyudar()">🤝 ${_ih('visor.prop.donde')}</button></div></div>`);
        $('modal-body').querySelectorAll('[data-pimg]').forEach(b => b.addEventListener('click', () => this.abrirImagen(b.dataset.pimg)));
    },
    /** 📷 Enviar fotos de hojas trabajadas («Pregunta en casa», «Dibuja la palabra»…). */
    abrirFotos(actividad = '') {
        if (!this.activo()) return UI.toast(_i('visor.prop.cerradas'), 4000);
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📷 ${_ih('visor.prop.fotos')}</h2>
            <p class="nota">${_ih('visor.prop.fotos.d')}</p>
            <label class="btn btn-secondary">📷 ${_ih('visor.prop.elegirFotos')}<input type="file" id="pf-f" accept="image/*" multiple class="oculto"></label>
            <div id="pf-prev" class="fila"></div>
            <label class="t" for="pf-act">🧩 ${_ih('visor.prop.actividad')}</label><input id="pf-act" class="input-box" maxlength="120" value="${esc(actividad)}" placeholder="${esc(_i('visor.prop.actividad.ph'))}">
            <label class="t" for="pf-n">${_ih('visor.prop.notaFotos')}</label><textarea id="pf-n" class="input-box" rows="2" maxlength="500"></textarea>
            <p class="nota">🧒 ${_ih('visor.prop.fotos.menores')}</p>
            ${this.comun(_i('visor.prop.asunto.fotos', { act: actividad }))}<div id="pf-msg"></div>
            <div class="acciones"><button class="btn btn-primary" id="pf-ok" disabled>📤 ${_ih('visor.prop.enviar')}</button></div></div>`);
        let fotos = [];
        $('pf-f').addEventListener('change', async e => {
            fotos = [];
            for (const f of [...e.target.files].slice(0, 8)) { try { fotos.push(await this.comprimir(f, 1600, 450 * 1024)); } catch (er) {} }
            $('pf-prev').innerHTML = fotos.map(d => `<img src="${d}" alt="" class="pi-mini">`).join('');
            $('pf-ok').disabled = !fotos.length;
        });
        $('pf-ok').addEventListener('click', () => this.mandar([{ tipo: 'fotos', fotos, actividad: $('pf-act').value.trim(), nota: $('pf-n').value.trim() }], _i('visor.prop.fotos.resumen', { n: fotos.length }), 'pf-msg'));
    },
    /** Pedir que se retire un nombre o una aportación (con o sin código; el correo siempre sirve). */
    abrirRetiro(codigo = '') {
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🙅 ${_ih('visor.prop.retiro')}</h2><p class="nota">${_ih('visor.prop.retiro.d')}</p>
            <label class="t" for="pr-c">${_ih('visor.prop.retiro.codigo')}</label><input id="pr-c" class="input-box" maxlength="12" value="${esc(codigo)}">
            <label class="t" for="pr-n">${_ih('visor.prop.retiro.que')}</label><textarea id="pr-n" class="input-box" rows="3" maxlength="500"></textarea>
            ${this.activo() ? this.comun(_i('visor.prop.asunto.retiro', { codigo })) : this.correoAlt(_i('visor.prop.asunto.retiro', { codigo }))}<div id="pr-msg"></div>
            ${this.activo() ? `<div class="acciones"><button class="btn btn-primary" id="pr-ok">📤 ${_ih('visor.prop.enviar')}</button></div>` : ''}</div>`);
        if ($('pr-ok')) $('pr-ok').addEventListener('click', () => this.mandar([{ tipo: 'retiro', nota: ($('pr-c').value.trim() ? $('pr-c').value.trim().toUpperCase() + ' · ' : '') + $('pr-n').value.trim(), texto: $('pr-c').value.trim().toUpperCase() }], _i('visor.prop.retiro'), 'pr-msg'));
    },
    /** 📬 Mis propuestas: lo enviado (por código) y lo que espera conexión. */
    async misPropuestas() {
        const m = this.mias(), cola = await this.cola();
        UI.abrir(`<div class="form"><h2 class="modal-titulo">📬 ${_ih('visor.prop.mias')}</h2>
            <div class="aviso-local"><p>${_ih('visor.prop.mias.d')}</p><p>${_ih('visor.prop.perdiste')}</p>${this.correoAlt(_i('visor.prop.asunto.perdi'))}</div>
            ${cola.length ? `<h3>📶 ${_ih('visor.prop.enCola')}</h3>${cola.map(e => `<div class="caja act-fila"><div><b>${esc(e.resumen || '')}</b><div class="nota">${esc(e.creado.slice(0, 10))}${e.error ? ' · ' + esc(_i('visor.prop.err.' + e.error, { n: '' })) : ''}</div></div>
                <div class="fila"><button class="btn btn-secondary btn-mini" data-q="${esc(e.id)}">🗑️ ${_ih('visor.prop.descartar')}</button></div></div>`).join('')}
                <button class="btn btn-primary" id="mp-reint">🔁 ${_ih('visor.prop.reintentar')}</button>` : ''}
            <div id="mp-lista">${m.length ? '⏳' : `<div class="caja">${_ih('visor.prop.mias.vacia')}</div>`}</div>
            <div class="acciones"><button class="btn btn-secondary" onclick="Propuestas.abrirRetiro()">🙅 ${_ih('visor.prop.retiro')}</button></div></div>`);
        $('modal-body').querySelectorAll('[data-q]').forEach(b => b.addEventListener('click', async () => { if (confirm(_i('visor.prop.descartar.c'))) { await this.quitar(b.dataset.q); this.misPropuestas(); } }));
        if ($('mp-reint')) $('mp-reint').addEventListener('click', async () => { await this.enviarCola(); this.misPropuestas(); });
        if (!m.length) return;
        let est = {};
        try { const r = await Datos.api({ action: 'propuestas', codigos: m.slice(0, 30).map(x => x.codigo).join(',') }); (r.propuestas || []).forEach(p => { est[p.codigo] = p.items; }); }
        catch (e) { est = null; }
        if (!$('mp-lista')) return;
        $('mp-lista').innerHTML = m.map(x => { const its = est && est[x.codigo];
            return `<div class="caja act-fila"><div><b class="codigo">${esc(x.codigo)}</b> · ${esc(x.fecha)}<div class="nota">${esc(x.resumen || '')}</div>
                ${its ? its.map(i => `<div>${esc(_i('visor.prop.tipo.' + i.tipo))}${i.texto ? ' · ' + esc(i.texto) : ''}: <b>${esc(_i('visor.prop.estado.' + (i.estado || 'Pendiente').replace(/\s/g, '')))}</b>${i.respuesta ? `<div class="nota">💬 ${esc(i.respuesta)}</div>` : ''}</div>`).join('') : `<div class="nota">${_ih(est ? 'visor.prop.sinDatos' : 'visor.prop.sinRed')}</div>`}</div>
                <div class="fila"><button class="btn btn-secondary btn-mini" data-ret="${esc(x.codigo)}">🙅 ${_ih('visor.prop.retirarEsta')}</button></div></div>`; }).join('');
        $('mp-lista').querySelectorAll('[data-ret]').forEach(b => b.addEventListener('click', () => this.abrirRetiro(b.dataset.ret)));
    },
    /** Palabras que necesitan algo: 'img' (sin imagen o con pocas) o 'desc' (sin descripción en una lengua). */
    necesitan(que, L = Datos.lengua) {
        const p = this.conf() || {};
        return Datos.words.filter(w => w.juegos !== false).filter(w => que === 'img' ? (!w.img || (w.imgs || 1) < Math.min(3, w.imagenesMax != null ? w.imagenesMax : (p.maxImagenes || 7))) && (w.imagenesMax !== 0)
            : (w.t[L] && !w.t[L].desc));
    },
    /** 🤝 Dónde ayudar: un mapa de lo que le falta al catálogo, por categoría. */
    dondeAyudar() {
        const L = Datos.lengua, sinImg = this.necesitan('img'), sinDesc = this.necesitan('desc', L);
        const porCat = l => { const g = new Map(); l.forEach(w => { const c = Datos.categoria(w.cat); if (!g.has(c)) g.set(c, []); g.get(c).push(w); }); return [...g].sort((a, b) => b[1].length - a[1].length); };
        const bloque = (l, que) => porCat(l).map(([c, ws]) => `<details class="caja"><summary><b>${esc(c)}</b> · ${ws.length}</summary><div class="chips" style="margin-top:6px">${ws.slice(0, 60).map(w => `<button class="chip-rel" data-${que}="${esc(w.id)}"${Texto.lang(w.t[L] ? L : '')}>${!w.img ? '∅ ' : ''}${esc(Datos.texto(w))}</button>`).join('')}</div></details>`).join('') || `<p class="nota">${_ih('visor.prop.nadaFalta')}</p>`;
        UI.abrir(`<div class="form"><h2 class="modal-titulo">🤝 ${_ih('visor.prop.donde')}</h2><p class="nota">${_ih('visor.prop.donde.d')}</p>
            <h3>🖼️ ${_ih('visor.prop.donde.img', { n: sinImg.length })}</h3>${bloque(sinImg, 'di')}
            <h3>📖 ${_ih('visor.prop.donde.desc', { n: sinDesc.length, lengua: Datos.nombreLengua(L) })}</h3>${bloque(sinDesc, 'dd')}
            <p class="nota">${_ih('visor.prop.donde.imprimir')}</p></div>`);
        $('modal-body').querySelectorAll('[data-di]').forEach(b => b.addEventListener('click', () => this.abrirImagen(b.dataset.di)));
        $('modal-body').querySelectorAll('[data-dd]').forEach(b => b.addEventListener('click', () => this.abrirDescripcion(b.dataset.dd, L)));
    },
    /** Botones de la ficha. */
    botonesFicha(w) {
        if (!this.activo()) return '';
        const c = this.cupo(w), L = Datos.lengua;
        return `<div class="caja prop-ficha"><b>🤝 ${_ih('visor.prop.ayuda')}</b><div class="fila" style="margin-top:6px">
            <button class="btn btn-secondary btn-mini" data-pf="corr">✏️ ${_ih('visor.prop.corr.corto')}</button>
            <button class="btn btn-secondary btn-mini" data-pf="desc">📖 ${_ih(w.t[L]?.desc ? 'visor.prop.desc.otra' : 'visor.prop.desc.corto')}</button>
            ${c.n < c.max ? `<button class="btn btn-secondary btn-mini" data-pf="img">🖼️ ${_ih('visor.prop.img.corto')} <small>(${c.n}/${c.max})</small></button>` : `<span class="nota">🖼️ ${_ih('visor.prop.cupo', { n: c.n, max: c.max })}</span>`}
            <button class="btn btn-secondary btn-mini" data-pf="var">📍 ${_ih('visor.prop.var.corto')}</button></div></div>`;
    },
    enlazarFicha(w) {
        $('modal-body').querySelectorAll('[data-pf]').forEach(b => b.addEventListener('click', () => {
            const L = Datos.lengua, a = b.dataset.pf;
            if (a === 'corr') this.abrirCorreccion(w.id, L); else if (a === 'desc') this.abrirDescripcion(w.id, L); else if (a === 'img') this.abrirImagen(w.id); else this.abrirVariante(w.id, L);
        }));
    },
    /** Enlace desde una hoja impresa (QR de «Pregunta en casa»): #proponer=<actividad> abre el formulario. */
    revisarEnlace() {
        const h = location.hash || '';
        if (!h.startsWith('#proponer')) return;
        const act = decodeURIComponent(h.split('=')[1] || '').slice(0, 120);
        history.replaceState(null, '', location.pathname + location.search);
        if (this.activo()) this.abrirPalabras(act); else UI.toast(_i('visor.prop.cerradas'), 5000);
    }
};
window.Propuestas = Propuestas;
</script>
