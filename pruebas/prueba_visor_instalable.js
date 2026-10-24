// Visor 1.2: instalable, sin conexión, copia del catálogo en IndexedDB, perfiles y progreso local.
// Uso, desde la raíz del repositorio:  python3 -m http.server 8766 &   y luego   node pruebas/prueba_visor_instalable.js
const { chromium } = require('playwright');
const BASE = 'http://localhost:8766/', A = 'https://script.google.com/macros/s/AAA/exec';
let F = 0; const ok = (c, m) => { console.log((c ? '✔ ' : '✘ ') + m); if (!c) F++; };
const svg = t => 'data:image/svg+xml,' + encodeURIComponent(`<svg xmlns='http://www.w3.org/2000/svg' width='200' height='200'><rect width='200' height='200' fill='#dbeafe'/><text x='100' y='120' font-size='60' text-anchor='middle'>${t}</text></svg>`);
const W = [['kiitɨ', 'casa'], ['tsɨkɨ', 'perro'], ['maxa', 'venado'], ['haramara', 'mar'], ['tukari', 'día'], ['hikuri', 'peyote'], ['tatewari', 'abuelo fuego'], ['naweri', 'tejuino'], ['wainu', 'pez'], ['kuka', 'chaquira'], ['tutu', 'flor'], ['hamuta', 'tortilla']];
const data = { config: { titulo: 'Acervo Wixárika', idCatalogo: 'cat-1', idiomaPrincipal: 'WIX' },
  idiomas: [{ id_idioma: 'WIX', nombre_propio: 'Wixárika', codigo_iso: 'hch' }, { id_idioma: 'ESP', nombre_propio: 'Español', codigo_iso: 'es' }],
  categorias: [{ id_categoria: 'c1', nombre_categoria: 'Casa' }, { id_categoria: 'c2', nombre_categoria: 'Campo' }],
  palabras: W.map((w, i) => ({ id_palabra: 'p' + i, id_categoria: i % 2 ? 'c1' : 'c2', imagen_referencia: svg(w[0].slice(0, 2)), juegos: true })),
  traducciones: W.flatMap((w, i) => [{ id_palabra: 'p' + i, id_idioma: 'WIX', texto: w[0], ...(i < 4 ? { audio: 'AUD' + i } : {}) }, { id_palabra: 'p' + i, id_idioma: 'ESP', texto: w[1] }]) };
const WAV = 'data:audio/wav;base64,' + Buffer.concat([Buffer.from('RIFF'), Buffer.from([36 + 800, 3, 0, 0]), Buffer.from('WAVEfmt '), Buffer.from([16, 0, 0, 0, 1, 0, 1, 0, 0x40, 0x1f, 0, 0, 0x80, 0x3e, 0, 0, 2, 0, 16, 0]), Buffer.from('data'), Buffer.from([0x20, 3, 0, 0]), Buffer.alloc(800)]).toString('base64');
async function api(route) {
  const u = new URL(route.request().url()), a = u.searchParams.get('action'), tipo = u.searchParams.get('type');
  const est = { revision: 0, avisos: [], mant: null, propuestas: { activo: false } };
  const body = tipo === 'audio' ? { status: 'ok', audioData: WAV } : a === 'manifiesto' ? { status: 200, version: 'v1', rol: 'publico', estado: est } : { status: 200, version: 'v1', data };
  await route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(body), headers: { 'Access-Control-Allow-Origin': '*' } });
}
async function contexto(b, extra = {}) {
  const c = await b.newContext(Object.assign({ viewport: { width: 420, height: 900 }, acceptDownloads: true, serviceWorkers: 'allow' }, extra));
  await c.route(u => u.href.startsWith('https://script.google.com'), api);
  await c.route(u => u.href.startsWith('https://fonts.'), r => r.fulfill({ status: 200, contentType: 'text/css', body: '' }));
  return c;
}
(async () => {
  const b = await chromium.launch({ args: ['--autoplay-policy=no-user-gesture-required'] });
  const c = await contexto(b), p = await c.newPage(), err = []; p.on('pageerror', e => err.push(e.message));
  // Una copia de la versión anterior (en localStorage) debe pasar a IndexedDB
  await p.goto(BASE + 'manifest.webmanifest');
  await p.evaluate(([A, data]) => { localStorage.setItem('cat:' + A, JSON.stringify({ version: 'v0', data, t: 1 })); localStorage.setItem('glosario_url', JSON.stringify(A)); }, [A, data]);
  await p.goto(BASE); await p.waitForTimeout(1800);
  ok((await p.textContent('#app-title')).includes('Acervo'), 'catálogo cargado');
  ok(await p.evaluate(() => localStorage.getItem('cat:' + Datos.base) === null), 'copia vieja quitada de localStorage');
  ok(await p.evaluate(async () => !!(await Copia.leer(Datos.base)).data), 'copia del catálogo en IndexedDB');
  ok(await p.evaluate(() => VERSION_VISOR) === '1.2', 'versión 1.2');
  const sw = await p.evaluate(async () => { const r = await navigator.serviceWorker.ready; return !!r.active; }); ok(sw, 'service worker activo');
  ok(await p.evaluate(() => !!document.querySelector('link[rel="manifest"]')), 'manifiesto enlazado');
  // Ficha: vista + favorita
  await p.evaluate(() => Ficha.abrir('p0')); await p.waitForSelector('#f-fav');
  await p.click('#f-fav'); ok((await p.textContent('#f-fav')).includes('Favorita'), 'marcar favorita');
  await p.click('#modal-body .audio-btn'); await p.waitForTimeout(800);
  await p.evaluate(() => UI.cerrarModal()); await p.evaluate(() => Ficha.abrir('p1')); await p.evaluate(() => UI.cerrarModal());
  let r = await p.evaluate(() => Progreso.resumen());
  ok(r.vistas === 2 && r.favoritas === 1, `vistas ${r.vistas}, favoritas ${r.favoritas}`);
  ok(r.escuchadas === 1, 'audio escuchado registrado: ' + r.escuchadas);
  // Partidas: rondas de aprender y un juego (motor)
  await p.evaluate(() => Rondas.final({ titulo: 'Escucha', L: 'WIX', total: 5, bien: 4, jugadas: [] }, () => {}));
  await p.evaluate(() => { UI.abrir('<div></div>'); Ahorcado.jugadores = [{ nombre: 'J', puntos: 2, tiempo: 30, color: '#000' }]; Ahorcado.registro = [{ id: 'p0', ok: true }, { id: 'p1', ok: false }]; Ahorcado.final('🏆 Ahorcado', 'WIX'); });
  await p.evaluate(() => { Basta.cfg = { L: 'WIX', modo: 'libre' }; Basta.nuevas = []; Basta.detener = () => {}; try { Basta.final(); } catch (e) {} });
  r = await p.evaluate(() => Progreso.resumen());
  ok(r.partidas === 3 && r.aciertos === Math.round(100 * 5 / 7), `partidas ${r.partidas}, aciertos ${r.aciertos} %`);
  // Perfiles
  await p.evaluate(() => UI.cerrarModal()); await p.click('#btn-perfil'); await p.waitForSelector('#pf-nuevo');
  ok((await p.textContent('#modal-body')).includes('nunca se envía'), 'aviso de privacidad');
  await p.click('#pf-nuevo'); await p.fill('#pf-apodo', 'Venado veloz'); await p.click('#pf-ok'); await p.waitForTimeout(300);
  ok(await p.evaluate(() => Perfiles.datos().apodo) === 'Venado veloz', 'perfil nuevo activo');
  ok(await p.evaluate(() => Progreso.resumen().vistas) === 0, 'perfil nuevo empieza en cero');
  ok(await p.evaluate(() => Repaso.clave('WIX').endsWith(Perfiles.sufijo()) && Perfiles.sufijo() !== ''), 'repaso separado por perfil');
  await p.evaluate(() => { UI.cerrarModal(); Ficha.abrir('p5'); UI.cerrarModal(); }); await p.click('#btn-perfil');
  await p.click('details summary'); const todos = await p.textContent('#modal-body');
  ok(todos.includes('Este aparato') && todos.includes('Venado veloz'), 'el docente ve todos los perfiles del aparato');
  await p.screenshot({ path: require('os').tmpdir() + '/visor_progreso.png', fullPage: true });
  // Compartir (intencional): sin navigator.share, se copia
  await c.grantPermissions(['clipboard-read', 'clipboard-write']);
  await p.evaluate(() => { delete navigator.share; });
  await p.click('#pf-compartir'); await p.waitForTimeout(300);
  const clip = await p.evaluate(() => navigator.clipboard.readText());
  ok(clip.includes('Acervo Wixárika') && clip.includes('palabras vistas') && clip.includes('Venado veloz'), 'resumen copiado solo al pedirlo, con el apodo del perfil');
  // Copia del progreso y recuperación en otro perfil
  const [dl] = await Promise.all([p.waitForEvent('download'), p.click('#pf-guardar')]);
  const ruta = require('os').tmpdir() + '/progreso.json'; await dl.saveAs(ruta);
  await p.evaluate(() => { Perfiles.elegir(Perfiles.crear('Pez azul', '🐟').id); Progreso.abrir(); });
  p.on('dialog', d => d.accept());
  await p.setInputFiles('#pf-archivo', ruta); await p.waitForTimeout(500);
  ok(await p.evaluate(() => Progreso.resumen().vistas) === 1, 'copia recuperada en otro perfil');
  // Botón «atrás»: cierra la ventana en vez de salir del visor
  await p.evaluate(() => UI.cerrarModal()); await p.evaluate(() => Ficha.abrir('p2'));
  await p.goBack(); await p.waitForTimeout(300);
  ok(await p.evaluate(() => $('modal').style.display !== 'flex') && p.url().startsWith(BASE), 'atrás cierra la ficha y sigue en el visor');
  await p.evaluate(() => { Ficha.abrir('p3'); UI.cerrarModal(); }); await p.waitForTimeout(300);
  ok(await p.evaluate(() => !(history.state && history.state.capa)), 'cerrar con ✕ no deja pasos de más en «atrás»');
  // App instalada: avisos adaptados y vista de impresión dentro de la página
  await p.evaluate(() => { Instalar.instalada = () => true; Actividades.abrir(); });
  ok((await p.textContent('#modal-body')).includes('desinstalas la app'), 'aviso de actividades para la app instalada');
  await p.evaluate(() => UI.cerrarModal());
  await p.evaluate(() => Impresos.abrir('glosario')); await p.waitForTimeout(200); await p.click('#im-ok'); await p.waitForTimeout(1200);
  const hojas = await p.evaluate(() => $('vista-imp') ? $('vista-imp').querySelector('iframe').contentDocument.querySelectorAll('section.hoja').length : -1);
  ok(hojas > 0, 'app instalada: vista de impresión dentro de la página (' + hojas + ' hoja/s)');
  await p.goBack(); await p.waitForTimeout(400);
  ok(await p.evaluate(() => !$('vista-imp')), 'atrás cierra la vista de impresión');
  await p.evaluate(() => { Instalar.instalada = () => false; Actividades.abrir(); });
  ok((await p.textContent('#modal-body')).includes('ventana privada'), 'en el navegador, el aviso de siempre'); await p.evaluate(() => UI.cerrarModal());
  ok(err.length === 0, 'sin errores: ' + err.join(' | '));
  // Sin conexión: abre desde lo guardado
  await p.evaluate(() => Perfiles.elegir('p0'));
  await c.setOffline(true); await c.unroute(u => true).catch(() => {});
  await c.route(u => u.href.startsWith('https://'), r => r.abort());
  await p.reload(); await p.waitForTimeout(1500);
  ok((await p.textContent('#app-title')).includes('Acervo') && await p.locator('.palabra, .card, [data-ficha]').count() > 0, 'sin conexión: abre con el catálogo guardado');
  await c.setOffline(false);
  // iPhone: botón instalar con instrucciones
  const ci = await contexto(b, { userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.5 Mobile/15E148 Safari/604.1', isMobile: true, hasTouch: true, viewport: { width: 390, height: 844 } });
  const q = await ci.newPage(); await q.goto(BASE + '?api=' + encodeURIComponent(A)); await q.waitForTimeout(1500);
  ok(await q.isVisible('#btn-instalar'), 'iPhone: botón instalar visible');
  await q.click('#btn-instalar'); ok((await q.textContent('#modal-body')).includes('Agregar a pantalla de inicio'), 'iPhone: instrucciones');
  // Pegar un enlace de invitación completo en 🔌: toma dirección y clave
  await q.evaluate(() => UI.cerrarModal()); await q.evaluate(() => UI.pantallaConexion());
  await q.fill('#cx-url', 'Te invito: https://joalesmu.github.io/Glosario/?api=' + encodeURIComponent('https://script.google.com/macros/s/BBB/exec') + '&token=clave123 ¡suerte!');
  await q.dispatchEvent('#cx-url', 'change');
  ok(await q.inputValue('#cx-url') === 'https://script.google.com/macros/s/BBB/exec' && await q.inputValue('#cx-tok') === 'clave123', 'enlace de invitación pegado: dirección y clave');
  await q.click('#cx-ok'); await q.waitForTimeout(1500);
  ok(await q.evaluate(() => Datos.base === 'https://script.google.com/macros/s/BBB/exec' && Datos.token === 'clave123' && Catalogos.lista().length === 2), 'cambia al catálogo del enlace y recuerda ambos');
  // Lista de catálogos recientes: nombre, lenguas y búsqueda
  await q.evaluate(() => { const l = Catalogos.lista(); l.forEach(c => { delete c.titulo; delete c.lenguas; }); for (let i = 0; i < 5; i++) l.push({ url: 'https://script.google.com/macros/s/Z' + i + '/exec', token: '' }); LS.set('glosario_catalogos', l); UI.pantallaConexion(); });
  await q.waitForTimeout(800);
  const filaA = await q.textContent('[data-rec="0"]');
  ok(filaA.includes('Acervo Wixárika') && filaA.includes('Wixárika'), 'recientes: nombre y lenguas tomados de la copia guardada: ' + filaA.trim().slice(0, 60));
  ok((await q.textContent('#vista')).includes('sin nombre todavía'), 'los que nunca se abrieron dicen «sin nombre todavía»');
  await q.fill('#cx-buscar', 'wixa'); await q.waitForTimeout(200);
  ok(await q.locator('[data-fila-rec]:visible').count() === 1, 'buscar entre muchos catálogos por nombre o lengua');
    // Dos catálogos en el mismo aparato no se mezclan
  ok(await p.evaluate(() => Progreso.clave('p0', 'https://script.google.com/macros/s/BBB/exec') !== Progreso.clave('p0')), 'progreso separado por catálogo');
  await b.close();
  console.log(F ? `FALLAS: ${F}` : 'TODO OK'); process.exitCode = F ? 1 : 0;
})();
