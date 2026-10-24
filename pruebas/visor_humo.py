import json, asyncio, sys, urllib.parse, os
from playwright.async_api import async_playwright
# Prueba de humo del visor: abre index.html con un catálogo simulado y genera cada material impreso.
# Uso, desde la raíz del repositorio:  python3 pruebas/visor_humo.py
# Requiere: pip install playwright  y  python -m playwright install chromium
RAIZ = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
VISOR = 'file://' + os.path.join(RAIZ, 'index.html')
OUT = os.path.join(RAIZ, 'pruebas', 'capturas'); os.makedirs(OUT, exist_ok=True)
def svg(t, c='#dbeafe'):
    s = f"<svg xmlns='http://www.w3.org/2000/svg' width='200' height='200'><rect width='200' height='200' fill='{c}'/><text x='100' y='120' font-size='60' text-anchor='middle'>{t}</text></svg>"
    return 'data:image/svg+xml,' + urllib.parse.quote(s)
W = [('kiitɨ','kii-tɨ','casa'),('tsɨkɨ','tsɨ-kɨ','perro'),('maxa','ma-xa','venado'),('haramara','ha-ra-ma-ra','mar'),('tukari','tu-ka-ri','día'),('hikuri','hi-ku-ri','peyote'),
     ('tatewari','ta-te-wa-ri','abuelo fuego'),('naweri','na-we-ri','tejuino'),('wainu','wai-nu','pez'),('kuka','ku-ka','chaquira'),('tutu','tu-tu','flor'),('hamuta','ha-mu-ta','tortilla'),
     ('muxa','mu-xa','borrego'),('kauyumari','kau-yu-ma-ri','venado azul'),('tsinari','tsi-na-ri','maíz'),('ikɨ','i-kɨ','esto')]
pal = [{'id_palabra': f'p{i}', 'id_categoria': 'c1' if i % 2 else 'c2', 'imagen_referencia': svg(w[0][:2]), 'juegos': True} for i, w in enumerate(W)]
tr = []
for i, (wx, sil, es) in enumerate(W):
    tr.append({'id_palabra': f'p{i}', 'id_idioma': 'WIX', 'texto': wx, 'silabas': sil, **({'audio': 'https://lh3.googleusercontent.com/d/AUD' + str(i)} if i < 6 else {}), **({'descripcion': f'Ki {wx} teteuyeika.'} if i < 5 else {})})
    tr.append({'id_palabra': f'p{i}', 'id_idioma': 'ESP', 'texto': es})
data = {'config': {'titulo': 'Acervo Wixárika', 'idCatalogo': 'cat-1', 'idiomaPrincipal': 'WIX', 'licencia': 'CC BY-NC-SA 4.0'},
        'idiomas': [{'id_idioma': 'WIX', 'nombre_propio': 'Wixárika', 'codigo_iso': 'hch'}, {'id_idioma': 'ESP', 'nombre_propio': 'Español', 'codigo_iso': 'es'}],
        'categorias': [{'id_categoria': 'c1', 'nombre_categoria': 'Casa'}, {'id_categoria': 'c2', 'nombre_categoria': 'Campo'}, {'id_categoria': 'c3', 'nombre_categoria': 'Otra'}],
        'palabras': pal, 'traducciones': tr}
async def ruta(route):
    q = dict(urllib.parse.parse_qsl(urllib.parse.urlparse(route.request.url).query)); a = q.get('action')
    est = {'revision': 0, 'avisos': [], 'mant': None, 'leyenda': 'Material comunitario', 'propuestas': {'activo': True, 'quien': 'todos', 'maxImagenes': 7}}
    body = {'status': 200, 'version': 'v1', 'rol': 'publico', 'estado': est} if a == 'manifiesto' else {'status': 200, 'version': 'v1', 'data': data}
    await route.fulfill(status=200, content_type='application/json', body=json.dumps(body), headers={'Access-Control-Allow-Origin': '*'})
A = 'https://script.google.com/macros/s/AAA/exec'
F = 0
def ok(c, m):
    global F; print(('✔ ' if c else '✘ ') + m); F += 0 if c else 1
async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(); ctx = await b.new_context(viewport={'width': 420, 'height': 900})
        await ctx.route(lambda u: u.startswith('https://script.google.com'), ruta)
        await ctx.route(lambda u: u.startswith('https://lh3.'), lambda r: r.fulfill(status=200, content_type='image/svg+xml', body=urllib.parse.unquote(svg('G', '#bbf7d0')[24:])))
        await ctx.route(lambda u: u.startswith('https://fonts.'), lambda r: r.fulfill(status=200, content_type='text/css', body=''))
        pg = await ctx.new_page(); errs = []; pg.on('pageerror', lambda e: errs.append(str(e)))
        await pg.goto(VISOR + '?api=' + A); await pg.wait_for_timeout(1200)
        async def imprimir(tipo, prep=None, nombre=None):
            await pg.evaluate(f"Impresos.abrir('{tipo}')"); await pg.wait_for_timeout(150)
            if prep: await prep()
            async with ctx.expect_page() as nueva: await pg.click('#im-ok')
            v = await nueva.value; ve = []; v.on('pageerror', lambda e: ve.append(str(e)))
            await v.wait_for_load_state(); await v.wait_for_timeout(500)
            await v.screenshot(path=f'{OUT}/{nombre or tipo}.png', full_page=True)
            n = await v.locator('section.hoja').count()
            # que ninguna hoja se desborde de su altura
            desb = await v.evaluate("[...document.querySelectorAll('section.hoja')].filter(s => s.scrollHeight > s.clientHeight + 4).length")
            html = await v.content()
            await v.close()
            return n, desb, html, ve
        for t in ['sopa', 'crucigrama', 'ordena', 'basta', 'memorama', 'preguntaCasa', 'recoleccion', 'inicial', 'silabas', 'armaSilabas', 'trazo', 'escucha', 'glosario', 'cotejo']:
            try:
                n, desb, html, ve = await imprimir(t)
                ok(n >= 1 and not ve and desb == 0, f'{t}: {n} hoja(s), desbordes {desb}, errores {ve[:1]}')
            except Exception as e:
                ok(False, f'{t}: {e}'); m = await pg.locator('#lot-msg').inner_text() if await pg.locator('#lot-msg').count() else ''; print('   msg:', m)
        # crucigrama con niveles
        async def niv():
            await pg.check('#im-niv')
        n, desb, html, ve = await imprimir('crucigrama', niv, 'crucigrama-niveles')
        ok(n == 6 and desb == 0, f'crucigrama con 3 niveles: alumno + docente = {n}')
        # memorama con reverso
        async def rev(): await pg.check('#op-rev')
        n, desb, html, ve = await imprimir('memorama', rev, 'memorama-rev'); ok('reverso' in html, 'memorama con reverso')
        # escucha docente con QR
        n, desb, html, ve = await imprimir('escucha', None, 'escucha2'); ok('qr-mini' in html, 'escucha: QR con enlace al audio'); ok(html.count('<svg') >= 1, 'QR dibujados')
        # inicial escribir
        async def esc_(): await pg.check('input[name="op-m"][value="escribir"]')
        n, desb, html, ve = await imprimir('inicial', esc_, 'inicial-escribir'); ok('caja-l' in html, 'inicial: cuadro para escribir')
        # sílabas docente
        n, desb, html, ve = await imprimir('silabas'); ok('ha · ra · ma · ra' in html or 'kii · tɨ' in html or ' · ' in html, 'sílabas: la hoja del docente muestra la división')
        # trazo con letras
        async def tr_():
            await pg.check('input[name="op-le"][value="k"]'); await pg.check('input[name="op-le"][value="m"]')
        n, desb, html, ve = await imprimir('trazo', tr_, 'trazo-km'); ok(n == 2 and 'tr-gran' in html, f'trazo: una hoja por letra ({n})')
        # glosario con español y descripción
        async def gl():
            await pg.select_option('#op-l2', 'ESP'); await pg.check('#op-desc')
        n, desb, html, ve = await imprimir('glosario', gl, 'glosario-esp'); ok('venado' in html and 'teteuyeika' in html and n == 2, f'glosario: con español y descripción, {n} hojas')
        # cotejo habilidades
        async def co():
            await pg.check('input[name="op-k"][value="habilidades"]'); await pg.fill('#op-hab', 'lee, escribe, pronuncia')
        n, desb, html, ve = await imprimir('cotejo', co, 'cotejo-hab'); ok('pronuncia' in html and html.count('<tr>') >= 26, 'cotejo: columnas propias y renglones')
        # letras conocidas: solo palabras con k, u, a, t, i
        await pg.evaluate("Impresos.abrir('inicial')"); await pg.wait_for_timeout(150)
        await pg.evaluate("document.querySelectorAll('#im-sel details').forEach(d => d.open = true)")
        for l in ['k', 'u', 'a', 't', 'i']: await pg.check(f'input[name="is-co"][value="{l}"]')
        sel = await pg.evaluate("Seleccion.palabras(Seleccion.leer('is'), 'WIX').map(w => w.t.WIX.texto)")
        ok(set(sel) == {'kuka', 'tutu'} or set(sel) <= {'kuka', 'tutu', 'kiitɨ'} and 'kuka' in sel, f'letras conocidas filtra: {sel}')
        # mayúsculas
        await pg.evaluate("Impresion.leer(); LS.set('imp:ajustes', Object.assign(Impresion.ajustes(), {mayus:'mayus'}))")
        n, desb, html, ve = await imprimir('glosario', None, 'glosario-mayus'); ok('uppercase' in html, 'ajuste MAYÚSCULAS')
        await pg.evaluate("LS.set('imp:ajustes', Object.assign(Impresion.ajustes(), {mayus:'como'}))")
        # En pantalla: ¿Con qué letra empieza?
        await pg.evaluate("UI.cerrarModal(); QueLetra.abrir()"); await pg.click('text=🚀')
        await pg.wait_for_timeout(200); ok(await pg.locator('#ql-ops [data-l]').count() >= 2, 'pantalla: opciones de letra')
        await pg.locator('#ql-ops [data-l]').first.click(); ok(await pg.locator('#ql-sig').count() == 1, 'pantalla: responde y sigue')
        await pg.screenshot(path=f'{OUT}/quelatra.png', full_page=True)
        for _ in range(12):
            if await pg.locator('#ql-sig').count(): await pg.click('#ql-sig')
            if await pg.locator('#ql-ops [data-l]').count(): await pg.locator('#ql-ops [data-l]').first.click()
            if await pg.locator('#r-otra').count(): break
        ok(await pg.locator('#r-otra').count() == 1, 'pantalla: resultado final')
        # Hub
        await pg.evaluate("UI.cerrarModal()")
        # enlace #oir
        pg2 = await ctx.new_page(); pg2.on('pageerror', lambda e: errs.append(str(e)))
        await pg2.goto(VISOR + '?api=' + A + '#oir=p3&l=WIX'); await pg2.wait_for_timeout(1500)
        ok('haramara' in await pg2.locator('#modal').inner_text(), 'QR #oir abre la ficha de la palabra')
        ok(errs == [], f'sin errores: {errs[:3]}')
        await b.close()
    print('TODO OK' if F == 0 else f'FALLAS {F}')
asyncio.run(main())
