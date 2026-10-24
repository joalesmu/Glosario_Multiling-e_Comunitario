"""Arma index.html (el visor) a partir de sus piezas.

Uso, desde la carpeta raíz del repositorio:
    python3 fuentes/visor/construir.py            → escribe index.html
    python3 fuentes/visor/construir.py otro.html  → escribe en otra ruta (para probar)

Qué hace:
  1. Lee los textos de la interfaz (textos/*.json; una clave puede estar en un solo archivo).
  2. Revisa que toda clave usada en el código exista (y avisa de las que ya no se usan).
  3. Inserta los textos en 03_core.js (marca /*V_ES*/{}) y une las piezas en orden.
Solo usa la biblioteca estándar de Python 3.
"""
import json, glob, os, re, sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, '..', '..'))

# Orden de las piezas: así quedan en el archivo final (el orden importa: cada pieza usa las anteriores)
PIEZAS = ['01_head.html', '02_body.html', '03_core.js', '04_juegos.js', '05_basta_charadas.js', '06_aprender.js',
          '07_loteria.js', '07a_qr.js', '07b_recursos.js', '07c_propuestas.js', '07d_impresos.js', '07e_letras.js',
          '08_hojas_inicio.js']

# Prefijos de claves que el código arma en tiempo de ejecución (p. ej. 'visor.act.' + id): no se pueden revisar literalmente
DINAMICAS = ['visor.imp2.', 'visor.prop.', 'visor.pista.', 'visor.imp.', 'visor.hoja.para.', 'visor.rel.', 'visor.act.',
             'visor.req.', 'visor.mem.cara.', 'visor.sopa.', 'visor.char.', 'visor.lot.var.', 'visor.hoja.', 'visor.char.instr.',
             'visor.aviso.tipo.', 'visor.mant.m.', 'visor.sel.largo.']


def textos():
    V = {}
    for f in sorted(glob.glob(os.path.join(AQUI, 'textos', '*.json'))):
        for k, v in json.load(open(f, encoding='utf-8')).items():
            if k in V and V[k] != v:
                print('CHOQUE de textos:', k, '|', V[k], '|', v)
            V[k] = v
    return V


def main():
    V = textos()
    src = {p: open(os.path.join(AQUI, p), encoding='utf-8').read() for p in PIEZAS}
    todo = ''.join(src.values())
    usadas = set(re.findall(r"""(?:_ih?|T|Tp)\(\s*'(visor\.[\w.]+)'""", todo)) | set(re.findall(r'data-i="(visor\.[\w.]+)"', todo))
    faltan = sorted(k for k in usadas if k not in V and not k.endswith('.'))
    if faltan:
        print('FALTAN TEXTOS (agrégalos a textos/visor_es.json):', faltan)
        sys.exit(1)
    sin = sorted(k for k in V if k not in usadas and not any(k.startswith(d) for d in DINAMICAS))
    if sin:
        print('Aviso: textos que parecen no usarse:', len(sin))
    js = json.dumps(dict(sorted(V.items())), ensure_ascii=False, indent=0).replace('</', '<\\/')
    assert src['03_core.js'].count('/*V_ES*/{}') == 1, 'Falta la marca /*V_ES*/{} en 03_core.js'
    src['03_core.js'] = src['03_core.js'].replace('/*V_ES*/{}', js)
    salida = sys.argv[1] if len(sys.argv) > 1 else os.path.join(RAIZ, 'index.html')
    out = ''.join(src[p] for p in PIEZAS)
    open(salida, 'w', encoding='utf-8').write(out)
    print('Listo:', salida, f'({len(out):,} caracteres, {len(V)} textos)')


if __name__ == '__main__':
    main()
