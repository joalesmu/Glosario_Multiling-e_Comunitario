"""Copia los textos del visor (textos/*.json) al bloque T_VISOR de Codigo.gs.

Así, quien traduce la interfaz desde la hoja (🌐 Idioma de los menús) también ve y traduce los textos del visor.
Uso, desde la raíz del repositorio:  python3 fuentes/visor/sincronizar_textos.py
"""
import json, glob, os

AQUI = os.path.dirname(os.path.abspath(__file__))
GS = os.path.abspath(os.path.join(AQUI, '..', '..', 'Codigo.gs'))

V = {}
for f in sorted(glob.glob(os.path.join(AQUI, 'textos', '*.json'))):
    V.update(json.load(open(f, encoding='utf-8')))
s = open(GS, encoding='utf-8').read()
i = s.index('const T_VISOR = {\n')
j = s.index('\n};\nObject.assign(T_ES, T_VISOR);', i)
s = s[:i] + 'const T_VISOR = {\n' + ',\n'.join('  %s: %s' % (json.dumps(k, ensure_ascii=False), json.dumps(v, ensure_ascii=False)) for k, v in sorted(V.items())) + s[j:]
open(GS, 'w', encoding='utf-8').write(s)
print('T_VISOR actualizado en Codigo.gs:', len(V), 'textos')
