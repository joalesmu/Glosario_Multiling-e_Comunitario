# Cambios

Cada versión indica qué cambió para quien usa el catálogo y qué debe hacer quien lo actualiza.
Números de versión: `CONFIG.VERSION` en `Codigo.gs`, `VERSION_VISOR` en el visor y `VERSION_ESTRUCTURA` (hojas y columnas).

---

## 1.1 — octubre de 2026

Primera versión publicada como modelo para replicar. Desde esta versión el programa se distribuye bajo **AGPL-3.0** con los términos adicionales de [NOTICE.md](NOTICE.md); las versiones anteriores se publicaron bajo MIT.

### Al actualizar desde una versión anterior
1. Pega el `Codigo.gs` nuevo en Apps Script y guarda.
2. **Implementar → Gestionar implementaciones → ✏️ → Versión: Nueva versión → Implementar** (no una implementación nueva: cambiaría la dirección).
3. En la hoja: **🔧 Mantenimiento → 🧩 Revisar y reparar el catálogo**. Registra el ID de la hoja (hazlo **antes** de copiarla para hacer una plantilla) y crea las hojas `PROPUESTAS`, `IMAGENES` y `TIPOS_RELACION`, y las columnas `descripcion` (TRADUCCIONES), `imagen_alt` e `imagenes_max` (PALABRAS).
4. En AppSheet: **Regenerate structure** en las tablas que cambiaron (ver [docs/APPSHEET.md](docs/APPSHEET.md)).
5. Reemplaza `index.html` en tu alojamiento.
6. Las propuestas desde el visor vienen **apagadas**: actívalas en **🔗 Visor: conexión y mantenimiento → 📥 Participación** cuando el equipo esté listo para revisarlas.

### Hoja del catálogo
- **Relaciones entre palabras** con vocabulario propio (🕸️ Tipos de relación): sinónimo, antónimo, tipo de, parte de, variante de, deriva de y los que defina la comunidad. Se publican en el visor y se exportan a grafos (SKOS, GraphML, CSV).
- **Descripciones en la lengua** (opcionales) en TRADUCCIONES, editables desde ✏️ Editar palabras, con plantilla de importación por lengua.
- **📥 Propuestas de la comunidad:** revisar palabras, descripciones, correcciones, variantes, imágenes y fotos de hojas que llegan desde el visor; aceptar, sumar, transcribir o no incluir, con respuesta para quien propuso y crédito con su consentimiento.
- **Galerías:** varias imágenes por palabra (hoja IMAGENES), una principal, texto alternativo y un máximo por palabra que decide el equipo.
- **📥 Participación:** quién puede proponer (todos o con clave), límites diarios, criterios de revisión, reconocimiento público de quienes aportan y avisos de novedades.
- **📜 Constancia de colaboración** imprimible, con firmantes que define el equipo.
- **Leyenda** que aparece en el pie de los materiales impresos.
- Aviso en el panel cuando falta la dirección del visor, el correo de contacto o hay poco espacio en Drive.
- Diagnóstico claro cuando la dirección del visor no responde como catálogo.
- **Copias con identidad propia:** cada hoja guarda su ID (`CONFIGURACION.ID_HOJA`). Una copia (Archivo → Hacer una copia, o «Copiar app» de AppSheet con datos) se reconoce sola y, la primera vez que el programa trabaja en ella, recibe código de catálogo y clave nuevos, olvida la dirección, las carpetas, la app, los ajustes del visor, la publicación y el directorio de voces del original, deja como administrador a quien la copió y desactiva (sin borrar) el equipo heredado. Nunca publica en el archivo del original. Las copias de trabajo quedan excluidas.
- **Plantilla para compartir:** «Dejar en blanco» tiene la opción «Preparar como plantilla para compartir» (sin nombre, contacto, claves ni equipo; no publica). Cada copia de la plantilla se vuelve un catálogo propio al abrir 🚀 Primeros pasos.
- 🚀 Primeros pasos aparece primero en el menú en una copia o una plantilla, avisa de dónde viene la hoja y no obliga a escribir lenguas si la hoja ya las trae.
- Quien es dueño de la hoja siempre puede administrar.
- Textos más claros sobre los atajos: funcionan desde la app sin configurar nada; el modo captura queda como opción avanzada (🧷 Conservar lo que se escribe).
- «Dejar en blanco» (🧹 Limpiar o recuperar datos) también vacía propuestas, galerías, avisos y copias de trabajo heredadas.

### Visor
- **Recursos para imprimir**, con hoja del estudiante y del docente, versiones A·B·C, código para regenerar cada hoja, pistas configurables (imagen, otra lengua, descripción, palabra, pista propia o espacio en blanco) y ajustes de impresión (papel, márgenes, blanco y negro, letra grande o escolar, minúsculas o MAYÚSCULAS):
  sopa de letras, crucigrama (con niveles para grupos multigrado), ordena las letras, Basta, memorama recortable con reverso, «Pregunta en casa» con QR y guía para el docente, hoja de recolección, ¿con qué letra empieza?, cuenta las sílabas, arma con sílabas, hoja de trazo, «Escucha y señala» con QR al audio, glosario de la unidad y lista de cotejo en blanco.
- Lotería impresa con espacio para un verso y con los mismos ajustes de impresión.
- **Selección de palabras:** al azar, por categorías o desde **Mis listas**; filtros por largo, con audio y **solo con las letras que el grupo ya conoce**.
- **📂 Mis actividades:** se guardan en el dispositivo, se respaldan en archivo y se comparten por enlace; sin cuentas.
- **Proponer** desde el visor (palabras, descripción, corrección, variante, imagen, fotos de hojas), con cola sin conexión, código para consultar el estado y retirar, y el correo siempre como alternativa. **🤝 Dónde ayudar** muestra lo que falta.
- Fichas con galería, descripción y palabras relacionadas; búsqueda también en las descripciones.
- Nuevas actividades en pantalla: ¿Qué es? (con descripciones) y ¿Con qué letra empieza?
- Sopa de letras con palabras cruzadas y tres niveles.
- **📥 Llevar sin conexión** también guarda imágenes y audios.
- Crédito de autoría y enlace al código en el pie y en «Sobre el visor».

### Base del sistema (también desde la 1.0)
- Funciona en una hoja en blanco: **🚀 Primeros pasos** crea hojas, columnas, carpetas por lengua, atajos y ajustes; la estructura se repara sola (agrega lo que falta, nunca borra ni reordena).
- **Panel de control** con estado del catálogo, lenguas, ajustes, contacto, estados, atajos de escritura, publicación, intercambio, respaldo y reparación.
- **Atajos de escritura** por lengua («buscar y reemplazar» permanente, p. ej. `+` → `ɨ`) que se aplican solos cada hora, sin cambiar nada en la app; modo captura opcional (`texto_captura` → `texto`).
- **Estados** Borrador, En revisión, Publicado y Retirado; solo lo publicado sale. Marca «usar en juegos» por palabra.
- **Equipo** (administradores y capturistas) que AppSheet lee para dar permisos.
- **Voces, fuentes, autorizaciones y comunidades:** cada audio con quién habla, de dónde y su consentimiento; cada texto con su fuente y sus permisos.
- **Publicación previa** en Drive con caché y manifiesto; la API nunca entrega datos internos.
- **Intercambio entre comunidades** sin duplicados y **exportación a grafos** (N-Triples OntoLex/SKOS, CSV para Neo4j y Apache AGE, GraphML).
- **Copias de trabajo**, limpieza con vista previa y respaldo, palabras que se escriben igual, menús en varios idiomas.
- **Visor:** mantenimiento («congelar» o «cerrar»), clave del equipo, avisos, lotería en línea, modo oscuro y alto contraste.

### Estructura (`VERSION_ESTRUCTURA` 1.3)
Hojas nuevas: `TIPOS_RELACION`, `PROPUESTAS`, `IMAGENES`. Columnas nuevas: `TRADUCCIONES.descripcion`, `PALABRAS.imagen_alt`, `PALABRAS.imagenes_max`, `CONFIGURACION.ID_HOJA`, `CONFIGURACION.COPIA_DE`, `CONFIGURACION.PLANTILLA`. Ninguna columna se renombró ni se borró.

---

## 1.0 — 2025

Primera versión: hoja del catálogo trabajada con AppSheet y un visor web con búsqueda y juegos. Publicada bajo licencia MIT.
