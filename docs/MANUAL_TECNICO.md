# Manual técnico

Para quien necesite **entender, corregir o ampliar** el programa: cómo está hecho, dónde está cada cosa, cómo hacer los cambios más comunes sin romper los catálogos que ya existen y cómo probar y publicar una versión nueva.

Para **usar** el catálogo no hace falta nada de esto: todo se hace desde el menú **💠 CATÁLOGO** de la hoja. Para instalarlo, lee el [README](../README.md); para la app de captura, [APPSHEET.md](APPSHEET.md).

Versión del programa que describe este manual: **1.1** (estructura 1.3).

---

## Contenido

1. [Principios que el código debe respetar](#1-principios-que-el-código-debe-respetar)
2. [Cómo está hecho](#2-cómo-está-hecho)
3. [El recorrido de un dato](#3-el-recorrido-de-un-dato)
4. [Las hojas (modelo de datos)](#4-las-hojas-modelo-de-datos)
5. [La dirección del catálogo (API)](#5-la-dirección-del-catálogo-api)
6. [Codigo.gs por dentro](#6-codigogs-por-dentro)
7. [El visor por dentro](#7-el-visor-por-dentro)
8. [Textos de la interfaz e idiomas](#8-textos-de-la-interfaz-e-idiomas)
9. [Recetas para cambios comunes](#9-recetas-para-cambios-comunes)
10. [Construir y probar](#10-construir-y-probar)
11. [Publicar una versión nueva](#11-publicar-una-versión-nueva)
12. [Seguridad y privacidad](#12-seguridad-y-privacidad)
13. [Límites conocidos](#13-límites-conocidos)
14. [Solución de problemas técnicos](#14-solución-de-problemas-técnicos)

---

## 1. Principios que el código debe respetar

Estas reglas no son de estilo: si un cambio las rompe, perjudica a comunidades que ya usan el programa.

| Principio | Qué significa en el código |
|---|---|
| **Nadie edita tablas a mano** | Todo ajuste tiene una ventana. Si agregas un dato configurable, agrega también dónde se cambia (panel o ventana). |
| **La estructura se repara, nunca se rompe** | `asegurarEstructura` crea hojas y columnas que faltan; **nunca** borra, renombra ni reordena. Una columna nueva se agrega al final del esquema. Una columna obsoleta se deja de usar, pero no se quita. |
| **Se lee y escribe por nombre de columna** | Nunca por posición. Usa `_tabla`, `_get`, `_agregar`, `_actualizar`. |
| **Solo sale lo publicado** | El visor recibe una copia preparada (`construirDatosPublicos`) con lo que está en estado *Publicado* y sin columnas privadas. |
| **Ninguna lengua es pivote** | No supongas que existe el español ni una «lengua principal» para que algo funcione. |
| **Nada se borra sin respaldo** | Toda operación destructiva hace antes una copia en `RESPALDOS` y muestra vista previa. |
| **Todo lo que llega del público pasa por revisión** | Las propuestas del visor entran a `PROPUESTAS` y no tocan el catálogo hasta que el equipo decide. |
| **Sin datos de menores ni cuentas en el visor** | El visor no pide nombres de estudiantes; lo de cada docente vive en su dispositivo. |
| **Gratis y sin permisos extra** | No agregues servicios que exijan pago o permisos de Google nuevos sin necesidad (por ejemplo `MailApp` pide un permiso nuevo y deja la dirección del visor sin funcionar hasta que alguien vuelva a autorizar). |
| **Crédito de autoría** | La AGPL-3.0 y [NOTICE.md](../NOTICE.md) piden conservar «Visor diseñado y desarrollado por Alejandro Estrada» visible y marcar las versiones modificadas. |

---

## 2. Cómo está hecho

```
┌───────────────────────────── Cuenta de Google de la comunidad ─────────────────────────────┐
│                                                                                             │
│   AppSheet (captura en el teléfono, opcional) ──lee y escribe──┐                            │
│                                                                ▼                            │
│   Google Sheets  ◄──── Codigo.gs (Apps Script ligado a la hoja) ────► Google Drive          │
│   (las hojas)          · menú 💠 CATÁLOGO y ventanas                  · imágenes y audios    │
│                        · tarea cada hora (atajos + publicar)          · catalogo_publicado.json│
│                        · aplicación web doGet / doPost                · RESPALDOS            │
│                                         │                                                   │
└─────────────────────────────────────────┼───────────────────────────────────────────────────┘
                                          │  https://script.google.com/macros/s/…/exec
                                          ▼
                     index.html (el visor, en GitHub Pages u otro alojamiento)
                     · un solo archivo, sin servidor propio ni dependencias externas
                     · guarda la última versión del catálogo en el dispositivo
```

| Pieza | Archivo | Lenguaje |
|---|---|---|
| Programa del catálogo | `Codigo.gs` (un solo archivo, ~12 500 líneas) | Google Apps Script (JavaScript V8) |
| Visor | `index.html` (generado) | HTML, CSS y JavaScript sin bibliotecas externas |
| Piezas del visor | `fuentes/visor/*.js`, `*.html` y `textos/visor_es.json` | — |
| Pruebas | `pruebas/` | Node.js (sin dependencias) y Python + Playwright |

**Por qué un solo archivo en cada lado:** quien replica el catálogo solo tiene que copiar y pegar `Codigo.gs` y subir `index.html`. No hay instalación, compilación ni paquetes. Las piezas de `fuentes/visor/` existen solo para que sea más fácil editar el visor; el archivo que se publica es `index.html`.

**Única biblioteca de terceros:** el generador de códigos QR `qrcode-generator` 1.4.4 (Kazuhiko Arase, licencia MIT), incluido dentro de `07a_qr.js` con su aviso de licencia.

---

## 3. El recorrido de un dato

1. **Captura.** Alguien escribe una palabra en AppSheet, en ✏️ Editar palabras, en una hoja para traductores (luego 📥 Traer) o la propone desde el visor (luego 📥 Propuestas). Entra como *Borrador* o *En revisión*.
2. **Atajos de escritura.** Cada hora, `tareaAutomatica` aplica las reglas de `REGLAS_ESCRITURA` (por ejemplo `+` → `ɨ`). Con el modo captura, lo escrito queda en `texto_captura` y la versión corregida en `texto`.
3. **Revisión.** El equipo cambia el estado a *Publicado* (📋 Revisar y publicar palabras).
4. **Publicación.** `publicarCatalogo` (cada hora, o al pulsar 📡 Mostrar los cambios en el visor) llama a `construirDatosPublicos`, que arma un JSON con lo publicado y sin columnas privadas, respetando permisos de fuentes y consentimiento de voces. Lo guarda en Drive (`catalogo_publicado.json`), en la caché y escribe un **manifiesto** (versión y conteos, ~1 KB). Comparte con «cualquiera con el enlace» las imágenes y audios publicados.
5. **Consulta.** El visor pide primero `?action=manifiesto`. Si la versión es la que ya tiene guardada, no descarga nada más; si cambió, pide `?action=glosario`.
6. **Sin conexión.** El visor guarda el catálogo en el dispositivo (`localStorage`) y, con 📥 Llevar sin conexión, imágenes y audios en IndexedDB.

Ninguna respuesta de la API lee las hojas: todo sale de la publicación previa. Por eso el visor responde rápido aunque el catálogo sea grande, y por eso **un cambio en la hoja no se ve hasta que se publica**.

---

## 4. Las hojas (modelo de datos)

La lista completa y exacta está en la función `_esquema()` de `Codigo.gs` (sección 10). Esa función es la fuente de verdad: `asegurarEstructura` crea lo que falte a partir de ella.

### Núcleo del catálogo

| Hoja | Clave | Columnas principales | Notas |
|---|---|---|---|
| `IDIOMAS` | `id_idioma` | `nombre_propio`, `nombre_completo`, `codigo_iso`, `glottocode`, `variante`, `region`, `alfabeto`, `carpeta_audios`, `estado` | `alfabeto`: letras separadas por espacios, con dígrafos; lo usan los juegos de letras. |
| `CATEGORIAS` | `id_categoria` | `nombre_categoria`, `imagen_categoria`, `estado` | |
| `TRADUCCIONES_CATS` | `id_trad_cat` | `id_categoria`, `id_idioma`, `texto` | Nombre de cada categoría en cada lengua. |
| `PALABRAS` | `id_palabra` | `id_categoria`, `imagen_referencia`, `estado`, `usar_en_juegos`, `imagen_credito`, `imagen_alt`, `imagenes_max` | Una fila = una idea, sin lengua. La imagen principal vive aquí. |
| `TRADUCCIONES` | `id_traduccion` | `id_palabra`, `id_idioma`, `texto`, `audio`, `silabas`, `nota_variante`, `descripcion`, `id_voz`, `id_fuente`, `estado` | Cómo se dice cada palabra en cada lengua. `texto` admite variantes separadas por comas. |
| `RELACIONES` | `id_relacion` | `id_origen`, `tipo`, `id_destino`, `estado` | Vínculos entre palabras. |
| `TIPOS_RELACION` | `id_tipo` | `nombre`, `nombre_inverso`, `simetrica`, `publica`, `sistema` | Vocabulario de relaciones; los de `sistema` no se borran. |
| `IMAGENES` | `id_imagen` | `id_palabra`, `archivo`, `tipo`, `credito`, `licencia`, `alt`, `orden`, `estado`, `solo_interna` | Galería: imágenes además de la principal. |

### Procedencia, permisos y voces

| Hoja | Para qué |
|---|---|
| `FUENTES` | De dónde viene cada texto (diccionario, persona, institución, aportación) y qué permite: `uso_publicar`, `uso_didactico`, `uso_adaptar`, `uso_compartir`, `uso_ia`. |
| `ACTORES` | Personas e instituciones (sin datos de contacto públicos). |
| `AUTORIZACIONES` | Permisos que una persona o comunidad da sobre una fuente; vigentes o impugnados. |
| `COMUNIDADES` | Comunidades, variantes y si hay disputa de representación. |
| `VOCES` | Quién presta su voz: `consentimiento`, `permiso_ia`, forma y fecha del permiso. El nombre real va en un directorio privado aparte. |

### Trabajo del equipo

| Hoja | Para qué |
|---|---|
| `EQUIPO` | `correo`, `nombre`, `rol` (Administrador / Capturista), `activo`. AppSheet la lee para dar permisos. |
| `REGLAS_ESCRITURA` | Atajos por lengua (`buscar` → `reemplazar`). |
| `PROPUESTAS` | Lo que llega desde el visor; estados *Pendiente*, *Aceptada*, *No incluida*, *Transcrita*. |
| `MEDIOS_PROPUESTOS` | Imágenes y audios que esperan revisión. |
| `AVISOS` | Avisos para el visor. |
| `HISTORIAL` | Quién cambió qué, cuándo, de qué a qué. |
| `COPIAS_TRABAJO` | Copias de trabajo («ramas») creadas desde este catálogo. |
| `MESA_DE_TRABAJO`, `AUDITORIA_ADMIN` | Reportes que genera el programa; se pueden borrar. |
| `CONFIGURACION` | Ajustes del catálogo: una fila de valores bajo encabezados (`NOMBRE_PROYECTO`, `API_TOKEN`, `URL_VISOR`, `VISOR_AJUSTES` en JSON, etc.). Se escribe solo con `_escribirConfig`. |

### Reglas del modelo
- **Identificadores** (`id_*`): se crean solos (🆔 Completar códigos que faltan). No se reutilizan ni se cambian; las relaciones dependen de ellos.
- **Estados:** `Borrador`, `En revisión`, `Publicado`, `Retirado` (`ESTADOS`). Una celda vacía cuenta como publicada, por compatibilidad con catálogos anteriores. El código compara con la forma canónica (`_estadoCanonico`), nunca con la etiqueta.
- **Valores de las listas** (estados, consentimientos, Sí/No): son datos, no se traducen. Las ayudas de las listas sí.
- **Columnas privadas** (`CONFIG.COLUMNAS_PRIVADAS`): nunca salen en la API aunque estén en la hoja.
- Las columnas `id_*` no llevan lista desplegable a propósito: AppSheet las interpretaría como Enum en lugar de Ref.

### Identidad de cada hoja (copias y plantillas)
Cuando alguien copia una hoja (Archivo → Hacer una copia, o «Copiar app» de AppSheet con datos), la copia se lleva el contenido de `CONFIGURACION` y normalmente también el programa y sus propiedades (que apuntan al archivo de publicación, al directorio de voces y a las carpetas del original). Para que una copia nunca publique encima del original ni comparta su código:

- Cada hoja guarda su propio ID en `CONFIGURACION.ID_HOJA` (y en la propiedad `HOJA_ID`). Lo escribe `_registrarHoja`, desde `asegurarEstructura`.
- `_copiaSinIdentidad(ss)`: el ID guardado en la hoja es de otra hoja y no es una copia de trabajo (`RAMA_DE`).
- `_asegurarIdentidad(ss)` le da identidad propia: `ID_CATALOGO` y clave nuevos; vacía `CONFIG_DE_IDENTIDAD` (dirección, carpetas, app, ajustes del visor); borra `PROPS_DEL_CATALOGO()`; deja `COPIA_DE`; pone como administrador a quien ejecuta y desactiva (sin borrar) al equipo heredado. No toca el contenido.
- Se llama al empezar `asegurarEstructura`, `publicarCatalogo`, `panelEstado`, `uiAsistenteInicial`, `limpiezaOpciones`, `_soyAdmin` y la API (`_identidadEnApi`, a lo más cada 6 horas). Así no importa por dónde entre la persona.
- **Catálogos anteriores** sin `ID_HOJA`: se registran como originales y nunca se reinician solos, aunque la propiedad no coincida (un falso positivo cambiaría la clave de un catálogo en uso).
- El archivo de publicación lleva en su descripción `hoja:<ID>`; `_guardarSnapshotDrive` no escribe en un archivo marcado por otra hoja.
- **Plantilla** (`CONFIGURACION.PLANTILLA = Sí`, desde «Dejar en blanco» con `op.plantilla`): sin código, clave, nombre, contacto ni equipo; `publicarCatalogo` no publica y `asegurarEstructura` no le crea código ni equipo. Cada copia la deja de ser al recibir identidad.
- El dueño de la hoja siempre es administrador (`_soyAdmin`), aunque no esté en `EQUIPO`.

---

## 5. La dirección del catálogo (API)

Es la aplicación web de Apps Script (`doGet` y `doPost`, sección 9 y 11 de `Codigo.gs`). Se implementa con **Ejecutar como: yo** y **Acceso: cualquier persona**.

### Consultas (GET)

| Parámetro | Responde |
|---|---|
| *(ninguno)* o `action=ping` | Estado de la API y lista de consultas. |
| `action=manifiesto` | Versión publicada, conteos, rol (`publico` o `equipo`) y estado del visor (avisos, mantenimiento, propuestas, leyenda). |
| `action=glosario` | El catálogo publicado completo. |
| `action=idiomas` | Solo las lenguas. |
| `action=paquete&ids=…&ancho=…` | Imágenes ligeras en base64 para jugar sin internet. |
| `type=audio&file=…` | Un audio publicado. |
| `action=propuestas&codigos=…` | Estado de las propuestas con esos códigos. |
| `action=lot_info`, `action=lot_estado` | Lotería en línea. |

Todas aceptan `token=…`. Sin clave en el catálogo, el acceso es público; con clave, hay que enviarla. La **clave del equipo** da rol `equipo` (ve el catálogo aunque esté en mantenimiento).

### Envíos (POST, cuerpo JSON)

| `accion` | Para qué |
|---|---|
| `proponer` | Propuestas del visor. Valida consentimiento, límites diarios por dispositivo y en total, campo trampa contra robots y el cupo de imágenes por palabra. |
| `lot_*` | Partidas de lotería en línea. |

### Respuestas
Siempre JSON con `status` (200, 400, 401, 403, 500 o 503 en mantenimiento). Si Google responde con HTML (implementación archivada, sin `doGet`, sin autorizar), el visor y el panel lo explican con `_diagnosticoNoJSON`.

### Mantenimiento
- **Congelar:** el público recibe la versión guardada al congelar; el equipo ve la actual.
- **Cerrar:** el público recibe 503 con el mensaje y el contacto del catálogo.
Se programa en 🔗 Visor: conexión y mantenimiento y lo revisa `tareaAutomatica`.

---

## 6. Codigo.gs por dentro

### Mapa de secciones
Las secciones están numeradas en comentarios `// N. TÍTULO`. Búscalas por número.

| Sección | Contenido |
|---|---|
| 1 | Menú (`_construirMenu`) |
| 2 | Configuración y acceso a tablas por nombre de columna |
| 3 | Reglas de escritura y `tareaAutomatica` |
| 4–6 | Mesa de trabajo, revisión de imágenes, plantillas para colaboradores e importación |
| 7 | Publicación previa (`construirDatosPublicos`, `publicarCatalogo`, imágenes y audios) |
| 8 | Mantenimiento y utilidades |
| 9 | API (`doGet`) |
| 10 | Esquema, asistente de Primeros pasos, panel de control y ajustes |
| 11 | Lotería en línea (`doPost`) |
| 12 | ¿Cómo va el catálogo? |
| 13 | Intercambio entre comunidades |
| 14 | Exportación a grafos |
| 15 | Equipo |
| 16 | Idioma de la interfaz (`T_ES`, `tx`) |
| 17–20 | Varias formas y coincidencias; fuentes, voces y permisos; limpiar y exportar por lengua; copias de trabajo |
| 21–23 | Imágenes y sonidos; ✏️ Editar palabras; visor: conexión, mantenimiento y avisos |
| 24 | Relaciones entre palabras |
| 25 | Propuestas de la comunidad, galerías y constancia |
| final | Textos de la interfaz: `T_ES` y `T_VISOR` |

### Funciones base (úsalas siempre)

| Función | Para qué |
|---|---|
| `_s(x)`, `_k(x)` | Texto limpio; texto en minúsculas para comparar. |
| `_tabla(ss, hoja)` | Lee una hoja una sola vez: `{ sh, headers, filas }`, con índice de columnas. |
| `_get(t, fila, 'columna')` | Valor por nombre de columna. |
| `_agregar(t, [objetos])` | Agrega filas por nombre. **Ignora columnas que la hoja no tiene**: antes asegura la columna con `_asegurarColumnas`. |
| `_actualizar(t, i, {cambios})` | Cambia solo esas celdas. |
| `readSheet(ss, hoja)` | Lista de objetos (solo lectura). |
| `getAppConfig(ss)` / `_escribirConfig(ss, {...})` | Ajustes del catálogo. |
| `_lock()` | Candado para operaciones que escriben en varias hojas. |
| `_historial(ss, hoja, id, antes, despues)` | Registra el cambio. |
| `_requerirAdmin(ss)` | Detiene la operación si quien la pide no es administrador. |
| `asegurarEstructura(ss)` | Crea hojas, columnas, formatos y listas que falten. Idempotente. |
| `responseJSON(obj)` | Respuesta de la API. |
| `tx(clave, vars)`, `txh(...)` | Texto de la interfaz en el idioma de cada persona (`txh` escapa HTML). |

### Convenciones de nombres

| Patrón | Qué es |
|---|---|
| `uiAlgo()` | Abre una ventana o barra lateral (`HtmlService`). Se llama desde el menú. |
| `servicioAlgo()` | Acción directa del menú, sin ventana o con un aviso. |
| `algoDatos()` | La ventana la llama con `google.script.run` para leer. |
| `algoGuardar(...)`, `algoAccion(...)` | La ventana la llama para escribir. Siempre valida y, si toca, exige administrador. |
| `_algo` | Uso interno. Apps Script no muestra en el menú de ejecución las funciones que terminan en `_`, pero las que empiezan con `_` sí se pueden llamar desde ventanas: **no expongas** datos privados en funciones que una ventana pueda llamar sin `_requerirAdmin`. |
| `COLS_ALGO` | Columnas de una hoja. |

### Ventanas
Cada ventana es HTML armado en una plantilla de texto dentro de su función `ui*`. Incluyen `_tScript()`, que pone en la página `tt(clave)`, `th(clave)` (escapado) y `_escT` para usar los mismos textos que el servidor. Los estilos comunes están en `_CSS_FICHAS`.

### Tareas automáticas
`_activarAutomatizacion` crea un único activador horario de `tareaAutomatica`: aplica atajos, publica si hubo cambios y revisa el mantenimiento programado. No agregues más activadores sin necesidad: Google limita su número y su tiempo diario.

---

## 7. El visor por dentro

### Piezas (en `fuentes/visor/`, en este orden)

| Pieza | Contiene |
|---|---|
| `01_head.html` | Encabezado, aviso de licencia y todo el CSS. |
| `02_body.html` | Estructura de la página (barra, secciones, ventana modal, pie). |
| `03_core.js` | Configuración (`API_URL`, `VERSION_VISOR`), textos (`_i`, `_ih`), `LS` (almacenamiento local), `Texto` (letras, alfabeto, grafemas), `Datos` (carga, caché, conteos), `UI`, `Ficha`, `Explorar`, `Sonido`, `Interfaz` (idioma), `Catalogos` (varios catálogos), `Colabora`, `ACTIVIDADES` y `Hub`. |
| `04_juegos.js` | `Comun` (piezas comunes), memorama, ahorcado, sopa de letras. |
| `05_basta_charadas.js` | Basta y charadas. |
| `06_aprender.js` | `Rondas`, repaso, escucha, ¿en qué lengua?, arma la palabra, dictado. |
| `07_loteria.js` | Lotería: en línea, impresa y sin conexión. |
| `07a_qr.js` | Generador de QR (biblioteca MIT). |
| `07b_recursos.js` | `Offline`, `Listas`, `Seleccion`, `Pistas`, `Impresion`, `Actividades`, `QueEs`. |
| `07c_propuestas.js` | `Propuestas`: formularios, cola sin conexión, Mis propuestas, Dónde ayudar. |
| `07d_impresos.js` | `Impresos` (registro de materiales para imprimir), sopa, crucigrama, ordena, Basta, memorama, QR, Pregunta en casa, recolección. |
| `07e_letras.js` | ¿Con qué letra empieza? (papel y pantalla), sílabas, arma con sílabas, trazo, Escucha y señala, glosario, cotejo. |
| `07f_progreso.js` | `Copia` (catálogo guardado en IndexedDB), `Perfiles` (locales, por apodo), `Progreso` (vistas, escuchadas, favoritas, partidas; envuelve `Ficha.abrir`, `Sonido.tocar`, `Motor.final`, `Rondas.final`, `Basta.final`, `Charadas.final`) e `Instalar` (app instalable, service worker, `storage.persist`). |
| `08_hojas_inicio.js` | Hojas de trabajo, etiquetas y tarjetas; arranque de la página. |

Los objetos son constantes globales. Las ventanas de vista previa de impresión (otra pestaña) acceden al visor por `window.opener`, así que lo que necesiten debe estar expuesto en `window` (`window.Pistas`, `window.Impresos`, `window.Propuestas`).

### Configuración de arranque
`API_URL` y `API_TOKEN` en `03_core.js` están **vacíos a propósito**: el mismo visor sirve para cualquier catálogo. La dirección llega por el enlace (`?api=…&token=…`), por la pantalla de conexión o por los catálogos recientes del dispositivo. Quien quiera un visor fijo para un solo catálogo puede escribirla ahí.

### Enlaces que entiende el visor

| Enlace | Hace |
|---|---|
| `?api=<dirección>&token=<clave>` | Conecta con un catálogo (la clave se quita de la barra después de leerla). |
| `?loteria=<nombre>&n=<número>` | Entra a una partida de lotería en línea. |
| `#act=<datos>` | Abre una actividad compartida desde 📂 Mis actividades. |
| `#proponer=<actividad>` | Abre el formulario para proponer (QR de Pregunta en casa y recolección). |
| `#oir=<id_palabra>&l=<id_idioma>` | Abre la ficha lista para escuchar (QR de Escucha y señala). |

### Lo que se guarda en el dispositivo
- `localStorage` (casi todo por catálogo, con la dirección como parte de la clave; los ajustes de impresión y el idioma de la interfaz son del dispositivo): catálogo guardado, lengua elegida, Mis listas, Mis actividades, pistas propias, ajustes de impresión, propuestas enviadas, avisos cerrados, idioma de la interfaz.
- IndexedDB `catalogo-lenguas-medios`: imágenes y audios de 📥 Llevar sin conexión, la cola de propuestas sin enviar y, desde la 1.2, la copia del catálogo (`<dirección>|cat|datos`; la de `localStorage` `cat:<dirección>` se migra sola).
- Progreso (1.2), en `localStorage`: `perfiles` y `perfil_actual` (del dispositivo), `prog:<dirección>:<perfil>` (vistas, escuchadas, favoritas, partidas, días) y el repaso `repaso:<dirección>:<lengua>[:<perfil>]`. El perfil `p0` («Este aparato») no lleva sufijo, para conservar el repaso de versiones anteriores.
- **Regla:** nada de esto se envía. Solo sale del aparato si la persona toca «Compartir mi resumen» o «Guardar copia de mi progreso». No se piden ni guardan nombres reales.
- La app instalada: `manifest.webmanifest`, `sw.js` (guarda solo los archivos del visor; las consultas al catálogo pasan directo) e `iconos/`.

Cambiar el nombre de una clave de almacenamiento hace que las personas «pierdan» lo guardado. Si hace falta, lee la clave vieja y migra.

### Agregar actividades: el registro `ACTIVIDADES`
Cada tarjeta del inicio es un objeto en `ACTIVIDADES` (`03_core.js`):

```js
{ sec: 'imprimir', ico: '🔤', id: 'inicialImp', d: '', req: { campo: 'textos', min: 6 }, soloSiHay: false, abrir: () => Impresos.abrir('inicial') }
```

- `sec`: `jugar`, `aprender` o `imprimir`.
- `req`: lo mínimo que la lengua necesita (`img`, `audio`, `imgAudio`, `largas`, `cats`, `textos`, `desc`, `silabas`, `lenguas`), contado en `Datos.conteo`. Si no alcanza, la tarjeta aparece desactivada con la explicación.
- `soloSiHay: true`: la tarjeta ni siquiera aparece si no alcanza (para lo que muchos catálogos no tendrán, como descripciones o sílabas).
- Textos: `visor.act.<id>` y `visor.act.<id>.d`.

---

## 8. Textos de la interfaz e idiomas

- **Hoja (menús y ventanas):** `T_ES` en `Codigo.gs`. Se usan con `tx('clave')`. Las traducciones a otros idiomas se instalan desde 🌐 Idioma de los menús y se guardan en la hoja `TEXTOS`.
- **Visor:** `fuentes/visor/textos/visor_es.json`. Se usan con `_i('visor.clave')` y `_ih(...)` (escapado). `construir.py` los mete en el visor; `sincronizar_textos.py` los copia al bloque `T_VISOR` de `Codigo.gs`, para que quien traduce los menús traduzca también el visor y el catálogo publique esa traducción.
- Variables: `{nombre}` dentro del texto, `_i('clave', { nombre: 'x' })`.
- **No escribas textos sueltos en el código.** Todo texto visible lleva clave, para que se pueda traducir.
- El crédito de autoría (`AUTOR`) no se traduce ni se quita.

---

## 9. Recetas para cambios comunes

### Agregar una columna a una hoja
1. Agrega el nombre **al final** de la lista de esa hoja en `_esquema()` (o de su `COLS_*`).
2. Si debe ser privada, agrégala a `CONFIG.COLUMNAS_PRIVADAS`.
3. Si debe salir al visor, agrégala en `construirDatosPublicos` y léela en `Datos.cargar` (`03_core.js`).
4. Si tiene lista desplegable, agrégala en el bloque de validaciones de la sección 10.
5. Sube `VERSION_ESTRUCTURA` y anótalo en `CHANGELOG.md`, incluyendo «Revisar y reparar» y *Regenerate structure* en AppSheet.

### Agregar una hoja
Agrega su entrada en `_esquema()` con `tecnica: true` si el equipo no debe editarla, define su `COLS_*` y, si debe vaciarse al «dejar en blanco» una copia, agrégala a `HOJAS_LIMPIABLES` o al caso `enBlanco` de la limpieza.

### Agregar una opción al menú
En `_construirMenu` (sección 1), con un texto nuevo `menu.algo` en `T_ES`. La función que llama debe existir y no terminar en `_`.

### Agregar un material para imprimir
En `07d_impresos.js` o `07e_letras.js`:

```js
Impresos.registrar('miMaterial', {
    ico: '🧩', resp: true, min: 4,              // resp: tiene hoja del docente; min: palabras mínimas
    pistas: ['imagen', 'lengua', 'descripcion'],// modos de pista que ofrece (o nada)
    opciones: op => `…HTML…`,                   // opciones propias (opcional)
    leer: () => ({ … }),                         // lee esas opciones
    filtrar: (pool, c) => pool,                  // qué palabras sirven
    modelo: (orden, c, r, i) => ({ ws: … }),     // datos de la hoja i (r: azar con semilla, reproducible por código)
    pagina: (m, c, docente) => `…HTML…`,         // la hoja (estudiante o docente)
    css: `…`                                     // estilos propios (usa mm y var(--k) para la letra grande)
});
```

Otras opciones: `versiones: false`, `sinSeleccion`, `sinPalabras`, `hojas: [1, 2, 5]` o `false`, `cuantas: (c, pool) => n` (cuántas hojas), `reverso`, `extra`, `niveles`, `ayudas`, `img: true`. Luego agrega su tarjeta en `ACTIVIDADES` y sus textos `visor.imp2.miMaterial`, `.d`, `.inst`, `.idea1`…`.idea3`, `visor.act.miMaterialImp` y `.d`. Las «ideas de uso» son sugerencias: el visor nunca decide el propósito didáctico por el docente.

Comprueba que **ninguna hoja se desborde** de la página (la prueba de humo lo revisa) en carta, oficio y A4.

### Agregar una actividad en pantalla
Sigue el patrón de `QueLetra` (`07e_letras.js`) o `Escucha` (`06_aprender.js`): `abrir()` con el formulario, `iniciar()` que arma las rondas con `Rondas.nuevo`, `ronda()` que pinta con `UI.cuerpo`, y `Rondas.final` al terminar. Marca `Juego.activo` mientras dure para que el visor no se actualice en medio.

### Agregar una consulta a la API
En `doGet` (o `doPost`), **después** de la revisión de clave y mantenimiento. Nunca leas las hojas en una consulta pública si puedes servirlo desde la publicación previa: las hojas son lentas y la cuota es limitada.

### Publicar un dato nuevo al visor
Agrégalo en `construirDatosPublicos` respetando estado, permisos de la fuente (`_usoFuente`) y consentimiento. Si es ajuste del visor (no dato del catálogo), va en `VISOR_AJUSTES` y sale en el `estado` del manifiesto (`_estadoParaVisor`).

---

## 10. Construir y probar

### Construir el visor
```bash
python3 fuentes/visor/construir.py           # escribe index.html
python3 fuentes/visor/sincronizar_textos.py  # copia los textos del visor a Codigo.gs
```
`construir.py` se detiene si el código usa un texto que no existe. Sube siempre **los dos archivos** cuando cambies textos del visor.

> Puedes editar `index.html` directamente para un arreglo rápido, pero la próxima vez que alguien lo construya desde `fuentes/` el cambio se perderá. Haz el cambio en la pieza.

### Pruebas del programa del catálogo (Node.js, sin instalar nada)
`pruebas/simulador_gas.js` imita lo necesario de Sheets, Drive y Apps Script para ejecutar `Codigo.gs` fuera de Google.

```bash
node pruebas/prueba_relaciones.js     Codigo.gs pruebas/simulador_gas.js
node pruebas/prueba_descripciones.js  Codigo.gs pruebas/simulador_gas.js
node pruebas/prueba_propuestas.js     Codigo.gs pruebas/simulador_gas.js
node pruebas/prueba_limpiar.js        Codigo.gs pruebas/simulador_gas.js
node pruebas/prueba_copia_en_blanco.js Codigo.gs pruebas/simulador_gas.js
node pruebas/prueba_replicar.js       Codigo.gs pruebas/simulador_gas.js
```
Cada una termina con `TODO OK` o con la lista de fallas. Revisa también la sintaxis:
```bash
node -e "new Function(require('fs').readFileSync('Codigo.gs','utf8'))"
```
(El simulador convierte `const` en `var`; esa revisión detecta, por ejemplo, una constante declarada dos veces, que en Apps Script impediría cargar todo el programa.)

### Prueba de humo del visor (Python + Playwright)
```bash
pip install playwright && python -m playwright install chromium
python3 pruebas/visor_humo.py
```
Abre el visor con un catálogo simulado, genera cada material impreso, revisa que ninguna hoja se desborde, prueba los filtros, una actividad en pantalla y el enlace de los QR. Deja capturas en `pruebas/capturas/`.

### Antes de publicar, revisa a mano
- En una **copia** de un catálogo real: Revisar y reparar, publicar, abrir el visor, proponer algo y resolverlo.
- El visor en un teléfono Android con datos lentos y sin conexión.
- Imprimir (o guardar en PDF) al menos un material en carta.

---

## 11. Publicar una versión nueva

1. Sube los números: `CONFIG.VERSION` (y `@version` en el encabezado de `Codigo.gs`), `VERSION_VISOR` en `03_core.js` y, si cambió el esquema, `VERSION_ESTRUCTURA`.
2. Escribe en `CHANGELOG.md` qué cambia y **qué debe hacer quien actualiza**.
3. Construye, sincroniza y corre las pruebas.
4. Sube `Codigo.gs`, `index.html` y `fuentes/` al repositorio.
5. Cada catálogo se actualiza así: pegar `Codigo.gs` → **Gestionar implementaciones → ✏️ → Nueva versión** → **Revisar y reparar** → *Regenerate structure* en AppSheet → reemplazar `index.html`.

**Compatibilidad:** un visor nuevo debe funcionar con un catálogo que todavía no se actualizó, y al revés. Por eso el visor trata como opcional todo dato nuevo (`w.t[L].desc || ''`) y el programa nunca exige columnas que «Revisar y reparar» no haya creado todavía.

**Versiones modificadas por otras personas:** la AGPL-3.0 pide publicar el código si se ofrece por internet, y [NOTICE.md](../NOTICE.md) pide conservar el crédito visible y marcar la versión como modificada (por ejemplo, «Basado en el visor diseñado y desarrollado por Alejandro Estrada; versión modificada por …»).

---

## 12. Seguridad y privacidad

- **La dirección es pública por diseño.** Solo entrega lo publicado. La clave del catálogo limita quién ve el visor; no es una contraseña fuerte (viaja en el enlace). No publiques nada que no deba ver cualquiera que tenga el enlace.
- **Columnas privadas** y estados nunca salen en la API. Correos, contactos, notas internas y nombres de voces sin permiso tampoco.
- **Audios:** solo los de traducciones publicadas y con consentimiento público; el nombre de quien habla, solo si lo autorizó.
- **Propuestas:** campo trampa contra robots (finge éxito y no guarda), límites diarios por dispositivo y en total, tamaño máximo de imagen, solo imágenes, y todo entra *Pendiente*. Las imágenes propuestas quedan en una carpeta privada hasta que se aceptan.
- **Funciones que llaman las ventanas:** cualquier función global sin `_` final puede llamarse desde una ventana del script. Las que leen o cambian datos sensibles deben empezar con `_requerirAdmin(ss)`.
- **Permisos de Google:** el programa usa Hojas, Drive, caché, propiedades y la aplicación web. Agregar un servicio nuevo (correo, calendario, solicitudes externas) cambia los permisos y obliga a reautorizar; mientras tanto la dirección del visor responde con una página de Google en lugar de JSON.
- **Sin datos de menores:** no agregues al visor campos para nombres de estudiantes ni seguimiento. Las listas de cotejo y hojas impresas se llenan a mano.

---

## 13. Límites conocidos

| Límite | Detalle | Qué hacer |
|---|---|---|
| Tiempo de ejecución de Apps Script | ~6 minutos por ejecución | Las operaciones grandes trabajan por bloques; importa por partes catálogos muy grandes. |
| Activadores y cuotas diarias | Las de una cuenta gratuita de Google | Un solo activador horario; la API sirve desde la publicación previa. |
| Caché | 100 KB por elemento | La publicación se parte en trozos (`CONFIG.CACHE`); el archivo en Drive es la copia completa. |
| AppSheet gratuito | Hasta 10 personas | El panel 👥 Equipo avisa al pasar el límite. |
| Espacio de Drive | 15 GB por cuenta gratuita (compartidos con Gmail) | El panel avisa con poco espacio; las imágenes propuestas se comprimen en el teléfono antes de enviarse. |
| Visor sin conexión | No es una app instalable (no usa *service worker*) | Abre la página una vez con internet; el catálogo y lo llevado sin conexión quedan guardados. Una versión instalable (PWA) es una mejora posible. |
| Un solo archivo de programa | `Codigo.gs` es largo | Busca por número de sección; respeta las convenciones de nombres. |
| Grafos | Hoy se exporta (N-Triples, CSV, GraphML) | La migración a una base de grafos usaría esas exportaciones y `RELACIONES`/`TIPOS_RELACION`. |

---

## 14. Solución de problemas técnicos

| Síntoma | Causa | Solución |
|---|---|---|
| El visor dice «La dirección no respondió como un catálogo» | Implementación archivada, sin `doGet` (código sin guardar) o falta autorizar | Guarda el código; en *Gestionar implementaciones* usa la activa; ejecuta cualquier función del menú una vez para autorizar. |
| Error «Identifier … has already been declared» al abrir la hoja | Dos `const` con el mismo nombre | La revisión de sintaxis con `new Function` lo detecta antes de publicar. |
| Una columna nueva no se llena | `_agregar` ignora columnas que la hoja no tiene | Llama antes a `_asegurarColumnas` o ejecuta Revisar y reparar. |
| AppSheet muestra columnas raras o le faltan | No se regeneró la estructura | *Regenerate structure* en esa tabla. |
| Las imágenes no se ven en el visor | El archivo no está compartido o la referencia es una ruta que no se encuentra | Publica de nuevo (comparte lo publicado) y revisa 📷 Revisar las imágenes. |
| Los cambios no aparecen | No se publicó, o el catálogo está congelado | 📡 Mostrar los cambios en el visor; revisa el mantenimiento. |
| Una copia sigue mostrando el catálogo original en su visor | Usa la dirección `/exec` de **su** implementación, no la del original | Implementa la copia (Nueva implementación) y pega su dirección en 🔗 Visor: conexión. |
| Tras actualizar, el visor de alguien sigue viejo | Su navegador conserva la versión anterior de `index.html` | Recargar la página; en GitHub Pages el cambio puede tardar unos minutos. |
