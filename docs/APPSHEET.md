# La app de captura en AppSheet

AppSheet es **opcional**: es la app del teléfono para que el equipo (hasta 10 personas en la versión gratuita) capture palabras, fotos y datos en comunidad, incluso sin señal, y sincronice al tener internet. El público nunca usa AppSheet: consulta el visor.

Todo lo demás (revisar, publicar, propuestas, permisos, limpieza) se hace en la hoja, desde el menú **💠 CATÁLOGO**.

> Los nombres de menús de AppSheet aparecen en inglés, con su traducción, porque el editor cambia de idioma y de diseño con el tiempo. Si no encuentras una opción exactamente donde se indica, búscala por su nombre en inglés.

---

## AppSheet gratuito: lo que hay que saber

- La app funciona **sin pagar mientras esté en modo prototipo**, con **hasta 10 personas** (incluida quien la creó). **No pulses «Deploy»**: al desplegarla, AppSheet pide un plan de pago («User signin not allowed with the FREE plan»). Los avisos de la sección *Deploy* (ícono, caché del servidor, plan) se pueden ignorar.
- El límite se cuenta por **cuenta de Google dueña de la app** y por personas que la usan en el mes. Si se pasa de 10, AppSheet puede bloquear **todas** las apps prototipo de esa cuenta. Por eso:
  - mantén el equipo en 10 personas o menos (el panel 👥 Equipo avisa);
  - deja la **app plantilla en una cuenta distinta** de la del catálogo real, para que quienes la copian no cuenten contra tu app de trabajo.
- AppSheet describe el plan gratuito como uso de prueba o personal. Para un uso estable con más personas, la opción es un plan de pago (algunas ediciones de Google Workspace lo incluyen). El catálogo funciona igual sin AppSheet: se puede capturar desde ✏️ Editar palabras y recibir propuestas desde el visor.

## 1. Antes de empezar

1. La hoja del catálogo ya pasó por **🚀 Primeros pasos** y **🧩 Revisar y reparar el catálogo** (todas las hojas y columnas existen).
2. Hay al menos una lengua, una categoría y una palabra de ejemplo (AppSheet deduce los tipos de columna mejor con datos).
3. En **👥 Equipo** están los correos de quienes usarán la app.

## 2. Crear la app

1. En [appsheet.com](https://www.appsheet.com) → **Create → App → Start with existing data** y elige la hoja del catálogo.
2. AppSheet agrega la primera hoja. Agrega las demás en **Data → Tables → Add table (+)**:

| Agregar | Para qué | Quién edita en la app |
|---|---|---|
| `PALABRAS` | Cada idea, con su imagen | Equipo |
| `TRADUCCIONES` | Cómo se dice en cada lengua, con audio | Equipo |
| `IDIOMAS` | Elegir la lengua | Solo lectura |
| `CATEGORIAS` | Elegir la categoría | Solo lectura (o administradores) |
| `VOCES` | Elegir quién habla en un audio | Administradores |
| `FUENTES` | Elegir de dónde viene un texto | Solo lectura |
| `EQUIPO` | La usan las fórmulas de permisos | **Solo lectura** (se cambia desde el panel 👥 Equipo) |
| `RELACIONES` *(opcional)* | Vincular palabras | Equipo |

**No agregues** `CONFIGURACION`, `REGLAS_ESCRITURA`, `HISTORIAL`, `PROPUESTAS`, `MEDIOS_PROPUESTOS`, `IMAGENES`, `AVISOS`, `TEXTOS`, `COPIAS_TRABAJO`, `AUTORIZACIONES`, `ACTORES`, `COMUNIDADES` ni los reportes: son del programa y se manejan desde la hoja. Cuantas menos tablas, más rápida la sincronización en el teléfono.

## 3. Tipos de columna

En **Data → Columns**, revisa cada tabla. Lo importante:

### PALABRAS
| Columna | Tipo | Ajustes |
|---|---|---|
| `id_palabra` | Text | **Key**. *Initial value*: `UNIQUEID()`. Oculta. |
| `id_categoria` | **Ref** → `CATEGORIAS` | Obligatoria si el equipo lo decide. |
| `imagen_referencia` | **Image** | Permite tomar foto o elegir de la galería. |
| `estado` | Enum: `Borrador`, `En revisión`, `Publicado`, `Retirado` | *Initial value*: `"Borrador"`. Ver permisos abajo. |
| `usar_en_juegos` | Enum: `Sí`, `No` | *Initial value*: `"Sí"`. |
| `imagen_credito`, `imagen_alt` | Text / LongText | `imagen_alt`: descripción breve de la imagen para quien no ve. |
| `creado_por` | Email | *Initial value*: `USEREMAIL()`. Oculta y no editable. |
| `imagen_id`, `imagen_ref`, `imagen_publica`, `origen`, `imagenes_max` | — | **Ocultas** (*Show?* apagado). Las llena el programa. |

### TRADUCCIONES
| Columna | Tipo | Ajustes |
|---|---|---|
| `id_traduccion` | Text | **Key**, `UNIQUEID()`, oculta. |
| `id_palabra` | **Ref** → `PALABRAS` | Activa **Is a part of** para capturar las traducciones dentro de la palabra. |
| `id_idioma` | **Ref** → `IDIOMAS` | |
| `texto` | Text | Admite variantes separadas por comas. Ver «Atajos» abajo. |
| `audio` | **File** | AppSheet no graba audio: se sube un archivo grabado con el teléfono. |
| `id_voz` | Ref → `VOCES` | Quién habla. |
| `consentimiento` | Enum: `Público con nombre`, `Público sin nombre`, `Solo comunidad`, `Pendiente` | *Initial value*: `"Pendiente"`. |
| `silabas` | Text | Separadas con guion: `ha-ra-ma-ra`. Activan los juegos de sílabas. |
| `descripcion` | LongText | Opcional: qué es, escrito en la misma lengua. |
| `nota_variante` | Text | Dónde se dice así. |
| `id_fuente` | Ref → `FUENTES` | |
| `estado` | Enum (los cuatro estados) | Igual que en PALABRAS. |
| `texto_captura`, `audio_archivo`, `lote`, `origen`, `audio_nota` | — | Ocultas (si existe `texto_captura`, ver la sección 8). |
| `creado_por` | Email | *Initial value*: `USEREMAIL()`. Oculta y no editable (sirve para los atajos «solo para mí»). |

> **Valores de las listas:** escríbelos exactamente como están (con acentos y mayúsculas). Son datos que el programa y el visor leen; no se traducen.
> **Columnas `id_*`:** deben ser **Ref**, nunca Enum ni Color. La hoja no les pone lista desplegable a propósito para que AppSheet no se confunda.

## 4. Permisos con la hoja EQUIPO

Quien está activo en `EQUIPO` como *Capturista* agrega y edita; como *Administrador* además borra y publica; cualquier otra persona solo ve. En **Data → Tables →** (cada tabla editable) **→ Are updates allowed?** pega esta fórmula:

```
IF(IN(LOWER(USEREMAIL()), SELECT(EQUIPO[correo], AND([rol] = "Administrador", [activo] = "Sí"))), "ALL_CHANGES",
  IF(IN(LOWER(USEREMAIL()), SELECT(EQUIPO[correo], [activo] = "Sí")), "ADDS_AND_UPDATES", "READ_ONLY"))
```

Para que solo administradores puedan poner una palabra en *Publicado*, en la columna `estado` → **Valid if**:

```
OR([_THIS] <> "Publicado", IN(LOWER(USEREMAIL()), SELECT(EQUIPO[correo], AND([rol] = "Administrador", [activo] = "Sí"))))
```

Las mismas fórmulas aparecen en la hoja, en **🗂️ Configurar → 👥 Equipo**.

Además: **Security → Require user sign-in** encendido, y comparte la app (**Share**) con el mismo correo que pusiste en el panel 👥 Equipo. AppSheet no permite automatizar este último paso en la versión gratuita.

## 5. Vistas

- **PALABRAS** como *Deck* o *Gallery* con la imagen, agrupadas por categoría.
- **Conserva los nombres automáticos** `PALABRAS_Detail` y `TRADUCCIONES_Form`: los botones «Abrir en la app» de la hoja (lista de lo que falta, revisión de imágenes) abren exactamente esas vistas.
- *Slices* útiles (**Data → Slices**):
  - **Por revisar:** `[estado] = "En revisión"`.
  - **Mis registros:** `[creado_por] = USEREMAIL()`.
  - **Sin audio:** en TRADUCCIONES, `ISBLANK([audio])`.

## 6. Trabajo sin señal y sincronización

| Dónde | Ajuste | Valor |
|---|---|---|
| **Settings → Offline mode** (editor anterior: *Behavior → Offline/Sync*) | The app can start when offline | ✅ |
| | Store content for offline use | ✅ (guarda también las imágenes) |
| **Settings → Performance** | Delayed sync | ✅ los cambios se guardan en el teléfono hasta enviarse |
| | Automatic updates | ✅ envía sola la cola cuando hay señal |
| | Sync on start | ❌ con poca señal bloquearía la app al abrirla |

Con **Delayed sync** y **Automatic updates** encendidos, la app envía sola lo pendiente cuando encuentra internet, pero puede tardar hasta unos 30 minutos en intentarlo. Por eso el equipo debe conocer también el **botón de sincronizar** (🔄, arriba en la app; ya viene incluido, no hay que crearlo): muestra cuántos cambios faltan por enviar y, al tocarlo, envía todo en ese momento.

Indicación para el equipo:
- **Al llegar a un lugar con señal, toca 🔄** y espera a que el número desaparezca. Así sabes que todo subió.
- Antes de salir a comunidad, abre la app con Wi-Fi y espera a que termine de sincronizar (para tener todo guardado).
- Si la sincronización falla, los cambios **no se pierden**: siguen en el teléfono. No borres la app ni sus datos mientras haya cambios pendientes; vuelve a tocar 🔄 con mejor señal.

## 7. Imágenes y audios

AppSheet guarda los archivos en una carpeta junto a la hoja (por ejemplo `PALABRAS_Images`) y escribe en la celda una ruta como `PALABRAS_Images/foto.jpg`. El programa reconoce esas rutas, encuentra el archivo en Drive y lo comparte al publicar: no hace falta convertir nada. Si prefieres que los archivos queden en las carpetas del catálogo, cámbialo en la columna → **Image/File folder path**.

Para ahorrar datos y espacio en Drive, limita el tamaño de las fotos que suben los teléfonos: **Settings → Views: General → Inputs → Image upload size** (en el editor anterior: *UX → Options*). Es un ancho máximo para toda la app:

| Opción | Ancho máximo | Recomendación |
|---|---|---|
| Low | 200 px | Demasiado pequeña para imprimir |
| Default | 600 px | Suficiente para el visor y los juegos |
| **Medium** | **800 px** | **Recomendada:** se ve bien en el visor y en los materiales impresos (lotería, memorama) |
| High | 1600 px | Si el equipo quiere conservar fotos de mejor calidad |
| Full | Original | Solo con buena señal y espacio de sobra |

Las ilustraciones o fotos de archivo que deban conservarse en alta calidad se pueden subir directamente a la carpeta de Drive del catálogo y asignarse desde 🖼️ Imágenes y sonidos.

## 8. Atajos de escritura (`+` → `ɨ`)

**No hay que hacer nada en AppSheet.** El equipo escribe en `texto` como siempre (con atajos o sin ellos) y cada hora el programa corrige lo escrito con los atajos de la hoja. Los atajos se agregan, cambian o apagan en cualquier momento desde **🗂️ Configurar → ✍️ Atajos de escritura**, y para ver el cambio al momento basta con **🔧 Mantenimiento → 🔄 Aplicar los atajos a todo lo escrito**. La app no cambia.

Para que funcionen los atajos «solo para mí» (de una persona), la columna `creado_por` debe tener *Initial value* `USEREMAIL()`, como se indica en la sección 3.

<details>
<summary>Opcional: modo captura (casi nunca hace falta)</summary>

**🔧 Mantenimiento → 🧷 Conservar lo que se escribe (modo captura)** guarda exactamente lo que se escribió en una columna aparte (`texto_captura`) y calcula `texto` a partir de ella. Solo sirve si se necesita recalcular todo desde lo escrito originalmente, o usar atajos que cambian su propio resultado (por ejemplo, «a» → «aa»). Requiere un cambio único en la app: *Regenerate structure* en `TRADUCCIONES`, `texto_captura` editable y `texto` de solo lectura. Si se activa, conviene hacerlo antes de crear la plantilla, para que las copias ya lo traigan.
</details>

## 9. Conectar la app con el catálogo

Copia el código de la app (en la dirección del editor, después de `appId=`) y pégalo en **🗂️ Configurar → 🛠️ Datos del catálogo → Código de tu app de AppSheet**. Así funcionan los botones «Abrir en la app».

## 10. Cuando el programa cambie

Cada vez que **🧩 Revisar y reparar** agregue hojas o columnas (lo avisa al terminar), abre esas tablas en AppSheet y pulsa **Regenerate structure** (*Data → Columns → Regenerate*). Después revisa que las columnas `id_*` sigan siendo **Ref** y que las nuevas columnas técnicas queden ocultas.

---

## 11. Dejar la app como plantilla para otras comunidades

Las comunidades copiarán tu app con la opción **Copy** de AppSheet. Al copiarla con sus datos, AppSheet le crea a cada quien **su propia copia de la hoja**, conectada a su app. El programa reconoce esa copia y, al abrir **🚀 Primeros pasos**, le da identidad propia: código y clave nuevos, sin las carpetas, la dirección, el directorio de voces ni el equipo del original. Por eso la plantilla debe estar vacía: todo lo que tenga, lo hereda cada copia.

1. Termina y prueba la app con tu catálogo real. Antes, ejecuta **🧩 Revisar y reparar** en ese catálogo con el programa actualizado.
2. Haz una copia de la hoja (Archivo → Hacer una copia) y, **en la copia**, usa **🔧 Mantenimiento → 🧹 Limpiar o recuperar datos → Dejar en blanco** con **«Preparar como plantilla para compartir»** marcado. Quita palabras, voces, propuestas, nombre, contacto, créditos, claves y equipo, y marca la hoja como plantilla (no publica). El programa no permite hacerlo en la hoja original.
3. Deja una lengua de ejemplo solo si la plantilla es para esa lengua. AppSheet ya conoce los tipos de columna por la app que copiaste.
4. En AppSheet: **Manage → Collaborate & Publish → Copy app** (en el editor anterior: *Manage → Author → Copy app*). En la copia, apunta cada tabla a la hoja plantilla (**Data → Tables →** la tabla **→ Source**) y comprueba que **ninguna** quede conectada a tu catálogo real.
5. En la app plantilla, deja vacío el código de AppSheet en la hoja: cada comunidad pondrá el de su copia.
6. Comparte la app plantilla con permiso **View definition** (ver definición). Quien la recibe la encuentra en *Shared with me* y la copia con **More → Copy**, con la opción de copiar los datos encendida.

> Nunca compartas para copiar la app de tu catálogo real: la copia incluye todos los datos, imágenes y archivos.

Los pasos completos para las comunidades, incluido lo que no se copia (la dirección del visor y los permisos de AppSheet), están en [REPLICAR.md](REPLICAR.md).

### Lista final antes de compartir la plantilla
- [ ] Ninguna tabla de la app plantilla apunta a la hoja de un catálogo real.
- [ ] La hoja plantilla dice «Plantilla de catálogo» en el panel y no tiene palabras, voces, correos ni contactos.
- [ ] `EQUIPO` está vacía.
- [ ] Las columnas `id_*` son Ref y las técnicas están ocultas.
- [ ] Las fórmulas de permisos están en las tablas editables.
- [ ] El modo sin conexión está activo.
- [ ] Probaste copiar la app con **otra cuenta de Google** y completar [REPLICAR.md](REPLICAR.md) de principio a fin.

Referencias de AppSheet: [Copiar una app](https://support.google.com/appsheet/answer/12491039) · [Dejar que otra persona copie tu app](https://support.google.com/appsheet/answer/10104387).
