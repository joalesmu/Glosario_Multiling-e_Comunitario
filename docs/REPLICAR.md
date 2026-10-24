# Crear el catálogo de tu comunidad

Esta guía es para quien quiere **su propio catálogo**: con sus lenguas, sus palabras, su equipo y su propio enlace para el visor. No hace falta saber programar ni tener cuenta de GitHub, y todo es gratis con una cuenta de Google.

Hay tres caminos. Todos terminan igual: una hoja propia, con su propio código y su propia clave, que nunca toca el catálogo del que salió.

| Camino | Para quién | Qué recibes |
|---|---|---|
| **A. Copiar la app de AppSheet** | Equipos que van a capturar desde el teléfono | La app y la hoja, ya conectadas |
| **B. Copiar la hoja plantilla** | Quien trabajará solo en la hoja, o hará su app después | La hoja con el programa |
| **C. Empezar en una hoja en blanco** | Quien no tiene acceso a la plantilla | La hoja; tú pegas el programa |

---

## Camino A · Copiar la app de AppSheet

1. Pide a quien comparte la plantilla que te la comparta en AppSheet (con permiso para **ver la definición**).
2. Entra a [appsheet.com](https://www.appsheet.com) con tu cuenta de Google. En **My apps → Shared with me** (compartidas contigo) busca la plantilla → **⋮ More → Copy**.
3. En la ventana de copia, deja encendida la opción de **copiar los datos**: así AppSheet te crea tu propia copia de la hoja.
4. Abre tu copia de la hoja en Google Drive (AppSheet suele guardarla en la carpeta `appsheet` → `data`).
5. Sigue con **[Los pasos de todos](#los-pasos-de-todos)**.

> Si al abrir la hoja no aparece el menú **💠 CATÁLOGO** después de recargar, el programa no se copió con ella: sigue el paso 2 del camino C y continúa.

## Camino B · Copiar la hoja plantilla

1. Abre el enlace de la hoja plantilla → **Archivo → Hacer una copia**.
2. Elige el nombre y la carpeta de tu Drive. Abre la copia.
3. Sigue con **[Los pasos de todos](#los-pasos-de-todos)**.

## Camino C · Empezar en una hoja en blanco

1. Crea una hoja de cálculo nueva en [Google Sheets](https://sheets.new).
2. **Extensiones → Apps Script**. Borra lo que haya, pega todo el contenido de [`Codigo.gs`](../Codigo.gs) y guarda (💾).
3. Vuelve a la hoja y recárgala.
4. Sigue con **[Los pasos de todos](#los-pasos-de-todos)**.

---

## Los pasos de todos

### 1. Primeros pasos (5 minutos)
1. Menú **💠 CATÁLOGO → 🚀 Primeros pasos** (aparece primero).
2. La primera vez, Google pide permiso para que el programa trabaje con tu hoja y tu Drive:
   - Elige tu cuenta.
   - Si aparece **«Google no verificó esta app»**, toca **Configuración avanzada → Ir a … (no seguro)**. Es normal: el programa no es de una empresa, es tuyo, vive en tu hoja y no envía tus datos a nadie.
   - Toca **Permitir**.
   - Vuelve a abrir **🚀 Primeros pasos**.
3. Si tu hoja vino de una copia (caminos A y B), verás: *«Esta hoja venía de… Ya tiene su propio código y quedaste como administrador»*. El programa le dio identidad propia: no comparte nada con el catálogo original.
4. Escribe el nombre de tu catálogo, tus lenguas (si la plantilla ya traía las tuyas, puedes dejar 0) y, si quieres, los grupos de palabras. Toca **🚀 Preparar el catálogo**.

### 2. La dirección del catálogo (una sola vez, 3 minutos)
Es lo que permite que el visor muestre tu catálogo.
1. **Extensiones → Apps Script → Implementar → Nueva implementación**.
2. En el engrane ⚙️ elige **Aplicación web**.
3. **Ejecutar como: Yo** · **Quién tiene acceso: Cualquier persona**. Toca **Implementar** (y autoriza si lo pide).
4. Copia la dirección que termina en **`/exec`**.
5. En la hoja: **💠 CATÁLOGO → 🤝 Colaborar → 🔗 Visor: conexión y mantenimiento**. Pégala, guarda y toca **Probar**.
6. Copia el **enlace para compartir el visor** que aparece ahí. Ese es el enlace de tu catálogo.

> ⚠️ Cuando actualices el programa en el futuro **no crees otra implementación**: usa **Gestionar implementaciones → ✏️ → Versión: Nueva versión**. Así el enlace no cambia.

### 3. Tu equipo
1. **💠 CATÁLOGO → 🗂️ Configurar el catálogo → 👥 Equipo**: agrega el correo de cada persona como *Capturista* o *Administrador* (hasta 10 con AppSheet gratuito).
2. Si usas AppSheet: abre tu app → **Share** (compartir) → escribe los mismos correos.

### 4. Si usas AppSheet
1. En tu app, copia el código que aparece en la dirección del editor, después de `appId=`.
2. Pégalo en **💠 CATÁLOGO → 🗂️ Configurar el catálogo → 🛠️ Datos del catálogo y contacto → Código de tu app de AppSheet**.
3. Instala la app en los teléfonos del equipo desde el enlace que da **Share**.

### 5. Comprueba que todo funciona
- [ ] Agrega una palabra (en la app o en ✏️ Editar palabras) y márcala como **Publicado**.
- [ ] **📡 Mostrar los cambios en el visor**.
- [ ] Abre el enlace de tu visor: la palabra aparece.
- [ ] El panel ya no muestra avisos amarillos.

---

## Qué se copia y qué no

| | Se copia | No se copia (lo hace cada catálogo) |
|---|---|---|
| Hoja | Estructura, lenguas y atajos que traiga la plantilla | Palabras, voces, permisos y equipo del original |
| Identidad | — | Código del catálogo, clave, carpetas, directorio de voces: el programa crea los tuyos |
| Programa | Normalmente viaja con la hoja | La **dirección** (implementación) y la tarea automática de cada hora |
| AppSheet | La app (camino A) | Las personas con acceso: se comparten de nuevo |

---

## Problemas frecuentes

> **AppSheet gratuito:** usa la app en modo prototipo, con hasta 10 personas, y **no pulses «Deploy»** (pide un plan de pago). Más detalles en [APPSHEET.md](APPSHEET.md#appsheet-gratuito-lo-que-hay-que-saber).

| Qué pasa | Qué hacer |
|---|---|
| No aparece el menú 💠 CATÁLOGO | Recarga la hoja. Si sigue sin aparecer, haz el paso 2 del camino C. |
| «Google no verificó esta app» | Es normal: **Configuración avanzada → Ir a … (no seguro) → Permitir**. |
| El visor dice «La dirección no respondió como un catálogo» | En **Quién tiene acceso** debe decir **Cualquier persona**; copia la dirección de la implementación **activa** (*Gestionar implementaciones*). |
| En la app no puedo editar | Tu correo debe estar activo en **🗂️ Configurar el catálogo → 👥 Equipo** y la app compartida contigo en **Share**. |
| Veo personas del equipo de otro catálogo | Venían de la copia y quedaron **inactivas**. Bórralas o actívalas en **👥 Equipo**. |
| AppSheet dice «User signin not allowed with the FREE plan» | Se intentó desplegar la app o se pasó de 10 personas. Deja la app como prototipo y el equipo en 10 o menos. |
| Quiero empezar otra vez | **🔧 Mantenimiento → 🧹 Limpiar o recuperar datos**: siempre hace un respaldo antes. |

---

## Para quien prepara la plantilla

Esto lo hace una sola vez quien comparte el modelo (no las comunidades).

1. En el catálogo real, con el programa actualizado, ejecuta **🧩 Revisar y reparar el catálogo**. Así la hoja guarda su ID y las copias se reconocen.
2. Termina y prueba la app de AppSheet sobre ese catálogo ([APPSHEET.md](APPSHEET.md)).
3. **Archivo → Hacer una copia** de la hoja y abre la copia. El programa la reconoce como copia y le da identidad propia en cuanto la usas.
4. En la copia: **🔧 Mantenimiento → 🧹 Limpiar o recuperar datos → Dejar en blanco**, con **«Preparar como plantilla para compartir»** marcado. Elige si conserva lenguas y atajos (útil para una plantilla de una lengua; para una plantilla general, desmárcalos).
   La hoja queda sin palabras, nombre, contacto, claves ni equipo, marcada como plantilla: no publica y no se implementa.
5. En AppSheet: copia la app (**Manage → Collaborate & Publish → Copy app**), conecta cada tabla de la copia a la hoja plantilla y comprueba que **ninguna** apunte al catálogo real.
6. Comparte la app plantilla con permiso **View definition**, y la hoja plantilla como **lector** («cualquier persona con el enlace»), para que se pueda copiar.
7. Haz la prueba completa con **otra cuenta de Google**, por cada camino (A, B y C), y anota lo que no haya quedado claro.

> Nunca implementes ni publiques la plantilla, y nunca compartas para copiar la app del catálogo real: la copia de AppSheet incluye todos sus datos.
