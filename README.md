# Catálogo de lenguas originarias

**Un sistema gratuito y replicable para que cada comunidad construya, cuide y comparta el catálogo de su lengua.**

Las palabras, audios e imágenes se capturan y revisan en Google Sheets (y, si se quiere, en AppSheet), y se muestran en un **visor web** de un solo archivo. El visor funciona en el teléfono, sin instalar nada y también sin internet, e incluye búsqueda, juegos, actividades para aprender y materiales para imprimir.

> Visor diseñado y desarrollado por **Alejandro Estrada** · © 2025 · Código abierto (AGPL-3.0)

---

## Contenido

- [Principios](#principios)
- [Cómo está hecho](#cómo-está-hecho)
- [Qué ofrece](#qué-ofrece)
- [Replicar el catálogo para tu comunidad](#replicar-el-catálogo-para-tu-comunidad)
- [Actualizar el programa](#actualizar-el-programa)
- [Privacidad, consentimiento y soberanía](#privacidad-consentimiento-y-soberanía)
- [Problemas frecuentes](#problemas-frecuentes)
- [Participar](#participar)
- [Licencia](#licencia)

---

## Principios

- **Soberanía comunitaria.** Cada catálogo pertenece a su comunidad, que decide qué se publica, con qué licencia y quién participa. El programa no recibe ni guarda nada fuera de la cuenta de cada comunidad.
- **Sin lengua puente.** Ninguna lengua es «la principal» por diseño: cada palabra existe como una idea con sus formas en cada lengua, y el español es una más.
- **Accesible y gratuito.** Funciona con una cuenta gratuita de Google y GitHub Pages, en teléfonos sencillos, con poca señal o sin ella, y sin instalar aplicaciones.
- **Nadie edita tablas a mano.** Todo se hace desde menús y ventanas en lenguaje sencillo; el programa crea y repara la estructura.
- **Cada comunidad lo ajusta a su manera.** Casi todo es configurable: lenguas, alfabetos, atajos de escritura, tipos de relación, avisos, mantenimiento, permisos.
- **Nada se borra sin respaldo.** Lo retirado se conserva, y la limpieza siempre ofrece vista previa y copia de seguridad.
- **Nunca se publica lo privado.** Correos, notas internas, estados y datos de contacto no salen del catálogo.

---

## Cómo está hecho

| Pieza | Qué es | Dónde vive |
|---|---|---|
| **Hoja del catálogo** | La base de datos: lenguas, palabras, traducciones, fuentes, voces, permisos… | Google Sheets, en la cuenta de la comunidad |
| **`Codigo.gs`** | El programa del catálogo: menús, ventanas, revisión, publicación y la dirección de internet que consulta el visor | Apps Script, dentro de la misma hoja |
| **`index.html`** | El visor: un solo archivo con todo lo necesario | GitHub Pages u otro alojamiento web |
| **AppSheet** *(opcional)* | Una app para capturar desde el teléfono | La cuenta de la comunidad |

El visor solo lee lo que la hoja publica: una copia preparada de lo que está **Publicado**, sin datos internos. Nunca lee la hoja directamente.

### Hojas principales

| Hoja | Para qué |
|---|---|
| IDIOMAS | Lenguas del catálogo, con su alfabeto, sus códigos y sus atajos |
| CATEGORIAS · TRADUCCIONES_CATS | Categorías y sus nombres en cada lengua |
| PALABRAS | Cada idea o concepto, con su categoría, su imagen principal y su estado |
| TRADUCCIONES | Cómo se dice cada palabra en cada lengua: texto, audio, sílabas, variante, fuente y descripción |
| RELACIONES · TIPOS_RELACION | Vínculos entre palabras (parte de, sinónimo, variante…) y su vocabulario |
| FUENTES · ACTORES · AUTORIZACIONES · COMUNIDADES · VOCES | De dónde viene cada dato, quién lo aporta, qué permisos tiene y quién presta su voz |
| EQUIPO | Quién administra y quién captura |
| AVISOS | Avisos que aparecen arriba del visor |
| HISTORIAL | Registro de cambios |
| CONFIGURACION | Datos del catálogo (los llena el asistente) |

---

## Qué ofrece

✅ = disponible · 🛠️ = en desarrollo

### En la hoja del catálogo (menú **💠 CATÁLOGO**)

- ✅ **Primeros pasos:** un asistente que prepara todo el catálogo.
- ✅ **Editar palabras:** buscar en cualquier lengua, corregir, agregar, cambios en bloque, aviso de palabras repetidas y relaciones entre palabras.
- ✅ **Imágenes y sonidos:** asignar, revisar y trabajar con quien ilustra, con créditos y notas internas.
- ✅ **Revisar y publicar:** estados Borrador, En revisión, Publicado y Retirado; el visor solo muestra lo publicado.
- ✅ **Colaborar:** hojas para quien traduce o ilustra, traer lo que llenaron, e intercambio con otras comunidades.
- ✅ **Visor: conexión y mantenimiento:** dirección y claves, clave del equipo, mantenimiento («congelar» o «cerrar», programado o inmediato), avisos, imagen del catálogo y frecuencia de actualización.
- ✅ **Configurar:** lenguas, equipo, voces y permisos, fuentes, atajos de escritura, tipos de relación y datos del catálogo.
- ✅ **Mantenimiento:** palabras que se escriben igual, copias de trabajo, reparar la estructura, copia de seguridad, limpiar y recuperar.
- ✅ **Exportar a grafo:** N-Triples (OntoLex-Lemon y SKOS), CSV para Neo4j y Apache AGE, y GraphML.
- ✅ **Menús en varios idiomas**, con traducciones que se pueden instalar.
- 🛠️ **Descripciones en la lengua:** una descripción de cada palabra, escrita en su propia lengua por la comunidad.
- 🛠️ **📥 Propuestas:** revisar lo que proponen docentes y hablantes desde el visor, con sus créditos.
- 🛠️ **Galerías:** varias imágenes por palabra, con una principal y un límite por palabra que decide el equipo.

### En el visor

- ✅ **Explorar:** búsqueda en todas las lenguas, fichas con audio, sílabas, fuente, crédito de la imagen y palabras relacionadas.
- ✅ **Aprender:** repaso, escucha, ¿en qué lengua?, arma la palabra y dictado.
- ✅ **Jugar:** lotería (también en línea), memorama, basta, charadas, ahorcado y sopa de letras con palabras cruzadas y tres niveles.
- ✅ **Imprimir:** lotería, hojas de trabajo, etiquetas y tarjetas.
- ✅ Funciona **sin conexión**, con modo oscuro, alto contraste e interfaz en varios idiomas.
- ✅ **Avisos** y **pantalla de mantenimiento** que configura el equipo.
- 🛠️ **Recursos para docentes:**
  - hoja del estudiante y hoja del docente con respuestas, varias versiones, y un código para regenerar cada hoja;
  - pistas configurables: imagen, otra lengua, descripción del catálogo, pista personalizada o espacio en blanco;
  - crucigrama, ordena las letras, basta en papel, memorama recortable, «Pregunta en casa», niveles para grupos multigrado;
  - letras que ya conocen, sílabas, trazo, «Escucha y señala», glosario y lista de cotejo.
- 🛠️ **📂 Mis actividades:** guardadas en el dispositivo, con respaldo en archivo y sin cuentas.
- 🛠️ **Proponer** palabras, descripciones, correcciones e imágenes, con crédito y siempre revisado por el equipo; el correo como alternativa.
- 🛠️ **🤝 Dónde ayudar:** palabras que todavía necesitan imagen o descripción.

---

## Replicar el catálogo para tu comunidad

Necesitas una cuenta de Google y, para publicar el visor, una cuenta de GitHub (o cualquier alojamiento web).

### 1. La hoja del catálogo
1. Crea una hoja de cálculo nueva en Google Sheets, o haz una copia de una plantilla.
2. Abre **Extensiones → Apps Script**, borra lo que haya, pega el contenido de `Codigo.gs` y guarda.
3. Vuelve a la hoja y recárgala. Aparecerá el menú **💠 CATÁLOGO**.
4. Elige **🚀 Primeros pasos** y sigue el asistente: nombre del catálogo, lenguas y categorías. La primera vez, Google pedirá autorizar el programa.

### 2. La dirección del catálogo (una sola vez)
1. En Apps Script: **Implementar → Nueva implementación → Aplicación web**.
2. Elige **Ejecutar como: Yo** y **Quién tiene acceso: Cualquier persona**, e implementa.
3. Copia la dirección que termina en `/exec`.
4. En la hoja: **💠 CATÁLOGO → 🤝 Colaborar → 🔗 Visor: conexión y mantenimiento**, pégala, guarda y pulsa **Probar**.

> La dirección es pública por diseño: solo entrega lo que está publicado. Si quieres limitar quién ve el visor, crea una **clave** en la misma ventana.

### 3. El visor
1. Sube `index.html` a un repositorio de GitHub y activa **GitHub Pages** en *Settings → Pages*, o súbelo a cualquier alojamiento.
2. Abre el visor y escribe la dirección del catálogo (y la clave, si la hay), o comparte el enlace directo que genera la ventana **Conexión y claves**.

Un mismo visor sirve para cualquier catálogo: no hay que modificarlo para cada comunidad.

### 4. AppSheet (opcional)
Crea una app desde la hoja y, cada vez que el programa agregue hojas o columnas, usa **Regenerate structure** en AppSheet.

---

## Actualizar el programa

1. Pega el `Codigo.gs` nuevo en Apps Script y guarda.
2. **Implementar → Gestionar implementaciones → ✏️ → Versión: Nueva → Implementar.**
   ⚠️ **No crees una «Nueva implementación»**: cambiaría la dirección y el visor dejaría de encontrar el catálogo.
3. Ejecuta **🔧 Mantenimiento → 🧩 Revisar y reparar el catálogo** para crear las hojas o columnas nuevas.
4. Reemplaza `index.html` en tu alojamiento.

Si necesitas hacer cambios grandes sin que el público vea el catálogo a medias, usa **mantenimiento**:
- **Congelar:** el público sigue viendo la versión anterior mientras el equipo trabaja.
- **Cerrar:** el público ve una pantalla de mantenimiento con tu mensaje y tus datos de contacto.

---

## Privacidad, consentimiento y soberanía

- **Solo sale lo publicado.** Lo que está en borrador, en revisión o retirado, los correos, las notas internas y los datos de contacto nunca llegan al visor.
- **Voces con consentimiento.** Un audio se publica solo con consentimiento público, y el nombre de quien habla solo si lo autorizó.
- **Fuentes con permisos.** Cada texto respeta lo que su fuente permite: publicar, adaptar, usar en juegos y dar crédito.
- **Lo que la comunidad restringe** (por ejemplo, palabras ceremoniales) puede quedar fuera de juegos e impresos.
- **Sin datos de menores.** Ni el visor ni los materiales piden datos de estudiantes. Las actividades de cada docente se guardan solo en su dispositivo.
- **Todo lo que llega desde el visor pasa por revisión** del equipo antes de entrar al catálogo.
- **Los contenidos son de la comunidad.** Cada catálogo elige su propia licencia para palabras, audios e imágenes; la licencia del programa no se aplica a ellos.

---

## Problemas frecuentes

| Síntoma | Causa probable | Solución |
|---|---|---|
| «La dirección no respondió como un catálogo» | La dirección es de una implementación vieja o archivada | Copia la dirección de la implementación **activa** en *Gestionar implementaciones* y pégala en **Conexión y claves** |
| Los cambios no aparecen en el visor | No se ha publicado, o el navegador conserva la versión anterior | **📡 Mostrar los cambios en el visor**; si sigue igual, **🔃 El visor no muestra los cambios** |
| El visor pide una clave | El catálogo tiene clave | Usa el enlace que genera **Conexión y claves**, que ya la incluye |
| Al importar, aparece una lengua duplicada | La lengua tiene otro código internacional en el archivo de origen | Revisa el **código internacional** (ISO 639-3) en **🗣️ Lenguas** antes de importar |
| Algo dejó de funcionar después de actualizar | Faltan hojas o columnas nuevas | **🧩 Revisar y reparar el catálogo**; si usas AppSheet, *Regenerate structure* |

El manual completo, paso a paso, está en preparación.

---

## Participar

- **¿Tu comunidad o institución usa este visor?** Nos ayuda mucho saberlo, de forma voluntaria y sin ningún rastreo: [cómo avisarnos].
- **Errores y sugerencias:** abre un *issue* en este repositorio.
- **Contacto:** [correo del proyecto].

---

## Licencia

© 2025 Alejandro Estrada. Visor diseñado y desarrollado por Alejandro Estrada.

El **programa** (el visor `index.html` y `Codigo.gs`) es software libre, bajo la licencia **GNU AGPL-3.0**, con términos adicionales que piden conservar el crédito de autoría visible e indicar cuándo una versión fue modificada. Consulta [LICENSE](LICENSE) y [NOTICE.md](NOTICE.md).

En pocas palabras, puedes usarlo, copiarlo, adaptarlo e instalarlo para tu comunidad o institución sin costo. Si lo modificas y lo ofreces por internet, comparte tu código con la misma licencia y conserva el crédito «Visor diseñado y desarrollado por Alejandro Estrada» en el pie de página o en «Créditos» o «Acerca de».

Los **contenidos de cada catálogo** (palabras, audios, imágenes, descripciones) pertenecen a sus comunidades y se rigen por la licencia que cada catálogo elija. La licencia del programa no se aplica a ellos.

Las versiones anteriores al [fecha del cambio] se publicaron bajo la licencia MIT.
