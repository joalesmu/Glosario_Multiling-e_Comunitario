# 🌐 Sistema de Gestión Lingüística

![Version](https://img.shields.io/badge/Versión-1.0-blue)
![Platform](https://img.shields.io/badge/Plataforma-Google_Workspace_%7C_AppSheet-green)
![License](https://img.shields.io/badge/Licencia-CECAN-orange)

Un sistema integral y replicable para la gestión, auditoría y recopilación de datos lingüísticos y multimedia. Diseñado para facilitar el trabajo colaborativo entre administradores, hablantes de lenguas originarias y diseñadores, utilizando **Google Sheets** como base de datos y **AppSheet** como interfaz de captura móvil.

---

## ✨ Características Principales

* **📊 Mesa de Trabajo y Auditoría Inteligente:** Detecta automáticamente traducciones, imágenes o audios faltantes cruzando los datos de la hoja con los archivos físicos en Google Drive (búsqueda *Fuzzy*).
* **📱 Integración Dinámica con AppSheet:** Genera *Deep Links* (hipervínculos dinámicos) en el panel de administración que abren directamente la aplicación móvil en el registro exacto para grabar audios o subir imágenes desde el celular.
* **📄 Plantillas "Conscientes" (ETL):** Genera hojas de cálculo temporales para los colaboradores. El sistema sabe qué falta y solo pide los huecos vacíos, reconciliando la información de regreso mediante un ID único para evitar duplicados.
* **🎨 Exportación para Diseño:** Crea listas limpias en un clic con los conceptos que requieren ilustración.
* **🌍 Arquitectura Replicable:** El código se adapta automáticamente a la configuración regional del usuario (uso de comas o puntos y comas) y vincula aplicaciones de AppSheet con solo cambiar un ID en la configuración.

---

## 🛠️ Requisitos del Sistema

Para desplegar tu propia instancia de este glosario, necesitas:
1. Una cuenta de **Google Workspace** (o Gmail gratuito).
2. Permisos para crear y editar **Google Apps Script**.
3. Una aplicación base generada en **Google AppSheet**.
4. Carpetas creadas en **Google Drive** para alojar Audios e Imágenes.

---

## 🚀 Guía de Instalación y Replicación

Sigue estos pasos para clonar el proyecto y configurarlo en tu entorno:

### Paso 1: Configurar el Entorno en Drive
1. Crea una carpeta principal en tu Google Drive.
2. Dentro, crea dos subcarpetas: una para **Audios** y otra para **Imágenes**.
3. Extrae el `ID` de ambas carpetas (la cadena de texto en la URL después de `folders/`).

### Paso 2: Base de Datos y AppSheet
1. Crea un nuevo Google Sheet y pega el código de `Código.gs` en **Extensiones > Apps Script**.
2. En la hoja de cálculo, crea una pestaña llamada `CONFIGURACION`.
3. Crea tu aplicación en AppSheet vinculada a este Google Sheet. Obtén tu **App ID** (visible en la URL del editor de AppSheet).

### Paso 3: Hoja de Configuración
Asegúrate de que tu pestaña `CONFIGURACION` tenga la siguiente estructura en las primeras dos filas:

| NOMBRE_PROYECTO | IDIOMA_ACTIVO | APPSHEET_APP_ID | ID_CARPETA_AUDIOS | ID_CARPETA_IMAGENES |
| :--- | :--- | :--- | :--- | :--- |
| Mi Glosario | Español | *[Tu-App-ID]* | *[ID-Carpeta-Audios]* | *[ID-Carpeta-Imágenes]* |

### Paso 4: Inicialización
1. En tu Google Sheet, recarga la página. Aparecerá un menú personalizado llamado **💠 ADMINISTRACIÓN GLOSARIO**.
2. Ve a **Mantenimiento > 🚀 Instalación de Carpetas y Hojas**. El script construirá automáticamente el resto de la base de datos relacional.
3. Ve a **Mantenimiento > 🆔 Reparar IDs Faltantes** para asegurar la integridad de la base de datos.

¡Listo! El sistema está operativo.

---

## ⚙️ Flujo de Trabajo Recomendado

1. **Gestión de Faltantes:** Usa el *Centro de Auditoría > Mesa de Trabajo* para ver qué audios o textos faltan.
2. **Grabación Móvil:** Haz clic en los hipervínculos dinámicos del Dashboard para abrir AppSheet en tu teléfono y grabar el audio directamente.
3. **Trabajo Asíncrono:** Genera una *Plantilla para Colaborador*, compártela con un traductor y, cuando termine, usa la *Importación Inteligente* para fusionar los datos sin crear duplicados.

---

## 📝 Créditos y Licencia

* **Autor:** Alejandro Estrada
* **Año:** 2026

*Creado para la preservación, documentación y revitalización de las lenguas comunitarias.*