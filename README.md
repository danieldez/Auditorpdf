# Auditor de Validaciones PDF/Excel 🛡️

Este proyecto es una herramienta de automatización diseñada para extraer, validar y auditar datos de documentos PDF (manifiestos, remesas, facturas) y cruzarlos con bases de datos en Excel. Utiliza OCR basado en coordenadas y lógica de procesamiento de lenguaje natural (regex) para garantizar la integridad de los datos en procesos industriales y contables.

## ✨ Características

- **Extracción de Datos Híbrida**: Combina extracción por coordenadas (plantillas) con escaneo global de tablas y búsqueda por expresiones regulares (Regex).
- **Entrenador de Plantillas**: Interfaz visual para mapear nuevos formatos de PDF arrastrando y soltando áreas de interés.
- **Auditoría Multi-página**: Capacidad para procesar documentos complejos donde la información está distribuida en varias páginas.
- **Panel de Control Intuitivo**: Dashboard con estadísticas en tiempo real sobre el progreso de la auditoría y detección de errores.
- **Modos Especializados**: Incluye un modo para auditoría de nómina y seguridad social.

## 🛠️ Tecnologías Utilizadas

- **Lenguaje**: Python 3.x
- **UI/UX**: `CustomTkinter` para una interfaz moderna y oscura.
- **Procesamiento PDF**: `pdfplumber` para extracción precisa de texto y tablas.
- **Gestión Excel**: `openpyxl` y `xlwings` para lectura y escritura de reportes.
- **Imágenes**: `Pillow` para el manejo de logos e iconos.

## 🚀 Instalación y Uso

1. **Clonar el repositorio**:
   ```bash
   git clone https://github.com/tu-usuario/auditor-validaciones.git
   cd auditor-validaciones
   ```

2. **Instalar dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configuración**:
   Renombra el archivo `config_auditor.json.example` a `config_auditor.json` y configura tus rutas iniciales si lo deseas (opcional, puedes hacerlo desde la UI).

4. **Ejecutar la aplicación**:
   ```bash
   python auditorPDF.py
   ```

## 📂 Estructura del Proyecto

- `auditorPDF.py`: Punto de entrada principal y lógica de la interfaz.
- `entrenador.py`: Módulo para la creación visual de plantillas OCR.
- `gestor.py`: Administrador de plantillas guardadas.
- `config.py`: Gestión de rutas y variables de entorno del sistema.
- `plantillas.json`: Base de datos de coordenadas para diferentes formatos de PDF.

## 📝 Nota sobre Privacidad

Este repositorio ha sido limpiado de datos sensibles y rutas locales. Para utilizarlo en producción, asegúrate de configurar tus propias rutas de archivos y plantillas según tus necesidades locales.

---
Desarrollado como parte de mi portafolio profesional. ¡Siéntete libre de contactarme para feedback o colaboración!
