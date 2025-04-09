# Proyecto de Gestión de Facturas

Este proyecto tiene como objetivo gestionar automáticamente archivos de facturas, incluyendo su conversión a PDF, extracción de información de clientes y organización por cliente. El código descarga las facturas de una carpeta específica, las convierte a PDF (si son DOCX), y las organiza en carpetas de clientes basándose en información contenida en los archivos.

## Funcionalidades

- **Conversión de archivos DOCX a PDF**: Si se encuentra un archivo `.docx`, se convierte automáticamente a PDF utilizando Microsoft Word.
- **Reconocimiento de texto en imágenes (OCR)**: El programa realiza un OCR sobre los archivos PDF que contienen imágenes para identificar el número de cliente.
- **Organización de las facturas**: Las facturas se mueven a una carpeta específica para cada cliente dentro de la carpeta `Clientes`, basándose en el número de cliente extraído del archivo.

## Requisitos

Antes de ejecutar el proyecto, asegúrate de tener instalados los siguientes requisitos:

### Requisitos del Sistema

- **Python 3.x** (Preferentemente Python 3.6 o superior).
- **Microsoft Word**: Se necesita tener Microsoft Word instalado para la conversión de archivos `.docx` a PDF.
- **Tesseract OCR**: Para el reconocimiento de texto en imágenes, es necesario tener Tesseract instalado.

### Bibliotecas necesarias

- `comtypes` (para la conversión de `.docx` a PDF).
- `PyMuPDF` (para la extracción de texto de archivos PDF).
- `Pillow` (para la manipulación de imágenes).
- `pytesseract` (para realizar OCR en las imágenes).

Puedes instalar todas las bibliotecas necesarias utilizando el siguiente comando:

```bash
pip install comtypes PyMuPDF Pillow pytesseract
