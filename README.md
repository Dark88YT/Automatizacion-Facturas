Proyecto de Gestión de Facturas
Este proyecto tiene como objetivo gestionar automáticamente archivos de facturas, incluyendo su conversión a PDF y organización por cliente. El código descarga las facturas de una carpeta específica, las convierte a PDF (si son DOCX), y las organiza en carpetas de clientes basándose en información contenida en los archivos.

Funcionalidades
Conversión de archivos DOCX a PDF: Si se encuentra un archivo DOCX, se convierte automáticamente a PDF utilizando Microsoft Word.

OCR para reconocer el cliente: El programa busca en el contenido de los archivos PDF y realiza un OCR sobre las imágenes (en caso de que el archivo sea una imagen escaneada) para identificar el número de cliente.

Organización de las facturas: Las facturas se mueven a una carpeta específica para cada cliente dentro de la carpeta Clientes, basándose en el número de cliente extraído del archivo.

Requisitos
Python 3.x

Bibliotecas necesarias:

comtypes (para la conversión de DOCX a PDF).

PyMuPDF (para la extracción de texto de PDFs).

Pillow (para la manipulación de imágenes).

pytesseract (para realizar OCR en las imágenes).

Puedes instalar todas las bibliotecas necesarias utilizando el siguiente comando:

bash
Copiar
Editar
pip install comtypes PyMuPDF Pillow pytesseract
Microsoft Word: Se necesita tener instalado Microsoft Word para la conversión de DOCX a PDF.

Tesseract OCR: Para el reconocimiento de texto en imágenes, se necesita Tesseract. Asegúrate de instalarlo y configurarlo correctamente. Puedes descargarlo desde aquí.

Uso
Configura las rutas: Asegúrate de modificar las rutas de las carpetas Facturas y Clientes para que apunten a las ubicaciones correctas en tu sistema.

python
Copiar
Editar
ruta_facturas = r"C:\Users\miguel-guinot\Documents\PRUEBA\Facturas"
ruta_clientes = r"C:\Users\miguel-guinot\Documents\PRUEBA\Clientes"
Asegúrate de tener Tesseract correctamente configurado:

Si no está en tu PATH, agrega la siguiente línea al principio del código (sustituyendo la ruta por la correcta en tu sistema):

python
Copiar
Editar
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
Ejecución del código:

Una vez configurado todo, puedes ejecutar el script para procesar las facturas:

bash
Copiar
Editar
python prueba.py
Resultado esperado:

Los archivos DOCX se convertirán a PDF.

Las facturas se moverán a las carpetas correspondientes de cada cliente.

Si el cliente no se encuentra en el archivo, se mostrará un mensaje de advertencia.

Limitaciones
Actualmente, la parte de descarga de correos electrónicos aún no está implementada. Una vez esté lista, el script podrá acceder a tu cuenta de correo, descargar los archivos adjuntos y procesarlos automáticamente.

Contribuciones
Si tienes sugerencias o mejoras, no dudes en hacer un pull request o abrir un issue.

