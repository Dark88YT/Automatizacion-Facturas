import os
import comtypes.client
import re
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
from io import BytesIO


# Especifica la ruta a Tesseract si no está en el PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\miguel-guinot\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

# El resto de tu código que usa pytesseract o cualquier otra librería

# Función para convertir un archivo DOCX a PDF
def docx_to_pdf(docx_path):
    try:
        # Crear una instancia de Word
        word = comtypes.client.CreateObject('Word.Application')
        
        # Abrir el documento .docx
        doc = word.Documents.Open(docx_path)
        
        # Guardar el documento como PDF (el archivo se guardará con el mismo nombre pero con extensión .pdf)
        pdf_path = docx_path.replace('.docx', '.pdf')
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 es el formato PDF
        
        # Cerrar el documento y Word
        doc.Close()
        word.Quit()
        
        print(f"[OK] {docx_path} convertido a PDF como {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"[ERROR] No se pudo convertir {docx_path} a PDF. Error: {e}")
        return None

# Función para extraer texto de un archivo PDF
def extraer_texto_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        texto = ""
        for pagina in doc:
            texto += pagina.get_text()
        return texto
    except Exception as e:
        print(f"[ERROR] No se pudo extraer texto del PDF: {pdf_path}. Error: {e}")
        return ""

# Función para realizar OCR en una imagen
def ocr_en_imagen(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        texto_imagen = ""
        for pagina in doc:
            # Extraer las imágenes de cada página
            imagenes = pagina.get_images(full=True)
            for img in imagenes:
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                
                # Convertir los bytes de la imagen a una imagen de PIL
                imagen_pil = Image.open(BytesIO(image_bytes))
                
                # Usar Tesseract OCR para extraer texto
                texto_imagen += pytesseract.image_to_string(imagen_pil)
        return texto_imagen
    except Exception as e:
        print(f"[ERROR] No se pudo hacer OCR en el PDF: {pdf_path}. Error: {e}")
        return ""

# Función para extraer número de cliente de un documento (ya sea texto o imagen)
def extraer_cliente(doc_path):
    cliente = None
    if doc_path.endswith('.docx'):
        try:
            doc = Document(doc_path)
            for p in doc.paragraphs:
                match = re.search(r'Cliente(\d+)', p.text)
                if match:
                    cliente = match.group(1)
                    break
        except Exception as e:
            print(f"[ERROR] No se pudo leer el archivo DOCX: {doc_path}. Error: {e}")
    
    elif doc_path.endswith('.pdf'):
        texto = extraer_texto_pdf(doc_path)
        if not texto.strip():  # Si no se extrajo texto, intentar con OCR
            texto = ocr_en_imagen(doc_path)
        
        match = re.search(r'Cliente(\d+)', texto)
        if match:
            cliente = match.group(1)
    
    return cliente

# Función para mover las facturas a las carpetas correspondientes
def mover_facturas():
    ruta_facturas = r"C:\Users\miguel-guinot\Documents\PRUEBA\Facturas"
    ruta_clientes = r"C:\Users\miguel-guinot\Documents\PRUEBA\Clientes"

    for archivo in os.listdir(ruta_facturas):
        if archivo.endswith('.docx') or archivo.endswith('.pdf'):
            ruta_factura = os.path.join(ruta_facturas, archivo)
            
            # Intentar extraer el número de cliente del archivo
            cliente = extraer_cliente(ruta_factura)
            
            if cliente:
                ruta_cliente = os.path.join(ruta_clientes, f'Cliente{cliente}')
                
                # Si la carpeta del cliente no existe, la creamos
                if not os.path.exists(ruta_cliente):
                    os.makedirs(ruta_cliente)

                # Convertir el archivo DOCX a PDF si es necesario
                if archivo.endswith('.docx'):
                    ruta_pdf = docx_to_pdf(ruta_factura)
                    if ruta_pdf:
                        # Mover el archivo PDF a la carpeta del cliente
                        os.rename(ruta_pdf, os.path.join(ruta_cliente, f'{archivo[:-5]}.pdf'))
                
                # Mover el archivo original (ya sea DOCX o PDF) a la carpeta del cliente
                os.rename(ruta_factura, os.path.join(ruta_cliente, archivo))
                print(f"[OK] {archivo} movido a {ruta_cliente}")
            else:
                print(f"[ADVERTENCIA] No se encontró número de cliente en el archivo: {archivo}")

# Llamar a la función principal
mover_facturas()

