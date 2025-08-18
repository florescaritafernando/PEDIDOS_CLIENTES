import os
import zipfile
import pandas as pd
from lxml import etree
import tempfile
from io import BytesIO

def xml_to_excel(xml_content):
    """Convierte contenido XML a DataFrame de pandas"""
    try:
        # Parsear XML
        root = etree.fromstring(xml_content)
        
        # Extraer datos (esto depende de tu estructura XML)
        data = []
        for element in root.xpath('//*[not(*)]'):  # Selecciona elementos hoja
            item = {
                'tag': element.tag,
                'text': element.text,
                **element.attrib
            }   
            data.append(item)
        
        return pd.DataFrame(data)
    
    except Exception as e:
        print(f"Error procesando XML: {str(e)}")
        return pd.DataFrame()

def procesar_zip_xml(input_zip_path, output_zip_path):
    """Procesa archivos XML en ZIP y genera ZIP con Excels"""
    with zipfile.ZipFile(input_zip_path, 'r') as zip_input:
        with zipfile.ZipFile(output_zip_path, 'w') as zip_output:
            for xml_file in zip_input.namelist():
                if xml_file.lower().endswith('.xml'):
                    try:
                        # Leer XML
                        with zip_input.open(xml_file) as f:
                            xml_content = f.read()
                        
                        # Convertir a Excel
                        df = xml_to_excel(xml_content)
                        
                        if not df.empty:
                            # Crear Excel en memoria
                            excel_buffer = BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                df.to_excel(writer, index=False, sheet_name='Datos')
                            
                            # Guardar en ZIP de salida
                            excel_name = os.path.splitext(xml_file)[0] + '.xlsx'
                            zip_output.writestr(excel_name, excel_buffer.getvalue())
                            print(f"Convertido: {xml_file} -> {excel_name}")
                        else:
                            print(f"Archivo {xml_file} vacío o con errores")
                    
                    except Exception as e:
                        print(f"Error procesando {xml_file}: {str(e)}")

if __name__ == "__main__":
    # Configuración
    INPUT_ZIP = 'comprobantes_XML_2025-08-18.zip'  # Cambia por tu archivo ZIP de entrada
    OUTPUT_ZIP = 'excels_resultado.zip'
    
    # Procesar
    print(f"Iniciando conversión de {INPUT_ZIP}...")
    procesar_zip_xml(INPUT_ZIP, OUTPUT_ZIP)
    print(f"\n✅ Conversión completada. Resultados en {OUTPUT_ZIP}")