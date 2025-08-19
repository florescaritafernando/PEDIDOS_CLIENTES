import pandas as pd
from lxml import etree
import os

def parse_xml_peru(xml_content):
    """Parser especializado para facturas electrónicas peruanas con direcciones correctas"""
    try:
        ns = {
            'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
            'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'
        }
        
        root = etree.fromstring(xml_content)
        
        # Función para extraer dirección en formato peruano
        def get_direccion(party):
            address = party.xpath('.//cac:RegistrationAddress', namespaces=ns)
            if not address:
                return "NO ESPECIFICADA"
            
            address = address[0]
            parts = [
                address.xpath('.//cbc:StreetName/text()', namespaces=ns)[0] if address.xpath('.//cbc:StreetName', namespaces=ns) else "",
                address.xpath('.//cac:AddressLine/cbc:Line/text()', namespaces=ns)[0] if address.xpath('.//cac:AddressLine/cbc:Line', namespaces=ns) else "",
                address.xpath('.//cbc:District/text()', namespaces=ns)[0] if address.xpath('.//cbc:District', namespaces=ns) else "",
                address.xpath('.//cbc:CityName/text()', namespaces=ns)[0] if address.xpath('.//cbc:CityName', namespaces=ns) else ""
            ]
            return " - ".join(filter(None, parts))
        
        # Datos del EMISOR
        emisor = root.xpath('//cac:AccountingSupplierParty/cac:Party', namespaces=ns)[0]
        ruc_emisor = emisor.xpath('.//cbc:ID/text()', namespaces=ns)[0]
        razon_social_emisor = emisor.xpath('.//cac:PartyLegalEntity/cbc:RegistrationName/text()', namespaces=ns)[0]
        direccion_emisor = get_direccion(emisor)
        
        # Datos del CLIENTE
        cliente = root.xpath('//cac:AccountingCustomerParty/cac:Party', namespaces=ns)[0]
        ruc_cliente = cliente.xpath('.//cbc:ID/text()', namespaces=ns)[0]
        razon_social_cliente = cliente.xpath('.//cac:PartyLegalEntity/cbc:RegistrationName/text()', namespaces=ns)[0]
        direccion_cliente = get_direccion(cliente)
        
        return {
            'RUC EMISOR': ruc_emisor,
            'RAZON SOCIAL EMISOR': razon_social_emisor,
            'DIRECCION EMISOR': direccion_emisor,
            'RUC CLIENTE': ruc_cliente,
            'RAZON SOCIAL CLIENTE': razon_social_cliente,
            'DIRECCION CLIENTE': direccion_cliente,
            'FECHA EMISION': root.xpath('//cbc:IssueDate/text()', namespaces=ns)[0],
            'MONEDA': root.xpath('//cbc:DocumentCurrencyCode/text()', namespaces=ns)[0],
            'MONTO TOTAL': root.xpath('//cbc:PayableAmount/text()', namespaces=ns)[0]
        }
        
    except Exception as e:
        print(f"Error parsing XML: {str(e)}")
        return None

def procesar_xmls(input_path, output_file='MANCHESTERTEX FACT_24-ENE-NOV1.xlsx'):
    """Procesa múltiples XML y genera consolidado"""
    datos = []
    
    if os.path.isdir(input_path):
        for filename in os.listdir(input_path):
            if filename.lower().endswith('.xml'):
                with open(os.path.join(input_path, filename), 'rb') as f:
                    xml_content = f.read()
                    if parsed := parse_xml_peru(xml_content):
                        parsed['ARCHIVO ORIGEN'] = filename
                        datos.append(parsed)
    
    elif input_path.endswith('.zip'):
        import zipfile
        with zipfile.ZipFile(input_path, 'r') as z:
            for filename in z.namelist():
                if filename.lower().endswith('.xml'):
                    with z.open(filename) as f:
                        xml_content = f.read()
                        if parsed := parse_xml_peru(xml_content):
                            parsed['ARCHIVO ORIGEN'] = filename
                            datos.append(parsed)
    
    if datos:
        df = pd.DataFrame(datos)
        column_order = [
            'ARCHIVO ORIGEN',
            'FECHA EMISION',
            'RUC EMISOR',
            'RAZON SOCIAL EMISOR',
            'DIRECCION EMISOR',
            'RUC CLIENTE',
            'RAZON SOCIAL CLIENTE',
            'DIRECCION CLIENTE',
            'MONEDA',
            'MONTO TOTAL'
        ]
        df[column_order].to_excel(output_file, index=False)
        print(f"✅ Consolidado generado: {output_file}")
        return True
    else:
        print("❌ No se encontraron datos válidos")
        return False

# Ejemplo de uso
if __name__ == "__main__":
    # Procesar directorio con XMLs
    #procesar_xmls('ruta/a/tus/xmls')
    
    # O procesar ZIP con XMLs
    procesar_xmls('MANCHESTERTEX FACT_24-ENE-NOV1.zip')