import json
import os
from typing import List
from xml.dom import minidom
import openpyxl
import xml.etree.ElementTree as ET
from dataclasses import dataclass
import requests

from logger import __logger
from dgii_client import DGIICFService
from rfce import RFCEClient
from ecf31 import ECF31
from ecf32 import ECF32
from ecf33 import ECF33
from ecf34 import ECF34
from ecf41 import ECF41
from ecf43 import ECF43
from ecf44 import ECF44
from ecf45 import ECF45
from ecf47 import ECF47

workbook = openpyxl.load_workbook("test_data.xlsx")
sheet = workbook["ECF"]

@dataclass
class Param:
    RNCEmisor: str = ""
    ENCF: str = ""
    CodigoSeguridadeCF: str = ""

param = Param()
invoice_name = ""

def generate_dummy_dgii_xml():
    """Generate fully compliant DGII dummy XML for immediate client demo"""
    root = ET.Element('ECF', {
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'xmlns:ds': 'http://www.w3.org/2000/09/xmldsig#',
        'xsi:schemaLocation': 'https://ecf.dgii.gov.do/esquemas/ecf/1.1'
    })

    # 1. Encabezado (Header) - Correct structure
    encabezado = ET.SubElement(root, 'Encabezado')
    
    # Required elements
    ET.SubElement(encabezado, 'Version').text = '1.1'
    
    id_doc = ET.SubElement(encabezado, 'IdDoc')
    ET.SubElement(id_doc, 'TipoeCF').text = "32"
    ET.SubElement(id_doc, 'eNCF').text = "E320000000011"
    ET.SubElement(id_doc, 'IndicadorMontoGravado').text = '1'  # 1=Yes, 0=No
    ET.SubElement(id_doc, 'FechaLimitePago').text = '2025-08-30'  # Optional
    # ET.SubElement(id_doc, 'TipoPago').text = ''  # "Contado" or "Crédito"
    ET.SubElement(id_doc, 'TerminoPago').text = ''  # "Contado" or "Crédito"
    TablaFormasPago = ET.SubElement(id_doc, 'TablaFormasPago')  # "Contado" or "Crédito"
    FormaDePago = ET.SubElement(TablaFormasPago, 'FormaDePago')
    ET.SubElement(FormaDePago, 'FormaPago').text = '01'  # Cash
    ET.SubElement(FormaDePago, 'MontoPago').text = '1180.00'  # Total amount paid    

    FormaDePago = ET.SubElement(TablaFormasPago, 'FormaDePago')
    ET.SubElement(FormaDePago, 'FormaPago').text = ''  # Cash
    ET.SubElement(FormaDePago, 'MontoPago').text = ''  # Total amount paid    

    
    emisor = ET.SubElement(encabezado, 'Emisor')
    ET.SubElement(emisor, 'RNCEmisor').text = '132641566'
    ET.SubElement(emisor, 'RazonSocialEmisor').text = 'EMPRESA DEMO SRL'
    ET.SubElement(emisor, 'NombreComercial').text = 'DEMO COMERCIAL'
    
    # Required address elements
    ET.SubElement(emisor, 'DireccionEmisor').text = 'Calle Principal #123'
    ET.SubElement(emisor, 'Municipio').text = 'Distrito Nacional'
    ET.SubElement(emisor, 'Provincia').text = 'Santo Domingo'
    
    # Contact information
    telefono = ET.SubElement(emisor, 'TablaTelefonoEmisor')
    
    # First phone number (landline)
    # telefono1 = ET.SubElement(telefono, 'TelefonoEmisor')
    # ET.SubElement(telefono1, 'NumeroTelefono').text = '8095551234'  # No hyphens
    # ET.SubElement(telefono1, 'TipoTelefono').text = '1'  # 1 for landline

    ET.SubElement(telefono, 'TelefonoEmisor').text = '8495555678'  # No hyphens

    ET.SubElement(emisor, 'CorreoEmisor').text = 'info@empresademo.com'
    ET.SubElement(emisor, 'WebSite').text = 'www.empresademo.com'
    
    # Economic information
    ET.SubElement(emisor, 'ActividadEconomica').text = 'VENTA AL POR MENOR'
    ET.SubElement(emisor, 'CodigoVendedor').text = 'VEN001'
    
    # Optional but recommended
    ET.SubElement(emisor, 'NumeroFacturaInterna').text = 'FAC-1001'
    ET.SubElement(emisor, 'NumeroPedidoInterno').text = 'PED-5001'
    ET.SubElement(emisor, 'ZonaVenta').text = 'ZONA 1'
    ET.SubElement(emisor, 'RutaVenta').text = 'RUTA A'
    
    # Additional info
    info_adicional = ET.SubElement(emisor, 'InformacionAdicionalEmisor')
    ET.SubElement(info_adicional, 'InformacionAdicional', {
        'nombre': 'Sucursal',
        'texto': 'Principal'
    })

    ET.SubElement(emisor, 'FechaEmision').text = "10-10-2020"
    
    comprador = ET.SubElement(encabezado, 'Comprador')
    ET.SubElement(comprador, 'RNCComprador').text = '987654321'
    ET.SubElement(comprador, 'RazonSocialComprador').text = 'CLIENTE DEMOSTRACION'

    # Critical fix: Correct order of elements
    # info_adicional = ET.SubElement(encabezado, 'InformacionesAdicionales')
    # ET.SubElement(info_adicional, 'InformacionAdicional', {
    #     'nombre': 'Observaciones',
    #     'texto': 'DEMO PARA CLIENTE - NO ES FACTURA REAL'
    # })
    
    ET.SubElement(encabezado, 'Transporte')  # Required empty element

    if not encabezado.find('Totales'):  # Only add if not present
        totales = ET.SubElement(encabezado, 'Totales')
        ET.SubElement(totales, 'MontoTotal').text = '1180.00'  # Required
        ET.SubElement(totales, 'ValorPagar').text = '1180.00'  # Required
        ET.SubElement(totales, 'TotalITBISRetenido').text = '0.00'  # Required for B2B
        ET.SubElement(totales, 'TotalISRRetencion').text = '0.00'  # Required for services
        ET.SubElement(totales, 'TotalITBISPercepcion').text = '0.00'  # Optional

    # 2. Detalles (Line Items)
    detalles = ET.SubElement(root, 'Detalles')
    
    # Sample product 1
    item1 = ET.SubElement(detalles, 'Item')
    ET.SubElement(item1, 'Descripcion').text = 'SERVICIO DE DEMOSTRACION'
    ET.SubElement(item1, 'Cantidad').text = '1'
    ET.SubElement(item1, 'PrecioUnitario').text = '1000.00'
    ET.SubElement(item1, 'ITBIS').text = '180.00'  # 18%
    ET.SubElement(item1, 'MontoItem').text = '1180.00'

    # 3. Totales (Must be direct child of ECF)
    totales = ET.SubElement(root, 'Totales')
    ET.SubElement(totales, 'MontoTotal').text = '1180.00'
    ET.SubElement(totales, 'ITBISTotal').text = '180.00'

    rough_string = ET.tostring(root, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    pretty_xml_as_str = reparsed.toprettyxml(indent="  ", encoding='utf-8')

    path = os.path.join(os.path.dirname(__file__), 'data/row_invoice.xml')
    
    with open(path, 'wb') as f:
        f.write(pretty_xml_as_str)

    # Return formatted XML
    xml_str = ET.tostring(root, encoding='utf-8').decode('utf-8')
    return xml_str

__trackId : str
__token : str

def main():

    try:
        dgii_service = DGIICFService(dgii_env='prod', company_id=1)

        # Step 1: Get the seed
        semilla = dgii_service.get_semilla()
        if not semilla:
            raise  __logger.error("Failed to retrieve semilla from DGII.")

        __logger.info(f"Received semilla: {semilla}")

        # Step 2: Sign the semilla
        signed_semilla = dgii_service.sign_semilla(semilla)

        # __logger.info(f"Signed semilla: {signed_semilla}")

        # Step 3: Validate the signed semilla and get the token
        token = dgii_service.validate_semilla()
        if not token:
            raise __logger.error("Failed to validate semilla with DGII and obtain token.")

        __logger.info(f"Received token: {token}")

        """ filetest for debuging"""
        # xml_file_path = os.path.join(os.path.dirname(__file__), 'data/row_invoice.xml')

        # with open(xml_file_path, 'r', encoding='utf-8') as file:
        #     xml_str = file.read()
        # Print or use the XML string
        # print(xml_str)

        # Step 4: Generate and sign e-CF XML     
        # xml_str = generate_dummy_dgii_xml()  
        # __logger.info(f"invoice xml: {xml_str}")

        #################################
        ###########   RFCE   ############ 
        #################################

        # rfce_client = RFCEClient()
        # xml_str = rfce_client.create_rfce_xml(5)

        # invoice_name = rfce_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")

        
        ##############################################
        ###########        E-CF-31        ############
        ###########   row : 7, 8, 9, 10   ############
        ##############################################
        # e_cf_31_client = ECF31()
        # xml_str = e_cf_31_client.create_e_cf_31(7)

        # invoice_name = e_cf_31_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")

        
        ##############################################
        ######        E-CF-32                  #######
        ######   row : 11, 12, 13, 14, 15, 3   #######
        ##############################################
        # e_cf_32_client = ECF32()
        # xml_str = e_cf_32_client.create_e_cf_32(3)

        # invoice_name = e_cf_32_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")

        
        ##############################################
        ######           E-CF-33               #######
        ######           row : 2               #######
        ##############################################
        # e_cf_33_client = ECF33()
        # xml_str = e_cf_33_client.create_e_cf_33(2)

        # invoice_name = e_cf_33_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")

                
        ##############################################
        ######           E-CF-34               #######
        ######           row : 4, 5            #######
        ##############################################
        # e_cf_34_client = ECF34()
        # xml_str = e_cf_34_client.create_e_cf_34(4)

        # invoice_name = e_cf_34_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")
        
                
        ##############################################
        ######           E-CF-41               #######
        ######           row : 6, 16           #######
        ##############################################
        # e_cf_41_client = ECF41()
        # xml_str = e_cf_41_client.create_e_cf_41(6)

        # invoice_name = e_cf_41_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")

        ##############################################
        ######           E-CF-43               #######
        ######           row : 17, 18          #######
        ##############################################
        # e_cf_43_client = ECF43()
        # xml_str = e_cf_43_client.create_e_cf_43(17)

        # invoice_name = e_cf_43_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")


        ##############################################
        ######           E-CF-44               #######
        ######           row : 19, 20          #######
        ##############################################
        # e_cf_44_client = ECF44()
        # xml_str = e_cf_44_client.create_e_cf_44(19)

        # invoice_name = e_cf_44_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")


        ##############################################
        ######           E-CF-45               #######
        ######           row : 21, 22          #######
        ##############################################
        # e_cf_45_client = ECF45()
        # xml_str = e_cf_45_client.create_e_cf_45(19)

        # invoice_name = e_cf_45_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")


        ##############################################
        ######           E-CF-46               #######
        ######           row : 23, 24          #######
        ##############################################
        # e_cf_46_client = ECF46()
        # xml_str = e_cf_46_client.create_e_cf_46(23)

        # invoice_name = e_cf_46_client.invoice_name
        # __logger.info(f"Invoice Name: {invoice_name}")


        ##############################################
        ######           E-CF-47               #######
        ######           row : 25, 26          #######
        ##############################################
        e_cf_47_client = ECF47()
        xml_str = e_cf_47_client.create_e_cf_47(25)

        invoice_name = e_cf_47_client.invoice_name
        __logger.info(f"Invoice Name: {invoice_name}")


        signed_xml = dgii_service.sign_xml(xml_str, invoice_name)
        # __logger.info(f"signed invoice xml: {signed_xml}")

        # Step 5: Submit the signed e-CF XML to DGII using the token
        # response = dgii_service.submit_rfce(token)
        response = dgii_service.submit_ecf(token)
        trackId = ""
        response_data = json.loads(response.content.decode('utf-8'))

        __logger.info(f"Submitting Response:  {response_data}")

        # Log the response and update the invoice status
        if response.status_code == 200:
            response_data = json.loads(response.content.decode('utf-8'))
            __logger.info(f"Submitted to DGII, response: trackId:  {response_data['trackId']}")
            trackId = response_data['trackId']
        else:
            __logger.error(f"Error submitting invoice to DGII: {response.text}")
            return

        # Step 6: Track the status of e-CF to DGII using the token
        response = dgii_service.track_ecf(trackId, token)
        # response = dgii_service.track_rfce(token, param)
        __logger.info(f"track result: {response.content}")

        response_data = json.loads(response.content.decode('utf-8'))
        __logger.info(f"dgii_track_id {trackId}, dgii_token : {token}")

    except Exception as e:
        __logger.error(f"Error processing invoice: {str(e)}")

if __name__ == "__main__":
    main()