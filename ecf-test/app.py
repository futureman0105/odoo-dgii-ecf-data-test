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

@dataclass
class RFCE_Data:
    Version: str = ""
    TipoeCF: str = ""
    eNCF: str = ""
    TipoIngresos: str = ""
    TipoPago: str = ""
    FormaPago: List[str] = None
    MontoPago: List[str] = None
    RNCEmisor: str = ""
    RazonSocialEmisor: str = ""
    FechaEmision: str = ""
    RNCComprador: str = ""
    IdentificadorExtranjero: str = ""
    RazonSocialComprador: str = ""
    MontoGravadoTotal: str = ""
    MontoGravadoI1: str = ""
    MontoGravadoI2: str = ""
    MontoGravadoI3: str = ""
    MontoExento: str = ""
    TotalITBIS: str = ""
    TotalITBIS1: str = ""
    TotalITBIS2: str = ""
    TotalITBIS3: str = ""
    MontoImpuestoAdicional: str = ""
    TipoImpuesto: List[str] = None
    MontoImpuestoSelectivoConsumoEspecifico: List[str] = None
    MontoImpuestoSelectivoConsumoAdvalorem: List[str] = None
    OtrosImpuestosAdicionales: List[str] = None
    MontoTotal: str = ""
    MontoNoFacturable: str = ""
    MontoPeriodo: str = ""
    CodigoSeguridadeCF: str = ""

workbook = openpyxl.load_workbook("test_data.xlsx")
sheet = workbook["ECF"]

def read_excel_create_rfce_xml(in_row : int) :
    rfce_data = RFCE_Data(
        FormaPago=[],
        MontoPago=[],
        TipoImpuesto=[],
        MontoImpuestoSelectivoConsumoEspecifico=[],
        MontoImpuestoSelectivoConsumoAdvalorem=[],
        OtrosImpuestosAdicionales=[]
    )

    # Read Version
    search_text = "Version" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.Version = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'Version : {rfce_data.Version}')

    # Read TipoeCF
    search_text = "TipoeCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.TipoeCF = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'TipoeCF : {rfce_data.TipoeCF}')

    # Read eNCF
    search_text = "ENCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.eNCF = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'ENCF : {rfce_data.eNCF}')

    # Read TipoIngresos
    search_text = "TipoIngresos" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.TipoIngresos = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'TipoIngresos : {rfce_data.TipoIngresos}')

    # Read TipoPago
    search_text = "TipoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.TipoPago = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'Version : {rfce_data.TipoPago}')

    # Read FormaDePago
    search_text = "FormaPago[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            index = 0

            while True :
                rfce_data.FormaPago.append(str(sheet.cell(in_row, col + index).value))
                __logger.info(f'FormaPago : {str(sheet.cell(in_row, col + index).value)}')

                rfce_data.MontoPago.append(str(sheet.cell(in_row, col + index + 1).value))
                __logger.info(f'MontoPago : {str(sheet.cell(in_row, col + index + 1).value)}')

                if index >= 6 :
                    break
                index += 1

    # Read RNCEmisor
    search_text = "RNCEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.RNCEmisor = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'RNCEmisor : {rfce_data.RNCEmisor}')

    # Read RazonSocialEmisor
    search_text = "RazonSocialEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.RazonSocialEmisor = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'RazonSocialEmisor : {rfce_data.RazonSocialEmisor}')

    # Read FechaEmision
    search_text = "FechaEmision" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.FechaEmision = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'FechaEmision : {rfce_data.FechaEmision}')

    # Read RNCComprador
    search_text = "RNCComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.RNCComprador = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'RNCComprador : {rfce_data.RNCComprador}')

    # Read IdentificadorExtranjero
    search_text = "IdentificadorExtranjero" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.IdentificadorExtranjero = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'IdentificadorExtranjero : {rfce_data.IdentificadorExtranjero}')

    # Read RazonSocialComprador
    search_text = "RazonSocialComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.RazonSocialComprador = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'RazonSocialComprador : {rfce_data.RazonSocialComprador}')

    # Read MontoGravadoTotal
    search_text = "MontoGravadoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoGravadoTotal = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoGravadoTotal : {rfce_data.MontoGravadoTotal}')

    # Read MontoGravadoI1
    search_text = "MontoGravadoI1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoGravadoI1 = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoGravadoI1 : {rfce_data.MontoGravadoI1}')

    # Read MontoGravadoI2
    search_text = "MontoGravadoI2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoGravadoI2 = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoGravadoI2 : {rfce_data.MontoGravadoI2}')

    # Read MontoGravadoI3
    search_text = "MontoGravadoI3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoGravadoI3 = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoGravadoI3 : {rfce_data.MontoGravadoI3}')

    # Read MontoExento
    search_text = "MontoExento" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoExento = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoExento : {rfce_data.MontoExento}')

    # Read TotalITBIS
    search_text = "TotalITBIS"
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.TotalITBIS = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'TotalITBIS : {rfce_data.TotalITBIS}')

    # Read TotalITBIS1
    search_text = "TotalITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.TotalITBIS1 = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'TotalITBIS1 : {rfce_data.TotalITBIS1}')

    # Read TotalITBIS2
    search_text = "TotalITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.TotalITBIS2 = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'TotalITBIS2 : {rfce_data.TotalITBIS2}')

    # Read TotalITBIS3
    search_text = "TotalITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.TotalITBIS3 = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'TotalITBIS3 : {rfce_data.TotalITBIS3}')

    # Read MontoImpuestoAdicional
    search_text = "MontoImpuestoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoImpuestoAdicional = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoImpuestoAdicional : {rfce_data.MontoImpuestoAdicional}')
            
    # Read ImpuestoAdicional
    search_text = "TipoImpuesto[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:

            index = 0;
            while True:
                cell_value = str(sheet.cell(in_row, col + index).value)
                rfce_data.TipoImpuesto.append(cell_value)
                __logger.info(f'TipoImpuesto : {cell_value}')

                cell_value = str(sheet.cell(in_row, col + index + 1).value)
                rfce_data.MontoImpuestoSelectivoConsumoEspecifico.append(cell_value)
                __logger.info(f'MontoImpuestoSelectivoConsumoEspecifico : {cell_value}')

                cell_value = str(sheet.cell(in_row, col + index + 2).value)
                rfce_data.MontoImpuestoSelectivoConsumoAdvalorem.append(cell_value)
                __logger.info(f'MontoImpuestoSelectivoConsumoAdvalorem : {cell_value}')

                cell_value = str(sheet.cell(in_row, col + index + 3).value)
                rfce_data.OtrosImpuestosAdicionales.append(cell_value)
                __logger.info(f'OtrosImpuestosAdicionales : {cell_value}')
                if index > 3 : 
                    break
                index += 1

    # Read MontoTotal
    search_text = "MontoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoTotal = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoTotal : {rfce_data.MontoTotal}')
  
    # Read MontoNoFacturable
    search_text = "MontoNoFacturable" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoNoFacturable = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoNoFacturable : {rfce_data.MontoNoFacturable}')
            
    # Read MontoPeriodo
    search_text = "MontoPeriodo" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.MontoPeriodo = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'MontoPeriodo : {rfce_data.MontoPeriodo}')

    # Read CodigoSeguridadeCF
    search_text = "CodigoSeguridadeCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            rfce_data.CodigoSeguridadeCF = str(sheet.cell(in_row, column=col).value)
            __logger.info(f'CodigoSeguridadeCF : {rfce_data.CodigoSeguridadeCF}')

    return _generate_rfce_xml(rfce_data)

def _generate_rfce_xml(rfce_data: RFCE_Data):
    """Generate DGII-compliant XML from an Odoo invoice."""
    # Create root element with namespaces
    root = ET.Element('RFCE', {
            'xmlns:xs': 'http://www.w3.org/2001/XMLSchema'
        })

    # 1. Encabezado
    encabezado = ET.SubElement(root, 'Encabezado')
    ET.SubElement(encabezado, 'Version').text = rfce_data.Version

    # IdDoc
    id_doc = ET.SubElement(encabezado, 'IdDoc')

    __logger.info(f'eNCF :  {rfce_data.eNCF}')

    ET.SubElement(id_doc, 'TipoeCF').text = rfce_data.TipoeCF
    ET.SubElement(id_doc, 'eNCF').text = rfce_data.eNCF  # must be filled from NCF sequence
    ET.SubElement(id_doc, 'TipoIngresos').text = rfce_data.TipoIngresos
    ET.SubElement(id_doc, 'TipoPago').text = rfce_data.TipoPago
    
    # TablaFormasPago = ET.SubElement(id_doc, 'TablaFormasPago')

    # for i in range(7):  
    #     FormaDePago = ET.SubElement(TablaFormasPago, 'FormaDePago')
    #     ET.SubElement(FormaDePago, 'FormaPago').text = rfce_data.FormaPago[i]
    #     ET.SubElement(FormaDePago, 'MontoPago').text = rfce_data.MontoPago[i]
       
    emisor = ET.SubElement(encabezado, 'Emisor')
    ET.SubElement(emisor, 'RNCEmisor').text = rfce_data.RNCEmisor or ''
    ET.SubElement(emisor, 'RazonSocialEmisor').text = rfce_data.RazonSocialEmisor or ''
    ET.SubElement(emisor, 'FechaEmision').text = rfce_data.FechaEmision or ''
    
    # Comprador (Customer)
    comprador = ET.SubElement(encabezado, 'Comprador')
    ET.SubElement(comprador, 'RNCComprador').text = rfce_data.RNCComprador or " "
    
    if rfce_data.IdentificadorExtranjero == "#e" : 
       ET.SubElement(comprador, 'IdentificadorExtranjero').text = " "
    else : 
       ET.SubElement(comprador, 'IdentificadorExtranjero').text = rfce_data.IdentificadorExtranjero

    ET.SubElement(comprador, 'RazonSocialComprador').text = rfce_data.RazonSocialComprador or ""

    # 3. Totales (must be direct child of ECF)
    totales_root = ET.SubElement(encabezado, 'Totales')
    ET.SubElement(totales_root, 'MontoGravadoTotal').text = "%.2f" %  float(rfce_data.MontoGravadoTotal)
    ET.SubElement(totales_root, 'MontoGravadoI1').text = "%.2f" %  float(rfce_data.MontoGravadoI1)

    if rfce_data.MontoGravadoI2 == "#e" : 
       ET.SubElement(totales_root, 'MontoGravadoI2').text = "%.2f" %  float(0)
    else :
       ET.SubElement(totales_root, 'MontoGravadoI2').text = "%.2f" %  float(rfce_data.MontoGravadoI2)  
        
    if rfce_data.MontoGravadoI3 == "#e" : 
       ET.SubElement(totales_root, 'MontoGravadoI3').text = "%.2f" %  float(0)
    else :
       ET.SubElement(totales_root, 'MontoGravadoI3').text = "%.2f" %  float(rfce_data.MontoGravadoI3)  

    if rfce_data.MontoExento == "#e" : 
       ET.SubElement(totales_root, 'MontoExento').text = "%.2f" %  float(0)
    else :
       ET.SubElement(totales_root, 'MontoExento').text = "%.2f" %  float(rfce_data.MontoExento)   
        
    if rfce_data.TotalITBIS == "#e" : 
        ET.SubElement(totales_root, 'TotalITBIS').text = "%.2f" %  float(0)
    else:
        ET.SubElement(totales_root, 'TotalITBIS').text = "%.2f" %  float(rfce_data.TotalITBIS)

    if rfce_data.TotalITBIS1 == "#e" : 
        ET.SubElement(totales_root, 'TotalITBIS1').text = "%.2f" %  float(0)
    else:
        ET.SubElement(totales_root, 'TotalITBIS1').text = "%.2f" %  float(rfce_data.TotalITBIS1)

    if rfce_data.TotalITBIS2 == "#e" : 
        ET.SubElement(totales_root, 'TotalITBIS2').text = "%.2f" %  float(0)
    else:
        ET.SubElement(totales_root, 'TotalITBIS2').text = "%.2f" %  float(rfce_data.TotalITBIS2)

    if rfce_data.TotalITBIS3 == "#e" : 
        ET.SubElement(totales_root, 'TotalITBIS3').text = "%.2f" %  float(0)
    else:
        ET.SubElement(totales_root, 'TotalITBIS3').text = "%.2f" %  float(rfce_data.TotalITBIS3)
    
    # if rfce_data.MontoImpuestoAdicional == "#e" : 
    #     ET.SubElement(totales_root, 'MontoImpuestoAdicional').text = "%.2f" %  float(0)
    # else:
    #     ET.SubElement(totales_root, 'MontoImpuestoAdicional').text = "%.2f" %  float(rfce_data.MontoImpuestoAdicional)

    # ImpuestosAdicionales = ET.SubElement(totales_root, 'ImpuestosAdicionales')

    # count_ImpuestoAdicional = len(rfce_data.TipoImpuesto)

    # for i in range(count_ImpuestoAdicional) : 
    #     ImpuestoAdicional = ET.SubElement(ImpuestosAdicionales, 'ImpuestoAdicional')
            
    #     if rfce_data.TipoImpuesto[i] == "#e" : 
    #         ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = " "
    #     else:
    #         ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = rfce_data.TipoImpuesto[i]

    #     if rfce_data.MontoImpuestoSelectivoConsumoEspecifico[i] == "#e" : 
    #         ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoEspecifico').text = "%.2f" %  float(0)
    #     else:
    #         ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoEspecifico').text = "%.2f" %  float(rfce_data.MontoImpuestoSelectivoConsumoEspecifico[i])
            
    #     if rfce_data.MontoImpuestoSelectivoConsumoAdvalorem[i] == "#e" : 
    #         ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoAdvalorem').text = "%.2f" %  float(0)
    #     else:
    #         ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoAdvalorem').text = "%.2f" %  float(rfce_data.MontoImpuestoSelectivoConsumoAdvalorem[i])
             
    #     if rfce_data.OtrosImpuestosAdicionales[i] == "#e" : 
    #         ET.SubElement(ImpuestoAdicional, 'OtrosImpuestosAdicionales').text = "%.2f" %  float(0)
    #     else:
    #         ET.SubElement(ImpuestoAdicional, 'OtrosImpuestosAdicionales').text = "%.2f" %  float(rfce_data.OtrosImpuestosAdicionales[i])

    if rfce_data.MontoTotal == "#e" : 
        ET.SubElement(totales_root, 'MontoTotal').text = "%.2f" %  float(0)
    else:
        ET.SubElement(totales_root, 'MontoTotal').text = "%.2f" %  float(rfce_data.MontoTotal)
                
    if rfce_data.MontoNoFacturable == "#e" : 
        ET.SubElement(totales_root, 'MontoNoFacturable').text = "%.2f" %  float(0)
    else:
        ET.SubElement(totales_root, 'MontoNoFacturable').text = "%.2f" %  float(rfce_data.MontoNoFacturable)

    # MontoPeriodo                    
    if rfce_data.MontoPeriodo == "#e" : 
        ET.SubElement(totales_root, 'MontoPeriodo').text = "%.2f" %  float(0)
    else:
        ET.SubElement(totales_root, 'MontoPeriodo').text = "%.2f" %  float(rfce_data.MontoPeriodo)

    value = "0234"
    value = value[:6].ljust(6, "0")
    print(f"Value : {value}")

    #CodigoSeguridadeCF    
    if rfce_data.CodigoSeguridadeCF == "#e" : 
        ET.SubElement(encabezado, 'CodigoSeguridadeCF').text = "000000"
    else:      
        value = rfce_data.CodigoSeguridadeCF
        value = value[:6].ljust(6, "0")
        ET.SubElement(encabezado, 'CodigoSeguridadeCF').text = value


    # Convert to XML string
    xml_str = ET.tostring(root, encoding='utf-8').decode('utf-8')
    
    rough_string = ET.tostring(root, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    pretty_xml_as_str = reparsed.toprettyxml(indent="  ", encoding='utf-8')

    path = os.path.join(os.path.dirname(__file__), 'data/row_invoice.xml')
    
    with open(path, 'wb') as f:
        f.write(pretty_xml_as_str)

    __logger.info("Finished the creating RFCE xml")

    # Convert to lxml for signing
    return xml_str

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

def read_excel_create_efc_xml(in_row : int) :

    """Generate DGII-compliant XML from an Odoo invoice."""
    # Create root element with namespaces
    root = ET.Element('ECF', {
         'xmlns:xs' : "http://www.w3.org/2001/XMLSchema",
    })

    # 1. Encabezado
    encabezado = ET.SubElement(root, 'Encabezado')

    # Write Version
    search_text = "Version" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(encabezado, 'Version').text = value
            __logger.info(f'Version : {value}')
            __logger.info(f'Version Col Index : {col}')
            break

    # IdDoc
    id_doc = ET.SubElement(encabezado, 'IdDoc')

    # Write TipoeCF
    search_text = "TipoeCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 
            ET.SubElement(id_doc, 'TipoeCF').text = value
            __logger.info(f'TipoeCF : {cell_value}')
            break

    # Write eNCF
    search_text = "ENCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 
            ET.SubElement(id_doc, 'eNCF').text = value
            __logger.info(f'eNCF : {cell_value}')
            break

    # Write FechaVencimientoSecuencia
    search_text = "FechaVencimientoSecuencia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaVencimientoSecuencia').text = value
            __logger.info(f'FechaVencimientoSecuencia : {value}')
            break
    
    
    # Write IndicadorNotaCredito
    search_text = "IndicadorNotaCredito" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'IndicadorNotaCredito').text = value
            __logger.info(f'IndicadorNotaCredito : {value}')
            break

    # Write IndicadorEnvioDiferido
    search_text = "IndicadorEnvioDiferido" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'IndicadorEnvioDiferido').text = value
            __logger.info(f'IndicadorEnvioDiferido : {value}')
            break

    # Write IndicadorMontoGravado
    search_text = "IndicadorMontoGravado" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'IndicadorMontoGravado').text = value
            __logger.info(f'IndicadorMontoGravado : {value}')
            break

    # Write TipoIngresos
    search_text = "TipoIngresos" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoIngresos').text = value
            __logger.info(f'TipoIngresos : {value}')
            break

    # Write TipoPago
    search_text = "TipoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoPago').text = value
            __logger.info(f'TipoPago : {value}')
            break

    # Write FechaLimitePago
    search_text = "FechaLimitePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaLimitePago').text = value
            __logger.info(f'FechaLimitePago : {value}')
            break


    # Write TerminoPago
    search_text = "TerminoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TerminoPago').text = value
            __logger.info(f'TerminoPago : {value}')
            break
    
    
    TablaFormasPago = ET.SubElement(id_doc, 'TablaFormasPago')

    search_text = "FormaPago[1]"
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            FormaDePago_count = 7
            col_index = col
            while True :
                FormaDePago_count -= 1
                if FormaDePago_count < 0 :
                    break            
                FormaDePago = ET.SubElement(TablaFormasPago, 'FormaDePago')

                value = str(sheet.cell(in_row, column= col_index ).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(FormaDePago, 'FormaPago').text = value
                col_index += 1

                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(FormaDePago, 'MontoPago').text = value
                col_index += 1
    
    # Write TipoCuentaPago
    search_text = "TipoCuentaPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoCuentaPago').text = value
            __logger.info(f'TipoCuentaPago : {value}')
            break
        
    # Write NumeroCuentaPago
    search_text = "NumeroCuentaPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'NumeroCuentaPago').text = value
            __logger.info(f'NumeroCuentaPago : {value}')
            break
        
    # Write BancoPago
    search_text = "BancoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'BancoPago').text = value
            __logger.info(f'BancoPago : {value}')
            break

    # Write FechaDesde
    search_text = "FechaDesde" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaDesde').text = value
            __logger.info(f'FechaDesde : {value}')
            break    

    # Write FechaHasta
    search_text = "FechaHasta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaHasta').text = value
            __logger.info(f'FechaHasta : {value}')
            break    

    # Write TotalPaginas
    search_text = "TotalPaginas" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TotalPaginas').text = value
            __logger.info(f'TotalPaginas : {value}')
            break

    Emisor = ET.SubElement(encabezado, 'Emisor')

    # Write RNCEmisor
    search_text = "RNCEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RNCEmisor').text = value
            __logger.info(f'RNCEmisor : {value}')
            break

    # Write RazonSocialEmisor
    search_text = "RazonSocialEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RazonSocialEmisor').text = value
            __logger.info(f'RazonSocialEmisor : {value}')
            break

    # Write NombreComercial
    search_text = "NombreComercial" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NombreComercial').text = value
            __logger.info(f'NombreComercial : {value}')
            break

    # Write Sucursal
    search_text = "Sucursal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Sucursal').text = value
            __logger.info(f'Sucursal : {value}')
            break

    # Write DireccionEmisor
    search_text = "DireccionEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'DireccionEmisor').text = value
            __logger.info(f'DireccionEmisor : {value}')
            break

    # Write Municipio
    search_text = "Municipio" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Municipio').text = value
            __logger.info(f'Municipio : {value}')
            break


    # Write Provincia
    search_text = "Provincia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Provincia').text = value
            __logger.info(f'Provincia : {value}')
            break

    TablaTelefonoEmisor = ET.SubElement(Emisor, 'TablaTelefonoEmisor')

    # Write TelefonoEmisor
    search_text = "TelefonoEmisor[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            TelefonoEmisor_count = 3
            col_index = col
            while True :
                TelefonoEmisor_count -= 1
                if TelefonoEmisor_count < 0 : 
                    break

                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(TablaTelefonoEmisor, 'TelefonoEmisor').text = value           
                __logger.info(f'TelefonoEmisor : {value}')

                col_index +=1
             
    # Write CorreoEmisor
    search_text = "CorreoEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'CorreoEmisor').text = value
            __logger.info(f'CorreoEmisor : {value}')
            break

    # Write WebSite
    search_text = "WebSite" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'WebSite').text = value
            __logger.info(f'WebSite : {value}')
            break

    # Write ActividadEconomica
    search_text = "ActividadEconomica" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'ActividadEconomica').text = value
            __logger.info(f'ActividadEconomica : {value}')
            break

    # Write NumeroFacturaInterna
    search_text = "NumeroFacturaInterna" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NumeroFacturaInterna').text = value
            __logger.info(f'NumeroFacturaInterna : {value}')
            break
 
    # Write NumeroPedidoInterno
    search_text = "NumeroPedidoInterno" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NumeroPedidoInterno').text = value
            __logger.info(f'NumeroPedidoInterno : {value}')
            break
 
    # Write ZonaVenta
    search_text = "ZonaVenta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'ZonaVenta').text = value
            __logger.info(f'ZonaVenta : {value}')
            break

    # Write RutaVenta
    search_text = "RutaVenta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RutaVenta').text = value
            __logger.info(f'RutaVenta : {value}')
            break

    # Write InformacionAdicionalEmisor
    search_text = "InformacionAdicionalEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'InformacionAdicionalEmisor').text = value
            __logger.info(f'InformacionAdicionalEmisor : {value}')
            break

    # Write FechaEmision
    search_text = "FechaEmision" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'FechaEmision').text = value
            __logger.info(f'FechaEmision : {value}')
            break

    Comprador = ET.SubElement(encabezado, 'Comprador')

    # Write RNCComprador
    search_text = "RNCComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'RNCComprador').text = value
            __logger.info(f'RNCComprador : {value}')
            break

    # Write IdentificadorExtranjero
    search_text = "IdentificadorExtranjero" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'IdentificadorExtranjero').text = value
            __logger.info(f'IdentificadorExtranjero : {value}')
            break

    # Write RazonSocialComprador
    search_text = "RazonSocialComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'RazonSocialComprador').text = value
            __logger.info(f'RazonSocialComprador : {value}')
            break

    # Write ContactoComprador
    search_text = "ContactoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ContactoComprador').text = value
            __logger.info(f'ContactoComprador : {value}')
            break

    # Write CorreoComprador
    search_text = "CorreoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'CorreoComprador').text = value
            __logger.info(f'CorreoComprador : {value}')
            break

    # Write DireccionComprador
    search_text = "DireccionComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'DireccionComprador').text = value
            __logger.info(f'DireccionComprador : {value}')
            break

    # Write ProvinciaComprador
    search_text = "ProvinciaComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ProvinciaComprador').text = value
            __logger.info(f'ProvinciaComprador : {value}')
            break

    # Write PaisComprador
    search_text = "PaisComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'PaisComprador').text = value
            __logger.info(f'PaisComprador : {value}')
            break

    # Write FechaEntrega
    search_text = "FechaEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'FechaEntrega').text = value
            __logger.info(f'FechaEntrega : {value}')
            break

    # Write ContactoEntrega
    search_text = "ContactoEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ContactoEntrega').text = value
            __logger.info(f'ContactoEntrega : {value}')
            break

    # Write DireccionEntrega
    search_text = "DireccionEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'DireccionEntrega').text = value
            __logger.info(f'DireccionEntrega : {value}')
            break

    # Write TelefonoAdicional
    search_text = "TelefonoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'TelefonoAdicional').text = value
            __logger.info(f'TelefonoAdicional : {value}')
            break

    # Write FechaOrdenCompra
    search_text = "FechaOrdenCompra" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'FechaOrdenCompra').text = value
            __logger.info(f'FechaOrdenCompra : {value}')
            break

    # Write NumeroOrdenCompra
    search_text = "NumeroOrdenCompra" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'NumeroOrdenCompra').text = value
            __logger.info(f'NumeroOrdenCompra : {value}')
            break

    # Write CodigoInternoComprador
    search_text = "CodigoInternoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'CodigoInternoComprador').text = value
            __logger.info(f'CodigoInternoComprador : {value}')
            break

    # Write ResponsablePago
    search_text = "ResponsablePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ResponsablePago').text = value
            __logger.info(f'ResponsablePago : {value}')
            break
    
    InformacionesAdicionales = ET.SubElement(encabezado, 'InformacionesAdicionales')
   
    # Write FechaEmbarque
    search_text = "FechaEmbarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'FechaEmbarque').text = value
            __logger.info(f'FechaEmbarque : {value}')
            break

    # Write NumeroEmbarque
    search_text = "NumeroEmbarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NumeroEmbarque').text = value
            __logger.info(f'NumeroEmbarque : {value}')
            break

    # Write NumeroReferencia
    search_text = "NumeroReferencia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NumeroReferencia').text = value
            __logger.info(f'NumeroReferencia : {value}')
            break

    # Write NombrePuertoEmbarque
    search_text = "NombrePuertoEmbarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NombrePuertoEmbarque').text = value
            __logger.info(f'NombrePuertoEmbarque : {value}')
            break

    # Write NombrePuertoEmbarque
    search_text = "CondicionesEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'CondicionesEntrega').text = value
            __logger.info(f'CondicionesEntrega : {value}')
            break

    # Write TotalFob
    search_text = "TotalFob" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'TotalFob').text = value
            __logger.info(f'TotalFob : {value}')
            break

    # Write Flete
    search_text = "Flete" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'Flete').text = value
            __logger.info(f'Flete : {value}')
            break

    # Write OtrosGastos
    search_text = "OtrosGastos" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'OtrosGastos').text = value
            __logger.info(f'OtrosGastos : {value}')
            break

    # Write TotalCif
    search_text = "TotalCif" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'TotalCif').text = value
            __logger.info(f'TotalCif : {value}')
            break

    # Write RegimenAduanero
    search_text = "RegimenAduanero" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'RegimenAduanero').text = value
            __logger.info(f'RegimenAduanero : {value}')
            break

    # Write NombrePuertoSalida
    search_text = "NombrePuertoSalida" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NombrePuertoSalida').text = value
            __logger.info(f'NombrePuertoSalida : {value}')
            break

    # Write NombrePuertoDesembarque
    search_text = "NombrePuertoDesembarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NombrePuertoDesembarque').text = value
            __logger.info(f'NombrePuertoDesembarque : {value}')
            break

    # Write PesoBruto
    search_text = "PesoBruto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'PesoBruto').text = value
            __logger.info(f'PesoBruto : {value}')
            break

    # Write PesoNeto
    search_text = "PesoNeto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'PesoNeto').text = value
            __logger.info(f'PesoNeto : {value}')
            break

    # Write UnidadPesoBruto
    search_text = "UnidadPesoBruto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadPesoBruto').text = value
            __logger.info(f'UnidadPesoBruto : {value}')
            break

    # Write UnidadPesoNeto
    search_text = "UnidadPesoNeto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadPesoNeto').text = value
            __logger.info(f'UnidadPesoNeto : {value}')
            break

    # Write CantidadBulto
    search_text = "CantidadBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'CantidadBulto').text = value
            __logger.info(f'CantidadBulto : {value}')
            break

    # Write UnidadBulto
    search_text = "UnidadBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadBulto').text = value
            __logger.info(f'UnidadBulto : {value}')
            break

    # Write VolumenBulto
    search_text = "VolumenBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'VolumenBulto').text = value
            __logger.info(f'VolumenBulto : {value}')
            break

    Transporte = ET.SubElement(encabezado, 'Transporte')
    
    # Write ViaTransporte
    search_text = "ViaTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'ViaTransporte').text = value
            __logger.info(f'ViaTransporte : {value}')
            break

    # Write PaisOrigen
    search_text = "PaisOrigen" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'PaisOrigen').text = value
            __logger.info(f'PaisOrigen : {value}')
            break

    # Write DireccionDestino
    search_text = "DireccionDestino" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'DireccionDestino').text = value
            __logger.info(f'DireccionDestino : {value}')
            break

    # Write PaisDestino
    search_text = "PaisDestino" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'PaisDestino').text = value
            __logger.info(f'PaisDestino : {value}')
            break

    # Write RNCIdentificacionCompaniaTransportista
    search_text = "RNCIdentificacionCompaniaTransportista" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'RNCIdentificacionCompaniaTransportista').text = value
            __logger.info(f'RNCIdentificacionCompaniaTransportista : {value}')
            break

    # Write NombreCompaniaTransportista
    search_text = "NombreCompaniaTransportista" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'NombreCompaniaTransportista').text = value
            __logger.info(f'NombreCompaniaTransportista : {value}')
            break

    # Write NumeroViaje
    search_text = "NumeroViaje" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'NumeroViaje').text = value
            __logger.info(f'NumeroViaje : {value}')
            break

    # Write Conductor
    search_text = "Conductor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Conductor').text = value
            __logger.info(f'Conductor : {value}')
            break

    # Write DocumentoTransporte
    search_text = "DocumentoTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'DocumentoTransporte').text = value
            __logger.info(f'DocumentoTransporte : {value}')
            break

    # Write Ficha
    search_text = "Ficha" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Ficha').text = value
            __logger.info(f'Ficha : {value}')
            break

    # Write Placa
    search_text = "Placa" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Placa').text = value
            __logger.info(f'Placa : {value}')
            break

    # Write RutaTransporte
    search_text = "RutaTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'RutaTransporte').text = value
            __logger.info(f'RutaTransporte : {value}')
            break

    # Write ZonaTransporte
    search_text = "ZonaTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'ZonaTransporte').text = value
            __logger.info(f'ZonaTransporte : {value}')
            break

    # Write NumeroAlbaran
    search_text = "NumeroAlbaran" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'NumeroAlbaran').text = value
            __logger.info(f'NumeroAlbaran : {value}')
            break

    Totales = ET.SubElement(encabezado, 'Totales')
   
    # Write MontoGravadoTotal
    search_text = "MontoGravadoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoTotal').text = value
            __logger.info(f'MontoGravadoTotal : {value}')
            break

   
    # Write MontoGravadoI1
    search_text = "MontoGravadoI1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI1').text = value
            __logger.info(f'MontoGravadoI1 : {value}')
            break
   
    # Write MontoGravadoI2
    search_text = "MontoGravadoI2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI2').text = value
            __logger.info(f'MontoGravadoI2 : {value}')
            break

   
    # Write MontoGravadoI3
    search_text = "MontoGravadoI3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI3').text = value
            __logger.info(f'MontoGravadoI3 : {value}')
            break
   
    # Write MontoExento
    search_text = "MontoExento" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoExento').text = value
            __logger.info(f'MontoExento : {value}')
            break
   
    # Write ITBIS1
    search_text = "ITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS1').text = value
            __logger.info(f'ITBIS1 : {value}')
            break
  
    # Write ITBIS2
    search_text = "ITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS2').text = value
            __logger.info(f'ITBIS2 : {value}')
            break
 
    # Write ITBIS3
    search_text = "ITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS3').text = value
            __logger.info(f'ITBIS3 : {value}')
            break
 
    # Write TotalITBIS
    search_text = "TotalITBIS" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS').text = value
            __logger.info(f'TotalITBIS : {value}')
            break
 
    # Write TotalITBIS1
    search_text = "TotalITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS1').text = value
            __logger.info(f'TotalITBIS1 : {value}')
            break

    # Write TotalITBIS2
    search_text = "TotalITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS2').text = value
            __logger.info(f'TotalITBIS2 : {value}')
            break

    # Write TotalITBIS3
    search_text = "TotalITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS3').text = value
            __logger.info(f'TotalITBIS3 : {value}')
            break

    # Write MontoImpuestoAdicional
    search_text = "MontoImpuestoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoImpuestoAdicional').text = value
            __logger.info(f'MontoImpuestoAdicional : {value}')
            break

    ImpuestosAdicionales = ET.SubElement(Totales, 'ImpuestosAdicionales')
    search_text = "TipoImpuesto[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:

            col_index = col
            ImpuestoAdicional_count = 4
            while True:
                ImpuestoAdicional_count -= 1
                if ImpuestoAdicional_count < 0:
                    break

                ImpuestoAdicional = ET.SubElement(ImpuestosAdicionales, 'ImpuestoAdicional')

                # TipoImpuesto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = value
                __logger.info(f'TipoImpuesto : {value}')
                col_index +=1

                # TasaImpuestoAdicional
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'TasaImpuestoAdicional').text = value
                __logger.info(f'TasaImpuestoAdicional : {value}')
                col_index +=1

                # MontoImpuestoSelectivoConsumoEspecifico
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoEspecifico').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoEspecifico : {value}')
                col_index +=1

                # MontoImpuestoSelectivoConsumoAdvalorem
                value = str(sheet.cell(in_row, column=col_index + 3).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoAdvalorem').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoAdvalorem : {value}')
                col_index +=1


                # OtrosImpuestosAdicionales
                value = str(sheet.cell(in_row, column=col_index + 4).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'OtrosImpuestosAdicionales').text = value
                __logger.info(f'OtrosImpuestosAdicionales : {value}')

    # Write MontoTotal
    search_text = "MontoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoTotal').text = value
            __logger.info(f'MontoTotal : {value}')
            break

    # Write MontoNoFacturable
    search_text = "MontoNoFacturable" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoNoFacturable').text = value
            __logger.info(f'MontoNoFacturable : {value}')
            break

    # Write MontoPeriodo
    search_text = "MontoPeriodo" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoPeriodo').text = value
            __logger.info(f'MontoPeriodo : {value}')
            break

    # Write SaldoAnterior
    search_text = "SaldoAnterior" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'SaldoAnterior').text = value
            __logger.info(f'SaldoAnterior : {value}')
            break

    # Write MontoAvancePago
    search_text = "MontoAvancePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoAvancePago').text = value
            __logger.info(f'MontoAvancePago : {value}')
            break

    # Write ValorPagar
    search_text = "ValorPagar" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ValorPagar').text = value
            __logger.info(f'ValorPagar : {value}')
            break

    # Write TotalITBISRetenido
    search_text = "TotalITBISRetenido" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBISRetenido').text = value
            __logger.info(f'TotalITBISRetenido : {value}')
            break

    # Write TotalISRRetencion
    search_text = "TotalISRRetencion" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalISRRetencion').text = value
            __logger.info(f'TotalISRRetencion : {value}')
            break

    # Write TotalITBISPercepcion
    search_text = "TotalITBISPercepcion" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBISPercepcion').text = value
            __logger.info(f'TotalITBISPercepcion : {value}')
            break

    # Write TotalISRPercepcion
    search_text = "TotalISRPercepcion" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalISRPercepcion').text = value
            __logger.info(f'TotalISRPercepcion : {value}')
            break

    OtraMoneda = ET.SubElement(encabezado, 'OtraMoneda')
  
    # Write TipoMoneda
    search_text = "TipoMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TipoMoneda').text = value
            __logger.info(f'TipoMoneda : {value}')
            break
  
    # Write TipoCambio
    search_text = "TipoCambio" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TipoCambio').text = value
            __logger.info(f'TipoCambio : {value}')
            break
  
    # Write MontoGravadoTotalOtraMoneda
    search_text = "MontoGravadoTotalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravadoTotalOtraMoneda').text = value
            __logger.info(f'MontoGravadoTotalOtraMoneda : {value}')
            break
  
    # Write MontoGravado1OtraMoneda
    search_text = "MontoGravado1OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado1OtraMoneda').text = value
            __logger.info(f'MontoGravado1OtraMoneda : {value}')
            break
  
    # Write MontoGravado2OtraMoneda
    search_text = "MontoGravado2OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado2OtraMoneda').text = value
            __logger.info(f'MontoGravado2OtraMoneda : {value}')
            break
  
    # Write MontoGravado3OtraMoneda
    search_text = "MontoGravado3OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado3OtraMoneda').text = value
            __logger.info(f'MontoGravado3OtraMoneda : {value}')
            break
  
    # Write MontoExentoOtraMoneda
    search_text = "MontoExentoOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoExentoOtraMoneda').text = value
            __logger.info(f'MontoExentoOtraMoneda : {value}')
            break
  
    # Write TotalITBISOtraMoneda
    search_text = "TotalITBISOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBISOtraMoneda').text = value
            __logger.info(f'TotalITBISOtraMoneda : {value}')
            break

    # Write TotalITBIS1OtraMoneda
    search_text = "TotalITBIS1OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS1OtraMoneda').text = value
            __logger.info(f'TotalITBIS1OtraMoneda : {value}')
            break

    # Write TotalITBIS2OtraMoneda
    search_text = "TotalITBIS2OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS2OtraMoneda').text = value
            __logger.info(f'TotalITBIS2OtraMoneda : {value}')
            break

    # Write TotalITBIS3OtraMoneda
    search_text = "TotalITBIS3OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS3OtraMoneda').text = value
            __logger.info(f'TotalITBIS3OtraMoneda : {value}')
            break

    # Write MontoImpuestoAdicionalOtraMoneda
    search_text = "MontoImpuestoAdicionalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoImpuestoAdicionalOtraMoneda').text = value
            __logger.info(f'MontoImpuestoAdicionalOtraMoneda : {value}')
            break

    ImpuestosAdicionalesOtraMoneda = ET.SubElement(OtraMoneda, 'ImpuestosAdicionalesOtraMoneda')

    search_text = "TipoImpuestoOtraMoneda[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            
            ImpuestoAdicionalOtraMoneda_count = 4
            col_index = col
            while True :
                ImpuestoAdicionalOtraMoneda_count -= 1
                if ImpuestoAdicionalOtraMoneda_count < 0 :
                    break

                ImpuestoAdicionalOtraMoneda = ET.SubElement(ImpuestosAdicionalesOtraMoneda, 'ImpuestoAdicionalOtraMoneda')

                # Write TipoImpuestoOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'TipoImpuestoOtraMoneda').text = value
                __logger.info(f'TipoImpuestoOtraMoneda : {value}')
                col_index +=1
               
                # Write TasaImpuestoAdicionalOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'TasaImpuestoAdicionalOtraMoneda').text = value
                __logger.info(f'TasaImpuestoAdicionalOtraMoneda : {value}')
                col_index +=1

                # Write MontoImpuestoSelectivoConsumoEspecificoOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'MontoImpuestoSelectivoConsumoEspecificoOtraMoneda').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoEspecificoOtraMoneda : {value}')
                col_index +=1

                # Write MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda : {value}')
                col_index +=1

                # Write OtrosImpuestosAdicionalesOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'OtrosImpuestosAdicionalesOtraMoneda').text = value
                __logger.info(f'OtrosImpuestosAdicionalesOtraMoneda : {value}')

    # Write MontoTotalOtraMoneda
    search_text = "MontoTotalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoTotalOtraMoneda').text = value
            __logger.info(f'MontoTotalOtraMoneda : {value}')
            break

    # """
    DetallesItems = ET.SubElement(root, 'DetallesItems')

    item_count = 62
    search_text = "NumeroLinea[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            col_index = col
            while True:
                item_count -= 1
                if item_count < 0 : 
                    break
                Item = ET.SubElement(DetallesItems, 'Item')

                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'NumeroLinea').text = value
                __logger.info(f'NumeroLinea : {value}')
                col_index += 1
                

                # TablaCodigosItem
                TablaCodigosItem = ET.SubElement(Item, 'TablaCodigosItem')
                CodigosItem_count = 5;
                while True : 
                    CodigosItem_count -= 1
                    if CodigosItem_count < 0 : 
                        break

                    CodigosItem = ET.SubElement(TablaCodigosItem, 'CodigosItem')
                    
                    # TipoCodigo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(CodigosItem, 'TipoCodigo').text = value
                    __logger.info(f'TipoCodigo : {value}')
                    col_index += 1

                    # CodigoItem
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(CodigosItem, 'CodigoItem').text = value
                    __logger.info(f'CodigoItem : {value}')
                    col_index += 1

                # IndicadorFacturacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'IndicadorFacturacion').text = value
                __logger.info(f'IndicadorFacturacion : {value}')
                col_index += 1

                Retencion = ET.SubElement(Item, 'Retencion')
                # IndicadorAgenteRetencionoPercepcion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'IndicadorAgenteRetencionoPercepcion').text = value
                __logger.info(f'IndicadorAgenteRetencionoPercepcion : {value}')
                col_index += 1
                
                # MontoITBISRetenido
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'MontoITBISRetenido').text = value
                __logger.info(f'MontoITBISRetenido : {value}')
                col_index += 1

                # MontoISRRetenido
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'MontoISRRetenido').text = value
                __logger.info(f'MontoISRRetenido : {value}')
                col_index += 1

                # NombreItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'NombreItem').text = value
                __logger.info(f'NombreItem : {value}')
                col_index += 1

                # IndicadorBienoServicio
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'IndicadorBienoServicio').text = value
                __logger.info(f'IndicadorBienoServicio : {value}')
                col_index += 1

                # DescripcionItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'DescripcionItem').text = value
                __logger.info(f'DescripcionItem : {value}')
                col_index += 1

                # CantidadItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'CantidadItem').text = value
                __logger.info(f'CantidadItem : {value}')
                col_index += 1

                # UnidadMedida
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'UnidadMedida').text = value
                __logger.info(f'UnidadMedida : {value}')
                col_index += 1

                # CantidadReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'CantidadReferencia').text = value
                __logger.info(f'CantidadReferencia : {value}')
                col_index += 1

                # UnidadReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'UnidadReferencia').text = value
                __logger.info(f'UnidadReferencia : {value}')
                col_index += 1

                TablaSubcantidad = ET.SubElement(Item, 'TablaSubcantidad')

                SubcantidadItem_count = 5
                while True : 
                    SubcantidadItem_count -= 1
                    if SubcantidadItem_count < 0 :
                        break

                    SubcantidadItem = ET.SubElement(TablaSubcantidad, 'SubcantidadItem')
                    
                    # Subcantidad
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubcantidadItem, 'Subcantidad').text = value
                    __logger.info(f'Subcantidad : {value}')
                    col_index += 1
                    
                    # CodigoSubcantidad
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubcantidadItem, 'CodigoSubcantidad').text = value
                    __logger.info(f'CodigoSubcantidad : {value}')
                    col_index += 1

                # GradosAlcohol
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'GradosAlcohol').text = value
                __logger.info(f'GradosAlcohol : {value}')
                col_index += 1

                # PrecioUnitarioReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'PrecioUnitarioReferencia').text = value
                __logger.info(f'PrecioUnitarioReferencia : {value}')
                col_index += 1

                # FechaElaboracion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'FechaElaboracion').text = value
                __logger.info(f'FechaElaboracion : {value}')
                col_index += 1

                # FechaVencimientoItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'FechaVencimientoItem').text = value
                __logger.info(f'FechaVencimientoItem : {value}')
                col_index += 1

                Mineria = ET.SubElement(Item, 'Mineria')
                
                # PesoNetoKilogramo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'PesoNetoKilogramo').text = value
                __logger.info(f'PesoNetoKilogramo : {value}')
                col_index += 1
                
                # PesoNetoMineria
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'PesoNetoMineria').text = value
                __logger.info(f'PesoNetoMineria : {value}')
                col_index += 1

                # TipoAfiliacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'TipoAfiliacion').text = value
                __logger.info(f'TipoAfiliacion : {value}')
                col_index += 1

                # Liquidacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'Liquidacion').text = value
                __logger.info(f'Liquidacion : {value}')
                col_index += 1

                
                # PrecioUnitarioItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(Item, 'PrecioUnitarioItem').text = value
                __logger.info(f'PrecioUnitarioItem : {value}')
                col_index += 1

                # DescuentoMonto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'DescuentoMonto').text = value
                __logger.info(f'DescuentoMonto : {value}')
                col_index += 1

                TablaSubDescuento = ET.SubElement(Item, 'TablaSubDescuento')

                SubDescuento_count = 5

                while True : 
                    SubDescuento_count -= 1
                    if SubDescuento_count < 0 : 
                        break

                    #SubDescuento
                    SubDescuento = ET.SubElement(TablaSubDescuento, 'SubDescuento')

                    # TipoSubDescuento
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'TipoSubDescuento').text = value
                    __logger.info(f'TipoSubDescuento : {value}')
                    col_index += 1

                    # SubDescuentoPorcentaje
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'SubDescuentoPorcentaje').text = value
                    __logger.info(f'SubDescuentoPorcentaje : {value}')
                    col_index += 1

                    # MontoSubDescuento
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'MontoSubDescuento').text = value
                    __logger.info(f'MontoSubDescuento : {value}')
                    col_index += 1


                # RecargoMonto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'RecargoMonto').text = value
                __logger.info(f'RecargoMonto : {value}')
                col_index += 1

                TablaSubRecargo = ET.SubElement(Item, 'TablaSubRecargo')

                SubRecargo_count = 5

                while True :
                    SubRecargo_count -= 1
                    if SubRecargo_count < 0 :
                        break    
                    SubRecargo = ET.SubElement(TablaSubRecargo, 'SubRecargo')

                    # TipoSubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'TipoSubRecargo').text = value
                    __logger.info(f'TipoSubRecargo : {value}')
                    col_index += 1

                    # SubRecargoPorcentaje
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'SubRecargoPorcentaje').text = value
                    __logger.info(f'SubRecargoPorcentaje : {value}')
                    col_index += 1

                    
                    # MontosubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'MontosubRecargo').text = value
                    __logger.info(f'MontosubRecargo : {value}')
                    col_index += 1

                TablaImpuestoAdicional = ET.SubElement(Item, 'TablaImpuestoAdicional')

                ImpuestoAdicional_count = 2
                while True :
                    ImpuestoAdicional_count -= 1
                    if ImpuestoAdicional_count < 0 : 
                        break
                    ImpuestoAdicional =  ET.SubElement(TablaImpuestoAdicional, 'ImpuestoAdicional')                                   
                    # MontosubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = value
                    __logger.info(f'TipoImpuesto : {value}')
                    col_index += 1
                
                OtraMonedaDetalle = ET.SubElement(Item, 'OtraMonedaDetalle')

                # PrecioOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'PrecioOtraMoneda').text = value
                __logger.info(f'PrecioOtraMoneda : {value}')
                col_index += 1

                # DescuentoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'DescuentoOtraMoneda').text = value
                __logger.info(f'DescuentoOtraMoneda : {value}')
                col_index += 1

                # RecargoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'RecargoOtraMoneda').text = value
                __logger.info(f'RecargoOtraMoneda : {value}')

                col_index += 1
                # MontoItemOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'MontoItemOtraMoneda').text = value
                __logger.info(f'MontoItemOtraMoneda : {value}')
                col_index += 1
                
                # MontoItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'MontoItem').text = value
                __logger.info(f'MontoItem : {value}')
                col_index += 1
    
    Subtotales = ET.SubElement(root, 'Subtotales')
    Subtotal = ET.SubElement(Subtotales, 'Subtotal')

    # Write NumeroSubTotal
    search_text = "NumeroSubTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'Subtotal').text = value
            __logger.info(f'Version : {cell_value}')
            break
    
    # Write DescripcionSubtotal
    search_text = "DescripcionSubtotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'DescripcionSubtotal').text = value
            __logger.info(f'DescripcionSubtotal : {cell_value}')
            break
    
    # Write Orden
    search_text = "Orden" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'Orden').text = value
            __logger.info(f'Orden : {cell_value}')
            break
    
    # Write SubTotalMontoGravadoTotal
    search_text = "SubTotalMontoGravadoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoTotal').text = value
            __logger.info(f'SubTotalMontoGravadoTotal : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI1
    search_text = "SubTotalMontoGravadoI1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI1').text = value
            __logger.info(f'SubTotalMontoGravadoI1 : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI2
    search_text = "SubTotalMontoGravadoI2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI2').text = value
            __logger.info(f'SubTotalMontoGravadoI2 : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI3
    search_text = "SubTotalMontoGravadoI3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI3').text = value
            __logger.info(f'SubTotalMontoGravadoI3 : {cell_value}')
            break
  
    # Write SubTotaITBIS
    search_text = "SubTotaITBIS" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS').text = value
            __logger.info(f'SubTotaITBIS : {cell_value}')
            break

    # Write SubTotaITBIS1
    search_text = "SubTotaITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS1').text = value
            __logger.info(f'SubTotaITBIS1 : {cell_value}')
            break

    # Write SubTotaITBIS2
    search_text = "SubTotaITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS2').text = value
            __logger.info(f'SubTotaITBIS2 : {cell_value}')
            break

    # Write SubTotaITBIS3
    search_text = "SubTotaITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS3').text = value
            __logger.info(f'SubTotaITBIS3 : {cell_value}')
            break

    # Write SubTotalImpuestoAdicional
    search_text = "SubTotalImpuestoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalImpuestoAdicional').text = value
            __logger.info(f'SubTotalImpuestoAdicional : {cell_value}')
            break

    # Write SubTotalExento
    search_text = "SubTotalExento" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalExento').text = value
            __logger.info(f'SubTotalExento : {cell_value}')
            break

    # Write MontoSubTotal
    search_text = "MontoSubTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'MontoSubTotal').text = value
            __logger.info(f'MontoSubTotal : {cell_value}')
            break

    # Write Lineas
    search_text = "Lineas" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'Lineas').text = value
            __logger.info(f'Lineas : {cell_value}')
            break

    DescuentosORecargos = ET.SubElement(root, 'DescuentosORecargos')
    DescuentoORecargo_count = 2

    search_text = "NumeroLineaDoR[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        
        if cell_value == search_text:
            col_index = col
            while True : 
                DescuentoORecargo_count -= 1
                if DescuentoORecargo_count < 0:
                    break

                DescuentoORecargo = ET.SubElement(DescuentosORecargos, 'DescuentoORecargo')

                # NumeroLineaDoR
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'NumeroLineaDoR').text = value
                __logger.info(f'Lineas : {cell_value}')
                col_index +=1

                # TipoAjuste
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'TipoAjuste').text = value
                __logger.info(f'TipoAjuste : {cell_value}')
                col_index +=1
                
                # IndicadorNorma1007
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'IndicadorNorma1007').text = value
                __logger.info(f'IndicadorNorma1007 : {cell_value}')
                col_index +=1
                         
                # DescripcionDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'DescripcionDescuentooRecargo').text = value
                __logger.info(f'DescripcionDescuentooRecargo : {cell_value}')
                col_index +=1
                                         
                # TipoValor
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'TipoValor').text = value
                __logger.info(f'TipoValor : {cell_value}')
                col_index +=1
                                                     
                # ValorDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'ValorDescuentooRecargo').text = value
                __logger.info(f'ValorDescuentooRecargo : {cell_value}')
                col_index +=1
                                                               
                # MontoDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'MontoDescuentooRecargo').text = value
                __logger.info(f'MontoDescuentooRecargo : {cell_value}')
                col_index +=1
                                                                        
                # MontoDescuentooRecargoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'MontoDescuentooRecargoOtraMoneda').text = value
                __logger.info(f'MontoDescuentooRecargoOtraMoneda : {cell_value}')
                col_index +=1        

                # IndicadorFacturacionDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'IndicadorFacturacionDescuentooRecargo').text = value
                __logger.info(f'IndicadorFacturacionDescuentooRecargo : {cell_value}')         
     

    Paginacion = ET.SubElement(root, 'Paginacion')
    Pagina_count = 2

    search_text = "PaginaNo[1]" 
    col_index : int
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:

            col_index = col
                
            while True:
                Pagina_count -= 1
                if Pagina_count < 0:
                    break

                Pagina = ET.SubElement(Paginacion, 'Pagina')

                # PaginNo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'PaginaNo').text = value
                __logger.info(f'PaginaNo : {cell_value}')
                __logger.info(f'PaginaNo col_index : {col_index}')
                col_index += 1

                # NoLineaDesde
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'NoLineaDesde').text = value
                __logger.info(f'NoLineaDesde : {cell_value}')
                col_index += 1

                # NoLineaHasta
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'NoLineaHasta').text = value
                __logger.info(f'NoLineaHasta : {value}')
                col_index += 1
                
                # SubtotalMontoGravadoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravadoPagina').text = value
                __logger.info(f'SubtotalMontoGravadoPagina : {value}')
                col_index += 1
                                
                # SubtotalMontoGravado1Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado1Pagina').text = value
                __logger.info(f'SubtotalMontoGravado1Pagina : {value}')
                col_index += 1
                                                
                # SubtotalMontoGravado2Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado2Pagina').text = value
                __logger.info(f'SubtotalMontoGravado2Pagina : {value}')
                col_index += 1
                                                   
                # SubtotalMontoGravado3Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado3Pagina').text = value
                __logger.info(f'SubtotalMontoGravado3Pagina : {value}')
                col_index += 1
                                                                   
                # SubtotalExentoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalExentoPagina').text = value
                __logger.info(f'SubtotalExentoPagina : {value}')
                col_index += 1
                                                               
                # SubtotalItbisPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbisPagina').text = value
                __logger.info(f'SubtotalItbisPagina : {value}')
                col_index += 1
                                                                      
                # SubtotalItbis1Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis1Pagina').text = value
                __logger.info(f'SubtotalItbis1Pagina : {value}')
                col_index += 1
                                                                                      
                # SubtotalItbis2Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis2Pagina').text = value
                __logger.info(f'SubtotalItbis2Pagina : {value}')
                col_index += 1
                                                                                         
                # SubtotalItbis3Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis3Pagina').text = value
                __logger.info(f'SubtotalItbis3Pagina : {value}')
                col_index += 1
                                                                                                         
                # SubtotalImpuestoAdicionalPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalImpuestoAdicionalPagina').text = value
                __logger.info(f'SubtotalImpuestoAdicionalPagina : {value}')
                col_index += 1
                                                                                                                 
                # SubtotalImpuestoAdicionalPaginaTabla
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalImpuestoAdicionalPaginaTabla').text = value
                __logger.info(f'SubtotalImpuestoAdicionalPaginaTabla : {value}')
                col_index += 1

                SubtotalImpuestoAdicional = ET.SubElement(Pagina, 'SubtotalImpuestoAdicional')
                
                # SubtotalImpuestoSelectivoConsumoEspecificoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(SubtotalImpuestoAdicional, 'SubtotalImpuestoSelectivoConsumoEspecificoPagina').text =value
                __logger.info(f'SubtotalImpuestoSelectivoConsumoEspecificoPagina : {value}')
                col_index += 1

                # SubtotalOtrosImpuesto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(SubtotalImpuestoAdicional, 'SubtotalOtrosImpuesto').text =value
                __logger.info(f'SubtotalOtrosImpuesto : {value}')
                col_index += 1

                # MontoSubtotalPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(Pagina, 'MontoSubtotalPagina').text =value
                __logger.info(f'MontoSubtotalPagina : {value}')
                col_index += 1

                # SubtotalMontoNoFacturablePagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(Pagina, 'SubtotalMontoNoFacturablePagina').text =value
                __logger.info(f'SubtotalMontoNoFacturablePagina : {value}')
                __logger.info(f'SubtotalMontoNoFacturablePagina _ index : {col_index}')
                # col_index += 1

    InformacionReferencia = ET.SubElement(root, 'InformacionReferencia')
    
    # NCFModificado
    col_index += 1
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'NCFModificado').text =value
    __logger.info(f'NCFModificado : {value}')
    col_index += 1
        
    # RNCOtroContribuyente
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'RNCOtroContribuyente').text =value
    __logger.info(f'RNCOtroContribuyente : {value}')
    col_index += 1
        
    # FechaNCFModificado
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'FechaNCFModificado').text =value
    __logger.info(f'FechaNCFModificado : {value}')
    col_index += 1
        
    # CodigoModificacion
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'CodigoModificacion').text =value
    __logger.info(f'CodigoModificacion : {value}')
    col_index += 1

    # RazonModificacion
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'RazonModificacion').text =value
    __logger.info(f'RazonModificacion : {value}')
    col_index += 1
    # """

    rough_string = ET.tostring(root, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    pretty_xml_as_str = reparsed.toprettyxml(indent="  ", encoding='utf-8')

    path = os.path.join(os.path.dirname(__file__), 'data/row_invoice.xml')
    
    with open(path, 'wb') as f:
        f.write(pretty_xml_as_str)
    
    xml_str = ET.tostring(root, encoding='utf-8').decode('utf-8')
    return xml_str

def create_e_cf_31(in_row : int) :

    """Generate DGII-compliant XML from an Odoo invoice."""
    # Create root element with namespaces
    root = ET.Element('ECF', {
         
        'xmlns:xs': 'http://www.w3.org/2001/XMLSchema',
    })

    # 1. Encabezado
    encabezado = ET.SubElement(root, 'Encabezado')

    # Write Version
    search_text = "Version" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(encabezado, 'Version').text = "1.1"
            __logger.info(f'Version : {value}')
            __logger.info(f'Version Col Index : {col}')
            break

    # IdDoc
    id_doc = ET.SubElement(encabezado, 'IdDoc')

    # Write TipoeCF
    search_text = "TipoeCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 
            ET.SubElement(id_doc, 'TipoeCF').text = value
            __logger.info(f'TipoeCF : {cell_value}')
            break

    # Write eNCF
    search_text = "ENCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 
            ET.SubElement(id_doc, 'eNCF').text = value
            __logger.info(f'eNCF : {cell_value}')
            break

    # Write FechaVencimientoSecuencia
    search_text = "FechaVencimientoSecuencia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaVencimientoSecuencia').text = value
            __logger.info(f'FechaVencimientoSecuencia : {value}')
            break

    # Write IndicadorEnvioDiferido
    search_text = "IndicadorEnvioDiferido" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'IndicadorEnvioDiferido').text = value
            __logger.info(f'IndicadorEnvioDiferido : {value}')
            break

    # Write IndicadorMontoGravado
    search_text = "IndicadorMontoGravado" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'IndicadorMontoGravado').text = value
            __logger.info(f'IndicadorMontoGravado : {value}')
            break

    # Write TipoIngresos
    search_text = "TipoIngresos" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoIngresos').text = value
            __logger.info(f'TipoIngresos : {value}')
            break

    # Write TipoPago
    search_text = "TipoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoPago').text = value
            __logger.info(f'TipoPago : {value}')
            break

    # Write FechaLimitePago
    search_text = "FechaLimitePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaLimitePago').text = value
            __logger.info(f'FechaLimitePago : {value}')
            break


    # Write TerminoPago
    search_text = "TerminoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TerminoPago').text = value
            __logger.info(f'TerminoPago : {value}')
            break
    
    
    TablaFormasPago = ET.SubElement(id_doc, 'TablaFormasPago')

    search_text = "FormaPago[1]"
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            FormaDePago_count = 7
            col_index = col
            while True :
                FormaDePago_count -= 1
                if FormaDePago_count < 0 :
                    break            
                FormaDePago = ET.SubElement(TablaFormasPago, 'FormaDePago')

                value = str(sheet.cell(in_row, column= col_index ).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(FormaDePago, 'FormaPago').text = value
                col_index += 1

                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(FormaDePago, 'MontoPago').text = value
                col_index += 1
    
    # Write TipoCuentaPago
    search_text = "TipoCuentaPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoCuentaPago').text = value
            __logger.info(f'TipoCuentaPago : {value}')
            break
        
    # Write NumeroCuentaPago
    search_text = "NumeroCuentaPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'NumeroCuentaPago').text = value
            __logger.info(f'NumeroCuentaPago : {value}')
            break
        
    # Write BancoPago
    search_text = "BancoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'BancoPago').text = value
            __logger.info(f'BancoPago : {value}')
            break

    # Write FechaDesde
    search_text = "FechaDesde" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaDesde').text = value
            __logger.info(f'FechaDesde : {value}')
            break    

    # Write FechaHasta
    search_text = "FechaHasta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaHasta').text = value
            __logger.info(f'FechaHasta : {value}')
            break    

    # Write TotalPaginas
    search_text = "TotalPaginas" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TotalPaginas').text = value
            __logger.info(f'TotalPaginas : {value}')
            break

    Emisor = ET.SubElement(encabezado, 'Emisor')

    # Write RNCEmisor
    search_text = "RNCEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RNCEmisor').text = value
            __logger.info(f'RNCEmisor : {value}')
            break

    # Write RazonSocialEmisor
    search_text = "RazonSocialEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RazonSocialEmisor').text = value
            __logger.info(f'RazonSocialEmisor : {value}')
            break

    # Write NombreComercial
    search_text = "NombreComercial" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NombreComercial').text = value
            __logger.info(f'NombreComercial : {value}')
            break

    # Write Sucursal
    search_text = "Sucursal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Sucursal').text = value
            __logger.info(f'Sucursal : {value}')
            break

    # Write DireccionEmisor
    search_text = "DireccionEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'DireccionEmisor').text = value
            __logger.info(f'DireccionEmisor : {value}')
            break

    # Write Municipio
    search_text = "Municipio" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Municipio').text = value
            __logger.info(f'Municipio : {value}')
            break

    # Write Provincia
    search_text = "Provincia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Provincia').text = value
            __logger.info(f'Provincia : {value}')
            break

    TablaTelefonoEmisor = ET.SubElement(Emisor, 'TablaTelefonoEmisor')

    # Write TelefonoEmisor
    search_text = "TelefonoEmisor[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            TelefonoEmisor_count = 3
            col_index = col
            while True :
                TelefonoEmisor_count -= 1
                if TelefonoEmisor_count < 0 : 
                    break

                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(TablaTelefonoEmisor, 'TelefonoEmisor').text = value           
                __logger.info(f'TelefonoEmisor : {value}')

                col_index +=1
             
    # Write CorreoEmisor
    search_text = "CorreoEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'CorreoEmisor').text = value
            __logger.info(f'CorreoEmisor : {value}')
            break

    # Write WebSite
    search_text = "WebSite" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'WebSite').text = value
            __logger.info(f'WebSite : {value}')
            break

    # Write ActividadEconomica
    search_text = "ActividadEconomica" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'ActividadEconomica').text = value
            __logger.info(f'ActividadEconomica : {value}')
            break

    # Write CodigoVendedor
    search_text = "CodigoVendedor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'CodigoVendedor').text = value
            __logger.info(f'CodigoVendedor : {value}')
            break

    # Write NumeroFacturaInterna
    search_text = "NumeroFacturaInterna" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NumeroFacturaInterna').text = value
            __logger.info(f'NumeroFacturaInterna : {value}')
            break
 
    # Write NumeroPedidoInterno
    search_text = "NumeroPedidoInterno" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NumeroPedidoInterno').text = value
            __logger.info(f'NumeroPedidoInterno : {value}')
            break
 
    # Write ZonaVenta
    search_text = "ZonaVenta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'ZonaVenta').text = value
            __logger.info(f'ZonaVenta : {value}')
            break

    # Write RutaVenta
    search_text = "RutaVenta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RutaVenta').text = value
            __logger.info(f'RutaVenta : {value}')
            break

    # Write InformacionAdicionalEmisor
    search_text = "InformacionAdicionalEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'InformacionAdicionalEmisor').text = value
            __logger.info(f'InformacionAdicionalEmisor : {value}')
            break

    # Write FechaEmision
    search_text = "FechaEmision" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'FechaEmision').text = value
            __logger.info(f'FechaEmision : {value}')
            break

    Comprador = ET.SubElement(encabezado, 'Comprador')

    # Write RNCComprador
    search_text = "RNCComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'RNCComprador').text = value
            __logger.info(f'RNCComprador : {value}')
            break

    # Write RazonSocialComprador
    search_text = "RazonSocialComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'RazonSocialComprador').text = value
            __logger.info(f'RazonSocialComprador : {value}')
            break
   
    # Write ContactoComprador
    search_text = "ContactoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ContactoComprador').text = value
            __logger.info(f'ContactoComprador : {value}')
            break

    # Write CorreoComprador
    search_text = "CorreoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'CorreoComprador').text = value
            __logger.info(f'CorreoComprador : {value}')
            break

    # Write DireccionComprador
    search_text = "DireccionComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'DireccionComprador').text = value
            __logger.info(f'DireccionComprador : {value}')
            break

    # Write MunicipioComprador
    search_text = "MunicipioComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'MunicipioComprador').text = value
            __logger.info(f'MunicipioComprador : {value}')
            break

    # Write ProvinciaComprador
    search_text = "ProvinciaComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ProvinciaComprador').text = value
            __logger.info(f'ProvinciaComprador : {value}')
            break

    # Write FechaEntrega
    search_text = "FechaEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'FechaEntrega').text = value
            __logger.info(f'FechaEntrega : {value}')
            break

    # Write ContactoEntrega
    search_text = "ContactoEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ContactoEntrega').text = value
            __logger.info(f'ContactoEntrega : {value}')
            break

    # Write DireccionEntrega
    search_text = "DireccionEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'DireccionEntrega').text = value
            __logger.info(f'DireccionEntrega : {value}')
            break

    # Write TelefonoAdicional
    search_text = "TelefonoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'TelefonoAdicional').text = value
            __logger.info(f'TelefonoAdicional : {value}')
            break

    # Write FechaOrdenCompra
    search_text = "FechaOrdenCompra" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'FechaOrdenCompra').text = value
            __logger.info(f'FechaOrdenCompra : {value}')
            break

    # Write NumeroOrdenCompra
    search_text = "NumeroOrdenCompra" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'NumeroOrdenCompra').text = value
            __logger.info(f'NumeroOrdenCompra : {value}')
            break

    # Write CodigoInternoComprador
    search_text = "CodigoInternoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'CodigoInternoComprador').text = value
            __logger.info(f'CodigoInternoComprador : {value}')
            break

    # Write ResponsablePago
    search_text = "ResponsablePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ResponsablePago').text = value
            __logger.info(f'ResponsablePago : {value}')
            break
    
    # Write InformacionAdicionalComprador
    search_text = "InformacionAdicionalComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'InformacionAdicionalComprador').text = value
            __logger.info(f'InformacionAdicionalComprador : {value}')
            break
    
    InformacionesAdicionales = ET.SubElement(encabezado, 'InformacionesAdicionales')
   
    # Write FechaEmbarque
    search_text = "FechaEmbarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'FechaEmbarque').text = value
            __logger.info(f'FechaEmbarque : {value}')
            break

    # Write NumeroEmbarque
    search_text = "NumeroEmbarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NumeroEmbarque').text = value
            __logger.info(f'NumeroEmbarque : {value}')
            break

    # Write NumeroContenedor
    search_text = "NumeroContene" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NumeroContenedor').text = value
            __logger.info(f'NumeroContenedor : {value}')
            break

    value = str(sheet.cell(in_row, column=73).value)
    if value == "#e" :
        value = "" 

    ET.SubElement(InformacionesAdicionales, 'NumeroContenedor').text = value
    __logger.info(f'NumeroContenedor : {value}')
    
    # Write NumeroReferencia
    search_text = "NumeroReferencia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NumeroReferencia').text = value
            __logger.info(f'NumeroReferencia : {value}')
            break

    # Write PesoBruto
    search_text = "PesoBruto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'PesoBruto').text = value
            __logger.info(f'PesoBruto : {value}')
            break

    # Write PesoNeto
    search_text = "PesoNeto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'PesoNeto').text = value
            __logger.info(f'PesoNeto : {value}')
            break

    # Write UnidadPesoBruto
    search_text = "UnidadPesoBruto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadPesoBruto').text = value
            __logger.info(f'UnidadPesoBruto : {value}')
            break

    # Write UnidadPesoNeto
    search_text = "UnidadPesoNeto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadPesoNeto').text = value
            __logger.info(f'UnidadPesoNeto : {value}')
            break

    # Write CantidadBulto
    search_text = "CantidadBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'CantidadBulto').text = value
            __logger.info(f'CantidadBulto : {value}')
            break

    # Write UnidadBulto
    search_text = "UnidadBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadBulto').text = value
            __logger.info(f'UnidadBulto : {value}')
            break

    # Write VolumenBulto
    search_text = "VolumenBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'VolumenBulto').text = value
            __logger.info(f'VolumenBulto : {value}')
            break

    # Write UnidadVolumen
    search_text = "UnidadVolumen" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadVolumen').text = value
            __logger.info(f'UnidadVolumen : {value}')
            break

    # Transporte
    Transporte = ET.SubElement(encabezado, 'Transporte')

    # Write Conductor
    search_text = "Conductor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Conductor').text = value
            __logger.info(f'Conductor : {value}')
            break

    # Write DocumentoTransporte
    search_text = "DocumentoTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'DocumentoTransporte').text = value
            __logger.info(f'DocumentoTransporte : {value}')
            break

    # Write Ficha
    search_text = "Ficha" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Ficha').text = value
            __logger.info(f'Ficha : {value}')
            break

    # Write Placa
    search_text = "Placa" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Placa').text = value
            __logger.info(f'Placa : {value}')
            break

    # Write RutaTransporte
    search_text = "RutaTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'RutaTransporte').text = value
            __logger.info(f'RutaTransporte : {value}')
            break

    # Write ZonaTransporte
    search_text = "ZonaTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'ZonaTransporte').text = value
            __logger.info(f'ZonaTransporte : {value}')
            break

    # Write NumeroAlbaran
    search_text = "NumeroAlbaran" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'NumeroAlbaran').text = value
            __logger.info(f'NumeroAlbaran : {value}')
            break

    Totales = ET.SubElement(encabezado, 'Totales')
   
    # Write MontoGravadoTotal
    search_text = "MontoGravadoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoTotal').text = value
            __logger.info(f'MontoGravadoTotal : {value}')
            break

    # Write MontoGravadoI1
    search_text = "MontoGravadoI1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI1').text = value
            __logger.info(f'MontoGravadoI1 : {value}')
            break
   
    # Write MontoGravadoI2
    search_text = "MontoGravadoI2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI2').text = value
            __logger.info(f'MontoGravadoI2 : {value}')
            break
   
    # Write MontoGravadoI3
    search_text = "MontoGravadoI3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI3').text = value
            __logger.info(f'MontoGravadoI3 : {value}')
            break
   
    # Write MontoExento
    search_text = "MontoExento" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoExento').text = value
            __logger.info(f'MontoExento : {value}')
            break
   
    # Write ITBIS1
    search_text = "ITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS1').text = value
            __logger.info(f'ITBIS1 : {value}')
            break
  
    # Write ITBIS2
    search_text = "ITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS2').text = value
            __logger.info(f'ITBIS2 : {value}')
            break
 
    # Write ITBIS3
    search_text = "ITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS3').text = value
            __logger.info(f'ITBIS3 : {value}')
            break
 
    # Write TotalITBIS
    search_text = "TotalITBIS" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS').text = value
            __logger.info(f'TotalITBIS : {value}')
            break
 
    # Write TotalITBIS1
    search_text = "TotalITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS1').text = value
            __logger.info(f'TotalITBIS1 : {value}')
            break

    # Write TotalITBIS2
    search_text = "TotalITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS2').text = value
            __logger.info(f'TotalITBIS2 : {value}')
            break

    # Write TotalITBIS3
    search_text = "TotalITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS3').text = value
            __logger.info(f'TotalITBIS3 : {value}')
            break

    # Write MontoImpuestoAdicional
    search_text = "MontoImpuestoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoImpuestoAdicional').text = value
            __logger.info(f'MontoImpuestoAdicional : {value}')
            break

    ImpuestosAdicionales = ET.SubElement(Totales, 'ImpuestosAdicionales')
    search_text = "TipoImpuesto[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:

            col_index = col
            ImpuestoAdicional_count = 4
            while True:
                ImpuestoAdicional_count -= 1
                if ImpuestoAdicional_count < 0:
                    break

                ImpuestoAdicional = ET.SubElement(ImpuestosAdicionales, 'ImpuestoAdicional')

                # TipoImpuesto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = value
                __logger.info(f'TipoImpuesto : {value}')
                col_index +=1

                # TasaImpuestoAdicional
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'TasaImpuestoAdicional').text = value
                __logger.info(f'TasaImpuestoAdicional : {value}')
                col_index +=1

                # MontoImpuestoSelectivoConsumoEspecifico
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoEspecifico').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoEspecifico : {value}')
                col_index +=1

                # MontoImpuestoSelectivoConsumoAdvalorem
                value = str(sheet.cell(in_row, column=col_index + 3).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoAdvalorem').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoAdvalorem : {value}')
                col_index +=1


                # OtrosImpuestosAdicionales
                value = str(sheet.cell(in_row, column=col_index + 4).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'OtrosImpuestosAdicionales').text = value
                __logger.info(f'OtrosImpuestosAdicionales : {value}')

    # Write MontoTotal
    search_text = "MontoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoTotal').text = value
            __logger.info(f'MontoTotal : {value}')
            break

    # Write MontoNoFacturable
    search_text = "MontoNoFacturable" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoNoFacturable').text = value
            __logger.info(f'MontoNoFacturable : {value}')
            break

    # Write MontoPeriodo
    search_text = "MontoPeriodo" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoPeriodo').text = value
            __logger.info(f'MontoPeriodo : {value}')
            break

    # Write SaldoAnterior
    search_text = "SaldoAnterior" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'SaldoAnterior').text = value
            __logger.info(f'SaldoAnterior : {value}')
            break

    # Write MontoAvancePago
    search_text = "MontoAvancePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoAvancePago').text = value
            __logger.info(f'MontoAvancePago : {value}')
            break

    # Write ValorPagar
    search_text = "ValorPagar" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ValorPagar').text = value
            __logger.info(f'ValorPagar : {value}')
            break

    # Write TotalITBISRetenido
    search_text = "TotalITBISRetenido" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBISRetenido').text = value
            __logger.info(f'TotalITBISRetenido : {value}')
            break

    # Write TotalISRRetencion
    search_text = "TotalISRRetencion" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalISRRetencion').text = value
            __logger.info(f'TotalISRRetencion : {value}')
            break

    # Write TotalITBISPercepcion
    search_text = "TotalITBISPercepcion" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBISPercepcion').text = value
            __logger.info(f'TotalITBISPercepcion : {value}')
            break

    # Write TotalISRPercepcion
    search_text = "TotalISRPercepcion" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalISRPercepcion').text = value
            __logger.info(f'TotalISRPercepcion : {value}')
            break

    OtraMoneda = ET.SubElement(encabezado, 'OtraMoneda')
  
    # Write TipoMoneda
    search_text = "TipoMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TipoMoneda').text = value
            __logger.info(f'TipoMoneda : {value}')
            break
  
    # Write TipoCambio
    search_text = "TipoCambio" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TipoCambio').text = value
            __logger.info(f'TipoCambio : {value}')
            break
  
    # Write MontoGravadoTotalOtraMoneda
    search_text = "MontoGravadoTotalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravadoTotalOtraMoneda').text = value
            __logger.info(f'MontoGravadoTotalOtraMoneda : {value}')
            break
  
    # Write MontoGravado1OtraMoneda
    search_text = "MontoGravado1OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado1OtraMoneda').text = value
            __logger.info(f'MontoGravado1OtraMoneda : {value}')
            break
  
    # Write MontoGravado2OtraMoneda
    search_text = "MontoGravado2OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado2OtraMoneda').text = value
            __logger.info(f'MontoGravado2OtraMoneda : {value}')
            break
  
    # Write MontoGravado3OtraMoneda
    search_text = "MontoGravado3OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado3OtraMoneda').text = value
            __logger.info(f'MontoGravado3OtraMoneda : {value}')
            break
  
    # Write MontoExentoOtraMoneda
    search_text = "MontoExentoOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoExentoOtraMoneda').text = value
            __logger.info(f'MontoExentoOtraMoneda : {value}')
            break
  
    # Write TotalITBISOtraMoneda
    search_text = "TotalITBISOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBISOtraMoneda').text = value
            __logger.info(f'TotalITBISOtraMoneda : {value}')
            break

    # Write TotalITBIS1OtraMoneda
    search_text = "TotalITBIS1OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS1OtraMoneda').text = value
            __logger.info(f'TotalITBIS1OtraMoneda : {value}')
            break

    # Write TotalITBIS2OtraMoneda
    search_text = "TotalITBIS2OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS2OtraMoneda').text = value
            __logger.info(f'TotalITBIS2OtraMoneda : {value}')
            break

    # Write TotalITBIS3OtraMoneda
    search_text = "TotalITBIS3OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS3OtraMoneda').text = value
            __logger.info(f'TotalITBIS3OtraMoneda : {value}')
            break

    # Write MontoImpuestoAdicionalOtraMoneda
    search_text = "MontoImpuestoAdicionalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoImpuestoAdicionalOtraMoneda').text = value
            __logger.info(f'MontoImpuestoAdicionalOtraMoneda : {value}')
            break

    ImpuestosAdicionalesOtraMoneda = ET.SubElement(OtraMoneda, 'ImpuestosAdicionalesOtraMoneda')

    search_text = "TipoImpuestoOtraMoneda[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            
            ImpuestoAdicionalOtraMoneda_count = 4
            col_index = col
            while True :
                ImpuestoAdicionalOtraMoneda_count -= 1
                if ImpuestoAdicionalOtraMoneda_count < 0 :
                    break

                ImpuestoAdicionalOtraMoneda = ET.SubElement(ImpuestosAdicionalesOtraMoneda, 'ImpuestoAdicionalOtraMoneda')

                # Write TipoImpuestoOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'TipoImpuestoOtraMoneda').text = value
                __logger.info(f'TipoImpuestoOtraMoneda : {value}')
                col_index +=1
               
                # Write TasaImpuestoAdicionalOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'TasaImpuestoAdicionalOtraMoneda').text = value
                __logger.info(f'TasaImpuestoAdicionalOtraMoneda : {value}')
                col_index +=1

                # Write MontoImpuestoSelectivoConsumoEspecificoOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'MontoImpuestoSelectivoConsumoEspecificoOtraMoneda').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoEspecificoOtraMoneda : {value}')
                col_index +=1

                # Write MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda : {value}')
                col_index +=1

                # Write OtrosImpuestosAdicionalesOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'OtrosImpuestosAdicionalesOtraMoneda').text = value
                __logger.info(f'OtrosImpuestosAdicionalesOtraMoneda : {value}')

    # Write MontoTotalOtraMoneda
    search_text = "MontoTotalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoTotalOtraMoneda').text = value
            __logger.info(f'MontoTotalOtraMoneda : {value}')
            break

    # """
    DetallesItems = ET.SubElement(root, 'DetallesItems')

    item_count = 62
    search_text = "NumeroLinea[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            col_index = col
            while True:
                item_count -= 1
                if item_count < 0 : 
                    break
                Item = ET.SubElement(DetallesItems, 'Item')

                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'NumeroLinea').text = value
                __logger.info(f'NumeroLinea : {value}')
                col_index += 1
                

                # TablaCodigosItem
                TablaCodigosItem = ET.SubElement(Item, 'TablaCodigosItem')
                CodigosItem_count = 5;
                while True : 
                    CodigosItem_count -= 1
                    if CodigosItem_count < 0 : 
                        break

                    CodigosItem = ET.SubElement(TablaCodigosItem, 'CodigosItem')
                    
                    # TipoCodigo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(CodigosItem, 'TipoCodigo').text = value
                    __logger.info(f'TipoCodigo : {value}')
                    col_index += 1

                    # CodigoItem
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(CodigosItem, 'CodigoItem').text = value
                    __logger.info(f'CodigoItem : {value}')
                    col_index += 1

                # IndicadorFacturacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'IndicadorFacturacion').text = value
                __logger.info(f'IndicadorFacturacion : {value}')
                col_index += 1

                Retencion = ET.SubElement(Item, 'Retencion')
                # IndicadorAgenteRetencionoPercepcion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'IndicadorAgenteRetencionoPercepcion').text = value
                __logger.info(f'IndicadorAgenteRetencionoPercepcion : {value}')
                col_index += 1
                
                # MontoITBISRetenido
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'MontoITBISRetenido').text = value
                __logger.info(f'MontoITBISRetenido : {value}')
                col_index += 1

                # MontoISRRetenido
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'MontoISRRetenido').text = value
                __logger.info(f'MontoISRRetenido : {value}')
                col_index += 1

                # NombreItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'NombreItem').text = value
                __logger.info(f'NombreItem : {value}')
                col_index += 1

                # IndicadorBienoServicio
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'IndicadorBienoServicio').text = value
                __logger.info(f'IndicadorBienoServicio : {value}')
                col_index += 1

                # DescripcionItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'DescripcionItem').text = value
                __logger.info(f'DescripcionItem : {value}')
                col_index += 1

                # CantidadItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'CantidadItem').text = value
                __logger.info(f'CantidadItem : {value}')
                col_index += 1

                # UnidadMedida
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'UnidadMedida').text = value
                __logger.info(f'UnidadMedida : {value}')
                col_index += 1

                # CantidadReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'CantidadReferencia').text = value
                __logger.info(f'CantidadReferencia : {value}')
                col_index += 1

                # UnidadReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'UnidadReferencia').text = value
                __logger.info(f'UnidadReferencia : {value}')
                col_index += 1

                TablaSubcantidad = ET.SubElement(Item, 'TablaSubcantidad')

                SubcantidadItem_count = 5
                while True : 
                    SubcantidadItem_count -= 1
                    if SubcantidadItem_count < 0 :
                        break

                    SubcantidadItem = ET.SubElement(TablaSubcantidad, 'SubcantidadItem')
                    
                    # Subcantidad
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubcantidadItem, 'Subcantidad').text = value
                    __logger.info(f'Subcantidad : {value}')
                    col_index += 1
                    
                    # CodigoSubcantidad
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubcantidadItem, 'CodigoSubcantidad').text = value
                    __logger.info(f'CodigoSubcantidad : {value}')
                    col_index += 1

                # GradosAlcohol
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'GradosAlcohol').text = value
                __logger.info(f'GradosAlcohol : {value}')
                col_index += 1

                # PrecioUnitarioReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'PrecioUnitarioReferencia').text = value
                __logger.info(f'PrecioUnitarioReferencia : {value}')
                col_index += 1

                # FechaElaboracion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'FechaElaboracion').text = value
                __logger.info(f'FechaElaboracion : {value}')
                col_index += 1

                # FechaVencimientoItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'FechaVencimientoItem').text = value
                __logger.info(f'FechaVencimientoItem : {value}')
                col_index += 1

                
                # Mineria = ET.SubElement(Item, 'Mineria')
                
                # PesoNetoKilogramo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                # ET.SubElement(Mineria, 'PesoNetoKilogramo').text = value
                __logger.info(f'PesoNetoKilogramo : {value}')
                col_index += 1
                
                # PesoNetoMineria
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                # ET.SubElement(Mineria, 'PesoNetoMineria').text = value
                __logger.info(f'PesoNetoMineria : {value}')
                col_index += 1

                # TipoAfiliacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                # ET.SubElement(Mineria, 'TipoAfiliacion').text = value
                __logger.info(f'TipoAfiliacion : {value}')
                col_index += 1

                # Liquidacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                # ET.SubElement(Mineria, 'Liquidacion').text = value
                __logger.info(f'Liquidacion : {value}')
                col_index += 1

                
                # PrecioUnitarioItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(Item, 'PrecioUnitarioItem').text = value
                __logger.info(f'PrecioUnitarioItem : {value}')
                col_index += 1

                # DescuentoMonto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'DescuentoMonto').text = value
                __logger.info(f'DescuentoMonto : {value}')
                col_index += 1

                TablaSubDescuento = ET.SubElement(Item, 'TablaSubDescuento')

                SubDescuento_count = 5

                while True : 
                    SubDescuento_count -= 1
                    if SubDescuento_count < 0 : 
                        break

                    #SubDescuento
                    SubDescuento = ET.SubElement(TablaSubDescuento, 'SubDescuento')

                    # TipoSubDescuento
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'TipoSubDescuento').text = value
                    __logger.info(f'TipoSubDescuento : {value}')
                    col_index += 1

                    # SubDescuentoPorcentaje
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'SubDescuentoPorcentaje').text = value
                    __logger.info(f'SubDescuentoPorcentaje : {value}')
                    col_index += 1

                    # MontoSubDescuento
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'MontoSubDescuento').text = value
                    __logger.info(f'MontoSubDescuento : {value}')
                    col_index += 1


                # RecargoMonto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'RecargoMonto').text = value
                __logger.info(f'RecargoMonto : {value}')
                col_index += 1

                TablaSubRecargo = ET.SubElement(Item, 'TablaSubRecargo')

                SubRecargo_count = 5

                while True :
                    SubRecargo_count -= 1
                    if SubRecargo_count < 0 :
                        break    
                    SubRecargo = ET.SubElement(TablaSubRecargo, 'SubRecargo')

                    # TipoSubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'TipoSubRecargo').text = value
                    __logger.info(f'TipoSubRecargo : {value}')
                    col_index += 1

                    # SubRecargoPorcentaje
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'SubRecargoPorcentaje').text = value
                    __logger.info(f'SubRecargoPorcentaje : {value}')
                    col_index += 1

                    
                    # MontosubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'MontosubRecargo').text = value
                    __logger.info(f'MontosubRecargo : {value}')
                    col_index += 1

                TablaImpuestoAdicional = ET.SubElement(Item, 'TablaImpuestoAdicional')

                ImpuestoAdicional_count = 2
                while True :
                    ImpuestoAdicional_count -= 1
                    if ImpuestoAdicional_count < 0 : 
                        break
                    ImpuestoAdicional =  ET.SubElement(TablaImpuestoAdicional, 'ImpuestoAdicional')                                   
                    # MontosubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = value
                    __logger.info(f'TipoImpuesto : {value}')
                    col_index += 1
                
                OtraMonedaDetalle = ET.SubElement(Item, 'OtraMonedaDetalle')

                # PrecioOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'PrecioOtraMoneda').text = value
                __logger.info(f'PrecioOtraMoneda : {value}')
                col_index += 1

                # DescuentoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'DescuentoOtraMoneda').text = value
                __logger.info(f'DescuentoOtraMoneda : {value}')
                col_index += 1

                # RecargoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'RecargoOtraMoneda').text = value
                __logger.info(f'RecargoOtraMoneda : {value}')

                col_index += 1
                # MontoItemOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'MontoItemOtraMoneda').text = value
                __logger.info(f'MontoItemOtraMoneda : {value}')
                col_index += 1
                
                # MontoItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'MontoItem').text = value
                __logger.info(f'MontoItem : {value}')
                col_index += 1
    
    Subtotales = ET.SubElement(root, 'Subtotales')
    Subtotal = ET.SubElement(Subtotales, 'Subtotal')

    # Write NumeroSubTotal
    search_text = "NumeroSubTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'NumeroSubTotal').text = value
            __logger.info(f'Version : {cell_value}')
            break
    
    # Write DescripcionSubtotal
    search_text = "DescripcionSubtotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'DescripcionSubtotal').text = value
            __logger.info(f'DescripcionSubtotal : {cell_value}')
            break
    
    # Write Orden
    search_text = "Orden" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'Orden').text = value
            __logger.info(f'Orden : {cell_value}')
            break
    
    # Write SubTotalMontoGravadoTotal
    search_text = "SubTotalMontoGravadoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoTotal').text = value
            __logger.info(f'SubTotalMontoGravadoTotal : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI1
    search_text = "SubTotalMontoGravadoI1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI1').text = value
            __logger.info(f'SubTotalMontoGravadoI1 : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI2
    search_text = "SubTotalMontoGravadoI2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI2').text = value
            __logger.info(f'SubTotalMontoGravadoI2 : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI3
    search_text = "SubTotalMontoGravadoI3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI3').text = value
            __logger.info(f'SubTotalMontoGravadoI3 : {cell_value}')
            break
  
    # Write SubTotaITBIS
    search_text = "SubTotaITBIS" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS').text = value
            __logger.info(f'SubTotaITBIS : {cell_value}')
            break

    # Write SubTotaITBIS1
    search_text = "SubTotaITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS1').text = value
            __logger.info(f'SubTotaITBIS1 : {cell_value}')
            break

    # Write SubTotaITBIS2
    search_text = "SubTotaITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS2').text = value
            __logger.info(f'SubTotaITBIS2 : {cell_value}')
            break

    # Write SubTotaITBIS3
    search_text = "SubTotaITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS3').text = value
            __logger.info(f'SubTotaITBIS3 : {cell_value}')
            break

    # Write SubTotalImpuestoAdicional
    search_text = "SubTotalImpuestoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalImpuestoAdicional').text = value
            __logger.info(f'SubTotalImpuestoAdicional : {cell_value}')
            break

    # Write SubTotalExento
    search_text = "SubTotalExento" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalExento').text = value
            __logger.info(f'SubTotalExento : {cell_value}')
            break

    # Write MontoSubTotal
    search_text = "MontoSubTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'MontoSubTotal').text = value
            __logger.info(f'MontoSubTotal : {cell_value}')
            break

    # Write Lineas
    search_text = "Lineas" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'Lineas').text = value
            __logger.info(f'Lineas : {cell_value}')
            break

    DescuentosORecargos = ET.SubElement(root, 'DescuentosORecargos')
    DescuentoORecargo_count = 2

    search_text = "NumeroLineaDoR[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        
        if cell_value == search_text:
            col_index = col
            while True : 
                DescuentoORecargo_count -= 1
                if DescuentoORecargo_count < 0:
                    break

                DescuentoORecargo = ET.SubElement(DescuentosORecargos, 'DescuentoORecargo')

                # NumeroLinea
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'NumeroLinea').text = value
                __logger.info(f'NumeroLinea : {cell_value}')
                col_index +=1

                # TipoAjuste
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'TipoAjuste').text = value
                __logger.info(f'TipoAjuste : {cell_value}')
                col_index +=1
                
                # IndicadorNorma1007
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'IndicadorNorma1007').text = value
                __logger.info(f'IndicadorNorma1007 : {cell_value}')
                col_index +=1
                         
                # DescripcionDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'DescripcionDescuentooRecargo').text = value
                __logger.info(f'DescripcionDescuentooRecargo : {cell_value}')
                col_index +=1
                                         
                # TipoValor
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'TipoValor').text = value
                __logger.info(f'TipoValor : {cell_value}')
                col_index +=1
                                                     
                # ValorDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'ValorDescuentooRecargo').text = value
                __logger.info(f'ValorDescuentooRecargo : {cell_value}')
                col_index +=1
                                                               
                # MontoDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'MontoDescuentooRecargo').text = value
                __logger.info(f'MontoDescuentooRecargo : {cell_value}')
                col_index +=1
                                                                        
                # MontoDescuentooRecargoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'MontoDescuentooRecargoOtraMoneda').text = value
                __logger.info(f'MontoDescuentooRecargoOtraMoneda : {cell_value}')
                col_index +=1        

                # IndicadorFacturacionDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'IndicadorFacturacionDescuentooRecargo').text = value
                __logger.info(f'IndicadorFacturacionDescuentooRecargo : {cell_value}')         
     

    Paginacion = ET.SubElement(root, 'Paginacion')
    Pagina_count = 2

    search_text = "PaginaNo[1]" 
    col_index : int
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:

            col_index = col
                
            while True:
                Pagina_count -= 1
                if Pagina_count < 0:
                    break

                Pagina = ET.SubElement(Paginacion, 'Pagina')

                # PaginNo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'PaginaNo').text = value
                __logger.info(f'PaginaNo : {cell_value}')
                __logger.info(f'PaginaNo col_index : {col_index}')
                col_index += 1

                # NoLineaDesde
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'NoLineaDesde').text = value
                __logger.info(f'NoLineaDesde : {cell_value}')
                col_index += 1

                # NoLineaHasta
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'NoLineaHasta').text = value
                __logger.info(f'NoLineaHasta : {value}')
                col_index += 1
                
                # SubtotalMontoGravadoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravadoPagina').text = value
                __logger.info(f'SubtotalMontoGravadoPagina : {value}')
                col_index += 1
                                
                # SubtotalMontoGravado1Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado1Pagina').text = value
                __logger.info(f'SubtotalMontoGravado1Pagina : {value}')
                col_index += 1
                                                
                # SubtotalMontoGravado2Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado2Pagina').text = value
                __logger.info(f'SubtotalMontoGravado2Pagina : {value}')
                col_index += 1
                                                   
                # SubtotalMontoGravado3Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado3Pagina').text = value
                __logger.info(f'SubtotalMontoGravado3Pagina : {value}')
                col_index += 1
                                                                   
                # SubtotalExentoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalExentoPagina').text = value
                __logger.info(f'SubtotalExentoPagina : {value}')
                col_index += 1
                                                               
                # SubtotalItbisPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbisPagina').text = value
                __logger.info(f'SubtotalItbisPagina : {value}')
                col_index += 1
                                                                      
                # SubtotalItbis1Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis1Pagina').text = value
                __logger.info(f'SubtotalItbis1Pagina : {value}')
                col_index += 1
                                                                                      
                # SubtotalItbis2Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis2Pagina').text = value
                __logger.info(f'SubtotalItbis2Pagina : {value}')
                col_index += 1
                                                                                         
                # SubtotalItbis3Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis3Pagina').text = value
                __logger.info(f'SubtotalItbis3Pagina : {value}')
                col_index += 1
                                                                                                         
                # SubtotalImpuestoAdicionalPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalImpuestoAdicionalPagina').text = value
                __logger.info(f'SubtotalImpuestoAdicionalPagina : {value}')
                col_index += 1

                SubtotalImpuestoAdicional = ET.SubElement(Pagina, 'SubtotalImpuestoAdicional')
                
                # SubtotalImpuestoSelectivoConsumoEspecificoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(SubtotalImpuestoAdicional, 'SubtotalImpuestoSelectivoConsumoEspecificoPagina').text =value
                __logger.info(f'SubtotalImpuestoSelectivoConsumoEspecificoPagina : {value}')
                col_index += 1

                # SubtotalOtrosImpuesto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(SubtotalImpuestoAdicional, 'SubtotalOtrosImpuesto').text =value
                __logger.info(f'SubtotalOtrosImpuesto : {value}')
                col_index += 1

                # MontoSubtotalPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(Pagina, 'MontoSubtotalPagina').text =value
                __logger.info(f'MontoSubtotalPagina : {value}')
                col_index += 1

                # SubtotalMontoNoFacturablePagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(Pagina, 'SubtotalMontoNoFacturablePagina').text =value
                __logger.info(f'SubtotalMontoNoFacturablePagina : {value}')
                __logger.info(f'SubtotalMontoNoFacturablePagina _ index : {col_index}')
                # col_index += 1

    InformacionReferencia = ET.SubElement(root, 'InformacionReferencia')
    
    # NCFModificado
    col_index += 1
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'NCFModificado').text =value
    __logger.info(f'NCFModificado : {value}')
    col_index += 1
        
    # RNCOtroContribuyente
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'RNCOtroContribuyente').text =value
    __logger.info(f'RNCOtroContribuyente : {value}')
    col_index += 1
        
    # FechaNCFModificado
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'FechaNCFModificado').text =value
    __logger.info(f'FechaNCFModificado : {value}')
    col_index += 1
        
    # CodigoModificacion
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'CodigoModificacion').text =value
    __logger.info(f'CodigoModificacion : {value}')
    col_index += 1

    ET.SubElement(root, 'FechaHoraFirma').text =""
    

    rough_string = ET.tostring(root, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    pretty_xml_as_str = reparsed.toprettyxml(indent="  ", encoding='utf-8')

    path = os.path.join(os.path.dirname(__file__), 'data/row_invoice.xml')
    
    with open(path, 'wb') as f:
        f.write(pretty_xml_as_str)
    
    xml_str = ET.tostring(root, encoding='utf-8').decode('utf-8')
    return xml_str

def create_e_cf_32(in_row : int) :

    """Generate DGII-compliant XML from an Odoo invoice."""
    # Create root element with namespaces
    root = ET.Element('ECF', {
         'xmlns:xs' : "http://www.w3.org/2001/XMLSchema",
    })

    # 1. Encabezado
    encabezado = ET.SubElement(root, 'Encabezado')

    # Write Version
    search_text = "Version" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value != "#e" :    
                ET.SubElement(encabezado, 'Version').text = '1.1'
                __logger.info(f'Version : {value}')
                __logger.info(f'Version Col Index : {col}')
            break

    # IdDoc
    id_doc = ET.SubElement(encabezado, 'IdDoc')

    # Write TipoeCF
    search_text = "TipoeCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 
            ET.SubElement(id_doc, 'TipoeCF').text = value
            __logger.info(f'TipoeCF : {cell_value}')
            break

    # Write eNCF
    search_text = "ENCF" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 
            ET.SubElement(id_doc, 'eNCF').text = value
            __logger.info(f'eNCF : {cell_value}')
            break

    # Write FechaVencimientoSecuencia
    search_text = "FechaVencimientoSecuencia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaVencimientoSecuencia').text = value
            __logger.info(f'FechaVencimientoSecuencia : {value}')
            break

    # Write IndicadorEnvioDiferido
    search_text = "IndicadorEnvioDiferido" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'IndicadorEnvioDiferido').text = value
            __logger.info(f'IndicadorEnvioDiferido : {value}')
            break

    # Write IndicadorMontoGravado
    search_text = "IndicadorMontoGravado" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'IndicadorMontoGravado').text = value
            __logger.info(f'IndicadorMontoGravado : {value}')
            break

    # Write TipoIngresos
    search_text = "TipoIngresos" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoIngresos').text = value
            __logger.info(f'TipoIngresos : {value}')
            break

    # Write TipoPago
    search_text = "TipoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoPago').text = value
            __logger.info(f'TipoPago : {value}')
            break

    # Write FechaLimitePago
    search_text = "FechaLimitePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaLimitePago').text = value
            __logger.info(f'FechaLimitePago : {value}')
            break


    # Write TerminoPago
    search_text = "TerminoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TerminoPago').text = value
            __logger.info(f'TerminoPago : {value}')
            break
    
    
    TablaFormasPago = ET.SubElement(id_doc, 'TablaFormasPago')

    search_text = "FormaPago[1]"
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            FormaDePago_count = 7
            col_index = col
            while True :
                FormaDePago_count -= 1
                if FormaDePago_count < 0 :
                    break            
                FormaDePago = ET.SubElement(TablaFormasPago, 'FormaDePago')

                value = str(sheet.cell(in_row, column= col_index ).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(FormaDePago, 'FormaPago').text = value
                col_index += 1

                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(FormaDePago, 'MontoPago').text = value
                col_index += 1
    
    # Write TipoCuentaPago
    search_text = "TipoCuentaPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TipoCuentaPago').text = value
            __logger.info(f'TipoCuentaPago : {value}')
            break
        
    # Write NumeroCuentaPago
    search_text = "NumeroCuentaPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'NumeroCuentaPago').text = value
            __logger.info(f'NumeroCuentaPago : {value}')
            break
        
    # Write BancoPago
    search_text = "BancoPago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'BancoPago').text = value
            __logger.info(f'BancoPago : {value}')
            break

    # Write FechaDesde
    search_text = "FechaDesde" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaDesde').text = value
            __logger.info(f'FechaDesde : {value}')
            break    

    # Write FechaHasta
    search_text = "FechaHasta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'FechaHasta').text = value
            __logger.info(f'FechaHasta : {value}')
            break    

    # Write TotalPaginas
    search_text = "TotalPaginas" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(id_doc, 'TotalPaginas').text = value
            __logger.info(f'TotalPaginas : {value}')
            break

    Emisor = ET.SubElement(encabezado, 'Emisor')

    # Write RNCEmisor
    search_text = "RNCEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RNCEmisor').text = value
            __logger.info(f'RNCEmisor : {value}')
            break

    # Write RazonSocialEmisor
    search_text = "RazonSocialEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RazonSocialEmisor').text = value
            __logger.info(f'RazonSocialEmisor : {value}')
            break

    # Write NombreComercial
    search_text = "NombreComercial" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NombreComercial').text = value
            __logger.info(f'NombreComercial : {value}')
            break

    # Write Sucursal
    search_text = "Sucursal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Sucursal').text = value
            __logger.info(f'Sucursal : {value}')
            break

    # Write DireccionEmisor
    search_text = "DireccionEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'DireccionEmisor').text = value
            __logger.info(f'DireccionEmisor : {value}')
            break

    # Write Municipio
    search_text = "Municipio" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Municipio').text = value
            __logger.info(f'Municipio : {value}')
            break

    # Write Provincia
    search_text = "Provincia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'Provincia').text = value
            __logger.info(f'Provincia : {value}')
            break

    TablaTelefonoEmisor = ET.SubElement(Emisor, 'TablaTelefonoEmisor')

    # Write TelefonoEmisor
    search_text = "TelefonoEmisor[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            TelefonoEmisor_count = 3
            col_index = col
            while True :
                TelefonoEmisor_count -= 1
                if TelefonoEmisor_count < 0 : 
                    break

                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""
                ET.SubElement(TablaTelefonoEmisor, 'TelefonoEmisor').text = value           
                __logger.info(f'TelefonoEmisor : {value}')

                col_index +=1
             
    # Write CorreoEmisor
    search_text = "CorreoEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'CorreoEmisor').text = value
            __logger.info(f'CorreoEmisor : {value}')
            break

    # Write WebSite
    search_text = "WebSite" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'WebSite').text = value
            __logger.info(f'WebSite : {value}')
            break

    # Write ActividadEconomica
    search_text = "ActividadEconomica" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'ActividadEconomica').text = value
            __logger.info(f'ActividadEconomica : {value}')
            break

    # Write CodigoVendedor
    search_text = "CodigoVendedor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'CodigoVendedor').text = value
            __logger.info(f'CodigoVendedor : {value}')
            break

    # Write NumeroFacturaInterna
    search_text = "NumeroFacturaInterna" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NumeroFacturaInterna').text = value
            __logger.info(f'NumeroFacturaInterna : {value}')
            break
 
    # Write NumeroPedidoInterno
    search_text = "NumeroPedidoInterno" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'NumeroPedidoInterno').text = value
            __logger.info(f'NumeroPedidoInterno : {value}')
            break
 
    # Write ZonaVenta
    search_text = "ZonaVenta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'ZonaVenta').text = value
            __logger.info(f'ZonaVenta : {value}')
            break

    # Write RutaVenta
    search_text = "RutaVenta" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'RutaVenta').text = value
            __logger.info(f'RutaVenta : {value}')
            break

    # Write InformacionAdicionalEmisor
    search_text = "InformacionAdicionalEmisor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'InformacionAdicionalEmisor').text = value
            __logger.info(f'InformacionAdicionalEmisor : {value}')
            break

    # Write FechaEmision
    search_text = "FechaEmision" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Emisor, 'FechaEmision').text = value
            __logger.info(f'FechaEmision : {value}')
            break

    Comprador = ET.SubElement(encabezado, 'Comprador')

    # Write RNCComprador
    search_text = "RNCComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'RNCComprador').text = value
            __logger.info(f'RNCComprador : {value}')
            break

    # Write IdentificadorExtranjero
    search_text = "IdentificadorExtranjero" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'IdentificadorExtranjero').text = value
            __logger.info(f'IdentificadorExtranjero : {value}')
            break

    # Write RazonSocialComprador
    search_text = "RazonSocialComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'RazonSocialComprador').text = value
            __logger.info(f'RazonSocialComprador : {value}')
            break
   
    # Write ContactoComprador
    search_text = "ContactoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ContactoComprador').text = value
            __logger.info(f'ContactoComprador : {value}')
            break

    # Write CorreoComprador
    search_text = "CorreoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'CorreoComprador').text = value
            __logger.info(f'CorreoComprador : {value}')
            break

    # Write DireccionComprador
    search_text = "DireccionComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'DireccionComprador').text = value
            __logger.info(f'DireccionComprador : {value}')
            break

    # Write MunicipioComprador
    search_text = "MunicipioComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'MunicipioComprador').text = value
            __logger.info(f'MunicipioComprador : {value}')
            break

    # Write ProvinciaComprador
    search_text = "ProvinciaComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ProvinciaComprador').text = value
            __logger.info(f'ProvinciaComprador : {value}')
            break

    # Write FechaEntrega
    search_text = "FechaEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'FechaEntrega').text = value
            __logger.info(f'FechaEntrega : {value}')
            break

    # Write ContactoEntrega
    search_text = "ContactoEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ContactoEntrega').text = value
            __logger.info(f'ContactoEntrega : {value}')
            break

    # Write DireccionEntrega
    search_text = "DireccionEntrega" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'DireccionEntrega').text = value
            __logger.info(f'DireccionEntrega : {value}')
            break

    # Write TelefonoAdicional
    search_text = "TelefonoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'TelefonoAdicional').text = value
            __logger.info(f'TelefonoAdicional : {value}')
            break

    # Write FechaOrdenCompra
    search_text = "FechaOrdenCompra" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'FechaOrdenCompra').text = value
            __logger.info(f'FechaOrdenCompra : {value}')
            break

    # Write NumeroOrdenCompra
    search_text = "NumeroOrdenCompra" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'NumeroOrdenCompra').text = value
            __logger.info(f'NumeroOrdenCompra : {value}')
            break

    # Write CodigoInternoComprador
    search_text = "CodigoInternoComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'CodigoInternoComprador').text = value
            __logger.info(f'CodigoInternoComprador : {value}')
            break

    # Write ResponsablePago
    search_text = "ResponsablePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'ResponsablePago').text = value
            __logger.info(f'ResponsablePago : {value}')
            break
    
    # Write InformacionAdicionalComprador
    search_text = "InformacionAdicionalComprador" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Comprador, 'InformacionAdicionalComprador').text = value
            __logger.info(f'InformacionAdicionalComprador : {value}')
            break
    
    InformacionesAdicionales = ET.SubElement(encabezado, 'InformacionesAdicionales')
   
    # Write FechaEmbarque
    search_text = "FechaEmbarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'FechaEmbarque').text = value
            __logger.info(f'FechaEmbarque : {value}')
            break

    # Write NumeroEmbarque
    search_text = "NumeroEmbarque" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NumeroEmbarque').text = value
            __logger.info(f'NumeroEmbarque : {value}')
            break

    # # Write NumeroContenedor
    # search_text = "NumeroContenedor" 
    # for col in range(1, sheet.max_column + 1):
    #     cell_value = sheet.cell(1, column=col).value
    #     if cell_value == search_text:
    #         value = str(sheet.cell(in_row, column=col).value)
    #         if value == "#e" :
    #             value = "" 

    #         ET.SubElement(InformacionesAdicionales, 'NumeroContenedor').text = value
    #         __logger.info(f'NumeroContenedor : {value}')
    #         break

    value = str(sheet.cell(in_row, column=73).value)
    if value == "#e" :
        value = "" 

    ET.SubElement(InformacionesAdicionales, 'NumeroContenedor').text = value
    __logger.info(f'NumeroContenedor : {value}')
    
    # Write NumeroReferencia
    search_text = "NumeroReferencia" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'NumeroReferencia').text = value
            __logger.info(f'NumeroReferencia : {value}')
            break

    # Write PesoBruto
    search_text = "PesoBruto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'PesoBruto').text = value
            __logger.info(f'PesoBruto : {value}')
            break

    # Write PesoNeto
    search_text = "PesoNeto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'PesoNeto').text = value
            __logger.info(f'PesoNeto : {value}')
            break

    # Write UnidadPesoBruto
    search_text = "UnidadPesoBruto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadPesoBruto').text = value
            __logger.info(f'UnidadPesoBruto : {value}')
            break

    # Write UnidadPesoNeto
    search_text = "UnidadPesoNeto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadPesoNeto').text = value
            __logger.info(f'UnidadPesoNeto : {value}')
            break

    # Write CantidadBulto
    search_text = "CantidadBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'CantidadBulto').text = value
            __logger.info(f'CantidadBulto : {value}')
            break

    # Write UnidadBulto
    search_text = "UnidadBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadBulto').text = value
            __logger.info(f'UnidadBulto : {value}')
            break

    # Write VolumenBulto
    search_text = "VolumenBulto" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'VolumenBulto').text = value
            __logger.info(f'VolumenBulto : {value}')
            break

    # Write UnidadVolumen
    search_text = "UnidadVolumen" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(InformacionesAdicionales, 'UnidadVolumen').text = value
            __logger.info(f'UnidadVolumen : {value}')
            break

    # Transporte
    Transporte = ET.SubElement(encabezado, 'Transporte')

    # Write Conductor
    search_text = "Conductor" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Conductor').text = value
            __logger.info(f'Conductor : {value}')
            break

    # Write DocumentoTransporte
    search_text = "DocumentoTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'DocumentoTransporte').text = value
            __logger.info(f'DocumentoTransporte : {value}')
            break

    # Write Ficha
    search_text = "Ficha" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Ficha').text = value
            __logger.info(f'Ficha : {value}')
            break

    # Write Placa
    search_text = "Placa" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'Placa').text = value
            __logger.info(f'Placa : {value}')
            break

    # Write RutaTransporte
    search_text = "RutaTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'RutaTransporte').text = value
            __logger.info(f'RutaTransporte : {value}')
            break

    # Write ZonaTransporte
    search_text = "ZonaTransporte" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'ZonaTransporte').text = value
            __logger.info(f'ZonaTransporte : {value}')
            break

    # Write NumeroAlbaran
    search_text = "NumeroAlbaran" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Transporte, 'NumeroAlbaran').text = value
            __logger.info(f'NumeroAlbaran : {value}')
            break

    Totales = ET.SubElement(encabezado, 'Totales')
   
    # Write MontoGravadoTotal
    search_text = "MontoGravadoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoTotal').text = value
            __logger.info(f'MontoGravadoTotal : {value}')
            break

    # Write MontoGravadoI1
    search_text = "MontoGravadoI1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI1').text = value
            __logger.info(f'MontoGravadoI1 : {value}')
            break
   
    # Write MontoGravadoI2
    search_text = "MontoGravadoI2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI2').text = value
            __logger.info(f'MontoGravadoI2 : {value}')
            break
   
    # Write MontoGravadoI3
    search_text = "MontoGravadoI3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoGravadoI3').text = value
            __logger.info(f'MontoGravadoI3 : {value}')
            break
   
    # Write MontoExento
    search_text = "MontoExento" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoExento').text = value
            __logger.info(f'MontoExento : {value}')
            break
   
    # Write ITBIS1
    search_text = "ITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS1').text = value
            __logger.info(f'ITBIS1 : {value}')
            break
  
    # Write ITBIS2
    search_text = "ITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS2').text = value
            __logger.info(f'ITBIS2 : {value}')
            break
 
    # Write ITBIS3
    search_text = "ITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ITBIS3').text = value
            __logger.info(f'ITBIS3 : {value}')
            break
 
    # Write TotalITBIS
    search_text = "TotalITBIS" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS').text = value
            __logger.info(f'TotalITBIS : {value}')
            break
 
    # Write TotalITBIS1
    search_text = "TotalITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS1').text = value
            __logger.info(f'TotalITBIS1 : {value}')
            break

    # Write TotalITBIS2
    search_text = "TotalITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS2').text = value
            __logger.info(f'TotalITBIS2 : {value}')
            break

    # Write TotalITBIS3
    search_text = "TotalITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'TotalITBIS3').text = value
            __logger.info(f'TotalITBIS3 : {value}')
            break

    # Write MontoImpuestoAdicional
    search_text = "MontoImpuestoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoImpuestoAdicional').text = value
            __logger.info(f'MontoImpuestoAdicional : {value}')
            break

    ImpuestosAdicionales = ET.SubElement(Totales, 'ImpuestosAdicionales')
    search_text = "TipoImpuesto[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:

            col_index = col
            ImpuestoAdicional_count = 4
            while True:
                ImpuestoAdicional_count -= 1
                if ImpuestoAdicional_count < 0:
                    break

                ImpuestoAdicional = ET.SubElement(ImpuestosAdicionales, 'ImpuestoAdicional')

                # TipoImpuesto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = value
                __logger.info(f'TipoImpuesto : {value}')
                col_index +=1

                # TasaImpuestoAdicional
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'TasaImpuestoAdicional').text = value
                __logger.info(f'TasaImpuestoAdicional : {value}')
                col_index +=1

                # MontoImpuestoSelectivoConsumoEspecifico
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoEspecifico').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoEspecifico : {value}')
                col_index +=1

                # MontoImpuestoSelectivoConsumoAdvalorem
                value = str(sheet.cell(in_row, column=col_index + 3).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'MontoImpuestoSelectivoConsumoAdvalorem').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoAdvalorem : {value}')
                col_index +=1


                # OtrosImpuestosAdicionales
                value = str(sheet.cell(in_row, column=col_index + 4).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(ImpuestoAdicional, 'OtrosImpuestosAdicionales').text = value
                __logger.info(f'OtrosImpuestosAdicionales : {value}')

    # Write MontoTotal
    search_text = "MontoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoTotal').text = value
            __logger.info(f'MontoTotal : {value}')
            break

    # Write MontoNoFacturable
    search_text = "MontoNoFacturable" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoNoFacturable').text = value
            __logger.info(f'MontoNoFacturable : {value}')
            break

    # Write MontoPeriodo
    search_text = "MontoPeriodo" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoPeriodo').text = value
            __logger.info(f'MontoPeriodo : {value}')
            break

    # Write SaldoAnterior
    search_text = "SaldoAnterior" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'SaldoAnterior').text = value
            __logger.info(f'SaldoAnterior : {value}')
            break

    # Write MontoAvancePago
    search_text = "MontoAvancePago" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'MontoAvancePago').text = value
            __logger.info(f'MontoAvancePago : {value}')
            break

    # Write ValorPagar
    search_text = "ValorPagar" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(Totales, 'ValorPagar').text = value
            __logger.info(f'ValorPagar : {value}')
            break

    OtraMoneda = ET.SubElement(encabezado, 'OtraMoneda')
  
    # Write TipoMoneda
    search_text = "TipoMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TipoMoneda').text = value
            __logger.info(f'TipoMoneda : {value}')
            break
  
    # Write TipoCambio
    search_text = "TipoCambio" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TipoCambio').text = value
            __logger.info(f'TipoCambio : {value}')
            break
  
    # Write MontoGravadoTotalOtraMoneda
    search_text = "MontoGravadoTotalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravadoTotalOtraMoneda').text = value
            __logger.info(f'MontoGravadoTotalOtraMoneda : {value}')
            break
  
    # Write MontoGravado1OtraMoneda
    search_text = "MontoGravado1OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado1OtraMoneda').text = value
            __logger.info(f'MontoGravado1OtraMoneda : {value}')
            break
  
    # Write MontoGravado2OtraMoneda
    search_text = "MontoGravado2OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado2OtraMoneda').text = value
            __logger.info(f'MontoGravado2OtraMoneda : {value}')
            break
  
    # Write MontoGravado3OtraMoneda
    search_text = "MontoGravado3OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoGravado3OtraMoneda').text = value
            __logger.info(f'MontoGravado3OtraMoneda : {value}')
            break
  
    # Write MontoExentoOtraMoneda
    search_text = "MontoExentoOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoExentoOtraMoneda').text = value
            __logger.info(f'MontoExentoOtraMoneda : {value}')
            break
  
    # Write TotalITBISOtraMoneda
    search_text = "TotalITBISOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBISOtraMoneda').text = value
            __logger.info(f'TotalITBISOtraMoneda : {value}')
            break

    # Write TotalITBIS1OtraMoneda
    search_text = "TotalITBIS1OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS1OtraMoneda').text = value
            __logger.info(f'TotalITBIS1OtraMoneda : {value}')
            break

    # Write TotalITBIS2OtraMoneda
    search_text = "TotalITBIS2OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS2OtraMoneda').text = value
            __logger.info(f'TotalITBIS2OtraMoneda : {value}')
            break

    # Write TotalITBIS3OtraMoneda
    search_text = "TotalITBIS3OtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'TotalITBIS3OtraMoneda').text = value
            __logger.info(f'TotalITBIS3OtraMoneda : {value}')
            break

    # Write MontoImpuestoAdicionalOtraMoneda
    search_text = "MontoImpuestoAdicionalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoImpuestoAdicionalOtraMoneda').text = value
            __logger.info(f'MontoImpuestoAdicionalOtraMoneda : {value}')
            break

    ImpuestosAdicionalesOtraMoneda = ET.SubElement(OtraMoneda, 'ImpuestosAdicionalesOtraMoneda')

    search_text = "TipoImpuestoOtraMoneda[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            
            ImpuestoAdicionalOtraMoneda_count = 4
            col_index = col
            while True :
                ImpuestoAdicionalOtraMoneda_count -= 1
                if ImpuestoAdicionalOtraMoneda_count < 0 :
                    break

                ImpuestoAdicionalOtraMoneda = ET.SubElement(ImpuestosAdicionalesOtraMoneda, 'ImpuestoAdicionalOtraMoneda')

                # Write TipoImpuestoOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'TipoImpuestoOtraMoneda').text = value
                __logger.info(f'TipoImpuestoOtraMoneda : {value}')
                col_index +=1
               
                # Write TasaImpuestoAdicionalOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'TasaImpuestoAdicionalOtraMoneda').text = value
                __logger.info(f'TasaImpuestoAdicionalOtraMoneda : {value}')
                col_index +=1

                # Write MontoImpuestoSelectivoConsumoEspecificoOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'MontoImpuestoSelectivoConsumoEspecificoOtraMoneda').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoEspecificoOtraMoneda : {value}')
                col_index +=1

                # Write MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda').text = value
                __logger.info(f'MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda : {value}')
                col_index +=1

                # Write OtrosImpuestosAdicionalesOtraMoneda
                value = str(sheet.cell(in_row, column= col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(ImpuestoAdicionalOtraMoneda, 'OtrosImpuestosAdicionalesOtraMoneda').text = value
                __logger.info(f'OtrosImpuestosAdicionalesOtraMoneda : {value}')

    # Write MontoTotalOtraMoneda
    search_text = "MontoTotalOtraMoneda" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = "" 

            ET.SubElement(OtraMoneda, 'MontoTotalOtraMoneda').text = value
            __logger.info(f'MontoTotalOtraMoneda : {value}')
            break

    # """
    DetallesItems = ET.SubElement(root, 'DetallesItems')

    item_count = 62
    search_text = "NumeroLinea[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            col_index = col
            while True:
                item_count -= 1
                if item_count < 0 : 
                    break
                Item = ET.SubElement(DetallesItems, 'Item')

                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'NumeroLinea').text = value
                __logger.info(f'NumeroLinea : {value}')
                col_index += 1
                

                # TablaCodigosItem
                TablaCodigosItem = ET.SubElement(Item, 'TablaCodigosItem')
                CodigosItem_count = 5;
                while True : 
                    CodigosItem_count -= 1
                    if CodigosItem_count < 0 : 
                        break

                    CodigosItem = ET.SubElement(TablaCodigosItem, 'CodigosItem')
                    
                    # TipoCodigo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(CodigosItem, 'TipoCodigo').text = value
                    __logger.info(f'TipoCodigo : {value}')
                    col_index += 1

                    # CodigoItem
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(CodigosItem, 'CodigoItem').text = value
                    __logger.info(f'CodigoItem : {value}')
                    col_index += 1

                # IndicadorFacturacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'IndicadorFacturacion').text = value
                __logger.info(f'IndicadorFacturacion : {value}')
                col_index += 1

                Retencion = ET.SubElement(Item, 'Retencion')
                # IndicadorAgenteRetencionoPercepcion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'IndicadorAgenteRetencionoPercepcion').text = value
                __logger.info(f'IndicadorAgenteRetencionoPercepcion : {value}')
                col_index += 1
                
                # MontoITBISRetenido
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'MontoITBISRetenido').text = value
                __logger.info(f'MontoITBISRetenido : {value}')
                col_index += 1

                # MontoISRRetenido
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Retencion, 'MontoISRRetenido').text = value
                __logger.info(f'MontoISRRetenido : {value}')
                col_index += 1

                # NombreItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'NombreItem').text = value
                __logger.info(f'NombreItem : {value}')
                col_index += 1

                # IndicadorBienoServicio
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'IndicadorBienoServicio').text = value
                __logger.info(f'IndicadorBienoServicio : {value}')
                col_index += 1

                # DescripcionItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'DescripcionItem').text = value
                __logger.info(f'DescripcionItem : {value}')
                col_index += 1

                # CantidadItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'CantidadItem').text = value
                __logger.info(f'CantidadItem : {value}')
                col_index += 1

                # UnidadMedida
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'UnidadMedida').text = value
                __logger.info(f'UnidadMedida : {value}')
                col_index += 1

                # CantidadReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'CantidadReferencia').text = value
                __logger.info(f'CantidadReferencia : {value}')
                col_index += 1

                # UnidadReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'UnidadReferencia').text = value
                __logger.info(f'UnidadReferencia : {value}')
                col_index += 1

                TablaSubcantidad = ET.SubElement(Item, 'TablaSubcantidad')

                SubcantidadItem_count = 5
                while True : 
                    SubcantidadItem_count -= 1
                    if SubcantidadItem_count < 0 :
                        break

                    SubcantidadItem = ET.SubElement(TablaSubcantidad, 'SubcantidadItem')
                    
                    # Subcantidad
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubcantidadItem, 'Subcantidad').text = value
                    __logger.info(f'Subcantidad : {value}')
                    col_index += 1
                    
                    # CodigoSubcantidad
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubcantidadItem, 'CodigoSubcantidad').text = value
                    __logger.info(f'CodigoSubcantidad : {value}')
                    col_index += 1

                # GradosAlcohol
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'GradosAlcohol').text = value
                __logger.info(f'GradosAlcohol : {value}')
                col_index += 1

                # PrecioUnitarioReferencia
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'PrecioUnitarioReferencia').text = value
                __logger.info(f'PrecioUnitarioReferencia : {value}')
                col_index += 1

                # FechaElaboracion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'FechaElaboracion').text = value
                __logger.info(f'FechaElaboracion : {value}')
                col_index += 1

                # FechaVencimientoItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'FechaVencimientoItem').text = value
                __logger.info(f'FechaVencimientoItem : {value}')
                col_index += 1

                
                Mineria = ET.SubElement(Item, 'Mineria')
                
                # PesoNetoKilogramo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'PesoNetoKilogramo').text = value
                __logger.info(f'PesoNetoKilogramo : {value}')
                col_index += 1
                
                # PesoNetoMineria
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'PesoNetoMineria').text = value
                __logger.info(f'PesoNetoMineria : {value}')
                col_index += 1

                # TipoAfiliacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'TipoAfiliacion').text = value
                __logger.info(f'TipoAfiliacion : {value}')
                col_index += 1

                # Liquidacion
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Mineria, 'Liquidacion').text = value
                __logger.info(f'Liquidacion : {value}')
                col_index += 1

                
                # PrecioUnitarioItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 
                ET.SubElement(Item, 'PrecioUnitarioItem').text = value
                __logger.info(f'PrecioUnitarioItem : {value}')
                col_index += 1

                # DescuentoMonto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'DescuentoMonto').text = value
                __logger.info(f'DescuentoMonto : {value}')
                col_index += 1

                TablaSubDescuento = ET.SubElement(Item, 'TablaSubDescuento')

                SubDescuento_count = 5

                while True : 
                    SubDescuento_count -= 1
                    if SubDescuento_count < 0 : 
                        break

                    #SubDescuento
                    SubDescuento = ET.SubElement(TablaSubDescuento, 'SubDescuento')

                    # TipoSubDescuento
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'TipoSubDescuento').text = value
                    __logger.info(f'TipoSubDescuento : {value}')
                    col_index += 1

                    # SubDescuentoPorcentaje
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'SubDescuentoPorcentaje').text = value
                    __logger.info(f'SubDescuentoPorcentaje : {value}')
                    col_index += 1

                    # MontoSubDescuento
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubDescuento, 'MontoSubDescuento').text = value
                    __logger.info(f'MontoSubDescuento : {value}')
                    col_index += 1


                # RecargoMonto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'RecargoMonto').text = value
                __logger.info(f'RecargoMonto : {value}')
                col_index += 1

                TablaSubRecargo = ET.SubElement(Item, 'TablaSubRecargo')

                SubRecargo_count = 5

                while True :
                    SubRecargo_count -= 1
                    if SubRecargo_count < 0 :
                        break    
                    SubRecargo = ET.SubElement(TablaSubRecargo, 'SubRecargo')

                    # TipoSubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'TipoSubRecargo').text = value
                    __logger.info(f'TipoSubRecargo : {value}')
                    col_index += 1

                    # SubRecargoPorcentaje
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'SubRecargoPorcentaje').text = value
                    __logger.info(f'SubRecargoPorcentaje : {value}')
                    col_index += 1

                    
                    # MontosubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(SubRecargo, 'MontosubRecargo').text = value
                    __logger.info(f'MontosubRecargo : {value}')
                    col_index += 1

                TablaImpuestoAdicional = ET.SubElement(Item, 'TablaImpuestoAdicional')

                ImpuestoAdicional_count = 2
                while True :
                    ImpuestoAdicional_count -= 1
                    if ImpuestoAdicional_count < 0 : 
                        break
                    ImpuestoAdicional =  ET.SubElement(TablaImpuestoAdicional, 'ImpuestoAdicional')                                   
                    # MontosubRecargo
                    value = str(sheet.cell(in_row, column=col_index).value)
                    if value == "#e" :
                        value = "" 

                    ET.SubElement(ImpuestoAdicional, 'TipoImpuesto').text = value
                    __logger.info(f'TipoImpuesto : {value}')
                    col_index += 1
                
                OtraMonedaDetalle = ET.SubElement(Item, 'OtraMonedaDetalle')

                # PrecioOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'PrecioOtraMoneda').text = value
                __logger.info(f'PrecioOtraMoneda : {value}')
                col_index += 1

                # DescuentoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'DescuentoOtraMoneda').text = value
                __logger.info(f'DescuentoOtraMoneda : {value}')
                col_index += 1

                # RecargoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'RecargoOtraMoneda').text = value
                __logger.info(f'RecargoOtraMoneda : {value}')

                col_index += 1
                # MontoItemOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(OtraMonedaDetalle, 'MontoItemOtraMoneda').text = value
                __logger.info(f'MontoItemOtraMoneda : {value}')
                col_index += 1
                
                # MontoItem
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = "" 

                ET.SubElement(Item, 'MontoItem').text = value
                __logger.info(f'MontoItem : {value}')
                col_index += 1
    
    Subtotales = ET.SubElement(root, 'Subtotales')
    Subtotal = ET.SubElement(Subtotales, 'Subtotal')

    # Write NumeroSubTotal
    search_text = "NumeroSubTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'NumeroSubTotal').text = value
            __logger.info(f'Version : {cell_value}')
            break
    
    # Write DescripcionSubtotal
    search_text = "DescripcionSubtotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'DescripcionSubtotal').text = value
            __logger.info(f'DescripcionSubtotal : {cell_value}')
            break
    
    # Write Orden
    search_text = "Orden" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'Orden').text = value
            __logger.info(f'Orden : {cell_value}')
            break
    
    # Write SubTotalMontoGravadoTotal
    search_text = "SubTotalMontoGravadoTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoTotal').text = value
            __logger.info(f'SubTotalMontoGravadoTotal : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI1
    search_text = "SubTotalMontoGravadoI1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI1').text = value
            __logger.info(f'SubTotalMontoGravadoI1 : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI2
    search_text = "SubTotalMontoGravadoI2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI2').text = value
            __logger.info(f'SubTotalMontoGravadoI2 : {cell_value}')
            break
  
    # Write SubTotalMontoGravadoI3
    search_text = "SubTotalMontoGravadoI3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalMontoGravadoI3').text = value
            __logger.info(f'SubTotalMontoGravadoI3 : {cell_value}')
            break
  
    # Write SubTotaITBIS
    search_text = "SubTotaITBIS" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS').text = value
            __logger.info(f'SubTotaITBIS : {cell_value}')
            break

    # Write SubTotaITBIS1
    search_text = "SubTotaITBIS1" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS1').text = value
            __logger.info(f'SubTotaITBIS1 : {cell_value}')
            break

    # Write SubTotaITBIS2
    search_text = "SubTotaITBIS2" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS2').text = value
            __logger.info(f'SubTotaITBIS2 : {cell_value}')
            break

    # Write SubTotaITBIS3
    search_text = "SubTotaITBIS3" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotaITBIS3').text = value
            __logger.info(f'SubTotaITBIS3 : {cell_value}')
            break

    # Write SubTotalImpuestoAdicional
    search_text = "SubTotalImpuestoAdicional" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalImpuestoAdicional').text = value
            __logger.info(f'SubTotalImpuestoAdicional : {cell_value}')
            break

    # Write SubTotalExento
    search_text = "SubTotalExento" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'SubTotalExento').text = value
            __logger.info(f'SubTotalExento : {cell_value}')
            break

    # Write MontoSubTotal
    search_text = "MontoSubTotal" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'MontoSubTotal').text = value
            __logger.info(f'MontoSubTotal : {cell_value}')
            break

    # Write Lineas
    search_text = "Lineas" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:
            value = str(sheet.cell(in_row, column=col).value)
            if value == "#e" :
                value = ""             
            ET.SubElement(Subtotal, 'Lineas').text = value
            __logger.info(f'Lineas : {cell_value}')
            break

    DescuentosORecargos = ET.SubElement(root, 'DescuentosORecargos')
    DescuentoORecargo_count = 2

    search_text = "NumeroLineaDoR[1]" 
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        
        if cell_value == search_text:
            col_index = col
            while True : 
                DescuentoORecargo_count -= 1
                if DescuentoORecargo_count < 0:
                    break

                DescuentoORecargo = ET.SubElement(DescuentosORecargos, 'DescuentoORecargo')

                # NumeroLinea
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'NumeroLinea').text = value
                __logger.info(f'NumeroLinea : {cell_value}')
                col_index +=1

                # TipoAjuste
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'TipoAjuste').text = value
                __logger.info(f'TipoAjuste : {cell_value}')
                col_index +=1
                
                # IndicadorNorma1007
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'IndicadorNorma1007').text = value
                __logger.info(f'IndicadorNorma1007 : {cell_value}')
                col_index +=1
                         
                # DescripcionDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'DescripcionDescuentooRecargo').text = value
                __logger.info(f'DescripcionDescuentooRecargo : {cell_value}')
                col_index +=1
                                         
                # TipoValor
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'TipoValor').text = value
                __logger.info(f'TipoValor : {cell_value}')
                col_index +=1
                                                     
                # ValorDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'ValorDescuentooRecargo').text = value
                __logger.info(f'ValorDescuentooRecargo : {cell_value}')
                col_index +=1
                                                               
                # MontoDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'MontoDescuentooRecargo').text = value
                __logger.info(f'MontoDescuentooRecargo : {cell_value}')
                col_index +=1
                                                                        
                # MontoDescuentooRecargoOtraMoneda
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'MontoDescuentooRecargoOtraMoneda').text = value
                __logger.info(f'MontoDescuentooRecargoOtraMoneda : {cell_value}')
                col_index +=1        

                # IndicadorFacturacionDescuentooRecargo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(DescuentoORecargo, 'IndicadorFacturacionDescuentooRecargo').text = value
                __logger.info(f'IndicadorFacturacionDescuentooRecargo : {cell_value}')         
     

    Paginacion = ET.SubElement(root, 'Paginacion')
    Pagina_count = 2

    search_text = "PaginaNo[1]" 
    col_index : int
    for col in range(1, sheet.max_column + 1):
        cell_value = sheet.cell(1, column=col).value
        if cell_value == search_text:

            col_index = col
                
            while True:
                Pagina_count -= 1
                if Pagina_count < 0:
                    break

                Pagina = ET.SubElement(Paginacion, 'Pagina')

                # PaginNo
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'PaginaNo').text = value
                __logger.info(f'PaginaNo : {cell_value}')
                __logger.info(f'PaginaNo col_index : {col_index}')
                col_index += 1

                # NoLineaDesde
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'NoLineaDesde').text = value
                __logger.info(f'NoLineaDesde : {cell_value}')
                col_index += 1

                # NoLineaHasta
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'NoLineaHasta').text = value
                __logger.info(f'NoLineaHasta : {value}')
                col_index += 1
                
                # SubtotalMontoGravadoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravadoPagina').text = value
                __logger.info(f'SubtotalMontoGravadoPagina : {value}')
                col_index += 1
                                
                # SubtotalMontoGravado1Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado1Pagina').text = value
                __logger.info(f'SubtotalMontoGravado1Pagina : {value}')
                col_index += 1
                                                
                # SubtotalMontoGravado2Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado2Pagina').text = value
                __logger.info(f'SubtotalMontoGravado2Pagina : {value}')
                col_index += 1
                                                   
                # SubtotalMontoGravado3Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalMontoGravado3Pagina').text = value
                __logger.info(f'SubtotalMontoGravado3Pagina : {value}')
                col_index += 1
                                                                   
                # SubtotalExentoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalExentoPagina').text = value
                __logger.info(f'SubtotalExentoPagina : {value}')
                col_index += 1
                                                               
                # SubtotalItbisPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbisPagina').text = value
                __logger.info(f'SubtotalItbisPagina : {value}')
                col_index += 1
                                                                      
                # SubtotalItbis1Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis1Pagina').text = value
                __logger.info(f'SubtotalItbis1Pagina : {value}')
                col_index += 1
                                                                                      
                # SubtotalItbis2Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis2Pagina').text = value
                __logger.info(f'SubtotalItbis2Pagina : {value}')
                col_index += 1
                                                                                         
                # SubtotalItbis3Pagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalItbis3Pagina').text = value
                __logger.info(f'SubtotalItbis3Pagina : {value}')
                col_index += 1
                                                                                                         
                # SubtotalImpuestoAdicionalPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""             
                ET.SubElement(Pagina, 'SubtotalImpuestoAdicionalPagina').text = value
                __logger.info(f'SubtotalImpuestoAdicionalPagina : {value}')
                col_index += 1

                SubtotalImpuestoAdicional = ET.SubElement(Pagina, 'SubtotalImpuestoAdicional')
                
                # SubtotalImpuestoSelectivoConsumoEspecificoPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(SubtotalImpuestoAdicional, 'SubtotalImpuestoSelectivoConsumoEspecificoPagina').text =value
                __logger.info(f'SubtotalImpuestoSelectivoConsumoEspecificoPagina : {value}')
                col_index += 1

                # SubtotalOtrosImpuesto
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(SubtotalImpuestoAdicional, 'SubtotalOtrosImpuesto').text =value
                __logger.info(f'SubtotalOtrosImpuesto : {value}')
                col_index += 1

                # MontoSubtotalPagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(Pagina, 'MontoSubtotalPagina').text =value
                __logger.info(f'MontoSubtotalPagina : {value}')
                col_index += 1

                # SubtotalMontoNoFacturablePagina
                value = str(sheet.cell(in_row, column=col_index).value)
                if value == "#e" :
                    value = ""     
                ET.SubElement(Pagina, 'SubtotalMontoNoFacturablePagina').text =value
                __logger.info(f'SubtotalMontoNoFacturablePagina : {value}')
                __logger.info(f'SubtotalMontoNoFacturablePagina _ index : {col_index}')
                # col_index += 1

    InformacionReferencia = ET.SubElement(root, 'InformacionReferencia')
    
    # NCFModificado
    col_index += 1
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'NCFModificado').text =value
    __logger.info(f'NCFModificado : {value}')
    col_index += 1
        
    # RNCOtroContribuyente
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'RNCOtroContribuyente').text =value
    __logger.info(f'RNCOtroContribuyente : {value}')
    col_index += 1
        
    # FechaNCFModificado
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'FechaNCFModificado').text =value
    __logger.info(f'FechaNCFModificado : {value}')
    col_index += 1
        
    # CodigoModificacion
    value = str(sheet.cell(in_row, column=col_index).value)
    if value == "#e" :
        value = ""     
    ET.SubElement(InformacionReferencia, 'CodigoModificacion').text =value
    __logger.info(f'CodigoModificacion : {value}')
    col_index += 1

    rough_string = ET.tostring(root, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    pretty_xml_as_str = reparsed.toprettyxml(indent="  ", encoding='utf-8')

    path = os.path.join(os.path.dirname(__file__), 'data/row_invoice.xml')
    
    with open(path, 'wb') as f:
        f.write(pretty_xml_as_str)
    
    xml_str = ET.tostring(root, encoding='utf-8').decode('utf-8')
    return xml_str

def main():
    try:
        dgii_service = DGIICFService(dgii_env='test', company_id=1)

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
        # xml_str = read_excel_create_rfce_xml(4)
        # xml_str = create_e_cf_31(9) # 7, 8, 9, 10
        xml_str = create_e_cf_32(14) # 11, 12, 13, 14
        # __logger.info(f"invoice xml: {xml_str}")

        signed_xml = dgii_service.sign_xml(xml_str)
        # __logger.info(f"signed invoice xml: {signed_xml}")

        # Step 5: Submit the signed e-CF XML to DGII using the token
        # response = dgii_service.submit_rcef(token)
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
        __logger.info(f"track result: {response.content}")

        response_data = json.loads(response.content.decode('utf-8'))
        __logger.info(f"dgii_track_id {trackId}, dgii_token : {token}")

    except Exception as e:
        __logger.error(f"Error processing invoice: {str(e)}")

if __name__ == "__main__":
    main()