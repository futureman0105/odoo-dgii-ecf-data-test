import requests
import logging
import os
import json
from lxml import etree

class DGIICFService:
    def __init__(self, dgii_env='test', company_id=None, env=None):
        """
        Initialize the DGII service with environment, company, and config data.
        """
        self._logger =  logging.getLogger(__name__)

        if not company_id:
            raise ValueError("Company ID is required")
             
        self.base_url = {
            'test': 'https://ecf.dgii.gov.do/testeCF',
            'prod': 'https://ecf.dgii.gov.do/CerteCF'
        }[dgii_env]

        
        self.fc_base_url = {
            'test': 'https://fc.dgii.gov.do/testeCF',
            'prod': 'https://fc.dgii.gov.do/CerteCF'
        }[dgii_env]
        

        self.env = env

        # Fetch DGII credentials from company settings
        # company = self.env['res.company'].browse(company_id)
        self.dgii_username = "132641566"
        self.dgii_password = "RICARDO123"
        self.invoice_name = ""
        # self.cert_password = company.dgii_cert_password

        # Password for private key if necessary
        # self.password = company.dgii_cert_password

    def get_semilla(self):
        """
        Fetch the seed (semilla) from DGII for authentication.
        """
        url = f"{self.base_url}/Autenticacion/api/Autenticacion/Semilla"
        res = requests.get(url, auth=(self.dgii_username, self.dgii_password))

        self._logger.info(f"Semilla response from DGII: {res.content}")

        if res.status_code == 200:
            tree = etree.fromstring(res.content)
            return res.content
        else:
            raise Exception("Error fetching semilla from DGII: " + res.text)

    def sign_semilla(self, semilla):
        """
        Sign the 'semilla' using the provided certificate and private key.
        """
        signed_xml = self.custom_sign_xml(semilla)

        return signed_xml

    # return Token
    def validate_semilla(self):
        """
        Validate the signed 'semilla' with DGII using their validation endpoint.
        """

        xml_file_path = os.path.join(os.path.dirname(__file__), 'data/signed.xml')

        with open(xml_file_path, 'rb') as f:
            files = {'xml': ('signed.xml', f, 'text/xml')}
            headers = {'accept': 'application/json'}
            
            res = requests.post(
                'https://ecf.dgii.gov.do/testecf/Autenticacion/api/Autenticacion/ValidarSemilla',
                files=files,
                headers=headers
            )

            self._logger.info(f"Validar Semilla response from DGII: {res.content}")

            if res.status_code == 200:
                response_data = json.loads(res.content.decode('utf-8'))
                self._logger.info(f"Token Data: {res.content}")
                token = response_data['token']
                return token
            else:
                raise Exception("Error validating certificate with DGII: " + res.text)
        
    def custom_sign_xml(self, semilla, type='semilla'):
        """
        Sign the XML for 'semilla' using the private key and certificate.
        """
        
        url = "http://localhost:8080/api/sign"
        headers = {
            "Content-Type": "application/xml",
        }

        if type == 'semilla':
            xml_str = semilla.decode('utf-8').strip()
        else:
            xml_str = semilla
            
        res = requests.post(url, data=xml_str, headers=headers)

        if type == 'semilla':
            path = os.path.join(os.path.dirname(__file__), 'data/signed.xml')
        else:
            path = os.path.join(os.path.dirname(__file__), f'data/{self.invoice_name}_signed.xml')
        
        with open(path, 'wb') as f:
                f.write(res.content)

        return res.content
    
    def sign_xml(self, xml_data, invoice_name):
        self.invoice_name = invoice_name
        signed_xml = self.custom_sign_xml(xml_data, type='invoice')
        return signed_xml
    
    def submit_rfce(self, token):
        """
        Submit the signed e-CF XML to DGII.
        """
        xml_file_path = os.path.join(os.path.dirname(__file__), f'data/{self.invoice_name}_signed.xml')

        with open(xml_file_path, 'rb') as f:
            # files = {'xml': ('signed.xml', f, 'text/xml')}
            files = {'xml': (f'{self.invoice_name}.xml', f, 'text/xml')}
            headers = {
                'accept': 'application/json',
                "Authorization": f"Bearer {token}"
            }

            try : 
                res = requests.post(
                    f"{self.fc_base_url}/RecepcionFC/api/recepcion/ecf",
                    files=files,
                    headers=headers
                )

                self._logger.info(f"RFCE response from DGII: {res.content}")
            except Exception as e:
                self._logger.error(f"Failed Validar Semilla response from DGII: {e}")
                return

            return res
      
    def submit_ecf(self, token):
        """
        Submit the signed e-CF XML to DGII.
        """

        self._logger.info(f"Submitting the signed e-CF XML to DGII.")

        xml_file_path = os.path.join(os.path.dirname(__file__), 'data/signed_invoice.xml')

        with open(xml_file_path, 'rb') as f:
            files = {'xml': ('signed.xml', f, 'text/xml')}
            headers = {
                'accept': 'application/json',
                "Authorization": f"Bearer {token}"
            }

            try : 
                res = requests.post(
                    f"{self.base_url}/Recepcion/api/FacturasElectronicas",
                    files=files,
                    headers=headers
                )

                self._logger.info(f"Validar Semilla response from DGII: {res.content}")
            except Exception as e:
                self._logger.error(f"Failed Validar Semilla response from DGII: {e}")
                return

            return res
        
    def track_ecf(self, trackid, token):
        """
        Track e-CF status from DGII.
        """

        data = {'trackid': trackid}
        headers = {
            'accept': 'application/json',
            "Authorization": f"Bearer {token}"
        }
        
        res = requests.get(
            f"{self.base_url}/consultaresultado/api/consultas/estado?trackid={trackid}",
            headers=headers
        )

        self._logger.info(f"Track response from DGII: {res.content}")

        return res
      
    def track_rfce(self, token, param):
        """
        Track e-CF status from DGII.
        """
        url = "https://fc.dgii.gov.do/CerteCF/consultarfce/api/Consultas/Consulta"

        headers = {
            'accept': 'application/json',
            "Authorization": f"Bearer {token}"
        }
        
        params = {
            "RNC_Emisor": param.RNCEmisor,
            "ENCF": param.ENCF,
            "Cod_Seguridad_eCF": param.CodigoSeguridadeCF
        }

        print(params, token)
       
        response = requests.get(f"{url}?RNC_Emisor={param.RNCEmisor}&ENCF={param.ENCF}&Cod_Seguridad_eCF={param.CodigoSeguridadeCF}", 
                                headers=headers)

        # Print response status and JSON data


        return response

