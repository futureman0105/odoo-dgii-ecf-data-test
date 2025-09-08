import xml.etree.ElementTree as ET
import logging

def clean_xml_safe(xml_input):
    """
    Safe version that handles various input types and errors
    """
    try:
        # Handle different input types
        if isinstance(xml_input, bytes):
            xml_str = xml_input.decode('utf-8')
        elif isinstance(xml_input, str):
            xml_str = xml_input
        else:
            raise ValueError("Input must be string or bytes")
        
        # Parse XML
        root = ET.fromstring(xml_str)
        
        # Iterative approach to find and remove empty elements
        changed = True
        while changed:
            changed = False
            # Find all empty elements
            empty_elements = []
            for element in root.iter():
                if (len(element) == 0 and 
                    (element.text is None or element.text.strip() == '')):
                    empty_elements.append(element)
            
            # Remove empty elements by searching for their parent
            for empty_element in empty_elements:
                for parent in root.iter():
                    if empty_element in list(parent):
                        parent.remove(empty_element)
                        changed = True
                        break
                                                     
        return ET.tostring(root, encoding='utf-8').decode('utf-8')
        
    except Exception as e:
        logging.error(f"Error in clean_xml_safe: {e}")
        return xml_input  # Return original input on error