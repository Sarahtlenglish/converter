# converters/xml_to_excel_csv.py

import xml.etree.ElementTree as ET
from openpyxl import Workbook
import csv
import logging
import zipfile
import re
from flask import send_file
import io
import os

def handle_conversion(files):
    """Handle one or multiple XML files to Excel/CSV conversion and zip them."""
    try:
        logging.debug("Handling XML to Excel/CSV conversions")
        
        # Create a zip buffer in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Ensure files is a list
            if not isinstance(files, list):
                files = [files]
            
            # Keep track of used filenames to avoid duplicates
            used_filenames = set()
            
            for xml_file in files:
                try:
                    logging.debug(f"Processing file: {xml_file.filename}")
                    xml_file.seek(0)
                    
                    # Parse XML and get PageName for filename
                    tree = ET.parse(xml_file)
                    root = tree.getroot()
                    namespace = {'ns': 'http://www.filemaker.com/fmpdsoresult'}
                    
                    # Get PageName or use original filename
                    page_name_element = root.find('.//ns:PageName', namespace)
                    if page_name_element is not None and page_name_element.text:
                        base_filename = page_name_element.text.strip()
                    else:
                        base_filename = os.path.splitext(xml_file.filename)[0]
                    
                    # Make filename unique if necessary
                    original_filename = base_filename
                    counter = 1
                    while base_filename in used_filenames:
                        base_filename = f"{original_filename} ({counter})"
                        counter += 1
                    used_filenames.add(base_filename)
                    
                    logging.debug(f"Using filename: {base_filename}")
                    
                    # Reset file pointer and convert data
                    xml_file.seek(0)
                    data = parse_xml_to_data(xml_file)
                    
                    # Create buffers for this file
                    excel_buffer = io.BytesIO()
                    csv_buffer = io.StringIO(newline='')
                    
                    # Save to Excel and CSV
                    save_to_excel(data, excel_buffer)
                    save_to_csv(data, zipf, base_filename)
                    
                    # Add to zip with unique filenames
                    excel_buffer.seek(0)
                    zipf.writestr(f'{base_filename}.xlsx', excel_buffer.getvalue())
                    
                except Exception as e:
                    logging.error(f"Error processing file {xml_file.filename}: {str(e)}")
                    continue

        # Prepare the zip file for download
        zip_buffer.seek(0)
        if zip_buffer.getvalue():  # Check if zip file has content
            return send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name='avis-data.zip'
            )
        else:
            raise ValueError("No files were successfully processed")
            
    except Exception as e:
        logging.error(f"Error during XML to Excel/CSV conversion: {str(e)}")
        raise e

def convert_br(text):
    """Convert ~br~ to a space followed by a bullet point."""
    return re.sub(r'~br~', ' •', text) if text else ""

def parse_xml_to_data(xml_file):
    """Parse the XML file and extract the relevant data into a list of dictionaries."""
    tree = ET.parse(xml_file)
    root = tree.getroot()

    namespace = {'ns': 'http://www.filemaker.com/fmpdsoresult'}

    data = []

    for row in root.findall('ns:ROW', namespace):
        varenr = convert_br(row.find('ns:Varenr', namespace).text if row.find('ns:Varenr', namespace) is not None else "")
        overskrift = convert_br(row.find('ns:Overskrift', namespace).text if row.find('ns:Overskrift', namespace) is not None else "")
        broedtekst = convert_br(row.find('ns:Broedtekst', namespace).text if row.find('ns:Broedtekst', namespace) is not None else "")
        broedtekstkort = convert_br(row.find('ns:Broedtekstkort', namespace).text if row.find('ns:Broedtekstkort', namespace) is not None else "")
        model = convert_br(row.find('ns:Model', namespace).text if row.find('ns:Model', namespace) is not None else "")
        undervaretekst1 = convert_br(row.find('ns:Undervaretekst1', namespace).text if row.find('ns:Undervaretekst1', namespace) is not None else "")
        kampprisinklmomsberegnet = convert_br(row.find('ns:Kampprisinklmomsberegnet', namespace).text if row.find('ns:Kampprisinklmomsberegnet', namespace) is not None else "")
        kampprisexmomsberegnet = convert_br(row.find('ns:Kampprisexmomsberegnet', namespace).text if row.find('ns:Kampprisexmomsberegnet', namespace) is not None else "")
        prioritet = convert_br(row.find('ns:Prioritet', namespace).text if row.find('ns:Prioritet', namespace) is not None else "")
        tilbudstype = convert_br(row.find('ns:Tilbudstype', namespace).text if row.find('ns:Tilbudstype', namespace) is not None else "")
        bureaukommentar = convert_br(row.find('ns:Bureaukommentar', namespace).text if row.find('ns:Bureaukommentar', namespace) is not None else "")
        urltilprimaerebillede = convert_br(row.find('ns:URLtilprimaerebillede', namespace).text if row.find('ns:URLtilprimaerebillede', namespace) is not None else "")
        picture1 = convert_br(row.find('ns:Picture1', namespace).text if row.find('ns:Picture1', namespace) is not None else "")
        logo = convert_br(row.find('ns:Logo', namespace).text if row.find('ns:Logo', namespace) is not None else "")
        urltillogo = convert_br(row.find('ns:URLtillogo', namespace).text if row.find('ns:URLtillogo', namespace) is not None else "")
        tekst_over_pris = convert_br(row.find('ns:incitoPriceFooter', namespace).text if row.find('ns:incitoPriceFooter', namespace) is not None else "")

        # Calculate "Pris" field
        pris = kampprisexmomsberegnet

        # Calculate "Broed" field - prioritize broedtekstkort
        broed = broedtekstkort if broedtekstkort else broedtekst

        # Calculate the @image field with the new path format
        image_path = f"/Volumes/work_creative/Bygma/BILLEDER/_BYGMA_PRODUKTBILLEDER/{varenr}.psd" if varenr else ""

        # Handle multiple logos
        logos = logo.split(',')
        logo_fields = {}
        for i, logo_path in enumerate(logos, start=1):
            logo_fields[f"logo{i}"] = logo_path.strip()

        # Append the row data to the list
        row_data = {
            "Varenr": varenr,
            "Overskrift": overskrift,
            "Broedtekst": broedtekst,
            "Broedtekstkort": broedtekstkort,
            "Model": model,
            "Undervaretekst1": undervaretekst1,
            "Kampprisinklmomsberegnet": kampprisinklmomsberegnet,
            "Kampprisexmomsberegnet": kampprisexmomsberegnet,
            "Prioritet": prioritet,
            "Tilbudstype": tilbudstype,
            "Bureaukommentar": bureaukommentar,
            "URLtilprimaerebillede": urltilprimaerebillede,
            "Picture1": picture1,
            "URLtillogo": urltillogo,
            "Pris": pris,
            "@image": image_path,
            "Broed": broed,
            "tekst over pris": tekst_over_pris
        }
        row_data.update(logo_fields)
        data.append(row_data)

    return data

def save_to_excel(data, excel_buffer):
    """Save the data to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    # Add headers
    headers = set()
    for row in data:
        headers.update(row.keys())
    headers = sorted(list(headers))
    ws.append(headers)

    # Add data rows
    for item in data:
        row = [item.get(header, "") for header in headers]
        ws.append(row)

    wb.save(excel_buffer)
    logging.debug("Excel data saved to buffer")

def clean_data(value):
    """Rens data for specielle tegn."""
    if value:
        return value.replace('\n', ' ').replace('\r', '').replace('"', '""')  # Erstat linjeskift og escape citationstegn
    return ""

def save_to_csv(data, zipf, base_filename):
    """Save the data to a CSV file."""
    headers = set()
    for row in data:
        headers.update(row.keys())
    headers = sorted(list(headers))

    csv_buffer = io.StringIO(newline='')  # Opret en StringIO buffer til CSV
    writer = csv.DictWriter(csv_buffer, fieldnames=headers, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
    writer.writeheader()
    for row in data:
        cleaned_row = {key: clean_data(value) for key, value in row.items()}  # Rens data
        writer.writerow(cleaned_row)

    # Gem CSV indholdet i zip-filen med 'utf-8-sig' encoding
    zipf.writestr(f'{base_filename}.csv', csv_buffer.getvalue().encode('utf-8-sig'))  # Brug 'utf-8-sig' encoding
    logging.debug("CSV data saved to zip")
