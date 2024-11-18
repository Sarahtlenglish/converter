import pandas as pd
from flask import send_file
import os
import logging
import re


def handle_conversion(file):
    try:
        logging.debug("Handling Excel to XML conversion")
        excel_data = pd.read_excel(file, dtype=str, header=None)
        xml_content = generate_xml(excel_data)

        kampagnenavn = excel_data.iat[1, 2]  # Campaign Name
        filename = f"{kampagnenavn.strip().replace(' ', '_')}.xml"

        temp_file_path = f"/tmp/{filename}"
        with open(temp_file_path, 'w', encoding='utf-8') as temp_file:
            temp_file.write(xml_content)

        response = send_file(temp_file_path, mimetype='application/xml', as_attachment=True, download_name=filename)
        os.remove(temp_file_path)

        return response
    except Exception as e:
        logging.error(f"Error during Excel to XML conversion: {str(e)}")
        raise e

def generate_xml(excel_data):
    kampagnenavn = excel_data.iat[1, 2]  # Campaign Name
    start_dato = pd.to_datetime(excel_data.iat[1, 0]).strftime('%d.%m.%Y')
    slut_dato = pd.to_datetime(excel_data.iat[1, 1]).strftime('%d.%m.%Y')
    moms_text = str(excel_data.iat[3, 2]).strip().lower()

    if re.search(r'inkl\.?\s*moms', moms_text):
        chain = "Byggecenter"
        price_tag = "KampagneprisInklMoms"
    elif re.search(r'eksl\.?\s*moms', moms_text):
        chain = "Proffcenter"
        price_tag = "KampagneprisExMoms"
    else:
        chain = "Proffcenter"
        price_tag = "KampagneprisExMoms"

    campaign_data = excel_data.iloc[6:, :]

    xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n'
    xml_str += '<Kampagne>\n'
    xml_str += f'    <OpgaveNavn><![CDATA[{kampagnenavn}]]></OpgaveNavn>\n'
    xml_str += '    <CampaignType><![CDATA[LOKALANNONCER]]></CampaignType>\n'
    xml_str += f'    <Chain><![CDATA[{chain}]]></Chain>\n'
    xml_str += f'    <CampaignStart><![CDATA[{start_dato}]]></CampaignStart>\n'
    xml_str += f'    <CampaignEnd><![CDATA[{slut_dato}]]></CampaignEnd>\n'

    previous_sidenr = None
    for index, row in campaign_data.iterrows():
        sidenr = row[0]
        m3_nr = row[1].zfill(6) if pd.notna(row[1]) else "000000"
        pris = f"{float(row[4].replace(',', '.')):.2f}".replace('.', ',') if pd.notna(row[4]) else "0,00"
        kommentarer = row[7] if pd.notna(row[7]) else ""

        if previous_sidenr is None or previous_sidenr != sidenr:
            if previous_sidenr is not None:
                xml_str += '        </Overskrift>\n'
                xml_str += '    </Side>\n'
            xml_str += '    <Side>\n'
            xml_str += f'        <Sidenavn><![CDATA[Side {sidenr}]]></Sidenavn>\n'
            xml_str += f'        <Sortering><![CDATA[{sidenr}]]></Sortering>\n'
            xml_str += '        <Overskrift>\n'
            xml_str += '            <OverskriftNavn><![CDATA[Overskrift]]></OverskriftNavn>\n'

        xml_str += '            <Vare>\n'
        xml_str += f'                <Varenummer><![CDATA[{m3_nr}]]></Varenummer>\n'

        if kommentarer:
            xml_str += '                <MoenstretVare>\n'
            xml_str += f'                    <Vare><![CDATA[{m3_nr}]]></Vare>\n'
            xml_str += '                    <Overskrift><![CDATA[Overskrift]]></Overskrift>\n'
            xml_str += f'                    <BureauKommentar><![CDATA[{kommentarer}]]></BureauKommentar>\n'
            xml_str += '                </MoenstretVare>\n'
        else:
            xml_str += '                <Overskrift><![CDATA[Overskrift]]></Overskrift>\n'

        xml_str += '                <Priser>\n'
        xml_str += f'                    <{price_tag}><![CDATA[{pris}]]></{price_tag}>\n'
        xml_str += '                </Priser>\n'
        xml_str += '            </Vare>\n'

        previous_sidenr = sidenr

    if previous_sidenr is not None:
        xml_str += '        </Overskrift>\n'
        xml_str += '    </Side>\n'

    xml_str += '</Kampagne>'
    return xml_str
