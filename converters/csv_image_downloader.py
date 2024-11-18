import csv
import requests
import zipfile
import io
import os
from flask import send_file
import logging

def handle_conversion(file):
    try:
        logging.debug("Handling CSV to Image ZIP conversion")
        
        # Read the CSV file
        csv_content = file.read().decode('utf-8')
        csv_reader = csv.DictReader(io.StringIO(csv_content))
        
        # Create a zip file in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
            for row in csv_reader:
                image_url = row.get('URLtilprimaerebillede')
                if image_url:
                    try:
                        # Download the image
                        response = requests.get(image_url, timeout=10)
                        if response.status_code == 200:
                            # Get the filename from the URL
                            filename = os.path.basename(image_url)
                            # Add the image to the zip file
                            zip_file.writestr(filename, response.content)
                    except requests.RequestException as e:
                        logging.error(f"Error downloading image {image_url}: {str(e)}")

        # Prepare the zip file for download
        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='product_images.zip'
        )
    except Exception as e:
        logging.error(f"Error during CSV to Image ZIP conversion: {str(e)}")
        raise e