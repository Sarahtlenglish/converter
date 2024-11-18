import pandas as pd
from flask import send_file
import os
import logging

def handle_conversion(file):
    try:
        logging.debug("Handling Excel to CSV conversion")
        
        # Determine the file extension
        file_extension = os.path.splitext(file.filename)[1].lower()
        
        # Choose the appropriate engine based on the file extension
        if file_extension == '.xlsx':
            engine = 'openpyxl'
        elif file_extension == '.xls':
            engine = 'xlrd'
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")
        
        # Read the Excel file with the specified engine
        df = pd.read_excel(file, engine=engine, dtype=str)
        
        # Generate a filename for the CSV
        filename = "converted_file.csv"
        temp_file_path = f"/tmp/{filename}"
        
        # Save as CSV
        df.to_csv(temp_file_path, index=False, encoding='utf-8')
        
        # Send the file
        response = send_file(temp_file_path, mimetype='text/csv', as_attachment=True, download_name=filename)
        
        # Clean up the temporary file
        os.remove(temp_file_path)
        
        return response
    except Exception as e:
        logging.error(f"Error during Excel to CSV conversion: {str(e)}")
        raise e