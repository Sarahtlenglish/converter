from flask import Flask, request, send_file, render_template, redirect, url_for, session
import logging
import os
from converters import excel_to_xml, xml_to_excel_csv, excel_to_csv, csv_image_downloader

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Configure logging
logging.basicConfig(level=logging.DEBUG)

PASSWORD = "nmickode"

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form.get('password') == PASSWORD:
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            return render_template('login.html', error="Incorrect password. Please try again.")

    return render_template('login.html')

@app.route('/', methods=['GET', 'POST'])
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if request.method == 'POST':
        converter_type = request.form.get('converter_type')
        
        try:
            # XML to Excel/CSV - Accepts multiple files
            if converter_type == 'xml_to_excel_csv':
                files = request.files.getlist('file')
                if not files or not files[0].filename:
                    logging.error("No files uploaded for XML conversion")
                    return "No files uploaded", 400
                return xml_to_excel_csv.handle_conversion(files)
            
            # Single file converters
            file = request.files.get('file')  # Ændret fra getlist til get for enkelte filer
            if not file or not file.filename:
                logging.error("No file uploaded")
                return "No file uploaded", 400

            # Route to appropriate converter
            if converter_type == 'excel_to_xml':
                logging.debug(f"Processing single file with excel_to_xml converter")
                return excel_to_xml.handle_conversion(file)
            elif converter_type == 'excel_to_csv':
                logging.debug(f"Processing single file with excel_to_csv converter")
                return excel_to_csv.handle_conversion(file)
            elif converter_type == 'csv_image_downloader':
                logging.debug(f"Processing single file with csv_image_downloader converter")
                try:
                    return csv_image_downloader.handle_conversion(file)
                except Exception as e:
                    logging.error(f"Error in csv_image_downloader: {str(e)}")
                    return str(e), 500
            else:
                logging.error(f"Unknown converter type: {converter_type}")
                return "Unknown converter type", 400

        except Exception as e:
            logging.error(f"Error during file processing: {str(e)}")
            return str(e), 500

    return render_template('index.html')

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_server_error(e):
    return render_template('500.html'), 500

if __name__ == '__main__':
    app.run(debug=True)
    app = app