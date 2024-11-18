from flask import Flask, request, send_file, render_template, redirect, url_for, session
import logging
import os
from converters import excel_to_xml, xml_to_excel_csv, excel_to_csv  # Add this import

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
        file = request.files.get('file')

        if not file:
            logging.error("No file uploaded")
            return "No file uploaded", 400

        try:
            if converter_type == 'excel_to_xml':
                return excel_to_xml.handle_conversion(file)
            elif converter_type == 'xml_to_excel_csv':
                return xml_to_excel_csv.handle_conversion(file)
            elif converter_type == 'excel_to_csv':  # Add this condition
                return excel_to_csv.handle_conversion(file)
            else:
                logging.error(f"Unknown converter type: {converter_type}")
                return "Unknown converter type", 400

        except Exception as e:
            logging.error(f"Error during file processing: {str(e)}")
            return str(e), 500

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)