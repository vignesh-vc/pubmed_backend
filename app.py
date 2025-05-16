from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile
from datetime import datetime
from pubmed_pdf_downloader import process_excel
import logging

app = Flask(__name__)
CORS(app)

# Setup logging
logging.basicConfig(level=logging.DEBUG)

# Create a permanent folder for processed files
PROCESSED_FOLDER = os.path.join(os.getcwd(), "processed_files")
os.makedirs(PROCESSED_FOLDER, exist_ok=True)  # Ensure folder exists

@app.route("/upload", methods=["POST"])
def upload_and_process():
    if 'file' not in request.files:
        logging.error("No file uploaded")
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']

    if file.filename == '':
        logging.error("Empty filename")
        return jsonify({'error': 'Empty filename'}), 400

    try:
        # Save uploaded file to a temp path
        temp_input_path = os.path.join(tempfile.gettempdir(), 'Pubs.xlsx')
        file.save(temp_input_path)
        logging.info(f"File saved to temporary path: {temp_input_path}")

        # Generate unique output filename using timestamp
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        output_filename = f'Pubs_Updated_{timestamp}.xlsx'
        output_path = os.path.join(PROCESSED_FOLDER, output_filename)

        # Process the file
        logging.info(f"Processing file: {temp_input_path}")
        process_excel(temp_input_path, output_path)

        # Confirm file exists
        if os.path.exists(output_path):
            logging.info(f"Processed file saved to: {output_path}")
            return send_file(
                output_path,
                as_attachment=True,
                download_name='Pubs_Updated.xlsx',  # Static download name
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            logging.error("Processed file not found")
            return jsonify({'error': 'Processed file not found'}), 500

    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        return jsonify({'error': str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
