from flask import Flask, request, send_file, render_template
import pandas as pd
import os
import uuid
import json

app = Flask(__name__)

# Define upload and converted folders
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

def flatten_json(y, sep='_'):
    """
    Recursively flattens a nested JSON object, including arrays.
    """
    out = {}

    def flatten(x, name=''):
        if isinstance(x, dict):
            for key in x:
                flatten(x[key], f"{name}{key}{sep}")
        elif isinstance(x, list):
            for i, item in enumerate(x):
                flatten(item, f"{name}{i}{sep}")
        else:
            out[name[:-1]] = x

    flatten(y)
    return out

def process_nested_json(data):
    """
    Processes nested JSON objects and arrays into a flat structure.
    """
    if isinstance(data, list):
        # If JSON is a list of dictionaries
        flattened_data = [flatten_json(item) for item in data]
    elif isinstance(data, dict):
        # If JSON is a single dictionary
        flattened_data = [flatten_json(data)]
    else:
        raise ValueError("Unsupported JSON structure")
    return flattened_data

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if a file is uploaded
        if 'json_file' not in request.files or request.files['json_file'].filename == '':
            return "No file uploaded", 400

        f = request.files['json_file']
        file_id = uuid.uuid4().hex
        json_path = os.path.join(UPLOAD_FOLDER, f"{file_id}.json")
        excel_path = os.path.join(CONVERTED_FOLDER, f"{file_id}.xlsx")
        f.save(json_path)

        try:
            # Read JSON and convert to Excel
            with open(json_path, 'r') as file:
                data = json.load(file)

            # Flatten the JSON data
            flattened_data = process_nested_json(data)

            # Convert to DataFrame
            df = pd.DataFrame(flattened_data)

            # Save to Excel
            df.to_excel(excel_path, index=False)
        except Exception as e:
            return f"Error processing file: {e}", 500

        return send_file(excel_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)