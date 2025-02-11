from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'txt_file' not in request.files or 'image_file' not in request.files:
        return "No file part", 400

    txt_file = request.files['txt_file']
    image_file = request.files['image_file']

    if txt_file.filename == '' or image_file.filename == '':
        return "No selected file", 400

    txt_filename = secure_filename(txt_file.filename)
    image_filename = secure_filename(image_file.filename)

    txt_file.save(os.path.join('/path/to/save', txt_filename))
    image_file.save(os.path.join('/path/to/save', image_filename))

    # Call the function to create the presentation
    criar_apresentacao(txt_filename, image_filename)

    return send_file('/path/to/save/generated_presentation.pptx', as_attachment=True)

@app.route('/criar_apresentacao', methods=['POST'])
def criar_apresentacao():
    data = request.get_json()
    # Process the data and create a presentation
    # For example, you can save the data to a file
    filename = secure_filename(data['filename'])
    with open(filename, 'w') as f:
        f.write(data['content'])
    return jsonify({"message": "Presentation created successfully"}), 201

if __name__ == '__main__':
    app.run(debug=True)