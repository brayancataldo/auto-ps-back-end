from flask import Flask, jsonify, send_from_directory, request
import os
from flask_cors import CORS
from photoshopy import Photoshopy
import pythoncom
from operator import itemgetter

# initialize our Flask application
app= Flask(__name__)
CORS(app)

DIRETORIO = "D:\\brayanHD\\wordspacePython\\auto-ps\\src\\export"
DIRETORIO_UPLOAD = "D:\\brayanHD\\wordspacePython\\auto-ps\\src\\import"

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    return response
    

@app.route("/files", methods=["GET"])
def generate():
    files = []
    
    try:
        for filename in os.listdir(DIRETORIO):
            file_path = os.path.join(DIRETORIO, filename)

            if os.path.isfile(file_path):
                files.append({
                    "nome": filename,
                    "data_criacao": os.path.getctime(file_path)
                })
    except OSError as e:
        return jsonify({"error": str(e)})
    
    # Sort the list of files by the ctime value
    sorted_files = sorted(files, key=itemgetter("data_criacao"))
    
    # Return only the sorted file names
    sorted_files_names = [a["nome"] for a in sorted_files]
    return jsonify(sorted_files_names)


@app.route("/files/<name>", methods=["GET"])
def get_file(name):
    return send_from_directory(DIRETORIO, name, as_attachment=True)

@app.route("/generate", methods=["POST"])
def generate_art():
    pythoncom.CoInitialize()
    data = request.get_json()
    art = Photoshopy()
    try:
        art.get_artwork_by_layers(new_name=data['new_name'],psd_filename='template.psd', layers=data['layers'])
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return "Ocorreu um erro ao gerar imagem.", 500
    finally:
        pythoncom.CoUninitialize()
    return 'A imagem foi gerada com sucesso!', 200

@app.route("/files", methods=["POST"])
def post_file():
    file1 = request.files.get("myFile")
    file_name = file1.filename
    file1.save(os.path.join(DIRETORIO_UPLOAD, file_name))
    print(file_name)
    return '', 201


app.run(port=5000, host='localhost', debug=True)

