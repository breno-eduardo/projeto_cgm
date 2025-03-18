from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

@app.route("/")
def home():
    return "API rodando!"

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400
    
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Nome do arquivo inválido"}), 400
    
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)

    return jsonify({"message": f"Arquivo {file.filename} salvo com sucesso!"})

if __name__ == "__main__":
    app.run(debug=True)
