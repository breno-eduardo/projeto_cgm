from flask import Flask, request, jsonify
from flask_cors import CORS
import os

app = Flask(__name__)
CORS(app)  # Habilita comunicação com o React

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"message": "Nenhum arquivo enviado!"}), 400

    file = request.files["file"]
    file.save(os.path.join(UPLOAD_FOLDER, file.filename))

    return jsonify({"message": f"Arquivo {file.filename} salvo com sucesso!"})

if __name__ == "__main__":
    app.run(debug=True)
