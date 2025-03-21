from flask import Flask, request, jsonify, send_from_directory, Response
from flask_cors import CORS
import os
import subprocess

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed_files"  # Pasta para arquivos processados
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)  # Garante que a pasta de processados exista

# Rota para upload de arquivos
@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"message": "Nenhum arquivo enviado!"}), 400

    file = request.files["file"]
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    return jsonify({"message": f"Arquivo {file.filename} salvo com sucesso!"})

# Rota para processar arquivos
@app.route("/process-file/<filename>", methods=["GET"])
def process_file(filename):
    """Executa o script e salva o arquivo processado na pasta correta"""
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    processed_path = os.path.join(PROCESSED_FOLDER, filename)

    if not os.path.exists(file_path):
        return jsonify({"error": "Arquivo não encontrado"}), 404

    try:
        result = subprocess.run(["python", "D1_convertido_otimizado.py", file_path, processed_path], capture_output=True, text=True)

        if os.path.exists(processed_path):
            return jsonify({"message": "Arquivo processado com sucesso!", "output": result.stdout})
        else:
            return jsonify({"error": "Erro no processamento do arquivo."}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Rota para baixar o arquivo processado
@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    """Retorna o arquivo processado para download com cabeçalhos para evitar cache"""
    file_path = os.path.join(PROCESSED_FOLDER, filename)

    if not os.path.exists(file_path):
        return jsonify({"error": "Arquivo não encontrado!"}), 404

    response = send_from_directory(PROCESSED_FOLDER, filename, as_attachment=True)
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"

    return response

# Rota para listar arquivos salvos
@app.route("/files", methods=["GET"])
def list_files():
    """Retorna a lista de arquivos salvos no backend"""
    files = os.listdir(UPLOAD_FOLDER)
    return jsonify({"files": files})

# Rota para visualizar arquivos diretamente
@app.route("/view/<path:filename>", methods=["GET"])
def view_file(filename):
    """Retorna o arquivo para visualização"""
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    if not os.path.exists(file_path):
        return jsonify({"error": "Arquivo não encontrado!"}), 404

    try:
        return send_from_directory(UPLOAD_FOLDER, filename)
    except Exception as e:
        return jsonify({"error": f"Erro ao acessar o arquivo: {str(e)}"}), 500


if __name__ == "__main__":
    app.run(debug=True)
