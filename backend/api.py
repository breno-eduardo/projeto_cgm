from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import subprocess

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed_files"  # Pasta para arquivos processados
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)  # Garante que a pasta de processados exista

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
        
        # Verifica se o processamento criou um arquivo na pasta correta
        if os.path.exists(processed_path):
            return jsonify({"message": "Arquivo processado com sucesso!", "output": result.stdout})
        else:
            return jsonify({"error": "Erro no processamento do arquivo."}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Rota para baixar o arquivo processado
@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    """Retorna o arquivo processado para download"""
    if not os.path.exists(PROCESSED_FOLDER):
        return jsonify({"error": "Pasta de arquivos processados não encontrada!"}), 500

    return send_from_directory(PROCESSED_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
