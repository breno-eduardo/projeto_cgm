import React, { useState, useEffect } from "react";
import axios from "axios";
import "../style.css"; // Certifique-se de que o CSS está correto

const UploadFile = () => {
  const [file, setFile] = useState(null);
  const [files, setFiles] = useState([]);

  useEffect(() => {
    fetchFiles(); // Carregar arquivos ao iniciar a página
  }, []);

  const fetchFiles = async () => {
    try {
      const response = await axios.get("http://127.0.0.1:5000/files");
      setFiles(response.data.files);
    } catch (error) {
      console.error("Erro ao buscar arquivos", error);
    }
  };

  const handleFileChange = (event) => {
    setFile(event.target.files[0]);
  };

  const handleUpload = async () => {
    if (!file) {
      alert("Selecione um arquivo!");
      return;
    }

    const formData = new FormData();
    formData.append("file", file);

    try {
      await axios.post("http://127.0.0.1:5000/upload", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      alert("Arquivo enviado com sucesso!");
      setFile(null);
      fetchFiles(); // Atualiza a lista após o upload
    } catch (error) {
      alert("Erro ao enviar arquivo");
      console.error(error);
    }
  };

  return (
    <div className="content">
      <h1>📁 Arquivos Enviados</h1>

      <div className="upload-container">
        <input type="file" id="file-upload" accept=".xlsx,.xls" onChange={handleFileChange} />
        <label htmlFor="file-upload" className="upload-label"> Escolher Arquivo</label>
        <button className="upload-btn" onClick={handleUpload}>Enviar</button>
      </div>

      {/* Lista de arquivos enviados */}
      <ul className="file-list">
        {files.length === 0 ? (
          <p className="empty-message">Nenhum arquivo enviado.</p>
        ) : (
          files.map((filename) => (
            <li key={filename} className="file-item">
              <span>📄 {filename}</span>
            </li>
          ))
        )}
      </ul>
    </div>
  );
};

export default UploadFile;
