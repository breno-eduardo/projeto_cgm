import React, { useState, useEffect } from "react";
import axios from "axios";
import "../style.css"; // Certifique-se de que esse CSS está atualizado

const VisualizarPlanilhas = () => {
  const [files, setFiles] = useState([]);

  useEffect(() => {
    fetchFiles();
  }, []);

  const fetchFiles = async () => {
    try {
      const response = await axios.get("http://127.0.0.1:5000/files");
      setFiles(response.data.files);
    } catch (error) {
      console.error("Erro ao buscar arquivos", error);
    }
  };

  const handleProcess = async (filename) => {
    try {
      const response = await axios.get(`http://127.0.0.1:5000/process-file/${filename}`);
      alert(response.data.message);
    } catch (error) {
      alert("Erro ao processar o arquivo");
      console.error(error);
    }
  };

  const handleDownload = (filename) => {
    window.open(`http://127.0.0.1:5000/download/${filename}`, "_blank");
  };

  const handleDelete = async (filename) => {
    try {
      await axios.delete(`http://127.0.0.1:5000/delete/${filename}`);
      alert(`Arquivo ${filename} excluído com sucesso!`);
      fetchFiles(); // Atualiza a lista de arquivos após a exclusão
    } catch (error) {
      alert("Erro ao excluir o arquivo");
      console.error(error);
    }
  };

  return (
    <div className="content1">
      <h1>📄 Visualizar Planilhas</h1>

      {/* Lista de arquivos */}
      <ul className="file-list">
        {files.map((filename) => (
          <li key={filename} className="file-item">
            <div className="file-name">{filename}</div>
            <div className="button-group">
              <button className="btn-processar" onClick={() => handleProcess(filename)}>Processar</button>
              <button className="btn-download" onClick={() => handleDownload(filename)}>Baixar Processado</button>
              <button className="btn-excluir" onClick={() => handleDelete(filename)}>Excluir</button>
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
};

export default VisualizarPlanilhas;
