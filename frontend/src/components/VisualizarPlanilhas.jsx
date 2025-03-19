import React, { useState, useEffect } from "react";
import axios from "axios";
import * as XLSX from "xlsx";
import Modal from "react-modal";
import "../App.css";
import "../style.css";

Modal.setAppElement("#root");

const VisualizarPlanilhas = () => {
  const [files, setFiles] = useState([]);
  const [modalIsOpen, setModalIsOpen] = useState(false);
  const [tableData, setTableData] = useState([]);

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

  const handleDelete = async (filename) => {
    try {
      await axios.delete(`http://127.0.0.1:5000/delete/${filename}`);
      alert("Arquivo excluído com sucesso!");
      fetchFiles();
    } catch (error) {
      alert("Erro ao excluir arquivo");
      console.error(error);
    }
  };

  const handleOpenFile = async (filename) => {
    try {
      const response = await axios.get(`http://127.0.0.1:5000/view/${filename}`, {
        responseType: "arraybuffer",
      });

      const data = new Uint8Array(response.data);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      setTableData(jsonData);
      setModalIsOpen(true);
    } catch (error) {
      alert("Erro ao visualizar arquivo");
      console.error(error);
    }
  };
  const handleProcess = async (filename) => {
    try {
      const response = await axios.get(`http://127.0.0.1:5000/process-file/${filename}`);
      alert(response.data.message);
      console.log(response.data.output); // Exibe o resultado do processamento no console
    } catch (error) {
      alert("Erro ao processar o arquivo");
      console.error(error);
    }
  };
  
  
  const handleDownload = (filename) => {
    window.open(`http://127.0.0.1:5000/download/${filename}`, "_blank");
  };
  
  return (
    <div className="content">
      <h1>Visualizar Planilhas</h1>
      <ul className="file-list">
        {files.map((filename) => (
          <li key={filename} className="file-item">
            <span>{filename}</span>
            <div className="button-group">
                <button className="btn-processar" onClick={() => handleProcess(filename)}>Processar</button>
                <button className="btn-visualizar" onClick={() => handleOpenFile(filename)}>Visualizar</button>
                <button className="btn-download" onClick={() => handleDownload(filename)}>Baixar Processado</button>
                <button className="btn-excluir" onClick={() => handleDelete(filename)}>Excluir</button>
            </div>



          </li>
        ))}
      </ul>

      {/* Modal para exibir os dados da planilha */}
      <Modal
        isOpen={modalIsOpen}
        onRequestClose={() => setModalIsOpen(false)}
        className="modal"
        overlayClassName="overlay"
      >
        <button className="btn-fechar" onClick={() => setModalIsOpen(false)}>Fechar</button>
        <table className="excel-table">
          <tbody>
            {tableData.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, cellIndex) => (
                  <td key={cellIndex}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </Modal>
    </div>
  );
};

export default VisualizarPlanilhas;
