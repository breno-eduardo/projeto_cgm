import React, { useState } from "react";
import axios from "axios";

function App() {
  const [file, setFile] = useState(null);

  const handleFileChange = (event) => {
    setFile(event.target.files[0]);
  };

  const handleUpload = async () => {
    if (!file) {
      alert("Selecione um arquivo primeiro!");
      return;
    }

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await axios.post("http://127.0.0.1:5000/upload", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      alert(response.data.message);
    } catch (error) {
      console.error("Erro no upload:", error);
      alert("Erro ao enviar o arquivo.");
    }
  };

  return (
    <div style={{ textAlign: "center", padding: "50px" }}>
      <h1>Upload de Arquivos</h1>
      <input type="file" onChange={handleFileChange} />
      <button onClick={handleUpload} style={{ marginLeft: "10px" }}>
        Enviar
      </button>
    </div>
  );
}

export default App;
