import React from "react";
import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import UploadFile from "./components/UploadFile";
import Layout from "./components/Layout";
import "./style.css";
//import "./checando.css";
import "./botoes.css";
import VisualizarPlanilhas from "./components/VisualizarPlanilhas";
import Menu from "./components/Menu"; // <- Adicionando o import do Menu.jsx



function App() {
  return (
    <Router>
      <Layout>
        <Routes>
          <Route path="/" element={<Menu />} />
          <Route path="/upload" element={<UploadFile />} />
          <Route path="/checklist1" element={<h1>Checklist 1</h1>} />
          <Route path="/checklist2" element={<h1>Checklist 2</h1>} />
          <Route path="/checklist3" element={<h1>Checklist 3</h1>} />
          <Route path="/checklist4" element={<h1>Checklist 4</h1>} />
          <Route path="/visualizar" element={<VisualizarPlanilhas />} />
        </Routes>
      </Layout>
    </Router>
  );
}

export default App;