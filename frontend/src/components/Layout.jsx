// src/components/Layout.jsx
import React from "react";
import { Link } from "react-router-dom";
import "../App.css"; // Importando o CSS original
import "../index.css";

const Layout = ({ children }) => {
  return (
    <div className="container">
      {/* Sidebar */}
      <div className="sidebar">
        <img src="/logo.png" alt="Logo" className="logo" />
        <ul>
          <li><h2><Link to="/">Menu</Link></h2></li>
          <li><Link to="/upload">Carregar Arquivos</Link></li>
          <li><Link to="/checklist1">Checklist D1</Link></li>
          <li><Link to="/checklist2">Checklist D2</Link></li>
          <li><Link to="/checklist3">Checklist D3</Link></li>
          <li><Link to="/checklist4">Checklist D4</Link></li>
          <li><Link to="/visualizar">Visualizar planilhas</Link></li>
        </ul>
      </div>

      {/* Conteúdo principal */}
      <div className="content">{children}</div>
    </div>
  );
};

export default Layout;
