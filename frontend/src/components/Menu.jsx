// src/components/Menu.jsx
import React from "react";
import "../style.css"; // Certifique-se de que o CSS existe

const Menu = () => {
  return (
    <div className="menu-container">
      {/* Imagem no topo */}
      <img src="/logo.png" className="menu-image" />

      <h1>Checklist Verificação Ranking dos Municípios</h1>
      <p>Bem-vindo ao sistema de verificação do ranking dos municípios.</p>
    </div>
  );
};

export default Menu;
