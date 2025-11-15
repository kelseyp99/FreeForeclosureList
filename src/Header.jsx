
import React from "react";
import fflIcon from './assets/FFL icon.png';

export default function Header() {
  return (
    <header style={{ marginBottom: 32, textAlign: 'center', width: '100%' }}>
      <img
        src={fflIcon}
        alt="FFL Icon"
        style={{
          height: 48,
          width: 48,
          verticalAlign: 'middle',
          marginRight: 16,
          borderRadius: 8,
          boxShadow: '0 2px 8px rgba(0,0,0,0.10)'
        }}
      />
      <span style={{
        display: 'inline-block',
        fontSize: '2.5em',
        fontWeight: 700,
        letterSpacing: '0.04em',
        color: '#f7c873',
        textShadow: '1px 2px 8px #2222, 0 1px 0 #fff',
        padding: '0.1em 0.4em',
        borderRadius: '12px',
        background: 'linear-gradient(90deg, #fffbe6 60%, #f7c873 100%)',
        border: '2px solid #f7c873',
        boxShadow: '0 2px 12px rgba(0,0,0,0.07)'
      }}>
        Free Foreclosure List
      </span>
    </header>
  );
}
