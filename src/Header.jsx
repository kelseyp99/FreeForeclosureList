
import React from "react";
import fflIcon from './assets/FFL icon.png';

export default function Header() {
  return (
    <header style={{
      marginBottom: 32,
      textAlign: 'left',
      width: '100%',
      display: 'flex',
      alignItems: 'center',
      padding: '24px 48px 24px 48px',
      background: 'linear-gradient(90deg, #fffbe6 0%, #f7c873 100%)',
      borderBottom: '4px solid #f7c873',
      boxShadow: '0 4px 24px rgba(0,0,0,0.10)',
      position: 'relative',
      zIndex: 20
    }}>
      <div style={{
        display: 'flex',
        alignItems: 'center',
        background: 'rgba(255,255,255,0.85)',
        borderRadius: 24,
        padding: '12px 32px 12px 12px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.06)'
      }}>
        <img
          src={fflIcon}
          alt="FFL Icon"
          style={{
            height: 120,
            width: 120,
            verticalAlign: 'middle',
            marginRight: 32,
            borderRadius: 20,
            boxShadow: '0 4px 16px rgba(0,0,0,0.13)',
            border: '3px solid #f7c873',
            background: '#fffbe6'
          }}
        />
        <span style={{
          display: 'inline-block',
          fontSize: '1.7em',
          fontWeight: 600,
          letterSpacing: '0.01em',
          color: '#7a5c1c',
          padding: '0.2em 0.6em',
          borderRadius: '12px',
          background: 'linear-gradient(90deg, #fffbe6 60%, #f7c873 100%)',
          border: '1.5px solid #f7c873',
          boxShadow: '0 2px 8px rgba(0,0,0,0.04)',
          maxWidth: 700,
          textShadow: '0 2px 8px #fffbe6, 0 1px 0 #fff'
        }}>
          AI-generated foreclosure information: timely, accurate, and a benefit to both investor-buyers and distressed property owners
        </span>
      </div>
    </header>
  );
}
