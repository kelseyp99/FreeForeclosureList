
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
      paddingLeft: 32,
      paddingTop: 16,
      paddingBottom: 16,
      background: '#fff',
      boxShadow: '0 2px 12px rgba(0,0,0,0.07)'
    }}>
      <img
        src={fflIcon}
        alt="FFL Icon"
        style={{
          height: 80,
          width: 80,
          verticalAlign: 'middle',
          marginRight: 24,
          borderRadius: 12,
          boxShadow: '0 2px 8px rgba(0,0,0,0.10)'
        }}
      />
      <span style={{
        display: 'inline-block',
        fontSize: '1.5em',
        fontWeight: 500,
        letterSpacing: '0.02em',
        color: '#333',
        padding: '0.1em 0.4em',
        borderRadius: '8px',
        background: 'none',
        border: 'none',
        boxShadow: 'none',
        maxWidth: 600
      }}>
        AI-generated foreclosure information: timely, accurate, and a benefit to both investor-buyers and distressed property owners
      </span>
    </header>
  );
}
