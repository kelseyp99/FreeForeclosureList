import React from "react";
import fflIcon from './assets/FFL icon.png';
import GoogleAuthButton from './GoogleAuthButton';

export default function Header() {
  return (
    <header style={{
      marginBottom: 32,
      textAlign: 'left',
      width: '100%',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '24px 0',
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
        boxShadow: '0 2px 8px rgba(0,0,0,0.06)',
        width: '100%',
        maxWidth: 1400,
        justifyContent: 'space-between',
      }}>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <img
            src={fflIcon}
            alt="FFL Icon"
            style={{
              height: 110,
              width: 110,
              verticalAlign: 'middle',
              marginRight: 28,
              borderRadius: 20,
              boxShadow: '0 4px 16px rgba(0,0,0,0.13)',
              border: '3px solid #f7c873',
              background: '#fffbe6'
            }}
          />
          <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
            <span style={{
              fontSize: '2.3em',
              fontWeight: 900,
              color: '#1a4d7a',
              letterSpacing: '0.04em',
              marginBottom: 2,
              textShadow: '0 2px 8px #fffbe6, 0 1px 0 #fff',
              fontFamily: 'inherit',
              display: 'inline-block',
              verticalAlign: 'middle',
              lineHeight: 1.1,
            }}>
              True<span style={{color:'#f7c873'}}>Foreclosure</span>
            </span>
            <span style={{
              display: 'inline-block',
              fontSize: '1.15em',
              fontWeight: 500,
              letterSpacing: '0.01em',
              color: '#7a5c1c',
              padding: '0.18em 0.6em',
              borderRadius: '10px',
              background: 'linear-gradient(90deg, #fffbe6 60%, #f7c873 100%)',
              border: '1.2px solid #f7c873',
              boxShadow: '0 2px 8px rgba(0,0,0,0.04)',
              maxWidth: 700,
              textShadow: '0 2px 8px #fffbe6, 0 1px 0 #fff',
              marginTop: 4
            }}>
              AI-generated foreclosure information: timely, accurate, and a benefit to both investor-buyers and distressed property owners
            </span>
          </div>
        </div>
        <div>
          <GoogleAuthButton />
        </div>
      </div>
    </header>
  );
}
