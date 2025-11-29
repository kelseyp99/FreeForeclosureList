
import React from "react";
import AuctionParametersPage from "./AuctionParameters";

export default function AuctionsPanel() {
  return (
    <div className="container" style={{ minHeight: '100vh', width: '100vw', background: '#fff' }}>
      <main className="main-content" style={{ padding: '40px 32px 0 32px', minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
        <h2 style={{ position: 'sticky', top: 0, background: '#fff', zIndex: 10, padding: '12px 0 8px 0', margin: 0, borderBottom: '1px solid #eee' }}>
          Auction Parameters (Control Table)
        </h2>
        <div style={{
          flex: 1,
          minHeight: 0,
          overflowY: 'auto',
          overflowX: 'auto',
          maxHeight: '70vh',
          marginTop: 8,
          width: '100%',
          boxSizing: 'border-box',
          position: 'relative',
          background: '#fff',
        }}>
          <AuctionParametersPage tableMinWidth={1200} />
        </div>
      </main>
      <footer className="footer">
        <div style={{display: 'flex', alignItems: 'center', gap: 8}}>
          ©{new Date().getFullYear()} by FreeForeclosureList.net
        </div>
      </footer>
    </div>
  );
}
