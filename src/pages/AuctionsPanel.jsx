import React from "react";
import Header from "../Header";
import AuctionParametersPage from "./AuctionParameters";

export default function AuctionsPanel() {
  return (
    <>
      <Header />
  <div className="container" style={{ display: 'flex', minHeight: '100vh', width: '100vw' }}>
        <aside style={{
          minWidth: 110,
          maxWidth: 180,
          background: '#f7f7f7',
          padding: '32px 8px 16px 8px',
          boxShadow: '2px 0 8px #eee',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'flex-start',
          position: 'sticky',
          top: 0,
          height: '100vh',
          zIndex: 100
        }}>
          <nav style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 32, width: '100%' }}>
            <a href="/" style={{ textDecoration: 'none', color: '#7a5c1c' }}>Home</a>
            <a href="/auctions" style={{ textDecoration: 'none', color: '#7a5c1c', fontWeight: 600 }}>Auctions</a>
          </nav>
        </aside>
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minHeight: '100vh' }}>
          <main className="main-content" style={{ padding: '40px 32px 0 32px', flex: 1, display: 'flex', flexDirection: 'column' }}>
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
      </div>
    </>
  );
}
