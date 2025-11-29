

import React, { useState } from "react";
import params from './config/params';
import { useNavigate } from "react-router-dom";
import { BrowserRouter as Router, Routes, Route, Link } from "react-router-dom";
import FloridaCountiesSidebar from "./components/FloridaCountiesSidebar";
import PascoCounty from "./pages/PascoCounty";
import AuctionsPanel from "./pages/AuctionsPanel";
import Header from "./Header";
import reactLogo from "./assets/react.svg";
import GoogleAuthButton from "./GoogleAuthButton";
import GlobalParameterTable from "./components/GlobalParameterTable";
import SalesReportPanel from "./components/SalesReportPanel";
import "./App.css";

// SalesMenu: Head menu item for Sales that toggles the counties menu

function AuctionsMenu({ onSelectReport }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ width: '100%' }}>
      <button
        onClick={() => setOpen((v) => !v)}
        style={{
          width: '100%',
          background: '#f7c873',
          color: '#7a5c1c',
          fontWeight: 600,
          border: '1px solid #e0b24d',
          borderRadius: 6,
          padding: '8px 10px',
          cursor: 'pointer',
          marginBottom: 4,
          fontSize: 16,
          textAlign: 'left',
        }}
      >
        Auctions
      </button>
      {open && (
        <FloridaCountiesSidebar onSelectReport={onSelectReport} />
      )}
    </div>
  );
}

function App() {
  const [selectedCounty, setSelectedCounty] = useState("");
  const [selectedSaleType, setSelectedSaleType] = useState("");
  const [ownerAssocFilter, setOwnerAssocFilter] = useState('include'); // 'exclude', 'include', 'only'
  const ownerAssocWords = (params.owner_assoc_words || '').split(',').map(w => w.trim()).filter(Boolean);
  const reportSrc = selectedCounty && selectedSaleType
    ? `/reports/sales_report_${selectedCounty.toLowerCase().replace(/\s/g, "_")}_${selectedSaleType.toLowerCase().replace(/\s/g, "")}.html`
    : null;
  console.log('APP STATE:', { selectedCounty, selectedSaleType, reportSrc });

  const navigate = useNavigate();
  return (
    <>
      <Header />
      <div className="container" style={{ display: 'flex', minHeight: '100vh' }}>
        <aside style={{ minWidth: 220, maxWidth: 280, background: '#f7f7f7', padding: '32px 8px 16px 8px', boxShadow: '2px 0 8px #eee', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
          {/* Home menu item at the top */}
          <nav style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 32, width: '100%' }}>
            <a
              href="/"
              style={{ color: '#0077cc', textDecoration: 'none', fontWeight: 600, fontSize: 17 }}
              onClick={() => {
                setSelectedCounty("");
                setSelectedSaleType("");
              }}
            >Home</a>
          </nav>
          {/* Auctions menu */}
          <AuctionsMenu onSelectReport={(county, saleType) => {
            setSelectedCounty(county);
            setSelectedSaleType(saleType);
          }} />
          {/* Owner Associations Filter below Auctions */}
          <div style={{ margin: '18px 0 0 0', width: '100%' }}>
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 8 }}>Certificate Holder Type</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4, marginBottom: 16 }}>
              <label style={{ fontWeight: 400, fontSize: 15 }}>
                <input
                  type="radio"
                  name="ownerAssocFilter"
                  value="exclude"
                  checked={ownerAssocFilter === 'exclude'}
                  onChange={() => setOwnerAssocFilter('exclude')}
                /> Exclude Owner Associations
              </label>
              <label style={{ fontWeight: 400, fontSize: 15 }}>
                <input
                  type="radio"
                  name="ownerAssocFilter"
                  value="include"
                  checked={ownerAssocFilter === 'include'}
                  onChange={() => setOwnerAssocFilter('include')}
                /> Include Owner Associations
              </label>
              <label style={{ fontWeight: 400, fontSize: 15 }}>
                <input
                  type="radio"
                  name="ownerAssocFilter"
                  value="only"
                  checked={ownerAssocFilter === 'only'}
                  onChange={() => setOwnerAssocFilter('only')}
                /> Show Only Owner Associations
              </label>
            </div>
            {/* Foreclosure Report Filters */}
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 8 }}>Foreclosure Report Filters</div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              <label style={{ fontWeight: 400, fontSize: 15 }}>
                <input
                  type="checkbox"
                  checked={localStorage.getItem('ffl_filter_timeshare') === '1'}
                  onChange={e => {
                    localStorage.setItem('ffl_filter_timeshare', e.target.checked ? '1' : '0');
                    window.dispatchEvent(new Event('storage'));
                  }}
                /> Hide Timeshare Parcel IDs
              </label>
              <label style={{ fontWeight: 400, fontSize: 15 }}>
                <input
                  type="checkbox"
                  checked={localStorage.getItem('ffl_filter_blank') === '1'}
                  onChange={e => {
                    localStorage.setItem('ffl_filter_blank', e.target.checked ? '1' : '0');
                    window.dispatchEvent(new Event('storage'));
                  }}
                /> Hide Blank Parcel IDs
              </label>
              <label style={{ fontWeight: 400, fontSize: 15, display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                Filter Status:
                <select
                  multiple
                  size={3}
                  style={{ minWidth: 160, maxWidth: 220, fontSize: '1em', marginTop: 4 }}
                  value={(() => {
                    try {
                      return JSON.parse(localStorage.getItem('ffl_filter_status') || '[]');
                    } catch { return []; }
                  })()}
                  onChange={e => {
                    const selected = Array.from(e.target.selectedOptions).map(opt => opt.value);
                    localStorage.setItem('ffl_filter_status', JSON.stringify(selected));
                    window.dispatchEvent(new Event('storage'));
                  }}
                >
                  {/* Status options will be injected by sortable-table.js on first load, so we leave this empty for now */}
                </select>
                <span style={{ fontSize: 12, color: '#888', marginTop: 2 }}>(Hold Ctrl/Cmd to select multiple)</span>
              </label>
            </div>
          </div>
          <nav style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 32, width: '100%' }}>
            <div style={{ marginTop: 18, marginBottom: 2, fontWeight: 700, color: '#7a5c1c', fontSize: 15, letterSpacing: 0.5 }}>Administration</div>
            <a
              href="/auction-parameters"
              style={{ color: '#0077cc', textDecoration: 'none', fontWeight: 600, fontSize: 16, marginLeft: 12 }}
              onClick={() => {
                setSelectedCounty("");
                setSelectedSaleType("");
              }}
            >Auction Parameters</a>
            <a
              href="/global-parameters"
              style={{ color: '#0077cc', textDecoration: 'none', fontWeight: 600, fontSize: 16, marginLeft: 12 }}
              onClick={() => {
                setSelectedCounty("");
                setSelectedSaleType("");
              }}
            >Global Parameters</a>
          </nav>
          {/* AdSense Ad below menu */}
          <div style={{ width: '100%', minWidth: 100, height: 120, background: '#f7f7f7', border: '1px solid #eee', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14, color: '#aaa', marginTop: 16 }}>
            AdSense Ad (Sidebar)
          </div>
        </aside>

        <div style={{ flex: 1, display: 'flex', flexDirection: 'row' }}>
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
            <main className="main-content" style={{ padding: '40px 32px 0 32px', flex: 1 }}>
              <Routes>
                <Route path="/" element={
                  reportSrc ? (
                    <div style={{ maxWidth: 1700, marginTop: 32, position: 'relative' }}>
                      <iframe
                        src={reportSrc}
                        title="County Sales Report"
                        style={{ width: '100%', minHeight: 1200, border: '1px solid #ccc', borderRadius: 8 }}
                      />
                      {/* Debug overlay for iframe src */}
                      <div style={{
                        position: 'absolute',
                        top: 0,
                        right: 0,
                        background: 'rgba(255,255,0,0.85)',
                        color: '#222',
                        padding: '4px 10px',
                        fontSize: 13,
                        borderBottomLeftRadius: 8,
                        zIndex: 10,
                        pointerEvents: 'none',
                      }}>
                        <strong>iframe src:</strong> {reportSrc}
                      </div>
                    </div>
                  ) : (
                    <div style={{ maxWidth: 900 }}>
                      <strong>Hello. We're FreeForeclosureList.net</strong>
                      <p>Welcome to FreeForeclosureList.net, your premier destination for accessing comprehensive real estate distressed property listings. Powered by cutting-edge AI and Robotic Process Automation, we revolutionize the way you explore foreclosure properties. Unlike traditional county foreclosure lists, we go above and beyond by curating additional insights sourced from the web, providing you with a one-stop solution for all your real estate investment needs.</p>
                      <p>Understanding the demands of modern investors, we offer invaluable features such as direct links to various real estate platforms, county property appraisers, and clerks of court. Our platform delivers more than just basic information; we provide estimated property values, judgment amounts for foreclosure cases, and opening bid amounts for Tax Deed sales. This empowers you to gauge potential equity and focus your efforts efficiently. By identifying properties where lenders are likely to halt bidding at the judgment amount, we save you valuable time. Moreover, you may discover opportunities to connect with property owners who owe less than the judgment amount, opening avenues for direct purchase.</p>
                      <p>In addition to our comprehensive foreclosure data, we also offer exclusive access to sales information from counties, including proprietary and hard-to-obtain lists.</p>
                      <p><em>Please note that FreeForeclosureList.net is currently in its prototype stage. Expect significant enhancements and updates in the coming months and weeks as we strive to provide you with an unparalleled user experience.</em></p>
                    </div>
                  )
                } />
                <Route path="/auctions" element={<AuctionsPanel />} />
                <Route path="/global-parameters" element={<GlobalParameterTable />} />
              </Routes>
            </main>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 24, minWidth: 160, marginLeft: 12, marginTop: 40 }}>
            {/* AdSense Ad 1 */}
            <div style={{ width: 160, height: 250, background: '#f7f7f7', border: '1px solid #eee', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14, color: '#aaa' }}>
              AdSense Ad 1
            </div>
            {/* AdSense Ad 2 */}
            <div style={{ width: 160, height: 250, background: '#f7f7f7', border: '1px solid #eee', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14, color: '#aaa' }}>
              AdSense Ad 2
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

export default App;
