import React, { useState, useEffect } from "react";
import params from './config/params';
import { useNavigate, Routes, Route } from "react-router-dom";
import FloridaCountiesSidebar from "./components/FloridaCountiesSidebar";
import AuctionsPanel from "./pages/AuctionsPanel";
import AuctionParametersPage from "./pages/AuctionParameters";
import Header from "./Header";
import GlobalParameterTable from "./components/GlobalParameterTable";
import SalesReportPanel from "./components/SalesReportPanel";
import "./App.css";
import { getStatusFilterArray } from "./utils/statusFilter";

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
  const [hideTimeshare, setHideTimeshare] = useState(() => localStorage.getItem('ffl_filter_timeshare') === '1');
  const [hideBlank, setHideBlank] = useState(() => localStorage.getItem('ffl_filter_blank') === '1');
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [statusFilter, setStatusFilter] = useState(() => {
    try {
      let val = JSON.parse(localStorage.getItem('ffl_filter_status') || '[]');
      if (typeof val === 'string') val = val ? [val] : [];
      if (Array.isArray(val)) return val.filter(Boolean).map(String);
      return [];
    } catch { return []; }
  });
  const ownerAssocWords = (params.owner_assoc_words || '').split(',').map(w => w.trim()).filter(Boolean);
  const reportSrc = selectedCounty && selectedSaleType
    ? `/reports/sales_report_${selectedCounty.toLowerCase().replace(/\s/g, "_")}_${selectedSaleType.toLowerCase().replace(/\s/g, "")}.html`
    : null;
  const navigate = useNavigate();

  function handleExternalClick(e) {
    // Only track left-clicks on anchor tags with target _blank
    if (e.target.tagName === 'A' && e.target.target === '_blank' && e.target.href && analytics) {
      logEvent(analytics, 'click_through', {
        url: e.target.href,
        text: e.target.innerText || undefined,
        location: window.location.pathname
      });
    }
  }

  // Add event listener for click-through tracking
  useEffect(() => {
    document.addEventListener('click', handleExternalClick);
    return () => document.removeEventListener('click', handleExternalClick);
  }, []);

  return (
    <>
      <Header />
      <div className="container" style={{ display: 'flex', minHeight: '100vh', position: 'relative' }}>
        {/* Collapse/Expand Button */}
        <button
          onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
          style={{
            position: 'fixed',
            left: sidebarCollapsed ? 0 : 280,
            top: '50%',
            transform: 'translateY(-50%)',
            zIndex: 1000,
            background: '#f7c873',
            border: '1px solid #e0b24d',
            borderRadius: '0 8px 8px 0',
            padding: '12px 6px',
            cursor: 'pointer',
            fontSize: 18,
            fontWeight: 'bold',
            color: '#7a5c1c',
            transition: 'left 0.3s ease'
          }}
          title={sidebarCollapsed ? 'Show Sidebar' : 'Hide Sidebar'}
        >
          {sidebarCollapsed ? '▶' : '◀'}
        </button>
        
        <aside style={{ 
          minWidth: sidebarCollapsed ? 0 : 220, 
          maxWidth: sidebarCollapsed ? 0 : 280, 
          width: sidebarCollapsed ? 0 : 280,
          background: '#f7f7f7', 
          padding: sidebarCollapsed ? 0 : '32px 8px 16px 8px', 
          boxShadow: '2px 0 8px #eee', 
          display: 'flex', 
          flexDirection: 'column', 
          alignItems: 'flex-start',
          overflow: 'hidden',
          transition: 'all 0.3s ease'
        }}>
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
              <label style={{ fontWeight: 400, fontSize: 15, color: '#7a5c1c' }}>
                <input
                  type="radio"
                  name="ownerAssocFilter"
                  value="exclude"
                  checked={ownerAssocFilter === 'exclude'}
                  onChange={() => setOwnerAssocFilter('exclude')}
                /> Exclude Owner Associations
              </label>
              <label style={{ fontWeight: 400, fontSize: 15, color: '#7a5c1c' }}>
                <input
                  type="radio"
                  name="ownerAssocFilter"
                  value="include"
                  checked={ownerAssocFilter === 'include'}
                  onChange={() => setOwnerAssocFilter('include')}
                /> Include Owner Associations
              </label>
              <label style={{ fontWeight: 400, fontSize: 15, color: '#7a5c1c' }}>
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
              <label style={{ fontWeight: 400, fontSize: 15, color: '#7a5c1c' }}>
                <input
                  type="checkbox"
                  checked={hideTimeshare}
                  onChange={e => {
                    setHideTimeshare(e.target.checked);
                    localStorage.setItem('ffl_filter_timeshare', e.target.checked ? '1' : '0');
                    window.dispatchEvent(new Event('storage'));
                  }}
                /> Hide Timeshare Parcel IDs
              </label>
              <label style={{ fontWeight: 400, fontSize: 15, color: '#7a5c1c' }}>
                <input
                  type="checkbox"
                  checked={hideBlank}
                  onChange={e => {
                    setHideBlank(e.target.checked);
                    localStorage.setItem('ffl_filter_blank', e.target.checked ? '1' : '0');
                    window.dispatchEvent(new Event('storage'));
                  }}
                /> Hide Blank Parcel IDs
              </label>
              <label style={{ fontWeight: 400, fontSize: 15, color: '#7a5c1c', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                Filter Status:
                <select
                  multiple
                  size={4}
                  style={{ minWidth: 160, maxWidth: 220, fontSize: '1em', marginTop: 4 }}
                  value={statusFilter}
                  onChange={e => {
                    const selected = Array.from(e.target.selectedOptions).map(opt => opt.value);
                    setStatusFilter(selected);
                    localStorage.setItem('ffl_filter_status', JSON.stringify(selected));
                    window.dispatchEvent(new Event('storage'));
                  }}
                >
                  <option value="Accepting Proxy">Accepting Proxy</option>
                  <option value="Canceled">Canceled</option>
                  <option value="Presale">Presale</option>
                  <option value="Running">Running</option>
                  <option value="Sold">Sold</option>
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
        </aside>
        {/* Main content area */}
        <main className="main-content" style={{ padding: '40px 32px 0 32px', flex: 1 }}>
          {/* Main iframe report view */}
          {reportSrc ? (
            <iframe
              title="Sales Report"
              src={reportSrc}
              style={{ width: '100%', minHeight: 800, border: 'none', background: '#fff', borderRadius: 8, boxShadow: '0 2px 12px #eee' }}
            />
          ) : (
            <div style={{ color: '#888', fontSize: 18, marginTop: 80, textAlign: 'center' }}>
              Select a county and sale type to view a report.
            </div>
          )}
        </main>
      </div>
    </>
  );
}

function RoutedApp() {
  return (
    <Routes>
      <Route path="/auction-parameters" element={<AuctionsPanel />} />
      <Route path="/global-parameters" element={<GlobalParameterTable />} />
      <Route path="/*" element={<App />} />
    </Routes>
  );
}

export default RoutedApp;
