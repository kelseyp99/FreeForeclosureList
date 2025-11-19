

import React, { useState } from "react";
import { BrowserRouter as Router, Routes, Route, Link } from "react-router-dom";
import FloridaCountiesSidebar from "./components/FloridaCountiesSidebar";
import AuctionsPanel from "./pages/AuctionsPanel";
import Header from "./Header";
import reactLogo from "./assets/react.svg";
import GoogleAuthButton from "./GoogleAuthButton";
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
          boxShadow: open ? '0 2px 8px #f7c87355' : 'none',
          transition: 'box-shadow 0.2s'
        }}
        aria-expanded={open}
        aria-controls="auctions-dropdown"
      >
        Auctions {open ? '▲' : '▼'}
      </button>
      {open && (
        <div id="auctions-dropdown" style={{
          maxHeight: 340,
          overflowY: 'auto',
          background: '#fffbe6',
          border: '1px solid #f7c873',
          borderRadius: 6,
          boxShadow: '0 2px 12px #f7c87333',
          marginTop: 2,
          padding: '4px 0',
          zIndex: 10,
          position: 'relative',
        }}>
          <FloridaCountiesSidebar onSelectReport={onSelectReport} />
        </div>
      )}
    </div>
  );
}

// List of all Florida counties (alphabetical, no 'County' in label)
const FLORIDA_COUNTIES = [
  "Alachua", "Baker", "Bay", "Bradford", "Brevard", "Broward", "Calhoun", "Charlotte", "Citrus", "Clay", "Collier", "Columbia", "DeSoto", "Dixie", "Duval", "Escambia", "Flagler", "Franklin", "Gadsden", "Gilchrist", "Glades", "Gulf", "Hamilton", "Hardee", "Hendry", "Hernando", "Highlands", "Hillsborough", "Holmes", "Indian River", "Jackson", "Jefferson", "Lafayette", "Lake", "Lee", "Leon", "Levy", "Liberty", "Madison", "Manatee", "Marion", "Martin", "Miami-Dade", "Monroe", "Nassau", "Okaloosa", "Okeechobee", "Orange", "Osceola", "Palm Beach", "Pasco", "Pinellas", "Polk", "Putnam", "St. Johns", "St. Lucie", "Santa Rosa", "Sarasota", "Seminole", "Sumter", "Suwannee", "Taylor", "Union", "Volusia", "Wakulla", "Walton", "Washington"
];

function toCountyPath(name) {
  // Convert county name to lowercase, remove spaces/dots, and use dashes for multi-part names
  return "/" + name.toLowerCase().replace(/\./g, '').replace(/ /g, '').replace(/-/g, '');
}

function Home() {
  const [selectedCounty, setSelectedCounty] = React.useState("");
  const [selectedSaleType, setSelectedSaleType] = React.useState("");

  // Build report file name
  let reportSrc = "";
  if (selectedCounty && selectedSaleType) {
    reportSrc = `/reports/sales_report_${selectedCounty.toLowerCase().replace(/\s/g, "_")}_${selectedSaleType.toLowerCase().replace(/\s/g, "")}.html`;
  }

  return (
    <>
      <Header />
      <div className="container" style={{ display: 'flex', minHeight: '100vh' }}>
        <aside style={{ minWidth: 220, maxWidth: 280, background: '#f7f7f7', padding: '32px 8px 16px 8px', boxShadow: '2px 0 8px #eee', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
          <nav style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 32, width: '100%' }}>
            <a href="/" style={{ color: '#0077cc', textDecoration: 'none', fontWeight: 600, fontSize: 17 }}>Home</a>
            <a href="/auctions" style={{ color: '#0077cc', textDecoration: 'none', fontWeight: 600, fontSize: 17 }}>Auction Parameters</a>
            <AuctionsMenu onSelectReport={(county, saleType) => {
              setSelectedCounty(county);
              setSelectedSaleType(saleType === 'UiPath' ? 'foreclosure' : 'taxdeed');
            }} />
          </nav>
          {/* AdSense Ad below menu */}
          <div style={{ width: '100%', minWidth: 100, height: 120, background: '#f7f7f7', border: '1px solid #eee', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 14, color: '#aaa', marginTop: 16 }}>
            AdSense Ad (Sidebar)
          </div>
        </aside>

        <div style={{ flex: 1, display: 'flex', flexDirection: 'row' }}>
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
            <main className="main-content" style={{ padding: '40px 32px 0 32px', flex: 1 }}>
              <div style={{ maxWidth: 900 }}>
                <strong>Hello. We're FreeForeclosureList.net</strong>
                <p>Welcome to FreeForeclosureList.net, your premier destination for accessing comprehensive real estate distressed property listings. Powered by cutting-edge AI and Robotic Process Automation, we revolutionize the way you explore foreclosure properties. Unlike traditional county foreclosure lists, we go above and beyond by curating additional insights sourced from the web, providing you with a one-stop solution for all your real estate investment needs.</p>
                <p>Understanding the demands of modern investors, we offer invaluable features such as direct links to various real estate platforms, county property appraisers, and clerks of court. Our platform delivers more than just basic information; we provide estimated property values, judgment amounts for foreclosure cases, and opening bid amounts for Tax Deed sales. This empowers you to gauge potential equity and focus your efforts efficiently. By identifying properties where lenders are likely to halt bidding at the judgment amount, we save you valuable time. Moreover, you may discover opportunities to connect with property owners who owe less than the judgment amount, opening avenues for direct purchase.</p>
                <p>In addition to our comprehensive foreclosure data, we also offer exclusive access to sales information from counties, including proprietary and hard-to-obtain lists.</p>
                <p><em>Please note that FreeForeclosureList.net is currently in its prototype stage. Expect significant enhancements and updates in the coming months and weeks as we strive to provide you with an unparalleled user experience.</em></p>
                {reportSrc && (
                  <div style={{ marginTop: 32 }}>
                    <h3>{selectedCounty} County {selectedSaleType === 'foreclosure' ? 'Foreclosure' : 'Tax Deed'} Report</h3>
                    <iframe
                      src={reportSrc}
                      title="County Sales Report"
                      style={{ width: '100%', minHeight: 600, border: '1px solid #ccc', borderRadius: 8 }}
                    />
                  </div>
                )}
              </div>
              {/* Auction Parameters table moved to Auctions page */}
            </main>
            <footer className="footer">
              <div style={{display: 'flex', alignItems: 'center', gap: 8}}>
                ©{new Date().getFullYear()} by FreeForeclosureList.net
                <img src={reactLogo} alt="React" style={{height: 24, width: 24, margin: '0 4px'}} />
                - Built with React
              </div>
            </footer>
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

function App() {
  const [selectedCounty, setSelectedCounty] = React.useState("");
  const [selectedSaleType, setSelectedSaleType] = React.useState("");

  // Build report file name
  let reportSrc = "";
  if (selectedCounty && selectedSaleType) {
    reportSrc = `/reports/sales_report_${selectedCounty.toLowerCase().replace(/\s/g, "_")}_${selectedSaleType.toLowerCase().replace(/\s/g, "")}.html`;
  }

  return (
    <Router>
      <Header />
      <div className="container" style={{ display: 'flex', minHeight: '100vh' }}>
        <aside style={{ minWidth: 220, maxWidth: 280, background: '#f7f7f7', padding: '32px 8px 16px 8px', boxShadow: '2px 0 8px #eee', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
          <nav style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 32, width: '100%' }}>
            <a href="/" style={{ color: '#0077cc', textDecoration: 'none', fontWeight: 600, fontSize: 17 }}>Home</a>
            <a href="/auctions" style={{ color: '#0077cc', textDecoration: 'none', fontWeight: 600, fontSize: 17 }}>Auction Parameters</a>
            <AuctionsMenu onSelectReport={(county, saleType) => {
              setSelectedCounty(county);
              setSelectedSaleType(saleType === 'UiPath' ? 'foreclosure' : 'taxdeed');
            }} />
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
                <Route path="/auctions" element={<AuctionsPanel />} />
                <Route path="/" element={
                  <div style={{ maxWidth: 900 }}>
                    <strong>Hello. We're FreeForeclosureList.net</strong>
                    <p>Welcome to FreeForeclosureList.net, your premier destination for accessing comprehensive real estate distressed property listings. Powered by cutting-edge AI and Robotic Process Automation, we revolutionize the way you explore foreclosure properties. Unlike traditional county foreclosure lists, we go above and beyond by curating additional insights sourced from the web, providing you with a one-stop solution for all your real estate investment needs.</p>
                    <p>Understanding the demands of modern investors, we offer invaluable features such as direct links to various real estate platforms, county property appraisers, and clerks of court. Our platform delivers more than just basic information; we provide estimated property values, judgment amounts for foreclosure cases, and opening bid amounts for Tax Deed sales. This empowers you to gauge potential equity and focus your efforts efficiently. By identifying properties where lenders are likely to halt bidding at the judgment amount, we save you valuable time. Moreover, you may discover opportunities to connect with property owners who owe less than the judgment amount, opening avenues for direct purchase.</p>
                    <p>In addition to our comprehensive foreclosure data, we also offer exclusive access to sales information from counties, including proprietary and hard-to-obtain lists.</p>
                    <p><em>Please note that FreeForeclosureList.net is currently in its prototype stage. Expect significant enhancements and updates in the coming months and weeks as we strive to provide you with an unparalleled user experience.</em></p>
                    {reportSrc && (
                      <div style={{ marginTop: 32 }}>
                        <h3>{selectedCounty} County {selectedSaleType === 'foreclosure' ? 'Foreclosure' : 'Tax Deed'} Report</h3>
                        <iframe
                          src={reportSrc}
                          title="County Sales Report"
                          style={{ width: '100%', minHeight: 600, border: '1px solid #ccc', borderRadius: 8 }}
                        />
                      </div>
                    )}
                  </div>
                } />
              </Routes>
            </main>
            <footer className="footer">
              <div style={{display: 'flex', alignItems: 'center', gap: 8}}>
                ©{new Date().getFullYear()} by FreeForeclosureList.net
                <img src={reactLogo} alt="React" style={{height: 24, width: 24, margin: '0 4px'}} />
                - Built with React
              </div>
            </footer>
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
    </Router>
  );
}

export default App;
