

import React from "react";
import { BrowserRouter as Router, Routes, Route, Link } from "react-router-dom";

import OrangeCounty from "./pages/OrangeCounty";
import OsceolaCounty from "./pages/OsceolaCounty";
import SeminoleCounty from "./pages/SeminoleCounty";
import Header from "./Header";
import "./App.css";

function Home() {
  return (
    <>
      <div style={{ width: '100vw', background: '#fff', boxShadow: '0 2px 12px rgba(0,0,0,0.07)', zIndex: 10 }}>
        <h1 style={{
          fontSize: '2.5em',
          fontWeight: 700,
          letterSpacing: '0.04em',
          color: '#f7c873',
          textShadow: '1px 2px 8px #2222, 0 1px 0 #fff',
          padding: '0.3em 0.4em 0.2em 0.4em',
          borderRadius: '0 0 18px 18px',
          background: 'linear-gradient(90deg, #fffbe6 60%, #f7c873 100%)',
          borderBottom: '2px solid #f7c873',
          boxShadow: '0 2px 12px rgba(0,0,0,0.07)',
          margin: 0,
          textAlign: 'center',
        }}>
          Free Foreclosure List
        </h1>
      </div>
      <div className="container" style={{ display: 'flex', minHeight: '100vh' }}>
        <aside style={{ width: 280, background: '#f7f7f7', padding: '32px 16px 16px 16px', boxShadow: '2px 0 8px #eee', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
          <h2 style={{ fontSize: 24, marginBottom: 16 }}>FreeForeclosureList.net</h2>
          <nav style={{ display: 'flex', flexDirection: 'column', gap: 12, marginBottom: 32, width: '100%' }}>
            <Link to="/">Home</Link>
            <Link to="/orange">Orange County</Link>
            <Link to="/osceola">Osceola County</Link>
            <Link to="/seminole">Seminole County</Link>
          </nav>
        </aside>
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
          <main className="main-content" style={{ padding: '40px 32px 0 32px', flex: 1 }}>
            <div style={{ maxWidth: 900 }}>
              <strong>Hello. We're FreeForeclosureList.net</strong>
              <p>Welcome to FreeForeclosureList.net, your premier destination for accessing comprehensive real estate distressed property listings. Powered by cutting-edge AI and Robotic Process Automation, we revolutionize the way you explore foreclosure properties. Unlike traditional county foreclosure lists, we go above and beyond by curating additional insights sourced from the web, providing you with a one-stop solution for all your real estate investment needs.</p>
              <p>Understanding the demands of modern investors, we offer invaluable features such as direct links to various real estate platforms, county property appraisers, and clerks of court. Our platform delivers more than just basic information; we provide estimated property values, judgment amounts for foreclosure cases, and opening bid amounts for Tax Deed sales. This empowers you to gauge potential equity and focus your efforts efficiently. By identifying properties where lenders are likely to halt bidding at the judgment amount, we save you valuable time. Moreover, you may discover opportunities to connect with property owners who owe less than the judgment amount, opening avenues for direct purchase.</p>
              <p>In addition to our comprehensive foreclosure data, we also offer exclusive access to sales information from counties, including proprietary and hard-to-obtain lists.</p>
              <p><em>Please note that FreeForeclosureList.net is currently in its prototype stage. Expect significant enhancements and updates in the coming months and weeks as we strive to provide you with an unparalleled user experience.</em></p>
            </div>
          </main>
          <footer className="footer">
            <div>FreeForeclosureList.net</div>
            <div>©{new Date().getFullYear()} by FreeForeclosureList. Proudly created with Wix.com</div>
          </footer>
        </div>
      </div>
    </>
  );
}

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/orange" element={<OrangeCounty />} />
        <Route path="/osceola" element={<OsceolaCounty />} />
        <Route path="/seminole" element={<SeminoleCounty />} />
      </Routes>
    </Router>
  );
}

export default App;
