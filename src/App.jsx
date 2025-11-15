

import React from "react";
import { BrowserRouter as Router, Routes, Route, Link } from "react-router-dom";


import OrangeCounty from "./pages/OrangeCounty";
import OsceolaCounty from "./pages/OsceolaCounty";
import SeminoleCounty from "./pages/SeminoleCounty";
import Header from "./Header";
import reactLogo from "./assets/react.svg";
import "./App.css";

function Home() {
  return (
    <>
      <Header />
      <div className="container" style={{ display: 'flex', minHeight: '100vh' }}>
  <aside style={{ minWidth: 110, maxWidth: 130, background: '#f7f7f7', padding: '32px 8px 16px 8px', boxShadow: '2px 0 8px #eee', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
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
            <div style={{display: 'flex', alignItems: 'center', gap: 8}}>
              ©{new Date().getFullYear()} by FreeForeclosureList.net
              <img src={reactLogo} alt="React" style={{height: 24, width: 24, margin: '0 4px'}} />
              - Built with React
            </div>
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
