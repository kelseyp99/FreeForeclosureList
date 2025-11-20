import React, { useEffect, useState } from "react";
import { db } from "../firebase";
import { collection, getDocs } from "firebase/firestore";

const SALE_TYPE_LABELS = {
  UiPath: "Foreclosures",
  UiPathTD: "Tax Deed"
};


export default function FloridaCountiesSidebar({ onSelectReport }) {
  const [params, setParams] = useState([]);
  const [expandedCounty, setExpandedCounty] = useState(null);

  useEffect(() => {
    fetchParams();
  }, []);

  async function fetchParams() {
    const querySnapshot = await getDocs(collection(db, "auction_parameters"));
    const data = querySnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
  setParams(data);
  console.log('Loaded auction_parameters:', data);
  }

  // Build menu structure: { county: ["UiPath", "UiPathTD"] }
  const countySaleTypes = {};
  params.forEach((param) => {
    const county = param.County;
    if (!county) return;
    countySaleTypes[county] = [];
    if (param.UiPath) countySaleTypes[county].push("UiPath");
    if (param.UiPathTD) countySaleTypes[county].push("UiPathTD");
  });

  function handleCountyClick(county) {
    setExpandedCounty(expandedCounty === county ? null : county);
  }

  function handleSelect(county, saleType) {
    if (onSelectReport) {
      // Map UiPath/UiPathTD to foreclosure/taxdeed for App.jsx
      let mappedType = saleType;
      if (saleType === 'UiPath') mappedType = 'foreclosure';
      else if (saleType === 'UiPathTD') mappedType = 'taxdeed';
      onSelectReport(county, mappedType);
    }
  }

  return (
    <nav style={{ width: 280, background: '#f7f7f7', padding: '32px 16px 16px 16px', boxShadow: '2px 0 8px #eee', display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
      <h2 style={{
        fontSize: 20,
        marginBottom: 16,
        background: '#f7c873',
        color: '#7a5c1c',
        padding: '8px 12px',
        borderRadius: 6,
        boxShadow: '0 1px 4px #f7c87333',
        letterSpacing: 0.5,
        width: '100%',
        textAlign: 'left',
      }}>
        Florida Counties
      </h2>
      <ul style={{ listStyle: 'none', padding: 0, margin: 0, width: '100%' }}>
        {Object.keys(countySaleTypes).map((county) => (
          <li key={county} style={{ marginBottom: 8 }}>
            <button
              style={{ background: 'none', border: 'none', color: '#222', cursor: 'pointer', padding: 0, fontWeight: 600, fontSize: 17, marginLeft: 4 }}
              onClick={() => handleCountyClick(county)}
            >
              {county}
            </button>
            {expandedCounty === county && (
              <ul style={{ listStyle: 'none', paddingLeft: 16, margin: 0 }}>
                {countySaleTypes[county].map((saleType) => (
                  <li key={saleType} style={{ marginBottom: 4 }}>
                    <button
                      style={{
                        background: 'none',
                        border: 'none',
                        color: '#222',
                        cursor: 'pointer',
                        padding: '6px 12px',
                        fontSize: 16,
                        borderRadius: 4,
                        textAlign: 'left',
                        width: '100%',
                        transition: 'background 0.2s',
                      }}
                      onMouseOver={e => e.currentTarget.style.background = '#f0f0f0'}
                      onMouseOut={e => e.currentTarget.style.background = 'none'}
                      onClick={() => handleSelect(county, saleType)}
                    >
                      {SALE_TYPE_LABELS[saleType]}
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </li>
        ))}
      </ul>
    </nav>
  );
}
