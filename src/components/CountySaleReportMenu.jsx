import React, { useEffect, useState } from "react";
import { getFirestore, collection, getDocs } from "firebase/firestore";
import { initializeApp, getApps } from "firebase/app";

// TODO: Replace with your actual Firebase config
const firebaseConfig = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_AUTH_DOMAIN",
  projectId: "foreclosure-15f09",
  storageBucket: "YOUR_STORAGE_BUCKET",
  messagingSenderId: "YOUR_MESSAGING_SENDER_ID",
  appId: "YOUR_APP_ID"
};

const app = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
const db = getFirestore(app);

const SALE_TYPE_LABELS = {
  UiPath: "Foreclosure",
  UiPathTD: "Tax Deed"
};

export default function CountySaleReportMenu() {
  const [params, setParams] = useState([]);
  const [selectedCounty, setSelectedCounty] = useState("");
  const [selectedSaleType, setSelectedSaleType] = useState("");

  useEffect(() => {
    fetchParams();
  }, []);

  async function fetchParams() {
    const querySnapshot = await getDocs(collection(db, "auction_parameters"));
    const data = querySnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
    setParams(data);
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

  function handleCountySelect(e) {
    setSelectedCounty(e.target.value);
    setSelectedSaleType("");
  }

  function handleSaleTypeSelect(e) {
    setSelectedSaleType(e.target.value);
  }

  // Build report file name
  let reportSrc = "";
  if (selectedCounty && selectedSaleType) {
    reportSrc = `/reports/sales_report_${selectedCounty.toLowerCase().replace(/\s/g, "_")}_${SALE_TYPE_LABELS[selectedSaleType].toLowerCase().replace(/\s/g, "")}.html`;
  }

  return (
    <div>
      <h2>Select County and Sale Type</h2>
      <div style={{ display: "flex", gap: 16, alignItems: "center" }}>
        <select value={selectedCounty} onChange={handleCountySelect}>
          <option value="">Select County</option>
          {Object.keys(countySaleTypes).map((county) => (
            <option key={county} value={county}>
              {county}
            </option>
          ))}
        </select>
        {selectedCounty && (
          <select value={selectedSaleType} onChange={handleSaleTypeSelect}>
            <option value="">Select Sale Type</option>
            {countySaleTypes[selectedCounty].map((saleType) => (
              <option key={saleType} value={saleType}>
                {SALE_TYPE_LABELS[saleType]}
              </option>
            ))}
          </select>
        )}
      </div>
      {reportSrc && (
        <div style={{ marginTop: 32 }}>
          <h3>
            {selectedCounty} County {SALE_TYPE_LABELS[selectedSaleType]} Report
          </h3>
          <iframe
            src={reportSrc}
            title="County Sale Report"
            style={{ width: "100%", minHeight: 600, border: "1px solid #ccc", borderRadius: 8 }}
          />
        </div>
      )}
    </div>
  );
}
