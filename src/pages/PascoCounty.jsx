
import React from "react";

export default function PascoCounty({ selectedCounty, selectedSaleType }) {
  let reportSrc = "";
  if (selectedCounty && selectedSaleType) {
    reportSrc = `/reports/sales_report_${selectedCounty.toLowerCase().replace(/\s/g, "_")}_${selectedSaleType.toLowerCase().replace(/\s/g, "")}.html`;
  }
  return (
    <div style={{ width: '100%', minHeight: '100vh', background: '#fff', display: 'flex', alignItems: 'flex-start', justifyContent: 'center', padding: 0 }}>
      {reportSrc ? (
        <iframe
          src={reportSrc}
          title="County Sales Report"
          style={{ width: '100%', minHeight: 800, border: 'none', borderRadius: 0 }}
        />
      ) : (
        <div style={{ margin: 40, fontSize: 20, color: '#888' }}>No report selected.</div>
      )}
    </div>
  );
}
