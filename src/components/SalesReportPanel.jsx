import React, { useEffect, useState } from "react";

const SalesReportPanel = ({ county = "pasco" }) => {
  const [html, setHtml] = useState("");
  const [error, setError] = useState(null);

  useEffect(() => {
    // Try to fetch the latest report for the county
    const fetchReport = async () => {
      try {
        // You may want to dynamically list files, but for now, use a convention
        const resp = await fetch(`/reports/sales_report_${county}.html`);
        if (!resp.ok) throw new Error("Report not found");
        const text = await resp.text();
        setHtml(text);
      } catch (err) {
        setError(err.message);
      }
    };
    fetchReport();
  }, [county]);

  if (error) {
    return <div style={{ color: "red" }}>Error: {error}</div>;
  }
  if (!html) {
    return <div>Loading report...</div>;
  }
  return (
    <div style={{ width: "100%", height: "100%", overflow: "auto" }}>
      <div dangerouslySetInnerHTML={{ __html: html }} />
    </div>
  );
};

export default SalesReportPanel;
