
import React, { useEffect, useState } from "react";
import params from '../config/params';

function parseTableFromHTML(html) {
  // Create a DOM parser
  const parser = new window.DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const table = doc.querySelector("table");
  if (!table) return { headers: [], rows: [] };
  // Parse headers
  const headerCells = Array.from(table.querySelectorAll("thead th"));
  const headers = headerCells.map((th) => th.textContent.trim());
  // Parse rows
  const rows = Array.from(table.querySelectorAll("tbody tr")).map((tr) =>
    Array.from(tr.querySelectorAll("td")).map((td) => td.textContent.trim())
  );
  return { headers, rows };
}

const SalesReportPanel = ({ county = "pasco", saleType = "foreclosure", ownerAssocFilter = 'include' }) => {
  const [tableData, setTableData] = useState({ headers: [], rows: [] });
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState("asc");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(true);

  // Compose the report filename
  const reportFile = `/reports/sales_report_${county.toLowerCase().replace(/\s/g, "_")}_${saleType.toLowerCase().replace(/\s/g, "")}.html`;

  useEffect(() => {
    setLoading(true);
    setError(null);
    setTableData({ headers: [], rows: [] });
    fetch(reportFile)
      .then((resp) => {
        if (!resp.ok) throw new Error("Report not found");
        return resp.text();
      })
      .then((html) => {
        const { headers, rows } = parseTableFromHTML(html);
        setTableData({ headers, rows });
        setLoading(false);
      })
      .catch((err) => {
        setError(err.message);
        setLoading(false);
      });
  }, [county, saleType, reportFile]);

  function handleSort(colIdx) {
    if (sortCol === colIdx) {
      setSortDir(sortDir === "asc" ? "desc" : "asc");
    } else {
      setSortCol(colIdx);
      setSortDir("asc");
    }
  }

  function getFilteredAndSortedRows() {
    const { rows, headers } = tableData;
    if (!rows.length) return [];
    // Find Certificate Holder Name column index
    const certIdx = headers.findIndex(h => h.toLowerCase() === 'certificate holder name');
    const ownerWords = (params.owner_assoc_words || '').split(',').map(w => w.trim().toLowerCase()).filter(Boolean);
    let filtered = rows;
    if (certIdx !== -1 && ownerAssocFilter !== 'include') {
      filtered = rows.filter(row => {
        const val = (row[certIdx] || '').toLowerCase();
        const isOwnerAssoc = ownerWords.some(word => val.includes(word));
        if (ownerAssocFilter === 'only') return isOwnerAssoc;
        if (ownerAssocFilter === 'exclude') return !isOwnerAssoc;
        return true;
      });
    }
    if (sortCol == null) return filtered;
    const sorted = [...filtered].sort((a, b) => {
      const valA = a[sortCol] || "";
      const valB = b[sortCol] || "";
      // Check if the column is 'Sale Date' or 'Add Date' (case-insensitive)
      const header = (headers[sortCol] || '').toLowerCase();
      if (header === 'sale date' || header === 'add date') {
        // Try to parse as date (MM/DD/YYYY or YYYY-MM-DD)
        const dateA = Date.parse(valA);
        const dateB = Date.parse(valB);
        if (!isNaN(dateA) && !isNaN(dateB)) {
          return sortDir === 'asc' ? dateA - dateB : dateB - dateA;
        }
      }
      // Try numeric sort if both values are numbers
      const numA = parseFloat(valA.replace(/[^\d.\-]/g, ""));
      const numB = parseFloat(valB.replace(/[^\d.\-]/g, ""));
      if (!isNaN(numA) && !isNaN(numB)) {
        return sortDir === "asc" ? numA - numB : numB - numA;
      }
      // Otherwise, string sort
      return sortDir === "asc"
        ? valA.localeCompare(valB)
        : valB.localeCompare(valA);
    });
    return sorted;
  }

  if (error) {
    return <div style={{ color: "red" }}>Error: {error}</div>;
  }
  if (loading) {
    return <div>Loading report...</div>;
  }
  if (!tableData.headers.length) {
    return <div>No data found in report.</div>;
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'row', alignItems: 'flex-start' }}>
  {/* Sidebar Filters removed: now controlled by App sidebar */}
      {/* Main Table */}
      <div className="report-scroll-container" style={{ maxWidth: 1700, height: 1200, overflow: 'auto', border: '1px solid #ccc', borderRadius: 8, background: '#fff', flex: 1 }}>
        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
          <thead className="sticky-table-header">
            <tr>
              {tableData.headers.map((header, idx) => (
                <th
                  key={header}
                  onClick={() => handleSort(idx)}
                  style={{ cursor: 'pointer', background: sortCol === idx ? '#ffe9b3' : undefined }}
                  title="Click to sort"
                >
                  {header}
                  {sortCol === idx ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {getFilteredAndSortedRows().map((row, i) => (
              <tr key={i}>
                {row.map((cell, j) => (
                  <td key={j}>{cell}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default SalesReportPanel;
