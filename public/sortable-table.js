// sortable-table.js
// Simple table sorting for static HTML tables
// Usage: just include this script in your HTML. All <th> in <thead> become sortable.


(function() {
  function sortTable(table, col, type, dir) {
    const tbody = table.tBodies[0];
    const rows = Array.from(tbody.rows);
    const compare = (a, b) => {
      let v1 = a.cells[col].textContent.trim();
      let v2 = b.cells[col].textContent.trim();
      if (type === 'number') {
        v1 = parseFloat(v1.replace(/[^\d.\-]/g, '')) || 0;
        v2 = parseFloat(v2.replace(/[^\d.\-]/g, '')) || 0;
      }
      return dir === 'asc' ? (v1 > v2 ? 1 : v1 < v2 ? -1 : 0) : (v1 < v2 ? 1 : v1 > v2 ? -1 : 0);
    };
    rows.sort(compare);
    rows.forEach(row => tbody.appendChild(row));
  }

  function detectType(val) {
    return /^\d+[.,\d]*$/.test(val.replace(/[^\d.\-]/g, '')) ? 'number' : 'string';
  }

  function makeTableSortableAndFilterable(table) {
    // Sorting
    const ths = table.querySelectorAll('thead th');
    ths.forEach((th, idx) => {
      let dir = 'asc';
      th.style.cursor = 'pointer';
      th.addEventListener('click', function() {
        const type = detectType(table.tBodies[0].rows[0]?.cells[idx]?.textContent || '');
        sortTable(table, idx, type, dir);
        ths.forEach(t => t.classList.remove('sorted-asc', 'sorted-desc'));
        th.classList.add(dir === 'asc' ? 'sorted-asc' : 'sorted-desc');
        dir = dir === 'asc' ? 'desc' : 'asc';
      });
    });

    // Filtering
    // Find the Parcel ID column index
    let parcelIdIdx = -1;
    ths.forEach((th, idx) => {
      if (th.textContent.trim().toLowerCase() === 'parcel id') {
        parcelIdIdx = idx;
      }
    });
    if (parcelIdIdx === -1) return; // No Parcel ID column

    // Create filter controls
    const filterDiv = document.createElement('div');
    filterDiv.style.margin = '12px 0 16px 0';
    filterDiv.style.display = 'flex';
    filterDiv.style.gap = '24px';
    filterDiv.style.alignItems = 'center';

    const timeshareBox = document.createElement('input');
    timeshareBox.type = 'checkbox';
    timeshareBox.id = 'filter-timeshare';
    const timeshareLabel = document.createElement('label');
    timeshareLabel.htmlFor = 'filter-timeshare';
    timeshareLabel.textContent = 'Hide Timeshare Parcel IDs';

    const blankBox = document.createElement('input');
    blankBox.type = 'checkbox';
    blankBox.id = 'filter-blank';
    const blankLabel = document.createElement('label');
    blankLabel.htmlFor = 'filter-blank';
    blankLabel.textContent = 'Hide Blank Parcel IDs';

    filterDiv.appendChild(timeshareBox);
    filterDiv.appendChild(timeshareLabel);
    filterDiv.appendChild(blankBox);
    filterDiv.appendChild(blankLabel);

    // Insert filterDiv before the table
    table.parentNode.insertBefore(filterDiv, table);

    function applyFilters() {
      const hideTimeshare = timeshareBox.checked;
      const hideBlank = blankBox.checked;
      Array.from(table.tBodies[0].rows).forEach(row => {
        const val = (row.cells[parcelIdIdx]?.textContent || '').trim();
        let hide = false;
        if (hideTimeshare && val.toUpperCase() === 'TIMESHARE') hide = true;
        if (hideBlank && val === '') hide = true;
        row.style.display = hide ? 'none' : '';
      });
    }
    timeshareBox.addEventListener('change', applyFilters);
    blankBox.addEventListener('change', applyFilters);
  }

  document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('table').forEach(makeTableSortableAndFilterable);
  });
})();
