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
    // Add persistent checkboxes to the far left of each row
    const ths = table.querySelectorAll('thead th');
    // Find Case Number column index (for unique key)
    let caseNumIdx = -1;
    ths.forEach((th, idx) => {
      if (th.textContent.trim().toLowerCase() === 'case number') {
        caseNumIdx = idx;
      }
    });
    if (caseNumIdx === -1) return;

    // Insert checkboxes in each row (if not already present)
    Array.from(table.tBodies[0].rows).forEach(row => {
      if (row.cells[0].querySelector('input[type="checkbox"]')) return;
      // Case Number is now shifted by 1 due to new checkbox column
      const caseNum = row.cells[caseNumIdx + 1]?.textContent.trim();
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.className = 'favorite-checkbox';
      cb.title = 'Mark as favorite';
      // Restore checked state from localStorage
      const checked = localStorage.getItem('ffl_fav_' + caseNum) === '1';
      cb.checked = checked;
      cb.addEventListener('change', function() {
        if (cb.checked) {
          localStorage.setItem('ffl_fav_' + caseNum, '1');
        } else {
          localStorage.removeItem('ffl_fav_' + caseNum);
        }
      });
      row.cells[0].appendChild(cb);
    });

    // Sorting
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
    // Find the Parcel ID and Status column indices
    let parcelIdIdx = -1;
    let statusIdx = -1;
    ths.forEach((th, idx) => {
      if (th.textContent.trim().toLowerCase() === 'parcel id') {
        parcelIdIdx = idx;
      }
      if (th.textContent.trim().toLowerCase() === 'status') {
        statusIdx = idx;
      }
    });
    if (parcelIdIdx === -1) return; // No Parcel ID column
    if (statusIdx === -1) return; // No Status column

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

    // Status dropdown (multi-select)
    const statusSelect = document.createElement('select');
    statusSelect.multiple = true;
    statusSelect.size = 1;
    statusSelect.style.minWidth = '160px';
    statusSelect.style.maxWidth = '220px';
    statusSelect.style.fontSize = '1em';
    statusSelect.title = 'Filter by Status (hold Ctrl/Cmd to select multiple)';
    statusSelect.id = 'filter-status';
    const statusLabel = document.createElement('label');
    statusLabel.htmlFor = 'filter-status';
    statusLabel.textContent = 'Filter Status:';

    // Get all unique status values from the table
    const statusSet = new Set();
    Array.from(table.tBodies[0].rows).forEach(row => {
      const val = (row.cells[statusIdx]?.textContent || '').trim();
      if (val) statusSet.add(val);
    });
    Array.from(statusSet).sort().forEach(val => {
      const opt = document.createElement('option');
      opt.value = val;
      opt.textContent = val;
      statusSelect.appendChild(opt);
    });

    filterDiv.appendChild(timeshareBox);
    filterDiv.appendChild(timeshareLabel);
    filterDiv.appendChild(blankBox);
    filterDiv.appendChild(blankLabel);
    filterDiv.appendChild(statusLabel);
    filterDiv.appendChild(statusSelect);

    // Insert filterDiv before the table
    table.parentNode.insertBefore(filterDiv, table);

    function applyFilters() {
      const hideTimeshare = timeshareBox.checked;
      const hideBlank = blankBox.checked;
      // Status filter
      const selectedStatuses = Array.from(statusSelect.selectedOptions).map(opt => opt.value);
      const statusFilterActive = selectedStatuses.length > 0;
      Array.from(table.tBodies[0].rows).forEach(row => {
        const parcelVal = (row.cells[parcelIdIdx]?.textContent || '').trim();
        const statusVal = (row.cells[statusIdx]?.textContent || '').trim();
        let hide = false;
        if (hideTimeshare && parcelVal.toUpperCase() === 'TIMESHARE') hide = true;
        if (hideBlank && parcelVal === '') hide = true;
        if (statusFilterActive && !selectedStatuses.includes(statusVal)) hide = true;
        row.style.display = hide ? 'none' : '';
      });
    }
    timeshareBox.addEventListener('change', applyFilters);
    blankBox.addEventListener('change', applyFilters);
    statusSelect.addEventListener('change', applyFilters);
  }

  document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('table').forEach(makeTableSortableAndFilterable);
  });
})();
