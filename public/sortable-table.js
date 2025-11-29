// sortable-table.js
// Simple table sorting for static HTML tables
// Usage: just include this script in your HTML. All <th> in <thead> become sortable.


function sortTable(table, col, type, dir) {
  const tbody = table.tBodies[0];
  const rows = Array.from(tbody.rows);
  // Detect if this column is a date column by header
  const ths = table.querySelectorAll('thead th');
  const header = ths[col]?.textContent.trim().toLowerCase();
  const isDateCol = header === 'sale date' || header === 'add date';
  const compare = (a, b) => {
    let v1 = a.cells[col].getAttribute('data-sort') || a.cells[col].textContent.trim();
    let v2 = b.cells[col].getAttribute('data-sort') || b.cells[col].textContent.trim();
    if (isDateCol) {
      // Try to parse as date
      const d1 = Date.parse(v1);
      const d2 = Date.parse(v2);
      if (!isNaN(d1) && !isNaN(d2)) {
        return dir === 'asc' ? d1 - d2 : d2 - d1;
      }
    }
    if (type === 'number') {
      v1 = parseFloat(v1.replace(/[^\d.\-]/g, '')) || 0;
      v2 = parseFloat(v2.replace(/[^\d.\-]/g, '')) || 0;
    }
    return dir === 'asc' ? (v1 > v2 ? 1 : v1 < v2 ? -1 : 0) : (v1 < v2 ? 1 : v1 > v2 ? -1 : 0);
  };
  // Only sort main data rows (skip .ffl-notes-row)
  const allRows = Array.from(tbody.rows);
  const dataRows = allRows.filter(row => !row.classList.contains('ffl-notes-row'));
  dataRows.sort(compare);
  // Re-attach each data row and its following .ffl-notes-row (if present)
  dataRows.forEach(row => {
    tbody.appendChild(row);
    const next = row.nextElementSibling;
    if (next && next.classList.contains('ffl-notes-row')) {
      tbody.appendChild(next);
    }
  });
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

    // Insert checkboxes and Notes button in each row (if not already present)
    Array.from(table.tBodies[0].rows).forEach((row, idx) => {
      if (!row.cells[0].querySelector('input[type="checkbox"]')) {
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
      }
      // Add Notes button if not already present
      if (!row.querySelector('.ffl-notes-btn')) {
        const caseNum = row.cells[caseNumIdx + 1]?.textContent.trim();
        const notesBtn = document.createElement('button');
        notesBtn.textContent = 'Notes';
        notesBtn.className = 'ffl-notes-btn';
        notesBtn.style.marginLeft = '6px';
        notesBtn.style.fontSize = '0.95em';
        notesBtn.style.padding = '2px 8px';
        notesBtn.style.cursor = 'pointer';
        // Set tooltip to current note (if any)
  notesBtn.title = localStorage.getItem('ffl_note_' + caseNum) || 'Add note';
  // Style tooltip for better visibility
  notesBtn.style.setProperty('color', '#155724'); // dark green
        row.cells[0].appendChild(notesBtn);

        // Insert expandable notes row after this row if not already present
        let nextRow = row.nextElementSibling;
        if (!nextRow || !nextRow.classList.contains('ffl-notes-row')) {
          const notesRow = document.createElement('tr');
          notesRow.className = 'ffl-notes-row';
          notesRow.style.display = 'none';
          const notesTd = document.createElement('td');
          notesTd.colSpan = row.cells.length;
          notesTd.style.background = '#f9f9f9';
          notesTd.style.borderTop = '1px solid #ddd';
          notesTd.style.padding = '10px 16px';
          const notesArea = document.createElement('textarea');
          notesArea.rows = 3;
          notesArea.style.width = '98%';
          notesArea.style.fontSize = '1em';
          notesArea.placeholder = 'Notes about this property...';
          notesArea.value = localStorage.getItem('ffl_note_' + caseNum) || '';
          notesArea.addEventListener('input', function() {
            localStorage.setItem('ffl_note_' + caseNum, notesArea.value);
            // Update tooltip on Notes button
            notesBtn.title = notesArea.value || 'Add note';
          });
          notesTd.appendChild(notesArea);
          notesRow.appendChild(notesTd);
          row.parentNode.insertBefore(notesRow, row.nextSibling);

          // Toggle notes row on button click
          notesBtn.addEventListener('click', function() {
            notesRow.style.display = notesRow.style.display === 'none' ? '' : 'none';
          });
        }
        // Update tooltip on mouseenter (in case note changed elsewhere)
        notesBtn.addEventListener('mouseenter', function() {
          notesBtn.title = localStorage.getItem('ffl_note_' + caseNum) || 'Add note';
        });
      }

      // Add Value Estimate button if not already present
      if (!row.querySelector('.ffl-value-btn')) {
        const caseNum = row.cells[caseNumIdx + 1]?.textContent.trim();
        const valueBtn = document.createElement('button');
        valueBtn.textContent = 'Value Estimate';
        valueBtn.className = 'ffl-value-btn';
        valueBtn.style.marginLeft = '6px';
        valueBtn.style.fontSize = '0.95em';
        valueBtn.style.padding = '2px 8px';
        valueBtn.style.cursor = 'pointer';
        // Set tooltip to current estimate (if any)
        function getEstimateTooltip() {
          const est = localStorage.getItem('ffl_est_' + caseNum);
          const fj = getFinalJudgment(row);
          if (est && fj !== null) {
            const diff = est - fj;
            return `Estimate: $${Number(est).toLocaleString()}\nFinal Judgment: $${fj.toLocaleString()}\nDifference: $${diff.toLocaleString()}`;
          } else if (est) {
            return `Estimate: $${Number(est).toLocaleString()}`;
          } else {
            return 'Add value estimate';
          }
        }
  valueBtn.title = getEstimateTooltip();
  // Style tooltip for better visibility
  valueBtn.style.setProperty('color', '#155724'); // dark green
        row.cells[0].appendChild(valueBtn);

        // Helper to get Final Judgment value from row
        function getFinalJudgment(row) {
          // Find Final Judgment column index
          let fjIdx = -1;
          Array.from(row.parentNode.parentNode.querySelectorAll('thead th')).forEach((th, idx) => {
            if (th.textContent.trim().toLowerCase() === 'final judgment') fjIdx = idx;
          });
          if (fjIdx === -1) return null;
          const val = row.cells[fjIdx]?.textContent.replace(/[^\d.\-]/g, '');
          return val ? parseFloat(val) : null;
        }

        // Popup for entering estimate
        valueBtn.addEventListener('click', function() {
          // Remove any existing modal
          const oldModal = document.getElementById('value-estimate-modal');
          if (oldModal) oldModal.remove();

          const modal = document.createElement('div');
          modal.id = 'value-estimate-modal';
          modal.style.position = 'fixed';
          modal.style.top = '50%';
          modal.style.left = '50%';
          modal.style.transform = 'translate(-50%, -50%)';
          modal.style.background = '#fff';
          modal.style.border = '1px solid #ccc';
          modal.style.boxShadow = '0 2px 12px rgba(0,0,0,0.2)';
          modal.style.zIndex = 10000;
          modal.style.padding = '20px';
          modal.style.minWidth = '320px';
          modal.style.borderRadius = '8px';

          const label = document.createElement('div');
          label.textContent = `Value Estimate for Case: ${caseNum}`;
          label.style.marginBottom = '8px';
          modal.appendChild(label);

          const input = document.createElement('input');
          input.type = 'number';
          input.placeholder = 'Enter your estimate ($)';
          input.style.width = '100%';
          input.style.fontSize = '1.1em';
          input.style.marginBottom = '12px';
          input.value = localStorage.getItem('ffl_est_' + caseNum) || '';
          modal.appendChild(input);

          // Show Final Judgment and difference
          const fj = getFinalJudgment(row);
          const fjDiv = document.createElement('div');
          if (fj !== null) {
            fjDiv.textContent = `Final Judgment: $${fj.toLocaleString()}`;
            fjDiv.style.marginBottom = '8px';
            modal.appendChild(fjDiv);
          }

          const diffDiv = document.createElement('div');
          diffDiv.style.marginBottom = '8px';
          modal.appendChild(diffDiv);

          function updateDiff() {
            const est = parseFloat(input.value);
            if (!isNaN(est) && fj !== null) {
              const diff = est - fj;
              diffDiv.textContent = `Difference: $${diff.toLocaleString()}`;
            } else {
              diffDiv.textContent = '';
            }
          }
          input.addEventListener('input', updateDiff);
          updateDiff();

          const btnRow = document.createElement('div');
          btnRow.style.textAlign = 'right';

          const saveBtn = document.createElement('button');
          saveBtn.textContent = 'Save';
          saveBtn.style.marginRight = '8px';
          saveBtn.onclick = () => {
            if (input.value) {
              localStorage.setItem('ffl_est_' + caseNum, input.value);
            } else {
              localStorage.removeItem('ffl_est_' + caseNum);
            }
            valueBtn.title = getEstimateTooltip();
            modal.remove();
          };
          btnRow.appendChild(saveBtn);

          const cancelBtn = document.createElement('button');
          cancelBtn.textContent = 'Cancel';
          cancelBtn.onclick = () => modal.remove();
          btnRow.appendChild(cancelBtn);

          modal.appendChild(btnRow);

          document.body.appendChild(modal);
          input.focus();
        });

        // Update tooltip on mouseenter (in case value changed elsewhere)
        valueBtn.addEventListener('mouseenter', function() {
          valueBtn.title = getEstimateTooltip();
        });
      }
    });

    // Remove notes column header if present
    const notesTh = table.querySelector('thead th.ffl-notes-header');
    if (notesTh) notesTh.remove();
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

    // Filtering removed: now handled by React sidebar
  }

  document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('table').forEach(makeTableSortableAndFilterable);
  });
