const fileInput = document.getElementById('file-input');
const sheetSelect = document.getElementById('sheet-select');
const headerSelect = document.getElementById('header-select');
const commonFilter = document.getElementById('common-filter');
const filterValue = document.getElementById('filter-value');
const applyFilter = document.getElementById('apply-filter');
const resetFilter = document.getElementById('reset-filter');
const downloadFilter = document.getElementById('download-filter');
const tableHead = document.querySelector('#data-table thead');
const tableBody = document.querySelector('#data-table tbody');
const summaryContainer = document.getElementById('summary-container');
const statusMessage = document.getElementById('status-message');
const tableInfo = document.getElementById('table-info');

let workbook = null;
let currentRows = [];
let currentHeaders = [];
let currentSheetName = '';
let displayedRows = [];

function setEnabled(enabled) {
  sheetSelect.disabled = !enabled;
  headerSelect.disabled = !enabled;
  commonFilter.disabled = !enabled;
  filterValue.disabled = !enabled;
  applyFilter.disabled = !enabled;
  resetFilter.disabled = !enabled;
  downloadFilter.disabled = !enabled;
}

function clearTable() {
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';
}

function renderTable(rows) {
  displayedRows = rows;
  clearTable();

  if (!rows.length) {
    tableInfo.textContent = 'No matching rows to display.';
    renderSummary([]);
    return;
  }

  const headerRow = document.createElement('tr');
  currentHeaders.forEach((header) => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });
  tableHead.appendChild(headerRow);

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    currentHeaders.forEach((header) => {
      const td = document.createElement('td');
      const value = row[header];
      td.textContent = value !== undefined && value !== null ? String(value) : '';
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });

  tableInfo.textContent = `${rows.length} row(s) displayed from "${currentSheetName}" sheet.`;
  renderSummary(rows);
}

function updateHeaderOptions(headers) {
  headerSelect.innerHTML = '<option value="">Choose a field</option>';
  headers.forEach((header) => {
    const option = document.createElement('option');
    option.value = header;
    option.textContent = header;
    headerSelect.appendChild(option);
  });
}

function normalizeHeader(name) {
  return String(name || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function findHeader(names) {
  const normalizedNames = names.map(normalizeHeader);
  return currentHeaders.find((header) => normalizedNames.includes(normalizeHeader(header)));
}

function renderSummary(rows) {
  if (!summaryContainer) return;
  summaryContainer.innerHTML = '';

  if (!rows.length) {
    summaryContainer.innerHTML = '<p class="summary-empty">No summary available for the current selection.</p>';
    return;
  }

  const contactField = findHeader(['Contact Person', 'Contact_Person', 'Contact', 'contactperson']);
  const uomField = findHeader(['UOM', 'Uom', 'uom']);
  const quantityField = findHeader(['Quantity', 'Qty', 'quantity']);
  const invoiceField = findHeader(['Invoice No', 'InvoiceNo', 'Invoice Number', 'Invoice']);
  const monthField = findHeader(['Month', 'month']);
  const yearField = findHeader(['Year', 'year']);

  if (!uomField) {
    summaryContainer.innerHTML = '<p class="summary-empty">UOM column not found in the selected sheet.</p>';
    return;
  }

  const groups = rows.reduce((acc, row) => {
    const uom = String(row[uomField] ?? 'Unknown').trim() || 'Unknown';
    acc[uom] = acc[uom] || { rows: [] };
    acc[uom].rows.push(row);
    return acc;
  }, {});

  Object.entries(groups).forEach(([uom, group]) => {
    const contactPersons = Array.from(new Set(group.rows.map((row) => String(row[contactField] ?? '').trim()).filter(Boolean)));
    const totalQuantity = group.rows.reduce((sum, row) => {
      const value = row[quantityField];
      const number = Number(String(value).replace(/,/g, ''));
      return sum + (Number.isFinite(number) ? number : 0);
    }, 0);

    const billedMonths = new Set();
    const notBilledMonths = new Set();

    group.rows.forEach((row) => {
      const month = monthField ? String(row[monthField] ?? '').trim() : '';
      const year = yearField ? String(row[yearField] ?? '').trim() : '';
      const monthYear = [month, year].filter(Boolean).join(' ').trim() || 'Unknown month/year';
      const invoice = String(row[invoiceField] ?? '').trim();
      if (invoice) {
        billedMonths.add(monthYear);
      } else {
        notBilledMonths.add(monthYear);
      }
    });

    const billedText = billedMonths.size ? Array.from(billedMonths).join(', ') : 'None';
    const notBilledText = notBilledMonths.size ? Array.from(notBilledMonths).join(', ') : 'None';

    const card = document.createElement('div');
    card.className = 'summary-card';
    card.innerHTML = `
      <div class="summary-card-content">
        <p><span>Contact Person:</span> ${contactPersons.length ? contactPersons.join(', ') : 'Unknown'}</p>
        <p><span>UOM:</span> ${uom}</p>
        <p><span>Quantity:</span> ${totalQuantity}</p>
        <p><span>How Many Month billed:</span> ${billedMonths.size}</p>
      </div>
      <div class="summary-card-right">
        <p><span>Billed Months:</span> ${billedText}</p>
        <p><span>Not billed Months:</span> ${notBilledText}</p>
      </div>
    `;
    summaryContainer.appendChild(card);
  });
}

function showStatus(message, isError = false) {
  if (!statusMessage) return;
  statusMessage.textContent = message;
  statusMessage.style.color = isError ? '#b91c1c' : '#334155';
}

function arrayBufferToBinary(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const slice = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, slice);
  }
  return binary;
}

function showSheet(sheetName) {
  currentSheetName = sheetName;
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
  currentRows = rows;
  currentHeaders = rows.length ? Object.keys(rows[0]) : [];

  updateHeaderOptions(currentHeaders);
  renderTable(currentRows);
}

function applyFilterAction() {
  const field = headerSelect.value;
  const query = filterValue.value.trim();
  const commonField = commonFilter.value;

  if (!field && !commonField) {
    renderTable(currentRows);
    return;
  }

  const activeField = commonField || field;
  const filteredRows = currentRows.filter((row) => {
    const value = row[activeField];
    if (value === undefined || value === null) {
      return false;
    }
    const cellText = String(value).trim();
    if (!query) {
      return true;
    }
    return cellText.toLowerCase().includes(query.toLowerCase());
  });
  renderTable(filteredRows);
}

function resetFilterAction() {
  headerSelect.selectedIndex = 0;
  commonFilter.selectedIndex = 0;
  filterValue.value = '';
  renderTable(currentRows);
}

function downloadFilteredSheet() {
  if (!currentSheetName) {
    showStatus('Select a sheet before downloading.', true);
    return;
  }

  if (!displayedRows.length) {
    showStatus('There is no data to download for the selected filter.', true);
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(displayedRows, { header: currentHeaders });
  const workbookToExport = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbookToExport, worksheet, currentSheetName);
  const safeName = currentSheetName.replace(/[^a-z0-9_\- ]/gi, '_').trim() || 'Sheet1';
  XLSX.writeFile(workbookToExport, `filtered-${safeName}.xlsx`);
}

fileInput.addEventListener('change', (event) => {
  const file = event.target.files[0];
  if (!file) {
    showStatus('No file selected.', true);
    return;
  }

  if (typeof XLSX === 'undefined') {
    showStatus('Excel library not loaded. Check your internet connection or script path.', true);
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const result = e.target.result;
      let readOptions = { type: 'array' };
      let workbookData = result;

      if (typeof result === 'string') {
        workbookData = result;
        readOptions = { type: 'binary' };
      }

      workbook = XLSX.read(workbookData, readOptions);
      sheetSelect.innerHTML = '<option value="">Choose sheet</option>';
      workbook.SheetNames.forEach((name) => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        sheetSelect.appendChild(option);
      });
      setEnabled(true);
      clearTable();
      renderSummary([]);
      tableInfo.textContent = 'Select a sheet to show data.';
      showStatus(`Loaded file: ${file.name}`);
    } catch (error) {
      console.error('Excel parse error:', error);
      showStatus('Unable to read this file. Try a valid .xlsx or .xls workbook.', true);
      setEnabled(false);
    }
  };

  reader.onerror = () => {
    console.error('File read error', reader.error);
    showStatus('Failed to read the file. Please try again.', true);
    setEnabled(false);
  };

  if (reader.readAsArrayBuffer) {
    reader.readAsArrayBuffer(file);
  } else if (reader.readAsBinaryString) {
    reader.readAsBinaryString(file);
  } else {
    showStatus('This browser does not support Excel file reading.', true);
  }
});

sheetSelect.addEventListener('change', (event) => {
  const sheetName = event.target.value;
  if (!sheetName) {
    clearTable();
    tableInfo.textContent = 'Select a sheet to show data.';
    return;
  }
  showSheet(sheetName);
});

commonFilter.addEventListener('change', (event) => {
  if (event.target.value) {
    const matchIndex = Array.from(headerSelect.options).findIndex((option) => option.value === event.target.value);
    if (matchIndex >= 0) {
      headerSelect.selectedIndex = matchIndex;
    }
  }
});

applyFilter.addEventListener('click', applyFilterAction);
resetFilter.addEventListener('click', resetFilterAction);
downloadFilter.addEventListener('click', downloadFilteredSheet);

setEnabled(false);
