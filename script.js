const fileInput = document.getElementById('file-input');
const dropZone = document.getElementById('drop-zone');
const sheetSelect = document.getElementById('sheet-select');
const headerSelect = document.getElementById('header-select');
const commonFilter = document.getElementById('common-filter');
const filterValue = document.getElementById('filter-value');
const filterValueList = document.getElementById('filter-value-list');
const applyFilter = document.getElementById('apply-filter');
const resetFilter = document.getElementById('reset-filter');
const downloadFilter = document.getElementById('download-filter');
const tableHead = document.querySelector('#data-table thead');
const tableBody = document.querySelector('#data-table tbody');
const summaryContainer = document.getElementById('summary-container');
const statusMessage = document.getElementById('status-message');
const tableInfo = document.getElementById('table-info');
const tableWrapper = document.getElementById('table-wrapper');
const scrollLeftBtn = document.getElementById('scroll-left');
const scrollRightBtn = document.getElementById('scroll-right');
const themeToggle = document.getElementById('theme-toggle');

let workbook = null;
let currentRows = [];
let currentHeaders = [];
let currentSheetName = '';
let displayedRows = [];
let currentSummaryRows = [];
let isDraggingTable = false;
let tableDragStartX = 0;
let tableScrollStartLeft = 0;

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
  updateScrollButtons();
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
  requestAnimationFrame(updateScrollButtons);
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

function getActiveFilterField() {
  return commonFilter.value || headerSelect.value;
}

function updateFilterValueOptions() {
  const activeField = getActiveFilterField();
  filterValueList.innerHTML = '';

  if (!activeField || !currentRows.length) {
    return;
  }

  const uniqueValues = Array.from(
    new Set(currentRows.map((row) => String(row[activeField] ?? '').trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

  uniqueValues.slice(0, 5000).forEach((value) => {
    const option = document.createElement('option');
    option.value = value;
    filterValueList.appendChild(option);
  });
}

function buildSummaryData(rows) {
  const contactField = findHeader(['Contact Person', 'Contact_Person', 'Contact', 'contactperson']);
  const uomField = findHeader(['UOM', 'Uom', 'uom']);
  const quantityField = findHeader(['Quantity', 'Qty', 'quantity']);
  const invoiceField = findHeader(['Invoice No', 'InvoiceNo', 'Invoice Number', 'Invoice']);
  const monthField = findHeader(['Month', 'month']);
  const yearField = findHeader(['Year', 'year']);

  if (!rows.length || !uomField) {
    return [];
  }

  const groups = rows.reduce((acc, row) => {
    const uom = String(row[uomField] ?? 'Unknown').trim() || 'Unknown';
    acc[uom] = acc[uom] || { rows: [] };
    acc[uom].rows.push(row);
    return acc;
  }, {});

  return Object.entries(groups).map(([uom, group]) => {
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

    return {
      'Contact Person': contactPersons.length ? contactPersons.join(', ') : 'Unknown',
      UOM: uom,
      Quantity: totalQuantity,
      'How Many Month billed': billedMonths.size,
      'Billed Months': billedMonths.size ? Array.from(billedMonths).join(', ') : 'None',
      'Not billed Months': notBilledMonths.size ? Array.from(notBilledMonths).join(', ') : 'None'
    };
  });
}

function renderSummary(rows) {
  if (!summaryContainer) return;
  summaryContainer.innerHTML = '';
  currentSummaryRows = buildSummaryData(rows);

  if (!rows.length) {
    summaryContainer.innerHTML = '<p class="summary-empty">No summary available for the current selection.</p>';
    return;
  }

  if (!currentSummaryRows.length) {
    summaryContainer.innerHTML = '<p class="summary-empty">UOM column not found in the selected sheet.</p>';
    return;
  }

  currentSummaryRows.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'summary-card';
    card.innerHTML = `
      <div class="summary-card-content">
        <p><span>Contact Person:</span> ${item['Contact Person']}</p>
        <p><span>UOM:</span> ${item.UOM}</p>
        <p><span>Quantity:</span> ${item.Quantity}</p>
        <p><span>How Many Month billed:</span> ${item['How Many Month billed']}</p>
      </div>
      <div class="summary-card-right">
        <p><span>Billed Months:</span> ${item['Billed Months']}</p>
        <p><span>Not billed Months:</span> ${item['Not billed Months']}</p>
      </div>
    `;
    summaryContainer.appendChild(card);
  });
}

function showStatus(message, isError = false) {
  if (!statusMessage) return;
  statusMessage.textContent = message;
  statusMessage.style.color = isError ? 'var(--danger)' : 'var(--muted)';
}

function showSheet(sheetName) {
  currentSheetName = sheetName;
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
  currentRows = rows;
  currentHeaders = rows.length ? Object.keys(rows[0]) : [];

  updateHeaderOptions(currentHeaders);
  updateFilterValueOptions();
  renderTable(currentRows);
}

function applyFilterAction() {
  const activeField = getActiveFilterField();
  const query = filterValue.value.trim();

  if (!activeField) {
    renderTable(currentRows);
    return;
  }

  const filteredRows = currentRows.filter((row) => {
    const value = row[activeField];
    if (value === undefined || value === null) return false;
    const cellText = String(value).trim();
    if (!query) return true;
    return cellText.toLowerCase() === query.toLowerCase();
  });

  renderTable(filteredRows);
}

function resetFilterAction() {
  headerSelect.selectedIndex = 0;
  commonFilter.selectedIndex = 0;
  filterValue.value = '';
  updateFilterValueOptions();
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

  const workbookToExport = XLSX.utils.book_new();
  const dataWorksheet = XLSX.utils.json_to_sheet(displayedRows, { header: currentHeaders });
  XLSX.utils.book_append_sheet(workbookToExport, dataWorksheet, 'Filtered Data');

  const summaryRows = currentSummaryRows.length ? currentSummaryRows : buildSummaryData(displayedRows);
  const summaryWorksheet = XLSX.utils.json_to_sheet(summaryRows.length ? summaryRows : [{ Summary: 'No summary available for the current selection.' }]);
  XLSX.utils.book_append_sheet(workbookToExport, summaryWorksheet, 'Summary');

  const safeName = currentSheetName.replace(/[^a-z0-9_\- ]/gi, '_').trim() || 'Sheet1';
  XLSX.writeFile(workbookToExport, `filtered-${safeName}-with-summary.xlsx`);
}

function processExcelFile(file) {
  if (!file) {
    showStatus('No file selected.', true);
    return;
  }

  const validExtension = /\.(xlsx|xls)$/i.test(file.name);
  if (!validExtension) {
    showStatus('Please upload only .xlsx or .xls file.', true);
    return;
  }

  if (typeof XLSX === 'undefined') {
    showStatus('Excel library not loaded. Check the script path.', true);
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      workbook = XLSX.read(e.target.result, { type: 'array', cellStyles: true, cellNF: true, cellDates: true });
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
      filterValue.value = '';
      filterValueList.innerHTML = '';
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

  reader.readAsArrayBuffer(file);
}

function updateScrollButtons() {
  if (!tableWrapper) return;
  const canScroll = tableWrapper.scrollWidth > tableWrapper.clientWidth + 2;
  scrollLeftBtn.disabled = !canScroll || tableWrapper.scrollLeft <= 0;
  scrollRightBtn.disabled = !canScroll || tableWrapper.scrollLeft + tableWrapper.clientWidth >= tableWrapper.scrollWidth - 2;
}

fileInput.addEventListener('change', (event) => processExcelFile(event.target.files[0]));

['dragenter', 'dragover'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    event.stopPropagation();
    dropZone.classList.add('drag-over');
  });
});

['dragleave', 'drop'].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    event.stopPropagation();
    dropZone.classList.remove('drag-over');
  });
});

dropZone.addEventListener('drop', (event) => {
  const file = event.dataTransfer.files && event.dataTransfer.files[0];
  processExcelFile(file);
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
    if (matchIndex >= 0) headerSelect.selectedIndex = matchIndex;
  }
  filterValue.value = '';
  updateFilterValueOptions();
});

headerSelect.addEventListener('change', () => {
  if (headerSelect.value) commonFilter.selectedIndex = 0;
  filterValue.value = '';
  updateFilterValueOptions();
});

filterValue.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    applyFilterAction();
  }
});

applyFilter.addEventListener('click', applyFilterAction);
resetFilter.addEventListener('click', resetFilterAction);
downloadFilter.addEventListener('click', downloadFilteredSheet);

scrollLeftBtn.addEventListener('click', () => {
  tableWrapper.scrollBy({ left: -Math.max(260, tableWrapper.clientWidth * 0.65), behavior: 'smooth' });
});

scrollRightBtn.addEventListener('click', () => {
  tableWrapper.scrollBy({ left: Math.max(260, tableWrapper.clientWidth * 0.65), behavior: 'smooth' });
});

tableWrapper.addEventListener('scroll', updateScrollButtons);
window.addEventListener('resize', updateScrollButtons);

tableWrapper.addEventListener('mousedown', (event) => {
  if (event.button !== 0) return;
  isDraggingTable = true;
  tableWrapper.classList.add('dragging');
  tableDragStartX = event.pageX;
  tableScrollStartLeft = tableWrapper.scrollLeft;
});

window.addEventListener('mousemove', (event) => {
  if (!isDraggingTable) return;
  event.preventDefault();
  const distance = event.pageX - tableDragStartX;
  tableWrapper.scrollLeft = tableScrollStartLeft - distance;
});

window.addEventListener('mouseup', () => {
  isDraggingTable = false;
  tableWrapper.classList.remove('dragging');
});

themeToggle.addEventListener('change', () => {
  document.body.classList.toggle('dark-mode', themeToggle.checked);
  localStorage.setItem('excelViewerTheme', themeToggle.checked ? 'dark' : 'light');
});

if (localStorage.getItem('excelViewerTheme') === 'dark') {
  themeToggle.checked = true;
  document.body.classList.add('dark-mode');
}

setEnabled(false);
updateScrollButtons();

