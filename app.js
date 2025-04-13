// Use global ExcelJS from CDN
const ExcelJS = window.ExcelJS;
const jsPDF = window.jspdf.jsPDF; // Added for PDF support

// State
let headers = JSON.parse(sessionStorage.getItem('headers')) || [];
let rows = JSON.parse(sessionStorage.getItem('rows')) || [];
let sessions = [];
let currentFileName = sessionStorage.getItem('currentFileName') || '';
let editingIndex = null;
let lastAction = null;
let presets = JSON.parse(localStorage.getItem('presets')) || [];
let validationRules = JSON.parse(localStorage.getItem('validationRules')) || {};
let offlineQueue = [];
let theme = localStorage.getItem('theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
let isSaving = false;
let currentPage = sessionStorage.getItem('currentPage') || 'home';

// DOM Elements
const homeScreen = document.getElementById('home-screen');
const dataEntryScreen = document.getElementById('data-entry-screen');
const fileInput = document.getElementById('fileInput');
const startNewBtn = document.getElementById('start-new');
const createTemplateBtn = document.getElementById('create-template');
const createCsvBtn = document.getElementById('create-csv');
const yourWorksBtn = document.getElementById('your-works-btn');
const backBtn = document.getElementById('back-btn');
const downloadBtn = document.getElementById('download-btn');
const dataForm = document.getElementById('data-form');
const addRowBtn = document.getElementById('add-row');
const rowTable = document.getElementById('row-table');
const fileNameSpan = document.getElementById('file-name');
const snackbar = document.getElementById('snackbar');
const saveStatus = document.getElementById('save-status');
const viewRowsBtn = document.getElementById('view-rows-btn');
const presetsBtn = document.getElementById('presets-btn');
const searchBtn = document.getElementById('search-btn');
const bulkEditBtn = document.getElementById('bulk-edit-btn');
const settingsBtn = document.getElementById('settings-btn');

// IndexedDB
let db;
const openDB = () => new Promise((resolve, reject) => {
  const request = indexedDB.open('ExcelApp', 1);
  request.onupgradeneeded = e => {
    db = e.target.result;
    db.createObjectStore('sessions', { keyPath: 'fileName' });
    db.createObjectStore('offlineQueue', { autoIncrement: true });
  };
  request.onsuccess = e => {
    db = e.target.result;
    resolve(db);
  };
  request.onerror = e => reject(e.target.error);
});

// Initialize DB and load sessions
openDB().then(() => {
  const transaction = db.transaction(['sessions'], 'readonly');
  const store = transaction.objectStore('sessions');
  store.getAll().onsuccess = e => {
    sessions = e.target.result;
  };
}).catch(err => console.error('DB Error:', err));

// Theme
document.body.classList.add(theme);
const updateTheme = newTheme => {
  document.body.classList.remove(theme);
  theme = newTheme;
  document.body.classList.add(theme);
  localStorage.setItem('theme', theme);
};

// Show Snackbar
function showSnackbar(message, action) {
  snackbar.innerHTML = action ? `${message} <span onclick="${action}" style="cursor: pointer; text-decoration: underline;">Undo</span>` : message;
  snackbar.classList.add('show');
  setTimeout(() => snackbar.classList.remove('show'), 5000);
}

// Parse CSV File
function parseCSV(text) {
  try {
    const lines = text.trim().split('\n');
    if (lines.length === 0) throw new Error('No data in CSV');
    headers = lines[0].split(',').map(h => h.trim().replace(/[^a-zA-Z0-9\s]/g, ''));
    if (headers.length === 0) throw new Error('No headers found');
    rows = lines.slice(1).map(line => {
      const values = line.split(',').map(v => v.trim());
      const row = {};
      headers.forEach((h, i) => row[h] = values[i] || '');
      return row;
    });
    sessionStorage.setItem('headers', JSON.stringify(headers));
    sessionStorage.setItem('rows', JSON.stringify(rows));
    return true;
  } catch (error) {
    showSnackbar(`Error parsing CSV: ${error.message}`);
    console.error('Parse CSV Error:', error);
    return false;
  }
}

// Parse Excel File
async function parseExcel(file, append = false) {
  if (!file || (!file.name.endsWith('.xlsx') && !file.name.endsWith('.csv'))) {
    showSnackbar('Error: Only .xlsx or .csv files are supported.');
    return false;
  }
  try {
    if (file.name.endsWith('.csv')) {
      const text = await file.text();
      return parseCSV(text);
    } else {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        showSnackbar('Error: No sheets found.');
        return false;
      }
      const firstRow = worksheet.getRow(1).values.slice(1).map(h => String(h || '').trim().replace(/[^a-zA-Z0-9\s]/g, '')).filter(h => h);
      if (firstRow.length === 0) {
        showSnackbar('Error: No valid headers found. Using row indices as headers.');
        headers = worksheet.getRow(1).values.slice(1).map((_, i) => `Column_${i + 1}`);
      } else {
        headers = firstRow;
      }
      const headerSet = new Set(headers);
      if (headerSet.size !== headers.length) {
        showSnackbar('Warning: Duplicate headers detected, using unique indices.');
        headers = headers.map((h, i) => `${h}_${i + 1}`).filter(h => h);
      }
      if (!append) {
        rows = [];
      }
      worksheet.eachRow((row, rowNum) => {
        if (rowNum > 1) {
          const rowData = {};
          headers.forEach((h, i) => {
            rowData[h] = String(row.values[i + 1] || '');
          });
          rows.push(rowData);
        }
      });
      sessionStorage.setItem('headers', JSON.stringify(headers));
      sessionStorage.setItem('rows', JSON.stringify(rows));
      return true;
    }
  } catch (error) {
    showSnackbar('Error: Failed to parse file.');
    console.error('Parse Error:', error);
    return false;
  }
}

// Generate Form (No Table)
function generateForm(container = dataForm) {
  container.innerHTML = '';
  headers.forEach((header, index) => {
    const div = document.createElement('div');
    const inputType = header.toLowerCase().includes('date') ? 'date' :
                      header.toLowerCase().includes('quantity') || header.toLowerCase().includes('price') ? 'number' : 'text';
    div.innerHTML = `
      <label for="${header}">${header}</label>
      <input type="${inputType}" id="${header}" name="${header}" aria-label="${header} input" data-index="${index}" value="${rows[editingIndex]?.[header] || ''}">
      <div class="error" style="display: none;"></div>
    `;
    container.appendChild(div);
  });
  container.querySelectorAll('input').forEach(input => {
    input.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const currentIndex = parseInt(input.dataset.index);
        const nextInput = container.querySelector(`input[data-index="${currentIndex + 1}"]`);
        if (nextInput) {
          nextInput.focus();
        } else if (currentIndex === headers.length - 1) {
          addRowBtn.focus();
        }
      }
    });
  });
  const buttons = [addRowBtn, viewRowsBtn, settingsBtn, bulkEditBtn, searchBtn, presetsBtn];
  buttons.forEach(btn => btn.style.display = 'block');
}

// Render Rows (Independent Editing)
// Render Rows (Independent Editing with Fixed Header Updates)
function renderRows(container, data = rows, editable = true) {
  if (!container) return;
  container.style.display = data.length ? 'block' : 'none';
  if (data.length === 0) {
    container.innerHTML = '<p>No rows added yet.</p>';
    return;
  }
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  thead.innerHTML = `
    <tr><th>Col #</th>${headers.map((_, i) => `<th>${i + 1}</th>`).join('')}</tr>
    <tr><th>Row #</th>${headers.map((h, i) => `<th><input type="text" value="${h}" class="header-edit" data-index="${i}" style="width: 100%;"></th>`).join('')}</tr>
  `;
  thead.querySelectorAll('.header-edit').forEach(input => {
    input.onchange = (e) => {
      const newValue = e.target.value.trim().replace(/[^a-zA-Z0-9\s]/g, '');
      const index = parseInt(e.target.dataset.index);
      if (newValue && !headers.includes(newValue)) {
        const oldHeader = headers[index];
        lastAction = { type: 'updateHeader', oldHeader, newHeader: newValue, index };
        // Preserve data by remapping only the affected column
        rows.forEach(row => {
          if (row.hasOwnProperty(oldHeader)) {
            row[newValue] = row[oldHeader]; // Copy data to new header
            delete row[oldHeader]; // Remove old header key
          } else {
            row[newValue] = ''; // Initialize new header with empty value if not present
          }
        });
        headers[index] = newValue; // Update the header
        sessionStorage.setItem('headers', JSON.stringify(headers));
        sessionStorage.setItem('rows', JSON.stringify(rows));
        saveSession();
        generateForm(); // Update form with new header
        renderRows(container); // Re-render only this panel
        showSnackbar('Header updated!', 'undoAction()');
      } else {
        e.target.value = headers[index]; // Revert if invalid or duplicate
        showSnackbar('Invalid or duplicate header name.');
      }
    };
    input.onblur = (e) => {
      if (!e.target.value.trim()) {
        e.target.value = headers[parseInt(e.target.dataset.index)];
        showSnackbar('Header cannot be empty.');
      }
    };
    input.parentElement.onclick = (e) => input.focus();
  });
  data.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${rowIndex + 1}</td>${headers.map((h, colIndex) => `<td><input type="text" value="${row[h] || ''}" data-row="${rowIndex}" data-col="${colIndex}" ${editable ? '' : 'readonly'} style="width: 100%;"></td>`).join('')}`;
    if (editable) {
      tr.querySelectorAll('input').forEach(input => {
        input.onchange = (e) => {
          const newValue = e.target.value;
          const rowIndex = parseInt(e.target.dataset.row);
          const colIndex = parseInt(e.target.dataset.col);
          const header = headers[colIndex];
          lastAction = { type: 'update', oldRow: { ...rows[rowIndex] }, row: { ...rows[rowIndex] }, index: rowIndex, col: colIndex, newValue };
          rows[rowIndex][header] = newValue; // Update only the specific cell
          sessionStorage.setItem('rows', JSON.stringify(rows));
          saveSession();
          renderRows(container); // Re-render only this panel
          showSnackbar('Cell updated!', 'undoAction()');
        };
        input.onblur = (e) => {
          if (!e.target.value.trim()) e.target.value = rows[parseInt(e.target.dataset.row)][headers[parseInt(e.target.dataset.col)]] || '';
        };
        input.parentElement.onclick = (e) => {
          if (!input.readOnly) input.focus();
          e.stopPropagation();
        };
      });
    }
    tbody.appendChild(tr);
  });
  table.appendChild(thead);
  table.appendChild(tbody);
  container.innerHTML = '';
  container.appendChild(table);
}

// Save Session
let saveTimeout;
function saveSession() {
  if (isSaving) {
    console.log('Save skipped: already in progress');
    return;
  }
  isSaving = true;
  clearTimeout(saveTimeout);
  saveTimeout = setTimeout(() => {
    saveStatus.textContent = 'Saving...';
    console.log('Saving session:', currentFileName);
    const session = { fileName: currentFileName, headers, rows, savedAt: Date.now() };
    const transaction = db.transaction(['sessions'], 'readwrite');
    const store = transaction.objectStore('sessions');
    store.put(session).onsuccess = () => {
      sessions = sessions.filter(s => s.fileName !== currentFileName).concat(session);
      sessionStorage.setItem('headers', JSON.stringify(headers));
      sessionStorage.setItem('rows', JSON.stringify(rows));
      sessionStorage.setItem('currentFileName', currentFileName);
      saveStatus.textContent = 'Saved';
      console.log('Session saved:', currentFileName);
      setTimeout(() => {
        if (saveStatus.textContent === 'Saved') saveStatus.textContent = '';
      }, 2000);
      if (navigator.onLine) syncOfflineQueue();
      isSaving = false;
    };
    store.onerror = () => {
      console.error('Save error');
      saveStatus.textContent = '';
      isSaving = false;
    };
  }, 500);
}

// Reset State
function resetState() {
  headers = [];
  rows = [];
  currentFileName = '';
  editingIndex = null;
  lastAction = null;
  dataForm.innerHTML = '';
  rowTable.style.display = 'none';
  rowTable.innerHTML = '';
  fileNameSpan.textContent = '';
  fileInput.value = '';
  addRowBtn.textContent = 'Add Row ➕';
  addRowBtn.classList.remove('btn-update');
  addRowBtn.classList.add('btn-primary');
  updateProgress();
  sessionStorage.clear();
}

// Export Session
function exportSession(index) {
  const session = sessions[index];
  if (session) {
    downloadFile(session.rows, session.headers, session.fileName);
    showSnackbar('Session exported!');
  } else {
    showSnackbar('Session not found.');
  }
}

// Download File (Fixed Excel Download)
async function downloadFile(rows, headers, fileName, format = 'xlsx') {
  try {
    if (!rows.length) {
      showSnackbar('No data to download.');
      return false;
    }
    const cleanRows = rows.map(row => {
      const cleanRow = {};
      headers.forEach(h => cleanRow[h] = String(row[h] || ''));
      return cleanRow;
    });
    showSnackbar('Downloading...');
    if (format === 'xlsx') {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      worksheet.columns = headers.map(h => ({ header: h, key: h }));
      cleanRows.forEach(row => worksheet.addRow(row));
      const buffer = await workbook.xlsx.writeBuffer();
      if (!buffer || buffer.byteLength === 0) throw new Error('Empty buffer');
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      showSnackbar('Downloaded!');
      return true;
    } else if (format === 'csv') {
      const csv = [headers.join(','), ...cleanRows.map(row => headers.map(h => `"${row[h] || ''}"`).join(','))].join('\n');
      const blob = new Blob([csv], { type: 'text/csv' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName.replace(/\.[^/.]+$/, '') + '.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
      showSnackbar('Downloaded!');
      return true;
    } else if (format === 'txt') {
      const txt = cleanRows.map(row => headers.map(h => row[h] || '').join('\t')).join('\n');
      const blob = new Blob([txt], { type: 'text/plain' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName.replace(/\.[^/.]+$/, '') + '.txt';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
      showSnackbar('Downloaded!');
      return true;
    } else if (format === 'pdf') {
      const doc = new jsPDF();
      doc.text('Data Export', 10, 10);
      doc.autoTable({ head: [headers], body: cleanRows.map(row => headers.map(h => row[h] || '')) });
      doc.save(fileName.replace(/\.[^/.]+$/, '') + '.pdf');
      showSnackbar('Downloaded!');
      return true;
    }
    return true;
  } catch (error) {
    console.error('Download Error:', error.message, error.stack);
    showSnackbar(`Download failed: ${error.message}`);
    return false;
  }
}

// Undo Action
function undoAction() {
  if (!lastAction) return;
  if (lastAction.type === 'add') rows.splice(lastAction.index, 1);
  else if (lastAction.type === 'update') {
    rows[lastAction.index][headers[lastAction.col]] = lastAction.oldRow[headers[lastAction.col]];
  }
  else if (lastAction.type === 'delete') rows.splice(lastAction.index, 0, lastAction.row);
  else if (lastAction.type === 'bulk') lastAction.rows.forEach(({ index, oldRow }) => rows[index] = oldRow);
  else if (lastAction.type === 'addAfter') rows.splice(lastAction.index + 1, 1);
  else if (lastAction.type === 'addColumn') {
    headers.splice(lastAction.index + 1, 1);
    rows.forEach(row => delete row[lastAction.newHeader]);
  } else if (lastAction.type === 'deleteColumn') {
    headers.splice(lastAction.index, 1);
    rows.forEach(row => delete row[lastAction.header]);
  } else if (lastAction.type === 'updateHeader') {
    headers[lastAction.index] = lastAction.oldHeader;
    rows.forEach(row => {
      row[lastAction.oldHeader] = row[lastAction.newHeader];
      delete row[lastAction.newHeader];
    });
  } else if (lastAction.type === 'border') {
    // No undo for borders yet (can be expanded if needed)
  }
  sessionStorage.setItem('rows', JSON.stringify(rows));
  sessionStorage.setItem('headers', JSON.stringify(headers));
  generateForm(); // Update form
  renderRows(rowTable); // Update main table if visible
  saveSession();
  showSnackbar('Action undone!');
  lastAction = null;
}

// Offline Queue
function queueAction(action) {
  offlineQueue.push(action);
  const transaction = db.transaction(['offlineQueue'], 'readwrite');
  const store = transaction.objectStore('offlineQueue');
  store.add(action);
}

async function syncOfflineQueue() {
  const transaction = db.transaction(['offlineQueue'], 'readwrite');
  const store = transaction.objectStore('offlineQueue');
  store.getAll().onsuccess = e => {
    offlineQueue = e.target.result;
    if (offlineQueue.length === 0) return;
    offlineQueue.forEach(action => {
      if (action.type === 'add') rows.push(action.row);
      else if (action.type === 'update') rows[action.index][headers[action.col]] = action.newValue;
      else if (action.type === 'delete') rows.splice(action.index, 1);
      else if (action.type === 'addAfter') rows.splice(action.index + 1, 0, action.row);
      else if (action.type === 'addColumn') {
        headers.splice(action.index + 1, 0, action.newHeader);
        rows.forEach(row => row[action.newHeader] = '');
      } else if (action.type === 'deleteColumn') {
        headers.splice(action.index, 1);
        rows.forEach(row => delete row[action.header]);
      } else if (action.type === 'updateHeader') {
        headers[action.index] = action.oldHeader;
        rows.forEach(row => {
          row[action.oldHeader] = row[action.newHeader];
          delete row[action.newHeader];
        });
      } else if (action.type === 'border') {
        // Handle border update if implemented
      }
    });
    sessionStorage.setItem('rows', JSON.stringify(rows));
    sessionStorage.setItem('headers', JSON.stringify(headers));
    saveSession();
    store.clear();
    offlineQueue = [];
  };
}

// Update Progress
let progressTimeout;
function updateProgress() {
  clearTimeout(progressTimeout);
  progressTimeout = setTimeout(() => {
    fileNameSpan.textContent = currentFileName ? `${currentFileName} (${rows.length} rows)` : '';
  }, 100);
}

// Slide-Up Panels
function createSlidePanel(id, title, contentGenerator, parent = dataEntryScreen) {
  let panel = document.getElementById(id);
  if (!panel) {
    panel = document.createElement('div');
    panel.id = id;
    panel.className = 'slide-panel';
    panel.innerHTML = `
      <div class="panel-header">
        <h2>${title}</h2>
        <div class="panel-actions">
          <button class="btn btn-secondary close-panel-btn" aria-label="Close ${title.toLowerCase()} panel">Close</button>
          <select class="btn btn-primary action-dropdown" aria-label="Actions">
            <option value="">Actions</option>
            <option value="delete">Delete Row</option>
            <option value="addAfter">Add Row After</option>
            <option value="addColumn">Add Column</option>
            <option value="deleteColumn">Delete Column</option>
            <option value="border">Add Border</option>
          </select>
          <select class="btn btn-primary border-sub-options" aria-label="Border options" style="display: none;">
            <option value="all">All Borders</option>
            <option value="none">No Borders</option>
            <option value="outer">Outer Only</option>
          </select>
        </div>
      </div>
      <div class="panel-content"></div>
    `;
    const container = parent.querySelector('.container');
    if (container) {
      container.appendChild(panel);
    } else {
      parent.appendChild(panel);
      console.warn(`No .container found in ${parent.id}, appending panel directly to ${parent.id}`);
    }
    panel.querySelector('.close-panel-btn').onclick = () => {
      panel.classList.remove('show');
      if (parent === dataEntryScreen) {
        dataForm.style.display = 'block';
        addRowBtn.style.display = 'block';
        viewRowsBtn.style.display = 'block';
        settingsBtn.style.display = 'block';
        bulkEditBtn.style.display = 'block';
        searchBtn.style.display = 'block';
        presetsBtn.style.display = 'block';
      }
    };
    const dropdown = panel.querySelector('.action-dropdown');
    const borderSubOptions = panel.querySelector('.border-sub-options');
    dropdown.onchange = (e) => {
      const action = e.target.value;
      if (action === 'border') {
        borderSubOptions.style.display = 'block';
        borderSubOptions.focus();
        e.target.value = '';
      } else {
        borderSubOptions.style.display = 'none';
        if (action === 'delete') {
          const rowNum = prompt('Enter row number to delete (1 to ' + rows.length + '):');
          if (rowNum && !isNaN(rowNum) && rowNum >= 1 && rowNum <= rows.length) {
            if (confirm('Are you sure you want to delete row ' + rowNum + '?')) {
              lastAction = { type: 'delete', row: rows[rowNum - 1], index: rowNum - 1 };
              rows.splice(rowNum - 1, 1);
              sessionStorage.setItem('rows', JSON.stringify(rows));
              saveSession();
              renderRows(panel.querySelector('.panel-content'));
              showSnackbar('Row deleted!', 'undoAction()');
            }
          } else {
            showSnackbar('Invalid row number.');
          }
          e.target.value = '';
        } else if (action === 'addAfter') {
          const rowNum = prompt('Enter row number to add after (1 to ' + rows.length + '):');
          if (rowNum && !isNaN(rowNum) && rowNum >= 0 && rowNum <= rows.length) {
            const newRow = {};
            headers.forEach(h => newRow[h] = '');
            lastAction = { type: 'addAfter', row: newRow, index: rowNum };
            rows.splice(rowNum, 0, newRow);
            sessionStorage.setItem('rows', JSON.stringify(rows));
            saveSession();
            renderRows(panel.querySelector('.panel-content'));
            showSnackbar('Row added after ' + rowNum + '!', 'undoAction()');
          } else {
            showSnackbar('Invalid row number.');
          }
          e.target.value = '';
        } else if (action === 'addColumn') {
          const colNum = prompt('Enter column number to add after (refer to Col # row) (1 to ' + headers.length + '):');
          const colName = prompt('Enter column name:');
          if (colNum && !isNaN(colNum) && colNum >= 0 && colNum <= headers.length && colName && colName.trim()) {
            const newHeader = colName.trim().replace(/[^a-zA-Z0-9\s]/g, '');
            if (headers.includes(newHeader)) {
              showSnackbar('Column name already exists.');
              return;
            }
            lastAction = { type: 'addColumn', index: colNum, newHeader };
            headers.splice(colNum, 0, newHeader);
            rows.forEach(row => row[newHeader] = '');
            sessionStorage.setItem('headers', JSON.stringify(headers));
            sessionStorage.setItem('rows', JSON.stringify(rows));
            saveSession();
            generateForm();
            renderRows(panel.querySelector('.panel-content'));
            showSnackbar('Column added!', 'undoAction()');
          } else {
            showSnackbar('Invalid column number or name.');
          }
          e.target.value = '';
        } else if (action === 'deleteColumn') {
          const colNum = prompt('Enter column number to delete (1 to ' + headers.length + '):');
          if (colNum && !isNaN(colNum) && colNum >= 1 && colNum <= headers.length) {
            const header = headers[colNum - 1];
            if (confirm('Are you sure you want to delete column ' + header + '?')) {
              lastAction = { type: 'deleteColumn', index: colNum - 1, header };
              headers.splice(colNum - 1, 1);
              rows.forEach(row => delete row[header]);
              sessionStorage.setItem('headers', JSON.stringify(headers));
              sessionStorage.setItem('rows', JSON.stringify(rows));
              saveSession();
              generateForm();
              renderRows(panel.querySelector('.panel-content'));
              showSnackbar('Column deleted!', 'undoAction()');
            }
          } else {
            showSnackbar('Invalid column number.');
          }
          e.target.value = '';
        }
      }
    };
    borderSubOptions.onchange = (e) => {
      const borderType = e.target.value;
      borderSubOptions.style.display = 'none';
      applyBorder(borderType, panel.querySelector('.panel-content'));
      lastAction = { type: 'border', borderType };
      showSnackbar(`Borders set to ${borderType}!`, 'undoAction()');
    };
    let startY = 0;
    panel.addEventListener('touchstart', e => {
      startY = e.touches[0].clientY;
    });
    panel.addEventListener('touchend', e => {
      const endY = e.changedTouches[0].clientY;
      if (endY - startY > 100) {
        panel.classList.remove('show');
        if (parent === dataEntryScreen) {
          dataForm.style.display = 'block';
          addRowBtn.style.display = 'block';
          viewRowsBtn.style.display = 'block';
          settingsBtn.style.display = 'block';
          bulkEditBtn.style.display = 'block';
          searchBtn.style.display = 'block';
          presetsBtn.style.display = 'block';
        }
      }
    });
  }
  const content = panel.querySelector('.panel-content');
  content.innerHTML = '';
  contentGenerator(content);
  panel.classList.add('show');
  if (id === 'rows-panel') renderRows(content, rows, true);
  if (parent === dataEntryScreen) {
    dataForm.style.display = 'none';
    addRowBtn.style.display = 'none';
    viewRowsBtn.style.display = 'none';
    settingsBtn.style.display = 'none';
    bulkEditBtn.style.display = 'none';
    searchBtn.style.display = 'none';
    presetsBtn.style.display = 'none';
  }
}

// Apply Borders
function applyBorder(borderType, container) {
  const table = container.querySelector('table');
  if (!table) return;
  const cells = table.querySelectorAll('th, td');
  cells.forEach(cell => {
    if (borderType === 'all') {
      cell.style.border = '1px solid var(--border)';
    } else if (borderType === 'none') {
      cell.style.border = 'none';
    } else if (borderType === 'outer') {
      cell.style.border = 'none';
      if (cell.parentElement.rowIndex === 0 || cell.cellIndex === 0 || cell.parentElement.rowIndex === table.rows.length - 1 || cell.cellIndex === table.rows[0].cells.length - 1) {
        cell.style.border = '1px solid var(--border)';
      }
    }
  });
  renderRows(container);
}

// Create Template
function createTemplatePanel() {
  createSlidePanel('template-panel', 'Create Template 📋', content => {
    const form = document.createElement('form');
    form.id = 'template-form';
    form.innerHTML = `
      <div id="header-inputs">
        <div class="header-input">
          <label for="header-0">Header 1</label>
          <input type="text" id="header-0" name="header-0" aria-label="Header 1 input">
          <button type="button" class="btn btn-secondary remove-header" style="display: none;" aria-label="Remove header">Remove</button>
          <div class="error" style="display: none;"></div>
        </div>
      </div>
      <button type="button" class="btn btn-primary add-header" aria-label="Add another header">Add Header ➕</button>
      <input type="text" id="template-name" placeholder="Template Name (optional)" aria-label="Template name">
      <button type="submit" class="btn btn-primary btn-full" aria-label="Create template">Create 📝</button>
    `;
    content.appendChild(form);

    let headerCount = 1;
    form.querySelector('.add-header').onclick = () => {
      const div = document.createElement('div');
      div.className = 'header-input';
      div.innerHTML = `
        <label for="header-${headerCount}">Header ${headerCount + 1}</label>
        <input type="text" id="header-${headerCount}" name="header-${headerCount}" aria-label="Header ${headerCount + 1} input">
        <button type="button" class="btn btn-secondary remove-header" aria-label="Remove header">Remove</button>
        <div class="error" style="display: none;"></div>
      `;
      form.querySelector('#header-inputs').appendChild(div);
      headerCount++;
      if (headerCount > 1) form.querySelectorAll('.remove-header')[0].style.display = 'block';
    };

    form.addEventListener('click', e => {
      if (e.target.classList.contains('remove-header')) {
        e.target.parentElement.remove();
        headerCount--;
        const inputs = form.querySelectorAll('.header-input');
        inputs.forEach((input, i) => {
          const label = input.querySelector('label');
          const inp = input.querySelector('input');
          label.setAttribute('for', `header-${i}`);
          label.textContent = `Header ${i + 1}`;
          inp.id = `header-${i}`;
          inp.name = `header-${i}`;
          inp.setAttribute('aria-label', `Header ${i + 1} input`);
        });
        if (inputs.length === 1) inputs[0].querySelector('.remove-header').style.display = 'none';
      }
    });

    form.onsubmit = e => {
      e.preventDefault();
      const inputs = form.querySelectorAll('#header-inputs input');
      const newHeaders = Array.from(inputs).map(input => input.value.trim().replace(/[^a-zA-Z0-9\s]/g, '')).filter(h => h);
      const headerSet = new Set(newHeaders);
      let valid = true;

      inputs.forEach(input => {
        const errorDiv = input.nextElementSibling.nextElementSibling;
        errorDiv.style.display = 'none';
        errorDiv.textContent = '';
      });

      if (newHeaders.length === 0) {
        showSnackbar('Error: At least one header is required.');
        valid = false;
      } else if (headerSet.size !== newHeaders.length) {
        inputs.forEach((input, i) => {
          const value = input.value.trim();
          if (value && newHeaders.indexOf(value) !== newHeaders.lastIndexOf(value)) {
            const errorDiv = input.nextElementSibling.nextElementSibling;
            errorDiv.textContent = 'Duplicate header';
            errorDiv.style.display = 'block';
            valid = false;
          }
        });
        showSnackbar('Error: Duplicate headers detected.');
      }

      if (valid) {
        headers = newHeaders;
        rows = [];
        const templateName = form.querySelector('#template-name').value.trim();
        currentFileName = templateName ? `${templateName}.xlsx` : `template_${Date.now()}.xlsx`;
        sessionStorage.setItem('headers', JSON.stringify(headers));
        sessionStorage.setItem('rows', JSON.stringify(rows));
        sessionStorage.setItem('currentFileName', currentFileName);
        updateProgress();
        generateForm();
        saveSession();
        homeScreen.style.display = 'none';
        dataEntryScreen.style.display = 'block';
        hideAllPanels();
        showSnackbar('Template created!');
        sessionStorage.setItem('currentPage', 'data-entry');
      }
    };
  }, homeScreen);
}

// Create CSV Panel
function createCsvPanel() {
  createSlidePanel('create-csv-panel', 'Create CSV 📝', content => {
    const textarea = document.createElement('textarea');
    textarea.placeholder = 'Enter or paste CSV content (e.g., header1,header2\nvalue1,value2)';
    textarea.rows = 10;
    textarea.style.width = '100%';
    textarea.style.boxSizing = 'border-box';
    content.appendChild(textarea);
    const saveBtn = document.createElement('button');
    saveBtn.className = 'btn btn-primary btn-full';
    saveBtn.textContent = 'Save as CSV 📥';
    saveBtn.onclick = () => {
      const csvText = textarea.value.trim();
      if (csvText) {
        const blob = new Blob([csvText], { type: 'text/csv' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `custom_csv_${Date.now()}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        showSnackbar('CSV saved!');
        panel.classList.remove('show');
      } else {
        showSnackbar('Please enter CSV content.');
      }
    };
    content.appendChild(saveBtn);
  }, homeScreen);
}

// Your Works Panel
function showYourWorksPanel() {
  createSlidePanel('your-works-panel', 'Your Works 📂', content => {
    if (sessions.length === 0) {
      content.innerHTML = '<p>No sessions yet.</p>';
      return;
    }
    sessions.forEach((session, index) => {
      const card = document.createElement('div');
      card.className = 'work-card';
      card.innerHTML = `
        <span>${session.fileName} (${session.rows.length} rows, ${new Date(session.savedAt).toLocaleString()})</span>
        <div class="work-card-buttons">
          <button class="btn btn-primary" onclick="resumeSession(${index})" aria-label="Resume ${session.fileName}">Resume 🔄</button>
          <button class="btn btn-primary" onclick="exportSession(${index})" aria-label="Export ${session.fileName}">Export 📤</button>
          <button class="btn btn-danger" onclick="deleteSession(${index})" aria-label="Delete ${session.fileName}">Delete ❌</button>
        </div>
      `;
      content.appendChild(card);
    });
  }, homeScreen);
}

function deleteSession(index) {
  if (confirm('Are you sure you want to delete ' + sessions[index].fileName + '? This action cannot be undone.')) {
    const fileName = sessions[index].fileName;
    sessions.splice(index, 1);
    const transaction = db.transaction(['sessions'], 'readwrite');
    const store = transaction.objectStore('sessions');
    store.delete(fileName);
    showYourWorksPanel();
    showSnackbar('Session deleted!');
  }
}

// Preset Functions
function applyPreset(i) {
  const preset = presets[i];
  headers.forEach(h => {
    const input = dataForm.querySelector(`[name="${h}"]`);
    if (input) input.value = preset[h] || '';
  });
  hideAllPanels();
  showSnackbar('Preset applied!');
}

function deletePreset(i) {
  presets.splice(i, 1);
  localStorage.setItem('presets', JSON.stringify(presets));
  showSnackbar('Preset deleted!');
  renderPresets();
}

function renderPresets(container) {
  container.innerHTML = presets.length ? '' : '<p>No presets yet.</p>';
  presets.forEach((p, i) => {
    const div = document.createElement('div');
    div.className = 'preset-item';
    div.innerHTML = `
      <span>${Object.values(p).filter(v => v).join(', ')}</span>
      <div>
        <button class="btn btn-primary" onclick="applyPreset(${i})" aria-label="Apply preset ${i + 1}">Apply ✅</button>
        <button class="btn btn-secondary" onclick="deletePreset(${i})" aria-label="Delete preset ${i + 1}">Delete ❌</button>
      </div>
    `;
    container.appendChild(div);
  });
}

// Event Listeners
startNewBtn.addEventListener('click', () => {
  fileInput.click();
  sessionStorage.setItem('currentPage', 'data-entry');
});

createTemplateBtn.addEventListener('click', () => {
  createTemplatePanel();
  sessionStorage.setItem('currentPage', 'home');
});

createCsvBtn.addEventListener('click', () => {
  createCsvPanel();
  sessionStorage.setItem('currentPage', 'home');
});

yourWorksBtn.addEventListener('click', () => {
  showYourWorksPanel();
  sessionStorage.setItem('currentPage', 'home');
});

fileInput.addEventListener('change', async () => {
  if (fileInput.files.length === 0) return;
  const file = fileInput.files[0];
  const success = await parseExcel(file);
  if (success) {
    currentFileName = file.name.replace(/\.csv$/, '.xlsx');
    sessionStorage.setItem('currentFileName', currentFileName);
    updateProgress();
    generateForm();
    homeScreen.style.display = 'none';
    dataEntryScreen.style.display = 'block';
    saveSession();
    showSnackbar('File loaded and converted to Excel!');
    sessionStorage.setItem('currentPage', 'data-entry');
  }
});

backBtn.addEventListener('click', () => {
  resetState();
  dataEntryScreen.style.display = 'none';
  homeScreen.style.display = 'block';
  hideAllPanels();
  sessionStorage.setItem('currentPage', 'home');
});

addRowBtn.addEventListener('click', () => {
  const formData = new FormData(dataForm);
  const row = {};
  let valid = true;
  headers.forEach(header => {
    const value = formData.get(header);
    row[header] = value ? String(value).trim() : '';
    if (validationRules[header] && value && !validationRules[header](value)) {
      const input = dataForm.querySelector(`[name="${header}"]`);
      const errorDiv = input.nextElementSibling;
      errorDiv.textContent = `Invalid ${header}`;
      errorDiv.style.display = 'block';
      valid = false;
    }
  });
  if (valid) {
    if (editingIndex !== null) {
      lastAction = { type: 'update', oldRow: { ...rows[editingIndex] }, row, index: editingIndex };
      rows[editingIndex] = row;
      editingIndex = null;
      addRowBtn.textContent = 'Add Row ➕';
      addRowBtn.classList.remove('btn-update');
      addRowBtn.classList.add('btn-primary');
      showSnackbar('Row updated!', 'undoAction()');
    } else {
      lastAction = { type: 'add', row, index: rows.length };
      rows.push(row);
      showSnackbar('Row added!', 'undoAction()');
    }
    sessionStorage.setItem('rows', JSON.stringify(rows));
    if (navigator.onLine) {
      saveSession();
    } else {
      queueAction(lastAction);
      showSnackbar('Action queued offline.');
    }
    updateProgress();
  }
});

downloadBtn.addEventListener('click', async () => {
  const fileName = prompt('Enter file name:', currentFileName || 'data');
  if (fileName) {
    const format = prompt('Enter file type (pdf/excel/csv/txt):', 'excel').toLowerCase();
    if (format && ['pdf', 'excel', 'csv', 'txt'].includes(format)) {
      if (format === 'excel') {
        const success = await downloadFile(rows, headers, fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`, 'xlsx');
        if (success) showSnackbar('Downloaded as .xlsx!');
      } else {
        const success = await downloadFile(rows, headers, fileName, format);
        if (success) showSnackbar('Downloaded!');
      }
    } else {
      showSnackbar('Invalid file type. Use pdf, excel, csv, or txt.');
    }
  }
});

viewRowsBtn.addEventListener('click', () => {
  createSlidePanel('rows-panel', 'Your Rows 👀', content => {
    renderRows(content, rows, true); // Ensure editable is true
  });
  sessionStorage.setItem('currentPage', 'data-entry');
});

presetsBtn.addEventListener('click', () => {
  createSlidePanel('presets-panel', 'Presets ⭐', content => {
    const form = document.createElement('form');
    generateForm(form);
    const saveBtn = document.createElement('button');
    saveBtn.className = 'btn btn-primary btn-full';
    saveBtn.textContent = 'Save Preset ⭐';
    saveBtn.type = 'button';
    saveBtn.onclick = e => {
      e.preventDefault();
      const formData = new FormData(form);
      const preset = {};
      headers.forEach(h => {
        const value = formData.get(h);
        preset[h] = value ? String(value).trim() : '';
      });
      presets.push(preset);
      localStorage.setItem('presets', JSON.stringify(presets));
      showSnackbar('Preset saved!');
      renderPresets(content);
    };
    content.appendChild(form);
    content.appendChild(saveBtn);
    const presetList = document.createElement('div');
    presetList.id = 'preset-list';
    content.appendChild(presetList);
    renderPresets(presetList);
  });
  sessionStorage.setItem('currentPage', 'data-entry');
});

searchBtn.addEventListener('click', () => {
  createSlidePanel('search-panel', 'Search Rows 🔍', content => {
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = 'Search rows...';
    searchInput.className = 'search-input';
    searchInput.setAttribute('aria-label', 'Search rows');
    content.appendChild(searchInput);
    const resultDiv = document.createElement('div');
    resultDiv.id = 'search-results';
    content.appendChild(resultDiv);
    searchInput.oninput = () => {
      const term = searchInput.value.toLowerCase();
      const filtered = rows.filter(row => headers.some(h => String(row[h]).toLowerCase().includes(term)));
      renderRows(resultDiv, filtered, true);
    };
  });
  sessionStorage.setItem('currentPage', 'data-entry');
});

bulkEditBtn.addEventListener('click', () => {
  createSlidePanel('bulk-edit-panel', 'Bulk Edit ✏️', content => {
    const selectedRows = [];
    const tableDiv = document.createElement('div');
    tableDiv.id = 'bulk-select-table';
    renderRows(tableDiv, rows, false);
    tableDiv.querySelectorAll('tr').forEach((tr, i) => {
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.onclick = (e) => {
        e.stopPropagation();
        if (checkbox.checked) selectedRows.push(i);
        else selectedRows.splice(selectedRows.indexOf(i), 1);
      };
      tr.insertBefore(document.createElement('td').appendChild(checkbox), tr.firstChild);
    });
    const form = document.createElement('form');
    generateForm(form);
    const applyBtn = document.createElement('button');
    applyBtn.className = 'btn btn-primary btn-full';
    applyBtn.textContent = 'Apply Changes ✅';
    applyBtn.type = 'button';
    applyBtn.onclick = e => {
      e.preventDefault();
      if (!selectedRows.length) {
        showSnackbar('No rows selected.');
        return;
      }
      const formData = new FormData(form);
      const changes = {};
      headers.forEach(h => {
        const value = formData.get(h);
        if (value) changes[h] = value;
      });
      lastAction = { type: 'bulk', rows: selectedRows.map(i => ({ index: i, oldRow: { ...rows[i] } })) };
      selectedRows.forEach(i => Object.assign(rows[i], changes));
      sessionStorage.setItem('rows', JSON.stringify(rows));
      renderRows(tableDiv);
      saveSession();
      hideAllPanels();
      showSnackbar('Rows updated!', 'undoAction()');
    };
    content.appendChild(tableDiv);
    content.appendChild(form);
    content.appendChild(applyBtn);
  });
  sessionStorage.setItem('currentPage', 'data-entry');
});

settingsBtn.addEventListener('click', () => {
  createSlidePanel('settings-panel', 'Settings ⚙️', content => {
    const validationForm = document.createElement('form');
    headers.forEach(h => {
      const div = document.createElement('div');
      div.innerHTML = `
        <label for="rule-${h}">${h} Rule</label>
        <select id="rule-${h}" name="${h}">
          <option value="">None</option>
          <option value="positive" ${validationRules[h] === positiveRule ? 'selected' : ''}>Positive Number</option>
          <option value="notFuture" ${validationRules[h] === notFutureRule ? 'selected' : ''}>Not Future Date</option>
        </select>
      `;
      validationForm.appendChild(div);
    });
    const saveRulesBtn = document.createElement('button');
    saveRulesBtn.className = 'btn btn-primary btn-full';
    saveRulesBtn.textContent = 'Save Rules ✅';
    saveRulesBtn.onclick = e => {
      e.preventDefault();
      const formData = new FormData(validationForm);
      validationRules = {};
      headers.forEach(h => {
        const rule = formData.get(h);
        if (rule === 'positive') validationRules[h] = positiveRule;
        else if (rule === 'notFuture') validationRules[h] = notFutureRule;
      });
      localStorage.setItem('validationRules', JSON.stringify(validationRules));
      showSnackbar('Rules saved!');
    };
    validationForm.appendChild(saveRulesBtn);

    const themeForm = document.createElement('form');
    themeForm.innerHTML = `
      <label for="theme">Theme</label>
      <select id="theme" name="theme">
        <option value="light" ${theme === 'light' ? 'selected' : ''}>Light ☀️</option>
        <option value="dark" ${theme === 'dark' ? 'selected' : ''}>Dark 🌙</option>
      </select>
    `;
    const saveThemeBtn = document.createElement('button');
    saveThemeBtn.className = 'btn btn-primary btn-full';
    saveThemeBtn.textContent = 'Save Theme ✅';
    saveThemeBtn.onclick = e => {
      e.preventDefault();
      updateTheme(themeForm.querySelector('#theme').value);
      showSnackbar('Theme updated!');
    };
    themeForm.appendChild(saveThemeBtn);

    const importInput = document.createElement('input');
    importInput.type = 'file';
    importInput.accept = '.xlsx,.csv';
    importInput.onchange = async () => {
      if (importInput.files.length) {
        const success = await parseExcel(importInput.files[0], true);
        if (success) {
          renderRows(rowTable);
          saveSession();
          showSnackbar('Rows imported!');
          importInput.value = '';
        }
      }
    };
    const importBtn = document.createElement('button');
    importBtn.className = 'btn btn-primary btn-full';
    importBtn.textContent = 'Import Rows 📥';
    importBtn.onclick = () => importInput.click();

    content.appendChild(validationForm);
    content.appendChild(themeForm);
    content.appendChild(importBtn);
    content.appendChild(importInput);
  });
  sessionStorage.setItem('currentPage', 'data-entry');
});

function positiveRule(value) {
  return !isNaN(value) && Number(value) > 0;
}

function notFutureRule(value) {
  return new Date(value) <= new Date();
}

// Initialize
updateProgress();
rowTable.style.display = 'none'; // Ensure table is hidden by default
if (currentPage === 'data-entry' && headers.length > 0) {
  homeScreen.style.display = 'none';
  dataEntryScreen.style.display = 'block';
  generateForm();
} else {
  homeScreen.style.display = 'block';
  dataEntryScreen.style.display = 'none';
}
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('/sw.js').catch(err => console.error('Service Worker Error:', err));
}