/**
 * Executive Sales Dashboard - Client Application with Manual Refresh, In-Table Deletion & File Manager
 */

// Global State
let allRecords = [];
let filteredRecords = [];
let availableFilters = {};
let currentSort = { column: 'date', direction: 'desc' };
let currentPage = 1;
let pageSize = 25;
let activeDateRange = 'all';

// File & Edit State
let currentDataVersion = null;
let selectedFile = 'AUTO';
let deletedCount = 0;
let cachedAvailableFiles = [];

// Chart instances
let monthlyChart = null;
let categoryChart = null;
let regionChart = null;
let repChart = null;

// Currency & Number Formatters
const curFormatter = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
  maximumFractionDigits: 0
});

const curExactFormatter = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

const numFormatter = new Intl.NumberFormat('en-US');

// DOM Elements
const elements = {
  kpiRevenue: document.getElementById('kpiRevenue'),
  kpiProfit: document.getElementById('kpiProfit'),
  kpiMargin: document.getElementById('kpiMargin'),
  kpiOrders: document.getElementById('kpiOrders'),
  kpiAOV: document.getElementById('kpiAOV'),
  kpiTopCategory: document.getElementById('kpiTopCategory'),
  kpiTopCategoryShare: document.getElementById('kpiTopCategoryShare'),
  kpiTopRegionSub: document.getElementById('kpiTopRegionSub'),
  kpiRevenueBadge: document.getElementById('kpiRevenueBadge'),
  kpiProfitMarginBadge: document.getElementById('kpiProfitMarginBadge'),
  kpiUnitsBadge: document.getElementById('kpiUnitsBadge'),

  searchInput: document.getElementById('searchInput'),
  btnClearSearch: document.getElementById('btnClearSearch'),
  regionSelect: document.getElementById('regionSelect'),
  categorySelect: document.getElementById('categorySelect'),
  segmentSelect: document.getElementById('segmentSelect'),
  repSelect: document.getElementById('repSelect'),
  btnResetFilters: document.getElementById('btnResetFilters'),

  displayedRecordsCount: document.getElementById('displayedRecordsCount'),
  totalRecordsCount: document.getElementById('totalRecordsCount'),
  lastUpdatedTime: document.getElementById('lastUpdatedTime'),
  liveStatusText: document.getElementById('liveStatusText'),
  sourceFileName: document.getElementById('sourceFileName'),
  sourceSheetName: document.getElementById('sourceSheetName'),
  activeSourceTag: document.getElementById('activeSourceTag'),
  excelFileSelect: document.getElementById('excelFileSelect'),
  btnDeleteSelectedFile: document.getElementById('btnDeleteSelectedFile'),
  btnOpenFileManager: document.getElementById('btnOpenFileManager'),

  fileManagerModal: document.getElementById('fileManagerModal'),
  btnCloseFileManager: document.getElementById('btnCloseFileManager'),
  btnCloseFileManagerFooter: document.getElementById('btnCloseFileManagerFooter'),
  fileManagerList: document.getElementById('fileManagerList'),
  btnUploadInModal: document.getElementById('btnUploadInModal'),

  liveNoticeBanner: document.getElementById('liveNoticeBanner'),
  btnDismissNotice: document.getElementById('btnDismissNotice'),
  unsavedBanner: document.getElementById('unsavedBanner'),
  unsavedCount: document.getElementById('unsavedCount'),
  btnSaveToExcel: document.getElementById('btnSaveToExcel'),
  saveBtnText: document.getElementById('saveBtnText'),
  btnUpload: document.getElementById('btnUpload'),
  fileInput: document.getElementById('fileInput'),

  toastNotification: document.getElementById('toastNotification'),
  toastMessage: document.getElementById('toastMessage'),

  tableBody: document.getElementById('tableBody'),
  pageSizeSelect: document.getElementById('pageSizeSelect'),
  paginationInfo: document.getElementById('paginationInfo'),
  pageNumbers: document.getElementById('pageNumbers'),
  btnPrevPage: document.getElementById('btnPrevPage'),
  btnNextPage: document.getElementById('btnNextPage'),

  btnRefresh: document.getElementById('btnRefresh'),
  btnExport: document.getElementById('btnExport'),
  btnThemeToggle: document.getElementById('btnThemeToggle'),
  themeIcon: document.getElementById('themeIcon'),
};

// ==========================================================================
// Initialization
// ==========================================================================
document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
  loadData(false, false);
});

function initEventListeners() {
  // Search input
  elements.searchInput.addEventListener('input', (e) => {
    const val = e.target.value.trim();
    elements.btnClearSearch.classList.toggle('hidden', val.length === 0);
    currentPage = 1;
    applyFilters();
  });

  elements.btnClearSearch.addEventListener('click', () => {
    elements.searchInput.value = '';
    elements.btnClearSearch.classList.add('hidden');
    currentPage = 1;
    applyFilters();
  });

  // Dropdown filters
  [elements.regionSelect, elements.categorySelect, elements.segmentSelect, elements.repSelect].forEach(sel => {
    sel.addEventListener('change', () => {
      currentPage = 1;
      applyFilters();
    });
  });

  // Date range pills
  document.querySelectorAll('.date-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.date-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      activeDateRange = pill.getAttribute('data-range');
      currentPage = 1;
      applyFilters();
    });
  });

  // Excel File selector change
  if (elements.excelFileSelect) {
    elements.excelFileSelect.addEventListener('change', (e) => {
      selectedFile = e.target.value;
      currentDataVersion = null;
      deletedCount = 0;
      updateUnsavedUI();
      loadData(true, true);
    });
  }

  // Delete Currently Selected File Button
  if (elements.btnDeleteSelectedFile) {
    elements.btnDeleteSelectedFile.addEventListener('click', () => {
      const activeFile = selectedFile === 'AUTO' ? elements.sourceFileName.textContent : selectedFile;
      if (!activeFile || activeFile === '-') {
        showToast('No active file selected to delete.', true);
        return;
      }
      deleteExcelFile(activeFile);
    });
  }

  // File Manager Modal Controls
  if (elements.btnOpenFileManager) {
    elements.btnOpenFileManager.addEventListener('click', openFileManager);
  }
  if (elements.btnCloseFileManager) {
    elements.btnCloseFileManager.addEventListener('click', closeFileManager);
  }
  if (elements.btnCloseFileManagerFooter) {
    elements.btnCloseFileManagerFooter.addEventListener('click', closeFileManager);
  }
  if (elements.fileManagerModal) {
    elements.fileManagerModal.addEventListener('click', (e) => {
      if (e.target === elements.fileManagerModal) closeFileManager();
    });
  }
  if (elements.btnUploadInModal) {
    elements.btnUploadInModal.addEventListener('click', () => {
      closeFileManager();
      elements.fileInput.click();
    });
  }

  // Dismiss notice banner
  if (elements.btnDismissNotice) {
    elements.btnDismissNotice.addEventListener('click', () => {
      elements.liveNoticeBanner.classList.add('hidden');
    });
  }

  // Reset Filters
  elements.btnResetFilters.addEventListener('click', resetAllFilters);

  // Manual Refresh Button (bypasses cache)
  elements.btnRefresh.addEventListener('click', handleManualRefresh);

  // Keyboard shortcut: Press 'R' to refresh (when not focused on search input)
  document.addEventListener('keydown', (e) => {
    if ((e.key === 'r' || e.key === 'R') && document.activeElement !== elements.searchInput) {
      handleManualRefresh();
    }
  });

  // Upload File Button
  if (elements.btnUpload && elements.fileInput) {
    elements.btnUpload.addEventListener('click', () => {
      elements.fileInput.click();
    });

    elements.fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleFileUpload(e.target.files[0]);
      }
    });
  }

  // Save to Excel Button
  if (elements.btnSaveToExcel) {
    elements.btnSaveToExcel.addEventListener('click', handleSaveToExcel);
  }

  // Export CSV
  elements.btnExport.addEventListener('click', exportCSV);

  // Theme toggle
  elements.btnThemeToggle.addEventListener('click', toggleTheme);

  // Page size select
  elements.pageSizeSelect.addEventListener('change', (e) => {
    pageSize = parseInt(e.target.value, 10);
    currentPage = 1;
    renderTable();
  });

  // Pagination buttons
  elements.btnPrevPage.addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      renderTable();
    }
  });

  elements.btnNextPage.addEventListener('click', () => {
    const totalPages = Math.ceil(filteredRecords.length / pageSize) || 1;
    if (currentPage < totalPages) {
      currentPage++;
      renderTable();
    }
  });

  // Table header sorting
  document.querySelectorAll('#ordersTable th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (currentSort.column === col) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
      } else {
        currentSort.column = col;
        currentSort.direction = 'desc';
      }
      updateSortIcons();
      applySorting();
      renderTable();
    });
  });
}

// ==========================================================================
// Data Fetching & Manual Refresh
// ==========================================================================
function handleManualRefresh() {
  elements.btnRefresh.classList.add('rotating');
  loadData(true, false, () => {
    setTimeout(() => elements.btnRefresh.classList.remove('rotating'), 600);
    showToast(`Data refreshed manually! Loaded ${numFormatter.format(allRecords.length)} records from ${elements.sourceFileName.textContent}.`);
  });
}

async function loadData(force = false, isLiveUpdate = false, callback = null) {
  try {
    let url = `/api/data?t=${Date.now()}`;
    if (selectedFile !== 'AUTO') {
      url += `&file=${encodeURIComponent(selectedFile)}`;
    }
    if (force) {
      url += `&force=1`;
    }

    const res = await fetch(url, { cache: 'no-store' });
    const data = await res.json();

    if (data.status === 'success') {
      allRecords = data.records;
      availableFilters = data.filters;
      currentDataVersion = data.data_version;
      cachedAvailableFiles = data.available_files || [];

      elements.lastUpdatedTime.textContent = data.last_updated || 'Just now';
      elements.sourceFileName.textContent = data.file || 'Excel File';
      if (elements.sourceSheetName) {
        elements.sourceSheetName.textContent = data.sheet || 'Data';
      }
      if (elements.activeSourceTag) {
        elements.activeSourceTag.textContent = `${data.file} (${data.sheet || 'Data'})`;
      }
      elements.totalRecordsCount.textContent = numFormatter.format(allRecords.length);

      updateExcelFilesList(cachedAvailableFiles, data.file);
      populateFilterOptions();
      applyFilters();

      if (isLiveUpdate) {
        showToast(`Loaded ${numFormatter.format(allRecords.length)} records from ${data.file}!`);
      }
    } else {
      console.error('Failed to load data:', data.error);
      elements.lastUpdatedTime.textContent = 'Error loading Excel';
      showToast('Error reading Excel: ' + (data.error || 'Check file'), true);
    }
  } catch (err) {
    console.error('Error fetching /api/data:', err);
    elements.lastUpdatedTime.textContent = 'Connection error';
  } finally {
    if (callback) callback();
  }
}

function updateExcelFilesList(availableFiles, activeFileName) {
  if (!elements.excelFileSelect) return;
  const currentVal = elements.excelFileSelect.value;
  elements.excelFileSelect.innerHTML = `<option value="AUTO">Auto (Newest: ${activeFileName})</option>`;

  availableFiles.forEach(f => {
    const opt = document.createElement('option');
    opt.value = f.name;
    const folderTag = f.folder === 'downloads' ? '[Downloads]' : '[Files]';
    opt.textContent = `${folderTag} ${f.name} (${f.last_updated})`;
    elements.excelFileSelect.appendChild(opt);
  });

  elements.excelFileSelect.value = currentVal;
}

// ==========================================================================
// File Manager & Delete File Operations
// ==========================================================================
async function deleteExcelFile(filename) {
  if (!confirm(`Are you sure you want to permanently delete '${filename}' from disk?`)) {
    return;
  }

  showToast(`Deleting '${filename}'...`);
  try {
    const res = await fetch('/api/delete_file', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: filename })
    });
    const result = await res.json();

    if (result.status === 'success') {
      showToast(`Deleted '${filename}' successfully!`);
      if (selectedFile === filename || selectedFile === 'AUTO') {
        selectedFile = 'AUTO';
      }
      deletedCount = 0;
      updateUnsavedUI();
      await loadData(true);
      renderFileManagerModal();
    } else {
      showToast(`Delete failed: ${result.message}`, true);
    }
  } catch (err) {
    showToast(`Error deleting file: ${err.message}`, true);
  }
}

function openFileManager() {
  renderFileManagerModal();
  elements.fileManagerModal.classList.remove('hidden');
}

function closeFileManager() {
  elements.fileManagerModal.classList.add('hidden');
}

function renderFileManagerModal() {
  const activeFile = selectedFile === 'AUTO' ? elements.sourceFileName.textContent : selectedFile;

  if (!cachedAvailableFiles || cachedAvailableFiles.length === 0) {
    elements.fileManagerList.innerHTML = `
      <p style="text-align: center; color: var(--text-muted); padding: 2rem;">No Excel or CSV data files found.</p>
    `;
    return;
  }

  elements.fileManagerList.innerHTML = cachedAvailableFiles.map(f => {
    const isActive = f.name === activeFile;
    const isCsv = f.name.toLowerCase().endsWith('.csv');
    const iconClass = isCsv ? 'fa-solid fa-file-csv' : 'fa-solid fa-file-excel';
    const iconColor = isCsv ? 'color: var(--accent-cyan);' : 'color: var(--accent-emerald);';
    const tagClass = f.folder === 'downloads' ? 'tag-downloads' : 'tag-files';
    const tagText = f.folder === 'downloads' ? 'Downloads' : 'Files';

    return `
      <div class="file-list-item ${isActive ? 'active' : ''}">
        <div class="file-info-group">
          <i class="${iconClass} file-type-icon" style="${iconColor}"></i>
          <div style="min-width: 0;">
            <div style="display: flex; align-items: center; flex-wrap: wrap; gap: 0.25rem;">
              <span class="file-name-text" title="${f.name}">${f.name}</span>
              <span class="file-tag ${tagClass}">${tagText}</span>
              ${isActive ? '<span style="font-size: 0.65rem; color: var(--accent-cyan); font-weight: 700; margin-left: 0.35rem;">[ACTIVE]</span>' : ''}
            </div>
            <div class="file-meta-sub">Last saved: ${f.last_updated}</div>
          </div>
        </div>
        <div class="file-actions-group">
          ${!isActive ? `<button class="btn-select-file" onclick="selectAndLoadFile('${f.name}')">Load</button>` : ''}
          <button class="btn-delete-file" title="Permanently delete ${f.name}" onclick="deleteExcelFile('${f.name}')">
            <i class="fa-solid fa-trash-can"></i> Delete
          </button>
        </div>
      </div>
    `;
  }).join('');
}

window.selectAndLoadFile = function(filename) {
  selectedFile = filename;
  elements.excelFileSelect.value = filename;
  deletedCount = 0;
  updateUnsavedUI();
  loadData(true, true);
  closeFileManager();
};

window.deleteExcelFile = deleteExcelFile;

// ==========================================================================
// In-Dashboard Deletion & Saving Back to Excel
// ==========================================================================
window.deleteOrder = function(orderId) {
  const idx = allRecords.findIndex(r => r.order_id === orderId);
  if (idx !== -1) {
    const deleted = allRecords.splice(idx, 1)[0];
    deletedCount++;
    updateUnsavedUI();
    applyFilters();
    showToast(`Deleted order ${orderId} (${curFormatter.format(deleted.revenue)})! Click "Save to Excel" to update file.`);
  }
};

function updateUnsavedUI() {
  if (deletedCount > 0) {
    if (elements.btnSaveToExcel) {
      elements.btnSaveToExcel.classList.remove('hidden');
      if (elements.saveBtnText) {
        elements.saveBtnText.textContent = `Save to Excel (${deletedCount} deleted)`;
      }
    }
    if (elements.unsavedBanner) {
      elements.unsavedBanner.classList.remove('hidden');
      if (elements.unsavedCount) {
        elements.unsavedCount.textContent = deletedCount;
      }
    }
  } else {
    if (elements.btnSaveToExcel) elements.btnSaveToExcel.classList.add('hidden');
    if (elements.unsavedBanner) elements.unsavedBanner.classList.add('hidden');
  }
}

async function handleSaveToExcel() {
  if (!elements.btnSaveToExcel) return;
  elements.btnSaveToExcel.classList.add('rotating');
  try {
    const res = await fetch('/api/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        records: allRecords,
        file: elements.sourceFileName.textContent || 'Sales_Dashboard.xlsx'
      })
    });
    const result = await res.json();
    if (result.status === 'success') {
      deletedCount = 0;
      updateUnsavedUI();
      showToast(`Success! ${result.message}`);
    } else {
      showToast(`Error saving: ${result.message}`, true);
    }
  } catch (err) {
    showToast(`Error saving to Excel: ${err.message}`, true);
  } finally {
    elements.btnSaveToExcel.classList.remove('rotating');
  }
}

async function handleFileUpload(file) {
  if (!file) return;
  showToast(`Uploading ${file.name}...`);
  try {
    const buffer = await file.arrayBuffer();
    const res = await fetch('/api/upload', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/octet-stream',
        'X-File-Name': file.name
      },
      body: buffer
    });
    const result = await res.json();
    if (result.status === 'success') {
      selectedFile = result.file;
      deletedCount = 0;
      updateUnsavedUI();
      showToast(`Uploaded ${result.file}! Loading data...`);
      await loadData(true);
    } else {
      showToast(`Upload failed: ${result.message}`, true);
    }
  } catch (err) {
    showToast(`Upload error: ${err.message}`, true);
  }
}

function showToast(msg, isError = false) {
  if (!elements.toastNotification) return;
  elements.toastMessage.textContent = msg;
  elements.toastNotification.style.background = isError ? 'rgba(244, 63, 94, 0.95)' : 'rgba(16, 185, 129, 0.95)';
  elements.toastNotification.classList.remove('hidden');

  setTimeout(() => {
    elements.toastNotification.classList.add('hidden');
  }, 4000);
}

function populateFilterOptions() {
  populateSelect(elements.regionSelect, availableFilters.regions || [], 'All Regions');
  populateSelect(elements.categorySelect, availableFilters.categories || [], 'All Categories');
  populateSelect(elements.segmentSelect, availableFilters.segments || [], 'All Segments');
  populateSelect(elements.repSelect, availableFilters.sales_reps || [], 'All Sales Reps');
}

function populateSelect(selectEl, items, defaultLabel) {
  const currentVal = selectEl.value;
  selectEl.innerHTML = `<option value="ALL">${defaultLabel}</option>`;
  items.forEach(item => {
    const opt = document.createElement('option');
    opt.value = item;
    opt.textContent = item;
    selectEl.appendChild(opt);
  });
  if (items.includes(currentVal)) {
    selectEl.value = currentVal;
  }
}

function resetAllFilters() {
  elements.searchInput.value = '';
  elements.btnClearSearch.classList.add('hidden');
  elements.regionSelect.value = 'ALL';
  elements.categorySelect.value = 'ALL';
  elements.segmentSelect.value = 'ALL';
  elements.repSelect.value = 'ALL';

  document.querySelectorAll('.date-pill').forEach(p => p.classList.remove('active'));
  const allPill = document.querySelector('.date-pill[data-range="all"]');
  if (allPill) allPill.classList.add('active');
  activeDateRange = 'all';

  currentPage = 1;
  applyFilters();
}

// ==========================================================================
// Filtering & Calculation
// ==========================================================================
function applyFilters() {
  const q = elements.searchInput.value.toLowerCase().trim();
  const region = elements.regionSelect.value;
  const category = elements.categorySelect.value;
  const segment = elements.segmentSelect.value;
  const rep = elements.repSelect.value;

  let minDate = null;
  let maxDate = null;
  const now = new Date('2026-10-30T00:00:00');

  if (activeDateRange === '2026') {
    minDate = new Date('2026-01-01');
    maxDate = new Date('2026-12-31');
  } else if (activeDateRange === '2025') {
    minDate = new Date('2025-01-01');
    maxDate = new Date('2025-12-31');
  } else if (activeDateRange === 'last90') {
    minDate = new Date(now.getTime() - 90 * 86400000);
  } else if (activeDateRange === 'last30') {
    minDate = new Date(now.getTime() - 30 * 86400000);
  }

  filteredRecords = allRecords.filter(row => {
    if (region !== 'ALL' && row.region !== region) return false;
    if (category !== 'ALL' && row.category !== category) return false;
    if (segment !== 'ALL' && row.segment !== segment) return false;
    if (rep !== 'ALL' && row.sales_rep !== rep) return false;

    if (minDate || maxDate) {
      const rowDate = new Date(row.date);
      if (minDate && rowDate < minDate) return false;
      if (maxDate && rowDate > maxDate) return false;
    }

    if (q) {
      const match =
        row.order_id.toLowerCase().includes(q) ||
        row.sales_rep.toLowerCase().includes(q) ||
        row.product.toLowerCase().includes(q) ||
        row.category.toLowerCase().includes(q) ||
        row.region.toLowerCase().includes(q) ||
        row.segment.toLowerCase().includes(q);
      if (!match) return false;
    }

    return true;
  });

  elements.displayedRecordsCount.textContent = numFormatter.format(filteredRecords.length);

  updateKPIs();
  applySorting();
  renderTable();
  renderCharts();
}

function updateKPIs() {
  const count = filteredRecords.length;
  let totalRevenue = 0;
  let totalProfit = 0;
  let totalUnits = 0;

  const categoryTotals = {};
  const regionTotals = {};

  filteredRecords.forEach(r => {
    totalRevenue += r.revenue;
    totalProfit += r.profit;
    totalUnits += r.units;

    categoryTotals[r.category] = (categoryTotals[r.category] || 0) + r.revenue;
    regionTotals[r.region] = (regionTotals[r.region] || 0) + r.revenue;
  });

  const overallMargin = totalRevenue > 0 ? (totalProfit / totalRevenue * 100) : 0;
  const aov = count > 0 ? (totalRevenue / count) : 0;

  let topCat = '-';
  let topCatRev = 0;
  for (const [cat, rev] of Object.entries(categoryTotals)) {
    if (rev > topCatRev) {
      topCatRev = rev;
      topCat = cat;
    }
  }
  const topCatShare = totalRevenue > 0 ? Math.round((topCatRev / totalRevenue) * 100) : 0;

  let topReg = '-';
  let topRegRev = 0;
  for (const [reg, rev] of Object.entries(regionTotals)) {
    if (rev > topRegRev) {
      topRegRev = rev;
      topReg = reg;
    }
  }

  elements.kpiRevenue.textContent = curFormatter.format(totalRevenue);
  elements.kpiProfit.textContent = curFormatter.format(totalProfit);
  elements.kpiMargin.textContent = overallMargin.toFixed(1) + '%';
  elements.kpiOrders.textContent = numFormatter.format(count);
  elements.kpiAOV.textContent = curFormatter.format(aov);
  elements.kpiTopCategory.textContent = topCat;
  elements.kpiTopCategoryShare.textContent = `${topCatShare}% Share`;
  elements.kpiTopRegionSub.textContent = `Top Region: ${topReg}`;

  elements.kpiProfitMarginBadge.textContent = `${overallMargin.toFixed(1)}% Margin`;
  elements.kpiUnitsBadge.textContent = `${numFormatter.format(totalUnits)} Units`;
}

// ==========================================================================
// Sorting & Table Rendering
// ==========================================================================
function updateSortIcons() {
  document.querySelectorAll('#ordersTable th.sortable').forEach(th => {
    const col = th.getAttribute('data-sort');
    th.classList.remove('sort-asc', 'sort-desc');
    const icon = th.querySelector('i');
    if (col === currentSort.column) {
      th.classList.add(currentSort.direction === 'asc' ? 'sort-asc' : 'sort-desc');
      icon.className = currentSort.direction === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';
    } else {
      icon.className = 'fa-solid fa-sort';
    }
  });
}

function applySorting() {
  const { column, direction } = currentSort;
  const mult = direction === 'asc' ? 1 : -1;

  filteredRecords.sort((a, b) => {
    let valA = a[column];
    let valB = b[column];

    if (typeof valA === 'string') {
      return valA.localeCompare(valB) * mult;
    }
    return (valA - valB) * mult;
  });
}

function renderTable() {
  const total = filteredRecords.length;
  const totalPages = Math.ceil(total / pageSize) || 1;

  if (currentPage > totalPages) currentPage = totalPages;
  if (currentPage < 1) currentPage = 1;

  const startIdx = (currentPage - 1) * pageSize;
  const endIdx = Math.min(startIdx + pageSize, total);
  const pageRows = filteredRecords.slice(startIdx, endIdx);

  if (pageRows.length === 0) {
    elements.tableBody.innerHTML = `
      <tr>
        <td colspan="12" class="loading-state">
          <p>No transactions found in this dataset.</p>
        </td>
      </tr>
    `;
    elements.paginationInfo.textContent = 'Showing 0 of 0 entries';
    elements.btnPrevPage.disabled = true;
    elements.btnNextPage.disabled = true;
    elements.pageNumbers.innerHTML = '';
    return;
  }

  const html = pageRows.map(r => {
    const catClass = getCategoryBadgeClass(r.category);
    const marginClass = r.margin >= 40 ? 'margin-high' : (r.margin >= 25 ? 'margin-mid' : 'margin-low');

    return `
      <tr>
        <td><strong>${r.date}</strong></td>
        <td><code>${r.order_id}</code></td>
        <td>${r.region}</td>
        <td><span class="badge-pill ${catClass}">${r.category}</span></td>
        <td>${r.product}</td>
        <td>${r.sales_rep}</td>
        <td><span class="badge-pill badge-segment">${r.segment}</span></td>
        <td class="text-right">${r.units}</td>
        <td class="text-right"><strong>${curExactFormatter.format(r.revenue)}</strong></td>
        <td class="text-right">${curExactFormatter.format(r.profit)}</td>
        <td class="text-right margin-indicator ${marginClass}">${r.margin.toFixed(1)}%</td>
        <td class="text-right">
          <button class="btn-delete-row" title="Delete order ${r.order_id}" onclick="deleteOrder('${r.order_id}')">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </td>
      </tr>
    `;
  }).join('');

  elements.tableBody.innerHTML = html;

  elements.paginationInfo.textContent = `Showing ${startIdx + 1} to ${endIdx} of ${numFormatter.format(total)} entries`;
  elements.btnPrevPage.disabled = currentPage === 1;
  elements.btnNextPage.disabled = currentPage === totalPages;

  renderPaginationNumbers(totalPages);
}

function renderPaginationNumbers(totalPages) {
  let pages = [];
  if (totalPages <= 7) {
    for (let i = 1; i <= totalPages; i++) pages.push(i);
  } else {
    pages = [1];
    if (currentPage > 3) pages.push('...');
    const start = Math.max(2, currentPage - 1);
    const end = Math.min(totalPages - 1, currentPage + 1);
    for (let i = start; i <= end; i++) pages.push(i);
    if (currentPage < totalPages - 2) pages.push('...');
    pages.push(totalPages);
  }

  elements.pageNumbers.innerHTML = pages.map(p => {
    if (p === '...') {
      return `<span style="padding: 0 4px; color: var(--text-muted);">...</span>`;
    }
    const isActive = p === currentPage ? 'active' : '';
    return `<button class="page-num ${isActive}" onclick="goToPage(${p})">${p}</button>`;
  }).join('');
}

window.goToPage = function(page) {
  currentPage = page;
  renderTable();
};

function getCategoryBadgeClass(cat) {
  const map = {
    'Laptops': 'badge-cat-laptop',
    'Accessories': 'badge-cat-accessories',
    'Monitors': 'badge-cat-monitors',
    'Software': 'badge-cat-software',
    'Services': 'badge-cat-services'
  };
  return map[cat] || 'badge-segment';
}

// ==========================================================================
// Chart.js Visualizations
// ==========================================================================
function renderCharts() {
  renderMonthlyTrendChart();
  renderCategoryDonutChart();
  renderRegionBarChart();
  renderSalesRepChart();
}

function renderMonthlyTrendChart() {
  const canvas = document.getElementById('monthlyTrendChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const monthlyAgg = {};
  filteredRecords.forEach(r => {
    const m = r.date.substring(0, 7);
    if (!monthlyAgg[m]) monthlyAgg[m] = { revenue: 0, profit: 0 };
    monthlyAgg[m].revenue += r.revenue;
    monthlyAgg[m].profit += r.profit;
  });

  const sortedMonths = Object.keys(monthlyAgg).sort();
  const labels = sortedMonths.map(m => {
    const [y, mon] = m.split('-');
    const dateObj = new Date(y, parseInt(mon, 10) - 1, 1);
    return dateObj.toLocaleString('en-US', { month: 'short', year: '2-digit' });
  });

  const revenueData = sortedMonths.map(m => Math.round(monthlyAgg[m].revenue));
  const profitData = sortedMonths.map(m => Math.round(monthlyAgg[m].profit));

  if (monthlyChart) monthlyChart.destroy();

  monthlyChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Revenue',
          data: revenueData,
          backgroundColor: 'rgba(59, 130, 246, 0.75)',
          hoverBackgroundColor: 'rgba(59, 130, 246, 1)',
          borderRadius: 6,
          order: 2,
          yAxisID: 'y'
        },
        {
          label: 'Profit',
          data: profitData,
          type: 'line',
          borderColor: '#f59e0b',
          backgroundColor: 'rgba(245, 158, 11, 0.1)',
          borderWidth: 3,
          pointBackgroundColor: '#f59e0b',
          pointBorderColor: '#ffffff',
          pointRadius: 4,
          pointHoverRadius: 6,
          tension: 0.3,
          fill: false,
          order: 1,
          yAxisID: 'y'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          titleFont: { family: 'Outfit', size: 13 },
          bodyFont: { family: 'Plus Jakarta Sans', size: 12 },
          padding: 12,
          borderColor: 'rgba(255, 255, 255, 0.15)',
          borderWidth: 1,
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${curFormatter.format(ctx.raw)}`
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: '#94a3b8', font: { family: 'Plus Jakarta Sans', size: 11 } }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: {
            color: '#94a3b8',
            font: { family: 'Plus Jakarta Sans', size: 11 },
            callback: (v) => '$' + (v / 1000).toFixed(0) + 'k'
          }
        }
      }
    }
  });
}

function renderCategoryDonutChart() {
  const canvas = document.getElementById('categoryDonutChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const catAgg = {};
  filteredRecords.forEach(r => {
    catAgg[r.category] = (catAgg[r.category] || 0) + r.revenue;
  });

  const categories = Object.keys(catAgg).sort((a, b) => catAgg[b] - catAgg[a]);
  const data = categories.map(c => Math.round(catAgg[c]));

  const colors = [
    '#3b82f6',
    '#a855f7',
    '#06b6d4',
    '#10b981',
    '#f59e0b',
    '#f43f5e'
  ];

  if (categoryChart) categoryChart.destroy();

  categoryChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: categories,
      datasets: [{
        data: data,
        backgroundColor: colors.slice(0, categories.length),
        borderWidth: 2,
        borderColor: 'rgba(15, 23, 42, 0.6)',
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '68%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: '#94a3b8',
            font: { family: 'Plus Jakarta Sans', size: 11 },
            padding: 14,
            usePointStyle: true,
            pointStyle: 'circle'
          }
        },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          padding: 12,
          borderColor: 'rgba(255, 255, 255, 0.15)',
          borderWidth: 1,
          callbacks: {
            label: (ctx) => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : 0;
              return `${ctx.label}: ${curFormatter.format(ctx.raw)} (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function renderRegionBarChart() {
  const canvas = document.getElementById('regionBarChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const regionAgg = {};
  filteredRecords.forEach(r => {
    regionAgg[r.region] = (regionAgg[r.region] || 0) + r.revenue;
  });

  const regions = Object.keys(regionAgg).sort((a, b) => regionAgg[b] - regionAgg[a]);
  const data = regions.map(r => Math.round(regionAgg[r]));

  if (regionChart) regionChart.destroy();

  regionChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: regions,
      datasets: [{
        data: data,
        backgroundColor: 'rgba(99, 102, 241, 0.75)',
        hoverBackgroundColor: 'rgba(99, 102, 241, 1)',
        borderRadius: 6
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          padding: 12,
          borderColor: 'rgba(255, 255, 255, 0.15)',
          borderWidth: 1,
          callbacks: {
            label: (ctx) => `Revenue: ${curFormatter.format(ctx.raw)}`
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: {
            color: '#94a3b8',
            font: { family: 'Plus Jakarta Sans', size: 10 },
            callback: (v) => '$' + (v / 1000).toFixed(0) + 'k'
          }
        },
        y: {
          grid: { display: false },
          ticks: { color: '#94a3b8', font: { family: 'Plus Jakarta Sans', size: 11 } }
        }
      }
    }
  });
}

function renderSalesRepChart() {
  const canvas = document.getElementById('salesRepChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  const repAgg = {};
  filteredRecords.forEach(r => {
    repAgg[r.sales_rep] = (repAgg[r.sales_rep] || 0) + r.revenue;
  });

  const reps = Object.keys(repAgg).sort((a, b) => repAgg[b] - repAgg[a]).slice(0, 8);
  const data = reps.map(r => Math.round(repAgg[r]));

  if (repChart) repChart.destroy();

  repChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: reps,
      datasets: [{
        data: data,
        backgroundColor: 'rgba(16, 185, 129, 0.75)',
        hoverBackgroundColor: 'rgba(16, 185, 129, 1)',
        borderRadius: 6
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          padding: 12,
          borderColor: 'rgba(255, 255, 255, 0.15)',
          borderWidth: 1,
          callbacks: {
            label: (ctx) => `Generated: ${curFormatter.format(ctx.raw)}`
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: {
            color: '#94a3b8',
            font: { family: 'Plus Jakarta Sans', size: 10 },
            callback: (v) => '$' + (v / 1000).toFixed(0) + 'k'
          }
        },
        y: {
          grid: { display: false },
          ticks: { color: '#94a3b8', font: { family: 'Plus Jakarta Sans', size: 11 } }
        }
      }
    }
  });
}

// ==========================================================================
// Export to CSV
// ==========================================================================
function exportCSV() {
  if (filteredRecords.length === 0) {
    alert('No records available to export.');
    return;
  }

  const headers = ['Date', 'Order ID', 'Region', 'Category', 'Product', 'Sales Rep', 'Segment', 'Units', 'Unit Price', 'Revenue', 'Cost', 'Profit', 'Margin %'];
  const rows = filteredRecords.map(r => [
    r.date,
    `"${r.order_id}"`,
    `"${r.region}"`,
    `"${r.category}"`,
    `"${r.product}"`,
    `"${r.sales_rep}"`,
    `"${r.segment}"`,
    r.units,
    r.unit_price,
    r.revenue,
    r.cost,
    r.profit,
    r.margin
  ]);

  const csvContent = [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', `sales_dashboard_export_${new Date().toISOString().slice(0, 10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// ==========================================================================
// Theme Switching
// ==========================================================================
function toggleTheme() {
  const isLight = document.body.classList.toggle('theme-light');
  if (isLight) {
    elements.themeIcon.className = 'fa-solid fa-moon';
  } else {
    elements.themeIcon.className = 'fa-solid fa-sun';
  }
  renderCharts();
}
