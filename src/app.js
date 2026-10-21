import {
  createInitialState,
  saveState,
  loadState,
  addProduct,
  removeProduct,
  duplicateProduct,
  updateProductRow,
  addRow,
  removeRow,
  reorderRows
} from './model.js';

import { renderTableToSVG } from './svg-renderer.js';
import { parseExcelWorkbook } from './excel-parser.js';
import {
  downloadSVG,
  downloadPNG,
  downloadJPG,
  copyImageToClipboard
} from './export-engine.js';

// Application State
let appState = loadState();

// DOM Elements
const productTabsBar = document.getElementById('productTabsBar');
const addNewProductBtn = document.getElementById('addNewProductBtn');
const uploadExcelBtn = document.getElementById('uploadExcelBtn');
const excelFileInput = document.getElementById('excelFileInput');
const backupBtn = document.getElementById('backupBtn');
const restoreBtn = document.getElementById('restoreBtn');
const restoreFileInput = document.getElementById('restoreFileInput');
const copyClipboardBtn = document.getElementById('copyClipboardBtn');
const footerCopyBtn = document.getElementById('footerCopyBtn');
const downloadSvgBtn = document.getElementById('downloadSvgBtn');
const downloadPngBtn = document.getElementById('downloadPngBtn');
const downloadJpgBtn = document.getElementById('downloadJpgBtn');

const headerTitleInput = document.getElementById('headerTitleInput');
const headerUnitInput = document.getElementById('headerUnitInput');
const rowsContainer = document.getElementById('rowsContainer');
const addRowBtn = document.getElementById('addRowBtn');

const outerBorderWidthSelect = document.getElementById('outerBorderWidthSelect');
const headerBorderSelect = document.getElementById('headerBorderSelect');
const verticalDividerSelect = document.getElementById('verticalDividerSelect');
const borderColorPicker = document.getElementById('borderColorPicker');
const fontFamilySelect = document.getElementById('fontFamilySelect');

const fontSizeSlider = document.getElementById('fontSizeSlider');
const fontSizeValue = document.getElementById('fontSizeValue');
const tableWidthSlider = document.getElementById('tableWidthSlider');
const tableWidthValue = document.getElementById('tableWidthValue');
const colRatioSlider = document.getElementById('colRatioSlider');
const colRatioValue = document.getElementById('colRatioValue');
const cellPaddingSlider = document.getElementById('cellPaddingSlider');
const cellPaddingValue = document.getElementById('cellPaddingValue');

const exportScaleSelect = document.getElementById('exportScaleSelect');
const exportBgSelect = document.getElementById('exportBgSelect');
const exportMarginSlider = document.getElementById('exportMarginSlider');
const exportMarginValue = document.getElementById('exportMarginValue');

const previewViewport = document.getElementById('previewViewport');
const svgContainer = document.getElementById('svgContainer');
const previewZoomSelect = document.getElementById('previewZoomSelect');
const dropzoneOverlay = document.getElementById('dropzoneOverlay');
const toastContainer = document.getElementById('toastContainer');

// Helper to get active product
function getActiveProduct() {
  let prod = appState.products.find(p => p.id === appState.activeProductId);
  if (!prod && appState.products.length > 0) {
    prod = appState.products[0];
    appState.activeProductId = prod.id;
  }
  return prod;
}

// Toast Notification
function showToast(message, type = 'success', duration = 3500) {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <span>${type === 'success' ? '✅' : (type === 'error' ? '❌' : 'ℹ️')}</span>
    <div>${message}</div>
  `;
  toastContainer.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// Render Tabs Bar
function renderTabs() {
  // Remove existing tab items (keeping addNewProductBtn)
  const existingTabs = productTabsBar.querySelectorAll('.tab-item');
  existingTabs.forEach(el => el.remove());

  appState.products.forEach(prod => {
    const isActive = prod.id === appState.activeProductId;
    const tabEl = document.createElement('div');
    tabEl.className = `tab-item ${isActive ? 'active' : ''}`;
    tabEl.dataset.id = prod.id;

    tabEl.innerHTML = `
      <span class="tab-name" title="${prod.name}">${prod.name}</span>
      <div class="tab-actions">
        <button class="btn btn-subtle btn-icon btn-rename" title="Premenovať produkt">✏️</button>
        <button class="btn btn-subtle btn-icon btn-duplicate" title="Duplikovať produkt">📄</button>
        ${appState.products.length > 1 ? '<button class="btn btn-subtle btn-icon btn-delete" title="Zmazať produkt">✕</button>' : ''}
      </div>
    `;

    // Click to activate
    tabEl.addEventListener('click', (e) => {
      if (e.target.closest('.tab-actions')) return;
      appState.activeProductId = prod.id;
      persistAndRefresh();
    });

    // Rename
    const renameBtn = tabEl.querySelector('.btn-rename');
    renameBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const newName = prompt('Zadajte nový názov produktu:', prod.name);
      if (newName && newName.trim()) {
        prod.name = newName.trim();
        persistAndRefresh();
        showToast(`Produkt bol premenovaný na "${prod.name}".`);
      }
    });

    // Duplicate
    const dupBtn = tabEl.querySelector('.btn-duplicate');
    dupBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const duplicated = duplicateProduct(appState, prod.id);
      persistAndRefresh();
      showToast(`Produkt bol skopírovaný ako "${duplicated.name}".`);
    });

    // Delete
    const delBtn = tabEl.querySelector('.btn-delete');
    if (delBtn) {
      delBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (confirm(`Naozaj chcete zmazať produkt "${prod.name}"?`)) {
          removeProduct(appState, prod.id);
          persistAndRefresh();
          showToast(`Produkt bol zmazaný.`, 'info');
        }
      });
    }

    productTabsBar.insertBefore(tabEl, addNewProductBtn);
  });
}

// Render Rows in Editor
function renderRowsEditor() {
  const prod = getActiveProduct();
  if (!prod) return;

  rowsContainer.innerHTML = '';

  prod.rows.forEach((row, index) => {
    const card = document.createElement('div');
    card.className = 'row-card';
    card.dataset.index = index;

    card.innerHTML = `
      <input type="text" class="input-text row-name-input" value="${row.name}" placeholder="Názov položky">
      <textarea class="input-text row-value-input" rows="2" style="resize: vertical; font-family: monospace;" placeholder="Hodnota">${row.value}</textarea>
      
      <div class="row-actions">
        <button class="toggle-chip ${row.indent ? 'active' : ''} btn-indent" title="Odsadenie doprava pre podkategórie (napr. z toho cukry)">Odsadiť</button>
        <button class="toggle-chip ${row.bold ? 'active' : ''} btn-bold" title="Tučné písmo">B</button>
      </div>

      <div style="min-width: 90px;">
        <select class="select-input row-border-select" title="Spodná oddeľovacia čiara">
          <option value="none" ${row.borderBottom === 'none' ? 'selected' : ''}>Bez čiary</option>
          <option value="thin" ${row.borderBottom === 'thin' ? 'selected' : ''}>Tenká čiara</option>
          <option value="thick" ${row.borderBottom === 'thick' ? 'selected' : ''}>Hrubá čiara</option>
          <option value="double" ${row.borderBottom === 'double' ? 'selected' : ''}>Dvojitá čiara</option>
        </select>
      </div>

      <div class="row-actions">
        <button class="btn btn-subtle btn-icon btn-move-up" ${index === 0 ? 'disabled style="opacity:0.3;"' : ''} title="Posunúť vyššie">↑</button>
        <button class="btn btn-subtle btn-icon btn-move-down" ${index === prod.rows.length - 1 ? 'disabled style="opacity:0.3;"' : ''} title="Posunúť nižšie">↓</button>
        <button class="btn btn-subtle btn-icon btn-del-row" style="color: var(--danger);" title="Zmazať riadok">✕</button>
      </div>
    `;

    // Name change
    const nameInput = card.querySelector('.row-name-input');
    nameInput.addEventListener('input', () => {
      row.name = nameInput.value;
      updateSvgPreview();
      saveState(appState);
    });

    // Value change
    const valueInput = card.querySelector('.row-value-input');
    valueInput.addEventListener('input', () => {
      row.value = valueInput.value;
      // Auto adjust height multiplier if newline present
      if (row.value.includes('\n')) {
        row.heightMultiplier = 1.8;
      } else {
        row.heightMultiplier = 1.0;
      }
      updateSvgPreview();
      saveState(appState);
    });

    // Indent toggle
    const indentBtn = card.querySelector('.btn-indent');
    indentBtn.addEventListener('click', () => {
      row.indent = !row.indent;
      indentBtn.classList.toggle('active', row.indent);
      updateSvgPreview();
      saveState(appState);
    });

    // Bold toggle
    const boldBtn = card.querySelector('.btn-bold');
    boldBtn.addEventListener('click', () => {
      row.bold = !row.bold;
      boldBtn.classList.toggle('active', row.bold);
      updateSvgPreview();
      saveState(appState);
    });

    // Border bottom change
    const borderSelect = card.querySelector('.row-border-select');
    borderSelect.addEventListener('change', () => {
      row.borderBottom = borderSelect.value;
      updateSvgPreview();
      saveState(appState);
    });

    // Move up
    const upBtn = card.querySelector('.btn-move-up');
    upBtn.addEventListener('click', () => {
      reorderRows(prod, index, index - 1);
      persistAndRefresh();
    });

    // Move down
    const downBtn = card.querySelector('.btn-move-down');
    downBtn.addEventListener('click', () => {
      reorderRows(prod, index, index + 1);
      persistAndRefresh();
    });

    // Delete row
    const delBtn = card.querySelector('.btn-del-row');
    delBtn.addEventListener('click', () => {
      if (prod.rows.length <= 1) {
        alert('Tabuľka musí obsahovať aspoň jeden riadok.');
        return;
      }
      removeRow(prod, index);
      persistAndRefresh();
    });

    rowsContainer.appendChild(card);
  });
}

// Sync Editor Controls with Active Product Styles
function syncControlsFromProduct() {
  const prod = getActiveProduct();
  if (!prod) return;

  headerTitleInput.value = prod.header.title || 'Výživové údaje na';
  headerUnitInput.value = prod.header.unit || '100g';

  const s = prod.styles || {};
  outerBorderWidthSelect.value = String(s.outerBorderWidth ?? 1.5);
  headerBorderSelect.value = s.headerBorderBottom || 'thick';
  verticalDividerSelect.value = s.verticalDivider ? 'true' : 'false';
  borderColorPicker.value = s.outerBorderColor || '#000000';
  fontFamilySelect.value = s.fontFamily || 'Calibri, Arial, sans-serif';

  fontSizeSlider.value = s.fontSize || 13;
  fontSizeValue.textContent = `${s.fontSize || 13} px`;

  tableWidthSlider.value = s.tableWidth || 320;
  tableWidthValue.textContent = `${s.tableWidth || 320} px`;

  colRatioSlider.value = Math.round((s.colRatio || 0.68) * 100);
  colRatioValue.textContent = `${Math.round((s.colRatio || 0.68) * 100)} %`;

  cellPaddingSlider.value = s.cellPaddingY || 4;
  cellPaddingValue.textContent = `${s.cellPaddingY || 4} px`;
}

// Update Live SVG Preview
function updateSvgPreview() {
  const prod = getActiveProduct();
  if (!prod) return;

  const bgChoice = exportBgSelect.value; // 'transparent' or 'white'
  const margin = parseInt(exportMarginSlider.value, 10) || 1;

  const result = renderTableToSVG(prod, {
    background: bgChoice,
    margin: margin
  });

  svgContainer.innerHTML = result.svgString;
}

// Full Persist & Refresh UI
function persistAndRefresh() {
  saveState(appState);
  renderTabs();
  syncControlsFromProduct();
  renderRowsEditor();
  updateSvgPreview();
}

// Tab Switching inside Left Panel (Values / Styling / Export)
document.querySelectorAll('.subtab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.subtab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.subtab-content').forEach(c => c.style.display = 'none');

    btn.classList.add('active');
    const targetId = btn.dataset.target;
    const targetEl = document.getElementById(targetId);
    if (targetEl) targetEl.style.display = 'block';
  });
});

// Event Listeners for Header Inputs
headerTitleInput.addEventListener('input', () => {
  const prod = getActiveProduct();
  if (prod) {
    prod.header.title = headerTitleInput.value;
    updateSvgPreview();
    saveState(appState);
  }
});

headerUnitInput.addEventListener('input', () => {
  const prod = getActiveProduct();
  if (prod) {
    prod.header.unit = headerUnitInput.value;
    updateSvgPreview();
    saveState(appState);
  }
});

// Add New Product Button
addNewProductBtn.addEventListener('click', () => {
  const name = prompt('Zadajte názov nového produktu:', `Produkt ${appState.products.length + 1}`);
  if (name && name.trim()) {
    addProduct(appState, name.trim());
    persistAndRefresh();
    showToast(`Nový produkt "${name.trim()}" bol vytvorený.`);
  }
});

// Add New Row Button
addRowBtn.addEventListener('click', () => {
  const prod = getActiveProduct();
  if (prod) {
    addRow(prod, prod.rows.length, {
      name: 'Vláknina',
      value: '0,0g',
      indent: false,
      bold: false,
      borderBottom: 'thin'
    });
    persistAndRefresh();
    showToast('Nový riadok bol pridaný.');
  }
});

// Styling Controls Event Listeners
outerBorderWidthSelect.addEventListener('change', () => {
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.outerBorderWidth = parseFloat(outerBorderWidthSelect.value);
    updateSvgPreview();
    saveState(appState);
  }
});

headerBorderSelect.addEventListener('change', () => {
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.headerBorderBottom = headerBorderSelect.value;
    updateSvgPreview();
    saveState(appState);
  }
});

verticalDividerSelect.addEventListener('change', () => {
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.verticalDivider = verticalDividerSelect.value === 'true';
    updateSvgPreview();
    saveState(appState);
  }
});

borderColorPicker.addEventListener('input', () => {
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.outerBorderColor = borderColorPicker.value;
    prod.styles.innerBorderColor = borderColorPicker.value;
    updateSvgPreview();
    saveState(appState);
  }
});

fontFamilySelect.addEventListener('change', () => {
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.fontFamily = fontFamilySelect.value;
    updateSvgPreview();
    saveState(appState);
  }
});

fontSizeSlider.addEventListener('input', () => {
  const val = parseFloat(fontSizeSlider.value);
  fontSizeValue.textContent = `${val} px`;
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.fontSize = val;
    updateSvgPreview();
    saveState(appState);
  }
});

tableWidthSlider.addEventListener('input', () => {
  const val = parseInt(tableWidthSlider.value, 10);
  tableWidthValue.textContent = `${val} px`;
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.tableWidth = val;
    updateSvgPreview();
    saveState(appState);
  }
});

colRatioSlider.addEventListener('input', () => {
  const val = parseInt(colRatioSlider.value, 10);
  colRatioValue.textContent = `${val} %`;
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.colRatio = val / 100;
    updateSvgPreview();
    saveState(appState);
  }
});

cellPaddingSlider.addEventListener('input', () => {
  const val = parseInt(cellPaddingSlider.value, 10);
  cellPaddingValue.textContent = `${val} px`;
  const prod = getActiveProduct();
  if (prod) {
    prod.styles.cellPaddingY = val;
    updateSvgPreview();
    saveState(appState);
  }
});

// Export Setting Listeners
exportBgSelect.addEventListener('change', () => {
  updateSvgPreview();
});

exportMarginSlider.addEventListener('input', () => {
  exportMarginValue.textContent = `${exportMarginSlider.value} px`;
  updateSvgPreview();
});

// Preview Zoom
previewZoomSelect.addEventListener('change', () => {
  const scale = parseFloat(previewZoomSelect.value);
  svgContainer.style.transform = `scale(${scale})`;
});

// Preview Background Switcher
document.querySelectorAll('.bg-mode-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.bg-mode-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    const mode = btn.dataset.bg;
    previewViewport.className = `preview-viewport viewport-${mode}`;
  });
});

// Clipboard Copy Handler
async function handleCopyClipboard() {
  const prod = getActiveProduct();
  if (!prod) return;

  const scale = parseInt(exportScaleSelect.value, 10) || 4;
  const bg = exportBgSelect.value;
  const margin = parseInt(exportMarginSlider.value, 10) || 1;

  const result = renderTableToSVG(prod, { background: bg, margin });

  try {
    showToast('Pripravujem obrázok vo vysokom rozlíšení...', 'info', 1500);
    await copyImageToClipboard(result.svgString, scale, bg);
    showToast('Obrázok tabuľky bol skopírovaný do schránky! Vložte ho do Avery stlačením Ctrl+V.', 'success', 5000);
  } catch (err) {
    console.error('Clipboard error', err);
    showToast('Nepodarilo sa skopírovať priamo do schránky. Použite tlačidlo Stiahnuť PNG.', 'error', 5000);
  }
}

copyClipboardBtn.addEventListener('click', handleCopyClipboard);
footerCopyBtn.addEventListener('click', handleCopyClipboard);

// Download SVG Handler
downloadSvgBtn.addEventListener('click', () => {
  const prod = getActiveProduct();
  if (!prod) return;
  const bg = exportBgSelect.value;
  const margin = parseInt(exportMarginSlider.value, 10) || 1;
  const result = renderTableToSVG(prod, { background: bg, margin });

  downloadSVG(result.svgString, prod.name);
  showToast(`Vektorové SVG "${prod.name}.svg" bolo úspešne stiahnuté!`);
});

// Download PNG Handler
downloadPngBtn.addEventListener('click', async () => {
  const prod = getActiveProduct();
  if (!prod) return;
  const scale = parseInt(exportScaleSelect.value, 10) || 4;
  const bg = exportBgSelect.value;
  const margin = parseInt(exportMarginSlider.value, 10) || 1;
  const result = renderTableToSVG(prod, { background: bg, margin });

  try {
    showToast('Generujem 300 DPI PNG obrázok...', 'info', 1500);
    await downloadPNG(result.svgString, prod.name, scale, bg);
    showToast(`Obrázok "${prod.name}.png" bol úspešne stiahnutý!`);
  } catch (e) {
    console.error(e);
    showToast('Chyba pri generovaní PNG.', 'error');
  }
});

// Download JPG Handler
downloadJpgBtn.addEventListener('click', async () => {
  const prod = getActiveProduct();
  if (!prod) return;
  const scale = parseInt(exportScaleSelect.value, 10) || 4;
  const margin = parseInt(exportMarginSlider.value, 10) || 1;
  const result = renderTableToSVG(prod, { background: 'white', margin });

  try {
    showToast('Generujem JPG obrázok...', 'info', 1500);
    await downloadJPG(result.svgString, prod.name, scale);
    showToast(`Obrázok "${prod.name}.jpg" bol úspešne stiahnutý!`);
  } catch (e) {
    console.error(e);
    showToast('Chyba pri generovaní JPG.', 'error');
  }
});

// Excel Upload Handler
function handleExcelFile(file) {
  if (!file) return;
  if (typeof XLSX === 'undefined') {
    showToast('Knižnica XLSX nie je načítaná.', 'error');
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const products = parseExcelWorkbook(workbook);

      if (products.length === 0) {
        showToast('V súbore sa nenašli žiadne platné nutričné tabuľky.', 'error');
        return;
      }

      // Add to products list
      products.forEach(newP => {
        // Replace existing if same name or add
        const existingIdx = appState.products.findIndex(p => p.name === newP.name);
        if (existingIdx !== -1) {
          appState.products[existingIdx] = newP;
        } else {
          appState.products.push(newP);
        }
      });

      appState.activeProductId = products[0].id;
      persistAndRefresh();
      showToast(`Úspešne načítaných ${products.length} produktov z Excelu!`, 'success', 5000);
    } catch (err) {
      console.error(err);
      showToast('Chyba pri spracovaní Excel súboru.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

uploadExcelBtn.addEventListener('click', () => excelFileInput.click());
excelFileInput.addEventListener('change', (e) => {
  if (e.target.files && e.target.files[0]) {
    handleExcelFile(e.target.files[0]);
    excelFileInput.value = '';
  }
});

// Drag & Drop on Window
window.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzoneOverlay.classList.add('drag-active');
});

dropzoneOverlay.addEventListener('dragleave', (e) => {
  e.preventDefault();
  dropzoneOverlay.classList.remove('drag-active');
});

dropzoneOverlay.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzoneOverlay.classList.remove('drag-active');
  if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
    handleExcelFile(e.dataTransfer.files[0]);
  }
});

// Backup & Restore
backupBtn.addEventListener('click', () => {
  const json = JSON.stringify(appState, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Vyzivove_udaje_zaloha_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('Záloha produktov bola stiahnutá.');
});

restoreBtn.addEventListener('click', () => restoreFileInput.click());
restoreFileInput.addEventListener('change', (e) => {
  if (e.target.files && e.target.files[0]) {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target.result);
        if (parsed && Array.isArray(parsed.products) && parsed.products.length > 0) {
          appState = parsed;
          persistAndRefresh();
          showToast('Záloha bola úspešne obnovená!');
        } else {
          showToast('Neplatný formát záložného súboru.', 'error');
        }
      } catch (err) {
        showToast('Chyba pri čítaní súboru zálohy.', 'error');
      }
    };
    reader.readAsText(e.target.files[0]);
    restoreFileInput.value = '';
  }
});

// Initial boot
persistAndRefresh();
