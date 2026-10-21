// app.bundle.js - Standalone, zero-dependency Excel-authentic spreadsheet and Avery export engine
(function() {
  'use strict';

  const STORAGE_KEY = 'oco_excel_nutrition_sheets_v2';
  const DRAWER_WIDTH_KEY = 'oco_avery_drawer_width';

  // --- Initial Default Products Matching User's Excel Sheets ---
  function createDefaultSheets() {
    return [
      {
        id: 's_papaya',
        name: 'Soft-Dried-Papaya',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B2': { v: '1372kJ /\n328kcal', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B3': { v: '0,12g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B4': { v: '0,08g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B5': { v: '80,9g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B6': { v: '70,4g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B7': { v: '1,54g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B8': { v: '0,23g', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_mango',
        name: 'Dried-Mango',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B2': { v: '1420kJ /\n339kcal', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B3': { v: '0,50g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B4': { v: '0,10g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B5': { v: '82,0g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B6': { v: '65,0g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B7': { v: '1,80g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B8': { v: '0,05g', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_passion',
        name: 'Soft-Dried-PassionFruit',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B2': { v: '1390kJ /\n332kcal', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B3': { v: '0,20g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B4': { v: '0,05g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B5': { v: '78,0g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B6': { v: '68,0g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B7': { v: '2,10g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B8': { v: '0,15g', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_mixed',
        name: 'Mixed-Dried-TropicalFruits',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thick' },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B2': { v: '1410kJ /\n337kcal', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thick', borderRight: 'thin' },
          'B3': { v: '0,35g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thick', borderLeft: 'thin', borderRight: 'thick' },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B4': { v: '0,07g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B5': { v: '81,5g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B6': { v: '69,0g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B7': { v: '1,65g', bold: false, align: 'right', borderBottom: 'thin', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thick', borderRight: 'thin' },
          'B8': { v: '0,18g', bold: false, align: 'right', borderBottom: 'thick', borderTop: 'thin', borderLeft: 'thin', borderRight: 'thick' }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      }
    ];
  }

  // --- State ---
  let appState = {
    sheets: createDefaultSheets(),
    activeSheetId: 's_papaya',
    selectedCell: 'A1',
    selectionRange: { startCol: 'A', startRow: 1, endCol: 'A', endRow: 1 },
    zoom: 1.0,
    fontFamily: 'Calibri',
    fontSize: 11
  };

  // Load from local storage
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed && Array.isArray(parsed.sheets) && parsed.sheets.length > 0) {
        appState = parsed;
      }
    }
  } catch (e) {
    console.warn('Could not load stored state', e);
  }

  function saveAppState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
    } catch (e) {
      console.warn('Could not save state', e);
    }
  }

  function getActiveSheet() {
    let s = appState.sheets.find(x => x.id === appState.activeSheetId);
    if (!s && appState.sheets.length > 0) {
      s = appState.sheets[0];
      appState.activeSheetId = s.id;
    }
    return s;
  }

  // --- Toast ---
  function showToast(msg, isSuccess = true) {
    const el = document.getElementById('toastNotification');
    if (!el) return;
    el.textContent = (isSuccess ? '✅ ' : 'ℹ️ ') + msg;
    el.style.backgroundColor = isSuccess ? '#107c41' : '#0f548c';
    el.style.display = 'block';
    el.style.opacity = '1';
    clearTimeout(el._timer);
    el._timer = setTimeout(() => {
      el.style.opacity = '0';
      setTimeout(() => el.style.display = 'none', 300);
    }, 3500);
  }

  // --- Grid Constants ---
  const COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
  const ROWS_COUNT = 24;

  function colIndex(col) {
    return COLUMNS.indexOf(col);
  }

  function parseAddr(addr) {
    const col = addr.replace(/[0-9]/g, '');
    const row = parseInt(addr.replace(/[^0-9]/g, ''), 10);
    return { col, row };
  }

  // --- DOM References ---
  const sheetTabsContainer = document.getElementById('sheetTabs');
  const addSheetBtn = document.getElementById('addSheetBtn');
  const gridTable = document.getElementById('gridTable');
  const formulaCellAddr = document.getElementById('formulaCellAddr');
  const formulaInput = document.getElementById('formulaInput');
  const previewSvgBox = document.getElementById('previewSvgBox');
  const statusBarSelection = document.getElementById('statusBarSelection');
  const statusBarRange = document.getElementById('statusBarRange');

  // Ribbon Tools
  const fontSelect = document.getElementById('ribbonFontFamily');
  const fontSizeSelect = document.getElementById('ribbonFontSize');
  const boldBtn = document.getElementById('btnBold');
  const italicBtn = document.getElementById('btnItalic');
  const borderDropdownBtn = document.getElementById('btnBorderDropdown');
  const borderMenu = document.getElementById('borderMenu');
  const alignLeftBtn = document.getElementById('btnAlignLeft');
  const alignCenterBtn = document.getElementById('btnAlignCenter');
  const alignRightBtn = document.getElementById('btnAlignRight');
  const indentIncBtn = document.getElementById('btnIndentInc');
  const indentDecBtn = document.getElementById('btnIndentDec');
  const btnSetExportRange = document.getElementById('btnSetExportRange');

  const btnCopyAvery = document.getElementById('btnCopyAvery');
  const btnExportSvg = document.getElementById('btnExportSvg');
  const btnExportPng = document.getElementById('btnExportPng');
  const btnUploadExcel = document.getElementById('btnUploadExcel');
  const excelFilePicker = document.getElementById('excelFilePicker');

  // Drawer & Splitter
  const averyPreviewDrawer = document.getElementById('averyPreviewDrawer');
  const workspaceSplitter = document.getElementById('workspaceSplitter');
  const btnToggleDrawerExpand = document.getElementById('btnToggleDrawerExpand');
  const previewZoomSelect = document.getElementById('previewZoomSelect');
  const previewViewport = document.getElementById('previewViewport');
  const btnBgWhite = document.getElementById('btnBgWhite');
  const btnBgCheck = document.getElementById('btnBgCheck');

  // Restore drawer width
  const savedDrawerWidth = localStorage.getItem(DRAWER_WIDTH_KEY);
  if (savedDrawerWidth && averyPreviewDrawer) {
    averyPreviewDrawer.style.width = savedDrawerWidth + 'px';
  }

  // --- Selection State ---
  let isDraggingSelection = false;
  let dragAnchor = { col: 'A', row: 1 };

  // --- Render Sheet Tabs at Bottom ---
  function renderSheetTabs() {
    sheetTabsContainer.innerHTML = '';
    appState.sheets.forEach(sheet => {
      const isActive = sheet.id === appState.activeSheetId;
      const tab = document.createElement('div');
      tab.className = `sheet-tab ${isActive ? 'active' : ''}`;
      tab.innerHTML = `
        <span class="sheet-tab-title" title="${sheet.name}">${sheet.name}</span>
        <span class="sheet-tab-close" title="Zmazať hárok">✕</span>
      `;

      tab.addEventListener('click', (e) => {
        if (e.target.classList.contains('sheet-tab-close')) return;
        appState.activeSheetId = sheet.id;
        saveAppState();
        renderSheetTabs();
        renderGrid();
        updatePreview();
      });

      // Double click to rename
      tab.querySelector('.sheet-tab-title').addEventListener('dblclick', () => {
        const newName = prompt('Zadajte nový názov hárku:', sheet.name);
        if (newName && newName.trim()) {
          sheet.name = newName.trim();
          saveAppState();
          renderSheetTabs();
          updatePreview();
        }
      });

      // Close tab
      tab.querySelector('.sheet-tab-close').addEventListener('click', (e) => {
        e.stopPropagation();
        if (appState.sheets.length <= 1) {
          alert('Zošit musí obsahovať aspoň jeden hárok.');
          return;
        }
        if (confirm(`Naozaj chcete zmazať hárok "${sheet.name}"?`)) {
          appState.sheets = appState.sheets.filter(s => s.id !== sheet.id);
          if (appState.activeSheetId === sheet.id) {
            appState.activeSheetId = appState.sheets[0].id;
          }
          saveAppState();
          renderSheetTabs();
          renderGrid();
          updatePreview();
        }
      });

      sheetTabsContainer.appendChild(tab);
    });
  }

  // --- Render Spreadsheet Grid ---
  function renderGrid() {
    const sheet = getActiveSheet();
    gridTable.innerHTML = '';

    // Colgroup for column widths
    const colgroup = document.createElement('colgroup');
    const thCol = document.createElement('col');
    thCol.style.width = '42px'; // Row numbers col
    colgroup.appendChild(thCol);

    COLUMNS.forEach(col => {
      const c = document.createElement('col');
      const w = (sheet.colWidths && sheet.colWidths[col]) || (col === 'A' ? 220 : (col === 'B' ? 110 : 70));
      c.style.width = w + 'px';
      colgroup.appendChild(c);
    });
    gridTable.appendChild(colgroup);

    // Header Row (A, B, C, D...)
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    
    // Top-left corner
    const cornerTh = document.createElement('th');
    cornerTh.className = 'grid-corner-cell';
    headerRow.appendChild(cornerTh);

    COLUMNS.forEach(col => {
      const th = document.createElement('th');
      th.className = 'grid-header-col';
      th.dataset.col = col;
      th.innerHTML = `
        <span>${col}</span>
        <div class="col-resize-handle" data-col="${col}"></div>
      `;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    gridTable.appendChild(thead);

    // Body Rows (1 to 24)
    const tbody = document.createElement('tbody');
    for (let r = 1; r <= ROWS_COUNT; r++) {
      const tr = document.createElement('tr');
      const rowHeight = (sheet.rowHeights && sheet.rowHeights[r]) || 24;
      tr.style.height = rowHeight + 'px';

      // Row Header
      const th = document.createElement('th');
      th.className = 'grid-header-row';
      th.textContent = r;
      th.dataset.row = r;
      tr.appendChild(th);

      // Data Cells
      COLUMNS.forEach(col => {
        const addr = col + r;
        const cellData = (sheet.cells && sheet.cells[addr]) || { v: '' };
        const td = document.createElement('td');
        td.className = 'grid-cell';
        td.dataset.addr = addr;
        td.dataset.col = col;
        td.dataset.row = r;
        td.contentEditable = 'true';
        td.spellcheck = false;

        // Apply borders & formatting
        applyCellStyles(td, cellData);

        // Mouse Drag Selection Handlers
        td.addEventListener('mousedown', (e) => {
          if (e.button !== 0) return; // Left click only
          isDraggingSelection = true;
          dragAnchor = { col, row: r };
          updateSelectionRange(col, r, col, r);
        });

        td.addEventListener('mouseenter', () => {
          if (isDraggingSelection) {
            updateSelectionRange(dragAnchor.col, dragAnchor.row, col, r);
          }
        });

        // Cell Selection on focus
        td.addEventListener('focus', () => {
          if (!isDraggingSelection) {
            updateSelectionRange(col, r, col, r);
          }
        });

        // Cell Typing / Editing
        td.addEventListener('input', () => {
          const text = td.innerText;
          setCellValue(addr, text);
          formulaInput.value = text;
          updatePreview();
        });

        // Keydown (Enter, Tab, Arrows)
        td.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            const nextRow = r + 1;
            if (nextRow <= ROWS_COUNT) {
              const nextTd = gridTable.querySelector(`td[data-addr="${col}${nextRow}"]`);
              if (nextTd) nextTd.focus();
            }
          }
        });

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    }
    gridTable.appendChild(tbody);

    setupColumnResize();
    refreshSelectionVisuals();
  }

  // Window mouseup ends drag selection
  window.addEventListener('mouseup', () => {
    if (isDraggingSelection) {
      isDraggingSelection = false;
    }
  });

  function applyCellStyles(td, cellData) {
    td.innerText = cellData.v || '';
    td.style.fontWeight = cellData.bold ? 'bold' : 'normal';
    td.style.fontStyle = cellData.italic ? 'italic' : 'normal';
    td.style.textAlign = cellData.align || 'left';
    td.style.paddingLeft = (cellData.indent ? (8 + cellData.indent) : 6) + 'px';

    // Borders
    td.style.borderBottom = getBorderCss(cellData.borderBottom);
    td.style.borderTop = getBorderCss(cellData.borderTop);
    td.style.borderLeft = getBorderCss(cellData.borderLeft);
    td.style.borderRight = getBorderCss(cellData.borderRight);
  }

  function getBorderCss(type) {
    if (type === 'thick') return '2px solid #000000';
    if (type === 'double') return '3px double #000000';
    if (type === 'thin') return '1px solid #000000';
    if (type === 'none') return '1px solid transparent';
    return '1px solid #e1dfdd'; // Default grid border
  }

  // --- Excel Drag Selection Range Logic ---
  function updateSelectionRange(startCol, startRow, endCol, endRow) {
    const sC = Math.min(colIndex(startCol), colIndex(endCol));
    const eC = Math.max(colIndex(startCol), colIndex(endCol));
    const sR = Math.min(startRow, endRow);
    const eR = Math.max(startRow, endRow);

    appState.selectionRange = {
      startCol: COLUMNS[sC],
      startRow: sR,
      endCol: COLUMNS[eC],
      endRow: eR
    };

    appState.selectedCell = `${startCol}${startRow}`;

    // Update Formula bar address box
    if (sC === eC && sR === eR) {
      formulaCellAddr.textContent = `${startCol}${startRow}`;
      if (statusBarSelection) statusBarSelection.textContent = `Bunka: ${startCol}${startRow}`;
    } else {
      formulaCellAddr.textContent = `${COLUMNS[sC]}${sR}:${COLUMNS[eC]}${eR}`;
      const count = (eC - sC + 1) * (eR - sR + 1);
      if (statusBarSelection) statusBarSelection.textContent = `Výber: ${COLUMNS[sC]}${sR}:${COLUMNS[eC]}${eR} (${count} buniek)`;
    }

    // Sync formula input with active anchor cell
    const sheet = getActiveSheet();
    const cellData = (sheet.cells && sheet.cells[appState.selectedCell]) || { v: '' };
    formulaInput.value = cellData.v || '';

    // Ribbon states
    boldBtn.classList.toggle('active', !!cellData.bold);
    italicBtn.classList.toggle('active', !!cellData.italic);

    refreshSelectionVisuals();
  }

  function refreshSelectionVisuals() {
    const range = appState.selectionRange || { startCol: 'A', startRow: 1, endCol: 'A', endRow: 1 };
    const sC = colIndex(range.startCol);
    const eC = colIndex(range.endCol);
    const sR = range.startRow;
    const eR = range.endRow;

    gridTable.querySelectorAll('td.grid-cell').forEach(td => {
      const c = colIndex(td.dataset.col);
      const r = parseInt(td.dataset.row, 10);
      const isAnchor = td.dataset.addr === appState.selectedCell;

      td.classList.toggle('selected', isAnchor);

      const inRange = (c >= sC && c <= eC && r >= sR && r <= eR);
      td.classList.toggle('in-range', inRange);
      td.classList.toggle('range-top', inRange && r === sR);
      td.classList.toggle('range-bottom', inRange && r === eR);
      td.classList.toggle('range-left', inRange && c === sC);
      td.classList.toggle('range-right', inRange && c === eC);
      td.classList.toggle('range-corner-handle', inRange && c === eC && r === eR);
    });
  }

  function getCellsInRange() {
    const range = appState.selectionRange || { startCol: 'A', startRow: 1, endCol: 'A', endRow: 1 };
    const sC = colIndex(range.startCol);
    const eC = colIndex(range.endCol);
    const sR = range.startRow;
    const eR = range.endRow;

    const list = [];
    for (let r = sR; r <= eR; r++) {
      for (let c = sC; c <= eC; c++) {
        list.push({ col: COLUMNS[c], row: r, addr: `${COLUMNS[c]}${r}` });
      }
    }
    return list;
  }

  function setCellValue(addr, value) {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    if (!sheet.cells[addr]) sheet.cells[addr] = {};
    sheet.cells[addr].v = value;
    saveAppState();
  }

  // --- Formula Bar Events ---
  formulaInput.addEventListener('input', () => {
    const addr = appState.selectedCell;
    setCellValue(addr, formulaInput.value);
    const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
    if (td) {
      td.innerText = formulaInput.value;
    }
    updatePreview();
  });

  // --- Ribbon Button Handlers Applied to Entire Selection Range ---
  boldBtn.addEventListener('click', () => {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const cells = getCellsInRange();
    const anchorData = sheet.cells[appState.selectedCell] || {};
    const targetBold = !anchorData.bold;

    cells.forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr].bold = targetBold;
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) applyCellStyles(td, sheet.cells[item.addr]);
    });

    boldBtn.classList.toggle('active', targetBold);
    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
  });

  italicBtn.addEventListener('click', () => {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const cells = getCellsInRange();
    const anchorData = sheet.cells[appState.selectedCell] || {};
    const targetItalic = !anchorData.italic;

    cells.forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr].italic = targetItalic;
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) applyCellStyles(td, sheet.cells[item.addr]);
    });

    italicBtn.classList.toggle('active', targetItalic);
    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
  });

  // Alignments for Selection
  alignLeftBtn.addEventListener('click', () => applyRangeProp('align', 'left'));
  alignCenterBtn.addEventListener('click', () => applyRangeProp('align', 'center'));
  alignRightBtn.addEventListener('click', () => applyRangeProp('align', 'right'));

  // Indents for Selection
  indentIncBtn.addEventListener('click', () => {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    getCellsInRange().forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr].indent = (sheet.cells[item.addr].indent || 0) + 16;
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) applyCellStyles(td, sheet.cells[item.addr]);
    });
    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
  });

  indentDecBtn.addEventListener('click', () => {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    getCellsInRange().forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr].indent = Math.max(0, (sheet.cells[item.addr].indent || 0) - 16);
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) applyCellStyles(td, sheet.cells[item.addr]);
    });
    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
  });

  function applyRangeProp(prop, value) {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    getCellsInRange().forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr][prop] = value;
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) applyCellStyles(td, sheet.cells[item.addr]);
    });
    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
  }

  // --- Borders Dropdown Menu Applied to Selected Range ---
  borderDropdownBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    borderMenu.classList.toggle('open');
  });

  document.addEventListener('click', () => {
    borderMenu.classList.remove('open');
  });

  document.querySelectorAll('.border-menu-item').forEach(item => {
    item.addEventListener('click', (e) => {
      e.stopPropagation();
      const type = item.dataset.border;
      applyBorderToSelection(type);
      borderMenu.classList.remove('open');
    });
  });

  function applyBorderToSelection(type) {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const range = appState.selectionRange || sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
    const sC = colIndex(range.startCol);
    const eC = colIndex(range.endCol);
    const sR = range.startRow;
    const eR = range.endRow;

    for (let r = sR; r <= eR; r++) {
      for (let c = sC; c <= eC; c++) {
        const addr = `${COLUMNS[c]}${r}`;
        if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
        const cell = sheet.cells[addr];

        if (type === 'all') {
          cell.borderTop = 'thin';
          cell.borderBottom = 'thin';
          cell.borderLeft = 'thin';
          cell.borderRight = 'thin';
        } else if (type === 'thick-box') {
          if (r === sR) cell.borderTop = 'thick';
          if (r === eR) cell.borderBottom = 'thick';
          if (c === sC) cell.borderLeft = 'thick';
          if (c === eC) cell.borderRight = 'thick';
        } else if (type === 'header-thick') {
          if (r === sR) cell.borderBottom = 'thick';
        } else if (type === 'thick-bottom') {
          if (r === eR) cell.borderBottom = 'thick';
        } else if (type === 'double-bottom') {
          if (r === eR) cell.borderBottom = 'double';
        } else if (type === 'none') {
          cell.borderTop = 'none';
          cell.borderBottom = 'none';
          cell.borderLeft = 'none';
          cell.borderRight = 'none';
        }

        const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
        if (td) applyCellStyles(td, cell);
      }
    }

    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
    showToast(`Orámovanie "${type}" bolo použité na výber.`);
  }

  // --- Button: Nastaviť výber ako tabuľku pre Avery ---
  if (btnSetExportRange) {
    btnSetExportRange.addEventListener('click', () => {
      const sheet = getActiveSheet();
      sheet.exportRange = Object.assign({}, appState.selectionRange);
      if (statusBarRange) {
        statusBarRange.textContent = `Tabuľka: ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}`;
      }
      saveAppState();
      updatePreview();
      showToast(`Oblasť pre Avery nastavená na ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}!`);
    });
  }

  // --- Column Resize Logic ---
  function setupColumnResize() {
    let activeHandle = null;
    let startX = 0;
    let startWidth = 0;
    let colName = '';

    gridTable.querySelectorAll('.col-resize-handle').forEach(handle => {
      handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        activeHandle = handle;
        colName = handle.dataset.col;
        startX = e.clientX;
        const th = handle.closest('th');
        startWidth = th.offsetWidth;

        document.body.style.cursor = 'col-resize';
      });
    });

    window.addEventListener('mousemove', (e) => {
      if (!activeHandle) return;
      const diff = e.clientX - startX;
      const newWidth = Math.max(40, startWidth + diff);
      const sheet = getActiveSheet();
      if (!sheet.colWidths) sheet.colWidths = {};
      sheet.colWidths[colName] = newWidth;

      // Update colgroup width directly
      const colIdx = COLUMNS.indexOf(colName) + 1; // +1 for row number col
      const colEl = gridTable.querySelectorAll('col')[colIdx];
      if (colEl) colEl.style.width = newWidth + 'px';
    });

    window.addEventListener('mouseup', () => {
      if (activeHandle) {
        activeHandle = null;
        document.body.style.cursor = '';
        saveAppState();
        updatePreview();
      }
    });
  }

  // --- Interactive Splitter for Avery Preview Drawer ---
  if (workspaceSplitter && averyPreviewDrawer) {
    let isDraggingSplitter = false;
    let startMouseX = 0;
    let startDrawerW = 0;

    workspaceSplitter.addEventListener('mousedown', (e) => {
      e.preventDefault();
      isDraggingSplitter = true;
      startMouseX = e.clientX;
      startDrawerW = averyPreviewDrawer.offsetWidth;
      workspaceSplitter.classList.add('dragging');
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';
    });

    window.addEventListener('mousemove', (e) => {
      if (!isDraggingSplitter) return;
      const deltaX = startMouseX - e.clientX; // Dragging left increases width
      const newWidth = Math.min(850, Math.max(320, startDrawerW + deltaX));
      averyPreviewDrawer.style.width = newWidth + 'px';
    });

    window.addEventListener('mouseup', () => {
      if (isDraggingSplitter) {
        isDraggingSplitter = false;
        workspaceSplitter.classList.remove('dragging');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        localStorage.setItem(DRAWER_WIDTH_KEY, averyPreviewDrawer.offsetWidth);
      }
    });
  }

  // Toggle Drawer Expand Button
  if (btnToggleDrawerExpand) {
    btnToggleDrawerExpand.addEventListener('click', () => {
      const isExpanded = averyPreviewDrawer.classList.toggle('expanded');
      btnToggleDrawerExpand.textContent = isExpanded ? '← Zúžiť' : '↔️ Roztiahnuť';
      if (!isExpanded) {
        averyPreviewDrawer.style.width = '420px';
      } else {
        averyPreviewDrawer.style.width = '640px';
      }
      localStorage.setItem(DRAWER_WIDTH_KEY, averyPreviewDrawer.offsetWidth);
    });
  }

  // Zoom Selector
  if (previewZoomSelect && previewSvgBox) {
    previewZoomSelect.addEventListener('change', () => {
      const scale = parseFloat(previewZoomSelect.value);
      previewSvgBox.style.transform = `scale(${scale})`;
    });
  }

  // Background Toggle (White vs. Checkerboard)
  if (btnBgWhite && btnBgCheck && previewViewport) {
    btnBgWhite.addEventListener('click', () => {
      btnBgWhite.classList.add('active');
      btnBgCheck.classList.remove('active');
      previewViewport.className = 'avery-drawer-body viewport-white';
    });

    btnBgCheck.addEventListener('click', () => {
      btnBgCheck.classList.add('active');
      btnBgWhite.classList.remove('active');
      previewViewport.className = 'avery-drawer-body viewport-checkerboard';
    });
  }

  // --- SVG Live Preview Generator ---
  function generateTableSVG(sheet) {
    const range = sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
    const colWidths = sheet.colWidths || {};
    const rowHeights = sheet.rowHeights || {};

    const sC = colIndex(range.startCol);
    const eC = colIndex(range.endCol);

    const activeCols = [];
    let totalW = 0;
    for (let c = sC; c <= eC; c++) {
      const colLetter = COLUMNS[c];
      const w = colWidths[colLetter] || (colLetter === 'A' ? 220 : (colLetter === 'B' ? 110 : 70));
      activeCols.push({ col: colLetter, width: w, x: totalW });
      totalW += w;
    }

    let totalH = 0;
    const rowPositions = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const h = rowHeights[r] || (r === 2 ? 38 : 24);
      rowPositions.push({ row: r, y: totalH, height: h });
      totalH += h;
    }

    const margin = 2;
    const svgW = totalW + margin * 2;
    const svgH = totalH + margin * 2;

    let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${svgW} ${svgH}" width="${svgW}" height="${svgH}">\n`;
    svg += `<defs>\n<style>\n`;
    svg += `  text { font-family: Calibri, Arial, sans-serif; font-size: 13px; fill: #000000; dominant-baseline: middle; }\n`;
    svg += `  .bold { font-weight: bold; }\n`;
    svg += `  .center { text-anchor: middle; }\n`;
    svg += `  .right { text-anchor: end; }\n`;
    svg += `</style>\n</defs>\n`;

    // Background
    svg += `<rect x="${margin}" y="${margin}" width="${totalW}" height="${totalH}" fill="#ffffff" />\n`;

    // Render cells
    rowPositions.forEach(rp => {
      const r = rp.row;
      const y = rp.y + margin;
      const h = rp.height;

      activeCols.forEach((ac, colIndexInTable) => {
        const addr = `${ac.col}${r}`;
        const cell = (sheet.cells && sheet.cells[addr]) || { v: '' };
        const cellX = margin + ac.x;
        const cellW = ac.width;

        const textY = y + h / 2;
        const indentX = cell.indent || 0;
        const textClass = cell.bold ? 'bold' : '';

        const valLines = String(cell.v || '').split('\n');
        if (valLines.length === 1) {
          if (cell.align === 'center') {
            svg += `<text x="${cellX + cellW / 2}" y="${textY}" class="center ${textClass}">${escapeXml(cell.v)}</text>\n`;
          } else if (cell.align === 'right') {
            svg += `<text x="${cellX + cellW - 8}" y="${textY}" class="right ${textClass}">${escapeXml(cell.v)}</text>\n`;
          } else {
            svg += `<text x="${cellX + 8 + indentX}" y="${textY}" class="${textClass}">${escapeXml(cell.v)}</text>\n`;
          }
        } else {
          // Multiline cell (kJ / kcal)
          const startY = textY - ((valLines.length - 1) * 7);
          const posX = cell.align === 'right' ? (cellX + cellW - 8) : (cellX + 8 + indentX);
          const anchorClass = cell.align === 'right' ? 'right ' + textClass : textClass;

          svg += `<text x="${posX}" y="${startY}" class="${anchorClass}">`;
          valLines.forEach((line, idx) => {
            svg += `<tspan x="${posX}" dy="${idx === 0 ? 0 : 15}">${escapeXml(line)}</tspan>`;
          });
          svg += `</text>\n`;
        }

        // Horizontal bottom border
        if (cell.borderBottom && cell.borderBottom !== 'none') {
          const strokeW = cell.borderBottom === 'thick' ? 2 : (cell.borderBottom === 'double' ? 3 : 1);
          svg += `<line x1="${cellX}" y1="${y + h}" x2="${cellX + cellW}" y2="${y + h}" stroke="#000000" stroke-width="${strokeW}" />\n`;
        }
      });
    });

    // Vertical Divider between columns
    for (let c = 1; c < activeCols.length; c++) {
      const divX = margin + activeCols[c].x;
      svg += `<line x1="${divX}" y1="${margin}" x2="${divX}" y2="${margin + totalH}" stroke="#000000" stroke-width="1" />\n`;
    }

    // Outer Box Border
    svg += `<rect x="${margin}" y="${margin}" width="${totalW}" height="${totalH}" fill="none" stroke="#000000" stroke-width="2" />\n`;

    svg += `</svg>`;
    return { svg, width: svgW, height: svgH };
  }

  function escapeXml(unsafe) {
    return String(unsafe || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function updatePreview() {
    const sheet = getActiveSheet();
    if (!sheet) return;
    const { svg } = generateTableSVG(sheet);
    if (previewSvgBox) {
      previewSvgBox.innerHTML = svg;
    }
    if (statusBarRange && sheet.exportRange) {
      statusBarRange.textContent = `Tabuľka pre Avery: ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}`;
    }
  }

  // --- Add Sheet (+) Button ---
  addSheetBtn.addEventListener('click', () => {
    const name = prompt('Zadajte názov nového hárku:', `Hárok ${appState.sheets.length + 1}`);
    if (name && name.trim()) {
      const newSheet = JSON.parse(JSON.stringify(createDefaultSheets()[0]));
      newSheet.id = 's_' + Date.now();
      newSheet.name = name.trim();
      appState.sheets.push(newSheet);
      appState.activeSheetId = newSheet.id;
      saveAppState();
      renderSheetTabs();
      renderGrid();
      updatePreview();
      showToast(`Hárok "${newSheet.name}" bol pridaný.`);
    }
  });

  // --- Export Handlers ---
  function sanitizeName(name) {
    return String(name || 'Vyzivove_udaje')
      .trim()
      .replace(/[<>:"/\\|?*]/g, '-')
      .replace(/\s+/g, '_');
  }

  // 1. Copy to Clipboard (for Avery Ctrl+V)
  btnCopyAvery.addEventListener('click', async () => {
    const sheet = getActiveSheet();
    const { svg, width, height } = generateTableSVG(sheet);

    try {
      showToast('Generujem obrázok pre Avery...');
      const canvas = document.createElement('canvas');
      const scale = 4; // 300 DPI high resolution
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';

      const img = new Image();
      const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);

      img.onload = async () => {
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);

        canvas.toBlob(async (pngBlob) => {
          if (!pngBlob) return;
          try {
            if (navigator.clipboard && navigator.clipboard.write && typeof ClipboardItem !== 'undefined') {
              const item = new ClipboardItem({ 'image/png': pngBlob });
              await navigator.clipboard.write([item]);
              showToast('Skopírované! V programe Avery stačí stlačiť Ctrl + V 🎉', true);
            } else {
              throw new Error('ClipboardItem nepodporované');
            }
          } catch (clipErr) {
            // Fallback download if clipboard access is blocked
            const a = document.createElement('a');
            a.href = URL.createObjectURL(pngBlob);
            a.download = sanitizeName(sheet.name) + '.png';
            a.click();
            showToast('Obrázok stiahnutý ako PNG (otvorte v Avery).', true);
          }
        }, 'image/png');
      };
      img.src = url;
    } catch (e) {
      console.error(e);
      showToast('Chyba pri kopírovaní do schránky.', false);
    }
  });

  // 2. Download Clean Vector SVG
  btnExportSvg.addEventListener('click', () => {
    const sheet = getActiveSheet();
    const { svg } = generateTableSVG(sheet);
    const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = sanitizeName(sheet.name) + '.svg';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    showToast(`Vektorový súbor "${sheet.name}.svg" bol stiahnutý!`);
  });

  // 3. Download PNG (300 DPI)
  btnExportPng.addEventListener('click', () => {
    const sheet = getActiveSheet();
    const { svg, width, height } = generateTableSVG(sheet);
    const canvas = document.createElement('canvas');
    const scale = 4; // 300 DPI
    canvas.width = width * scale;
    canvas.height = height * scale;
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';

    const img = new Image();
    const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    img.onload = () => {
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(url);
      canvas.toBlob((pngBlob) => {
        if (!pngBlob) return;
        const a = document.createElement('a');
        a.href = URL.createObjectURL(pngBlob);
        a.download = sanitizeName(sheet.name) + '_300DPI.png';
        a.click();
        showToast(`Tlačový obrázok "${sheet.name}_300DPI.png" bol stiahnutý!`);
      }, 'image/png');
    };
    img.src = url;
  });

  // 4. Excel Upload
  btnUploadExcel.addEventListener('click', () => {
    excelFilePicker.click();
  });

  excelFilePicker.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    if (typeof XLSX === 'undefined') {
      alert('Knižnica SheetJS nie je načítaná.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        
        const importedSheets = [];
        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          if (!ws) return;

          const newSheet = JSON.parse(JSON.stringify(createDefaultSheets()[0]));
          newSheet.id = 's_imp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4);
          newSheet.name = sheetName;
          newSheet.cells = {};

          // Read cells
          for (const key in ws) {
            if (key.startsWith('!')) continue;
            const val = ws[key] ? (ws[key].w || ws[key].v || '') : '';
            if (val !== '') {
              const isIndent = String(val).toLowerCase().startsWith('z toho');
              newSheet.cells[key] = {
                v: String(val),
                bold: key.endsWith('1'),
                align: key.startsWith('B') ? 'right' : 'left',
                indent: isIndent ? 18 : 0,
                borderBottom: key.endsWith('1') ? 'thick' : 'thin',
                borderTop: 'thin',
                borderLeft: key.startsWith('A') ? 'thick' : 'thin',
                borderRight: key.startsWith('B') ? 'thick' : 'thin'
              };
            }
          }

          importedSheets.push(newSheet);
        });

        if (importedSheets.length > 0) {
          appState.sheets = importedSheets;
          appState.activeSheetId = importedSheets[0].id;
          saveAppState();
          renderSheetTabs();
          renderGrid();
          updatePreview();
          showToast(`Úspešne načítaných ${importedSheets.length} hárkov z Excelu!`);
        }
      } catch (err) {
        console.error(err);
        alert('Chyba pri načítaní Excel súboru: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
    excelFilePicker.value = '';
  });

  // --- Initial Boot ---
  renderSheetTabs();
  renderGrid();
  updatePreview();
  updateSelectionRange('A', 1, 'B', 8);

})();
