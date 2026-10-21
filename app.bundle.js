// app.bundle.js - Standalone, zero-dependency Excel-authentic spreadsheet and high-DPI export engine
(function() {
  'use strict';

  const STORAGE_KEY = 'oco_excel_nutrition_sheets_v3';
  const DRAWER_WIDTH_KEY = 'oco_avery_drawer_width_v3';

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

  // --- Border Normalizer ---
  function normalizeBorder(b) {
    if (!b) return { width: 0, style: 'none', color: 'transparent' };
    if (typeof b === 'object') {
      const w = Number(b.width) || 0;
      const s = b.style || (w > 0 ? 'solid' : 'none');
      const c = b.color || '#000000';
      return { width: w, style: s, color: c };
    }
    if (typeof b === 'string') {
      if (b === 'thick') return { width: 3, style: 'solid', color: '#000000' };
      if (b === 'double') return { width: 3, style: 'double', color: '#000000' };
      if (b === 'thin') return { width: 1, style: 'solid', color: '#000000' };
      if (b === 'none') return { width: 0, style: 'none', color: 'transparent' };
      const match = b.match(/^(\d+)px\s+([a-z]+)\s+(.+)$/i);
      if (match) {
        return { width: parseInt(match[1], 10), style: match[2], color: match[3] };
      }
    }
    return { width: 0, style: 'none', color: 'transparent' };
  }

  function getBorderCss(b) {
    const norm = normalizeBorder(b);
    if (norm.width <= 0 || norm.style === 'none') {
      return '1px solid #e1dfdd'; // Default subtle Excel grid border
    }
    return `${norm.width}px ${norm.style} ${norm.color}`;
  }

  // --- Initial Default Products Matching User's Excel Sheets ---
  function createDefaultSheets() {
    const thickBox = { width: 2, style: 'solid', color: '#000000' };
    const thinInner = { width: 1, style: 'solid', color: '#000000' };

    return [
      {
        id: 's_papaya',
        name: 'Soft-Dried-Papaya',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B2': { v: '1372kJ /\n328kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: thinInner, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B3': { v: '0,12g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B4': { v: '0,08g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B5': { v: '80,9g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B6': { v: '70,4g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B7': { v: '1,54g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B8': { v: '0,23g', bold: false, align: 'right', borderBottom: thickBox, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_mango',
        name: 'Dried-Mango',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B2': { v: '1420kJ /\n339kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: thinInner, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B3': { v: '0,50g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B4': { v: '0,10g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B5': { v: '82,0g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B6': { v: '65,0g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B7': { v: '1,80g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B8': { v: '0,05g', bold: false, align: 'right', borderBottom: thickBox, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_passion',
        name: 'Soft-Dried-PassionFruit',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B2': { v: '1390kJ /\n332kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: thinInner, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B3': { v: '0,20g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B4': { v: '0,05g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B5': { v: '78,0g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B6': { v: '68,0g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B7': { v: '2,10g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B8': { v: '0,15g', bold: false, align: 'right', borderBottom: thickBox, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_mixed',
        name: 'Mixed-Dried-TropicalFruits',
        colWidths: { A: 220, B: 110, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'B1': { v: '100g', bold: true, align: 'center', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B2': { v: '1410kJ /\n337kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderBottom: thinInner, borderTop: thickBox, borderLeft: thickBox, borderRight: thinInner },
          'B3': { v: '0,35g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thickBox, borderLeft: thinInner, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B4': { v: '0,07g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B5': { v: '81,5g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B6': { v: '69,0g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B7': { v: '1,65g', bold: false, align: 'right', borderBottom: thinInner, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderTop: thinInner, borderLeft: thickBox, borderRight: thinInner },
          'B8': { v: '0,18g', bold: false, align: 'right', borderBottom: thickBox, borderTop: thinInner, borderLeft: thinInner, borderRight: thickBox }
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

  // --- Clean Toast (No emojis) ---
  function showToast(msg, isSuccess = true) {
    const el = document.getElementById('toastNotification');
    if (!el) return;
    el.textContent = msg;
    el.style.backgroundColor = isSuccess ? '#107c41' : '#0f548c';
    el.style.display = 'block';
    el.style.opacity = '1';
    clearTimeout(el._timer);
    el._timer = setTimeout(() => {
      el.style.opacity = '0';
      setTimeout(() => el.style.display = 'none', 300);
    }, 3500);
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
  const btnFontInc = document.getElementById('btnFontInc');
  const btnFontDec = document.getElementById('btnFontDec');
  const boldBtn = document.getElementById('btnBold');
  const italicBtn = document.getElementById('btnItalic');
  const ribbonTextColor = document.getElementById('ribbonTextColor');
  const ribbonFillColor = document.getElementById('ribbonFillColor');
  const btnTextColorWrap = document.getElementById('btnTextColorWrap');
  const btnFillColorWrap = document.getElementById('btnFillColorWrap');
  const textColorBar = document.getElementById('textColorBar');
  const fillColorBar = document.getElementById('fillColorBar');

  // Border Controls
  const borderDropdownBtn = document.getElementById('btnBorderDropdown');
  const borderMenu = document.getElementById('borderMenu');
  const borderWidthInput = document.getElementById('borderWidthInput');
  const borderStyleSelect = document.getElementById('borderStyleSelect');
  const borderColorPicker = document.getElementById('borderColorPicker');

  // Alignments & Indent
  const alignLeftBtn = document.getElementById('btnAlignLeft');
  const alignCenterBtn = document.getElementById('btnAlignCenter');
  const alignRightBtn = document.getElementById('btnAlignRight');
  const indentIncBtn = document.getElementById('btnIndentInc');
  const indentDecBtn = document.getElementById('btnIndentDec');

  // Data & Table Manipulation
  const btnInsertRow = document.getElementById('btnInsertRow');
  const btnDeleteRow = document.getElementById('btnDeleteRow');
  const btnInsertCol = document.getElementById('btnInsertCol');
  const btnDeleteCol = document.getElementById('btnDeleteCol');
  const btnMoveUp = document.getElementById('btnMoveUp');
  const btnMoveDown = document.getElementById('btnMoveDown');
  const btnMoveLeft = document.getElementById('btnMoveLeft');
  const btnMoveRight = document.getElementById('btnMoveRight');
  const btnClearContent = document.getElementById('btnClearContent');
  const btnClearFormat = document.getElementById('btnClearFormat');

  // Export & Files
  const btnSetExportRange = document.getElementById('btnSetExportRange');
  const btnCopyAvery = document.getElementById('btnCopyAvery');
  const btnExportSvg = document.getElementById('btnExportSvg');
  const btnExportPng = document.getElementById('btnExportPng');
  const btnExportJpg = document.getElementById('btnExportJpg');
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
        <span class="sheet-tab-close" title="Zmazať hárok">&times;</span>
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
    td.style.color = cellData.textColor || '#000000';
    td.style.backgroundColor = cellData.fillColor || '';
    td.style.paddingLeft = (cellData.indent ? (8 + cellData.indent) : 6) + 'px';
    const fontSize = cellData.fontSize || 11;
    td.style.fontSize = fontSize + 'pt';
    if (cellData.fontFamily) {
      td.style.fontFamily = cellData.fontFamily;
    } else {
      td.style.fontFamily = '';
    }

    // Borders
    td.style.borderBottom = getBorderCss(cellData.borderBottom);
    td.style.borderTop = getBorderCss(cellData.borderTop);
    td.style.borderLeft = getBorderCss(cellData.borderLeft);
    td.style.borderRight = getBorderCss(cellData.borderRight);
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
    if (fontSizeSelect) {
      fontSizeSelect.value = String(cellData.fontSize || 11);
    }
    if (fontSelect && cellData.fontFamily) {
      fontSelect.value = cellData.fontFamily;
    }
    if (textColorBar) {
      textColorBar.setAttribute('fill', cellData.textColor || '#c00000');
    }
    if (fillColorBar) {
      fillColorBar.setAttribute('fill', cellData.fillColor || '#ffff00');
    }
    if (ribbonTextColor && cellData.textColor) {
      ribbonTextColor.value = cellData.textColor;
    }
    if (ribbonFillColor && cellData.fillColor) {
      ribbonFillColor.value = cellData.fillColor;
    }

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

  // Font Size & Font Family
  const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48];

  function applyFontSize(size) {
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const cells = getCellsInRange();
    cells.forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr].fontSize = size;
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) applyCellStyles(td, sheet.cells[item.addr]);
    });
    if (fontSizeSelect) fontSizeSelect.value = String(size);
    saveAppState();
    updatePreview();
  }

  if (fontSizeSelect) {
    fontSizeSelect.addEventListener('change', () => {
      applyFontSize(parseInt(fontSizeSelect.value, 10) || 11);
    });
  }

  if (btnFontInc) {
    btnFontInc.addEventListener('click', () => {
      const cur = parseInt(fontSizeSelect ? fontSizeSelect.value : '11', 10) || 11;
      const next = FONT_SIZES.find(s => s > cur) || (cur + 2);
      applyFontSize(next);
    });
  }

  if (btnFontDec) {
    btnFontDec.addEventListener('click', () => {
      const cur = parseInt(fontSizeSelect ? fontSizeSelect.value : '11', 10) || 11;
      const prev = [...FONT_SIZES].reverse().find(s => s < cur) || Math.max(6, cur - 2);
      applyFontSize(prev);
    });
  }

  if (fontSelect) {
    fontSelect.addEventListener('change', () => {
      const fam = fontSelect.value;
      const sheet = getActiveSheet();
      if (!sheet.cells) sheet.cells = {};
      getCellsInRange().forEach(item => {
        if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
        sheet.cells[item.addr].fontFamily = fam;
        const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
        if (td) applyCellStyles(td, sheet.cells[item.addr]);
      });
      saveAppState();
      updatePreview();
    });
  }

  // Text Color & Fill Color with dynamic ribbon bars
  if (btnTextColorWrap && ribbonTextColor) {
    btnTextColorWrap.addEventListener('click', (e) => {
      if (e.target !== ribbonTextColor) {
        ribbonTextColor.click();
      }
    });
  }

  if (btnFillColorWrap && ribbonFillColor) {
    btnFillColorWrap.addEventListener('click', (e) => {
      if (e.target !== ribbonFillColor) {
        ribbonFillColor.click();
      }
    });
  }

  if (ribbonTextColor) {
    ribbonTextColor.addEventListener('input', () => {
      const color = ribbonTextColor.value;
      if (textColorBar) textColorBar.setAttribute('fill', color);
      const sheet = getActiveSheet();
      if (!sheet.cells) sheet.cells = {};
      getCellsInRange().forEach(item => {
        if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
        sheet.cells[item.addr].textColor = color;
        const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
        if (td) applyCellStyles(td, sheet.cells[item.addr]);
      });
      saveAppState();
      updatePreview();
    });
  }

  if (ribbonFillColor) {
    ribbonFillColor.addEventListener('input', () => {
      const color = ribbonFillColor.value;
      if (fillColorBar) fillColorBar.setAttribute('fill', color);
      const sheet = getActiveSheet();
      if (!sheet.cells) sheet.cells = {};
      getCellsInRange().forEach(item => {
        if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
        sheet.cells[item.addr].fillColor = color;
        const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
        if (td) applyCellStyles(td, sheet.cells[item.addr]);
      });
      saveAppState();
      updatePreview();
    });
  }

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

    const borderWidth = parseInt(borderWidthInput ? borderWidthInput.value : '1', 10) || 1;
    const borderStyle = borderStyleSelect ? borderStyleSelect.value : 'solid';
    const borderColor = borderColorPicker ? borderColorPicker.value : '#000000';

    const userBorder = { width: borderWidth, style: borderStyle, color: borderColor };
    const thickBorder = { width: 3, style: 'solid', color: borderColor };
    const doubleBorder = { width: 3, style: 'double', color: borderColor };
    const noneBorder = { width: 0, style: 'none', color: 'transparent' };

    for (let r = sR; r <= eR; r++) {
      for (let c = sC; c <= eC; c++) {
        const addr = `${COLUMNS[c]}${r}`;
        if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
        const cell = sheet.cells[addr];

        if (type === 'all') {
          cell.borderTop = Object.assign({}, userBorder);
          cell.borderBottom = Object.assign({}, userBorder);
          cell.borderLeft = Object.assign({}, userBorder);
          cell.borderRight = Object.assign({}, userBorder);
        } else if (type === 'outer') {
          if (r === sR) cell.borderTop = Object.assign({}, userBorder);
          if (r === eR) cell.borderBottom = Object.assign({}, userBorder);
          if (c === sC) cell.borderLeft = Object.assign({}, userBorder);
          if (c === eC) cell.borderRight = Object.assign({}, userBorder);
        } else if (type === 'inner') {
          if (r > sR) cell.borderTop = Object.assign({}, userBorder);
          if (c > sC) cell.borderLeft = Object.assign({}, userBorder);
        } else if (type === 'inner-horizontal') {
          if (r < eR) cell.borderBottom = Object.assign({}, userBorder);
        } else if (type === 'inner-vertical') {
          if (c < eC) cell.borderRight = Object.assign({}, userBorder);
        } else if (type === 'top') {
          if (r === sR) cell.borderTop = Object.assign({}, userBorder);
        } else if (type === 'bottom') {
          if (r === eR) cell.borderBottom = Object.assign({}, userBorder);
        } else if (type === 'left') {
          if (c === sC) cell.borderLeft = Object.assign({}, userBorder);
        } else if (type === 'right') {
          if (c === eC) cell.borderRight = Object.assign({}, userBorder);
        } else if (type === 'thick-outer') {
          if (r === sR) cell.borderTop = Object.assign({}, thickBorder);
          if (r === eR) cell.borderBottom = Object.assign({}, thickBorder);
          if (c === sC) cell.borderLeft = Object.assign({}, thickBorder);
          if (c === eC) cell.borderRight = Object.assign({}, thickBorder);
        } else if (type === 'thick-bottom') {
          if (r === eR) cell.borderBottom = Object.assign({}, thickBorder);
        } else if (type === 'double-bottom') {
          if (r === eR) cell.borderBottom = Object.assign({}, doubleBorder);
        } else if (type === 'none') {
          cell.borderTop = Object.assign({}, noneBorder);
          cell.borderBottom = Object.assign({}, noneBorder);
          cell.borderLeft = Object.assign({}, noneBorder);
          cell.borderRight = Object.assign({}, noneBorder);
        }

        const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
        if (td) applyCellStyles(td, cell);
      }
    }

    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
    showToast(`Orámovanie bolo aplikované na vybrané bunky (${borderWidth}px, ${borderStyle}).`);
  }

  // --- Row & Column Manipulation ---
  if (btnInsertRow) {
    btnInsertRow.addEventListener('click', () => {
      const sheet = getActiveSheet();
      const range = appState.selectionRange || { startRow: 1 };
      const targetRow = range.startRow;

      const newCells = {};
      for (const addr in sheet.cells) {
        const { col, row } = parseAddr(addr);
        if (row < targetRow) {
          newCells[addr] = sheet.cells[addr];
        } else if (row >= targetRow && row < ROWS_COUNT) {
          newCells[`${col}${row + 1}`] = sheet.cells[addr];
        }
      }
      sheet.cells = newCells;

      if (sheet.exportRange && sheet.exportRange.endRow >= targetRow) {
        sheet.exportRange.endRow = Math.min(ROWS_COUNT, sheet.exportRange.endRow + 1);
      }

      saveAppState();
      renderGrid();
      updatePreview();
      showToast(`Riadok bol vložený na pozíciu ${targetRow}.`);
    });
  }

  if (btnDeleteRow) {
    btnDeleteRow.addEventListener('click', () => {
      const sheet = getActiveSheet();
      const range = appState.selectionRange || { startRow: 1 };
      const targetRow = range.startRow;

      const newCells = {};
      for (const addr in sheet.cells) {
        const { col, row } = parseAddr(addr);
        if (row < targetRow) {
          newCells[addr] = sheet.cells[addr];
        } else if (row > targetRow) {
          newCells[`${col}${row - 1}`] = sheet.cells[addr];
        }
      }
      sheet.cells = newCells;

      if (sheet.exportRange && sheet.exportRange.endRow > sheet.exportRange.startRow) {
        if (targetRow <= sheet.exportRange.endRow) {
          sheet.exportRange.endRow--;
        }
      }

      saveAppState();
      renderGrid();
      updatePreview();
      showToast(`Riadok ${targetRow} bol zmazaný.`);
    });
  }

  if (btnInsertCol) {
    btnInsertCol.addEventListener('click', () => {
      const sheet = getActiveSheet();
      const range = appState.selectionRange || { startCol: 'A' };
      const targetColIdx = colIndex(range.startCol);
      if (targetColIdx >= COLUMNS.length - 1) return;

      const newCells = {};
      for (const addr in sheet.cells) {
        const { col, row } = parseAddr(addr);
        const cIdx = colIndex(col);
        if (cIdx < targetColIdx) {
          newCells[addr] = sheet.cells[addr];
        } else if (cIdx >= targetColIdx && cIdx < COLUMNS.length - 1) {
          newCells[`${COLUMNS[cIdx + 1]}${row}`] = sheet.cells[addr];
        }
      }
      sheet.cells = newCells;
      saveAppState();
      renderGrid();
      updatePreview();
      showToast(`Nový stĺpec bol vložený pred stĺpec ${range.startCol}.`);
    });
  }

  if (btnDeleteCol) {
    btnDeleteCol.addEventListener('click', () => {
      const sheet = getActiveSheet();
      const range = appState.selectionRange || { startCol: 'A' };
      const targetColIdx = colIndex(range.startCol);
      if (targetColIdx < 0) return;

      const newCells = {};
      for (const addr in sheet.cells) {
        const { col, row } = parseAddr(addr);
        const cIdx = colIndex(col);
        if (cIdx < targetColIdx) {
          newCells[addr] = sheet.cells[addr];
        } else if (cIdx > targetColIdx) {
          newCells[`${COLUMNS[cIdx - 1]}${row}`] = sheet.cells[addr];
        }
      }
      sheet.cells = newCells;
      saveAppState();
      renderGrid();
      updatePreview();
      showToast(`Stĺpec ${range.startCol} bol zmazaný.`);
    });
  }

  // --- Clear Tools ---
  function clearContent() {
    const sheet = getActiveSheet();
    getCellsInRange().forEach(item => {
      if (sheet.cells && sheet.cells[item.addr]) {
        sheet.cells[item.addr].v = '';
        const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
        if (td) td.innerText = '';
      }
    });
    formulaInput.value = '';
    saveAppState();
    updatePreview();
    showToast('Obsah označených buniek bol vymazaný.');
  }

  function clearFormat() {
    const sheet = getActiveSheet();
    getCellsInRange().forEach(item => {
      if (sheet.cells && sheet.cells[item.addr]) {
        const val = sheet.cells[item.addr].v || '';
        sheet.cells[item.addr] = {
          v: val,
          bold: false,
          italic: false,
          align: 'left',
          indent: 0,
          textColor: '#000000',
          fillColor: '',
          borderTop: { width: 0, style: 'none' },
          borderBottom: { width: 0, style: 'none' },
          borderLeft: { width: 0, style: 'none' },
          borderRight: { width: 0, style: 'none' }
        };
        const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
        if (td) applyCellStyles(td, sheet.cells[item.addr]);
      }
    });
    saveAppState();
    updatePreview();
    showToast('Formátovanie a orámovanie bolo odstránené.');
  }

  if (btnClearContent) btnClearContent.addEventListener('click', clearContent);
  if (btnClearFormat) btnClearFormat.addEventListener('click', clearFormat);

  // Keyboard Delete key
  window.addEventListener('keydown', (e) => {
    if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'SELECT')) {
      return;
    }
    if (e.key === 'Delete') {
      clearContent();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b') {
      e.preventDefault();
      boldBtn.click();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i') {
      e.preventDefault();
      italicBtn.click();
    }
  });

  // --- Move Table Across Spreadsheet Grid ---
  function moveTable(deltaRow, deltaCol) {
    const sheet = getActiveSheet();
    if (!sheet) return;
    const range = sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
    const sC = colIndex(range.startCol) + deltaCol;
    const eC = colIndex(range.endCol) + deltaCol;
    const sR = range.startRow + deltaRow;
    const eR = range.endRow + deltaRow;

    if (sC < 0 || eC >= COLUMNS.length || sR < 1 || eR > ROWS_COUNT) {
      showToast('Tabuľku nie je možné posunúť mimo rozsah hárku (A1 až J24).', false);
      return;
    }

    const oldSC = colIndex(range.startCol);
    const oldEC = colIndex(range.endCol);
    const oldSR = range.startRow;
    const oldER = range.endRow;

    // Collect all cells within the table
    const movingCells = [];
    for (let r = oldSR; r <= oldER; r++) {
      for (let c = oldSC; c <= oldEC; c++) {
        const fromAddr = `${COLUMNS[c]}${r}`;
        const toAddr = `${COLUMNS[c + deltaCol]}${r + deltaRow}`;
        const data = (sheet.cells && sheet.cells[fromAddr]) ? JSON.parse(JSON.stringify(sheet.cells[fromAddr])) : null;
        movingCells.push({ fromAddr, toAddr, data });
      }
    }

    // Delete old positions from sheet.cells
    movingCells.forEach(item => {
      if (sheet.cells) delete sheet.cells[item.fromAddr];
    });

    // Write to new positions
    if (!sheet.cells) sheet.cells = {};
    movingCells.forEach(item => {
      if (item.data) {
        sheet.cells[item.toAddr] = item.data;
      }
    });

    // Shift row heights if custom
    if (sheet.rowHeights) {
      const newRowHeights = {};
      for (const rStr in sheet.rowHeights) {
        const r = parseInt(rStr, 10);
        if (r >= oldSR && r <= oldER) {
          newRowHeights[r + deltaRow] = sheet.rowHeights[r];
        } else {
          newRowHeights[r] = sheet.rowHeights[r];
        }
      }
      sheet.rowHeights = newRowHeights;
    }

    // Update table exportRange
    sheet.exportRange = {
      startCol: COLUMNS[sC],
      startRow: sR,
      endCol: COLUMNS[eC],
      endRow: eR
    };

    // Update selection to cover the new moved table
    updateSelectionRange(COLUMNS[sC], sR, COLUMNS[eC], eR);

    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Tabuľka presunutá na ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}.`);
  }

  if (btnMoveUp) btnMoveUp.addEventListener('click', () => moveTable(-1, 0));
  if (btnMoveDown) btnMoveDown.addEventListener('click', () => moveTable(1, 0));
  if (btnMoveLeft) btnMoveLeft.addEventListener('click', () => moveTable(0, -1));
  if (btnMoveRight) btnMoveRight.addEventListener('click', () => moveTable(0, 1));

  // --- Button: Nastaviť výber ako oblasť pre export ---
  if (btnSetExportRange) {
    btnSetExportRange.addEventListener('click', () => {
      const sheet = getActiveSheet();
      sheet.exportRange = Object.assign({}, appState.selectionRange);
      if (statusBarRange) {
        statusBarRange.textContent = `Export: ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}`;
      }
      saveAppState();
      updatePreview();
      showToast(`Oblasť exportu nastavená na ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}.`);
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

      const colIdx = COLUMNS.indexOf(colName) + 1;
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

  // --- Interactive Splitter for Preview Drawer ---
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
      const deltaX = startMouseX - e.clientX;
      const newWidth = Math.min(1200, Math.max(320, startDrawerW + deltaX));
      averyPreviewDrawer.style.width = newWidth + 'px';
      if (previewZoomSelect && previewZoomSelect.value === 'fit') {
        applyFitZoom();
      }
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

  // Toggle Drawer Expand Button (Clean labels, no emojis)
  if (btnToggleDrawerExpand) {
    btnToggleDrawerExpand.addEventListener('click', () => {
      const isExpanded = averyPreviewDrawer.classList.toggle('expanded');
      btnToggleDrawerExpand.textContent = isExpanded ? 'Zúžiť panel' : 'Roztiahnuť panel';
      if (!isExpanded) {
        averyPreviewDrawer.style.width = '440px';
      } else {
        averyPreviewDrawer.style.width = '750px';
      }
      localStorage.setItem(DRAWER_WIDTH_KEY, averyPreviewDrawer.offsetWidth);
      if (previewZoomSelect && previewZoomSelect.value === 'fit') {
        applyFitZoom();
      }
    });
  }

  function applyFitZoom() {
    if (!previewViewport || !previewSvgBox) return;
    const svgEl = previewSvgBox.querySelector('svg');
    if (!svgEl) return;
    const availW = previewViewport.clientWidth - 48;
    const svgW = parseFloat(svgEl.getAttribute('width')) || 330;
    const fitScale = Math.min(2.0, Math.max(0.4, availW / svgW));
    previewSvgBox.style.transform = `scale(${fitScale.toFixed(2)})`;
  }

  // Zoom Selector
  if (previewZoomSelect && previewSvgBox) {
    previewZoomSelect.addEventListener('change', () => {
      if (previewZoomSelect.value === 'fit') {
        applyFitZoom();
      } else {
        const scale = parseFloat(previewZoomSelect.value);
        previewSvgBox.style.transform = `scale(${scale})`;
      }
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

  // --- SVG Live Preview Generator With Custom Pixel Borders ---
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
      let maxCellPt = 11;
      let maxLines = 1;
      for (let c = sC; c <= eC; c++) {
        const cell = (sheet.cells && sheet.cells[`${COLUMNS[c]}${r}`]) || {};
        const pt = cell.fontSize || 11;
        if (pt > maxCellPt) maxCellPt = pt;
        const lines = String(cell.v || '').split('\n').length;
        if (lines > maxLines) maxLines = lines;
      }
      const cellPx = Math.round(maxCellPt * 1.33);
      const neededH = Math.round(maxLines * (cellPx * 1.25) + 8);
      const baseH = rowHeights[r] || (r === 2 ? 38 : 24);
      const h = Math.max(baseH, neededH);
      rowPositions.push({ row: r, y: totalH, height: h });
      totalH += h;
    }

    const margin = 2;
    const svgW = totalW + margin * 2;
    const svgH = totalH + margin * 2;

    let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${svgW} ${svgH}" width="${svgW}" height="${svgH}">\n`;
    svg += `<defs>\n<style>\n`;
    svg += `  text { font-family: Calibri, Arial, sans-serif; font-size: 13px; dominant-baseline: middle; }\n`;
    svg += `  .bold { font-weight: bold; }\n`;
    svg += `  .italic { font-style: italic; }\n`;
    svg += `  .center { text-anchor: middle; }\n`;
    svg += `  .right { text-anchor: end; }\n`;
    svg += `</style>\n</defs>\n`;

    // Background base
    svg += `<rect x="${margin}" y="${margin}" width="${totalW}" height="${totalH}" fill="#ffffff" />\n`;

    // 1. Cell fills
    rowPositions.forEach(rp => {
      const r = rp.row;
      const y = rp.y + margin;
      const h = rp.height;
      activeCols.forEach(ac => {
        const addr = `${ac.col}${r}`;
        const cell = (sheet.cells && sheet.cells[addr]) || {};
        if (cell.fillColor && cell.fillColor !== '#ffffff' && cell.fillColor !== 'transparent') {
          const cellX = margin + ac.x;
          svg += `<rect x="${cellX}" y="${y}" width="${ac.width}" height="${h}" fill="${cell.fillColor}" />\n`;
        }
      });
    });

    // 2. Cell texts
    rowPositions.forEach(rp => {
      const r = rp.row;
      const y = rp.y + margin;
      const h = rp.height;

      activeCols.forEach(ac => {
        const addr = `${ac.col}${r}`;
        const cell = (sheet.cells && sheet.cells[addr]) || { v: '' };
        const cellX = margin + ac.x;
        const cellW = ac.width;

        const cellFontSize = cell.fontSize || 11;
        const cellFontPx = Math.round(cellFontSize * 1.33);
        const cellFontFam = cell.fontFamily || 'Calibri, Arial, sans-serif';
        const lineH = Math.round(cellFontPx * 1.25);

        const textY = y + h / 2;
        const indentX = cell.indent || 0;
        const textClass = (cell.bold ? 'bold ' : '') + (cell.italic ? 'italic ' : '');
        const textColor = cell.textColor || '#000000';

        const valLines = String(cell.v || '').split('\n');
        if (valLines.length === 1) {
          if (cell.align === 'center') {
            svg += `<text x="${cellX + cellW / 2}" y="${textY}" class="center ${textClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}">${escapeXml(cell.v)}</text>\n`;
          } else if (cell.align === 'right') {
            svg += `<text x="${cellX + cellW - 8}" y="${textY}" class="right ${textClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}">${escapeXml(cell.v)}</text>\n`;
          } else {
            svg += `<text x="${cellX + 8 + indentX}" y="${textY}" class="${textClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}">${escapeXml(cell.v)}</text>\n`;
          }
        } else {
          // Multiline cell (kJ / kcal)
          const startY = textY - ((valLines.length - 1) * (lineH / 2));
          const posX = cell.align === 'right' ? (cellX + cellW - 8) : (cell.align === 'center' ? (cellX + cellW / 2) : (cellX + 8 + indentX));
          const anchorClass = (cell.align === 'right' ? 'right ' : (cell.align === 'center' ? 'center ' : '')) + textClass;

          svg += `<text x="${posX}" y="${startY}" class="${anchorClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}">`;
          valLines.forEach((line, idx) => {
            svg += `<tspan x="${posX}" dy="${idx === 0 ? 0 : lineH}">${escapeXml(line)}</tspan>`;
          });
          svg += `</text>\n`;
        }
      });
    });

    // 3. Exact Cell Borders (Custom px thickness, style, color)
    function renderSvgBorder(x1, y1, x2, y2, b) {
      const norm = normalizeBorder(b);
      if (norm.width <= 0 || norm.style === 'none') return '';
      if (norm.style === 'double') {
        const offset = 1.5;
        const isHoriz = Math.abs(y1 - y2) < 0.1;
        if (isHoriz) {
          return `<line x1="${x1}" y1="${y1 - offset}" x2="${x2}" y2="${y2 - offset}" stroke="${norm.color}" stroke-width="1" />\n` +
                 `<line x1="${x1}" y1="${y1 + offset}" x2="${x2}" y2="${y2 + offset}" stroke="${norm.color}" stroke-width="1" />\n`;
        } else {
          return `<line x1="${x1 - offset}" y1="${y1}" x2="${x2 - offset}" y2="${y2}" stroke="${norm.color}" stroke-width="1" />\n` +
                 `<line x1="${x1 + offset}" y1="${y1}" x2="${x2 + offset}" y2="${y2}" stroke="${norm.color}" stroke-width="1" />\n`;
        }
      }
      let dash = '';
      if (norm.style === 'dashed') dash = ' stroke-dasharray="4,3"';
      if (norm.style === 'dotted') dash = ' stroke-dasharray="1,2"';
      return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${norm.color}" stroke-width="${norm.width}"${dash} />\n`;
    }

    rowPositions.forEach(rp => {
      const r = rp.row;
      const y = rp.y + margin;
      const h = rp.height;

      activeCols.forEach(ac => {
        const addr = `${ac.col}${r}`;
        const cell = (sheet.cells && sheet.cells[addr]) || {};
        const cellX = margin + ac.x;
        const cellW = ac.width;

        if (cell.borderBottom) {
          svg += renderSvgBorder(cellX, y + h, cellX + cellW, y + h, cell.borderBottom);
        }
        if (cell.borderTop) {
          svg += renderSvgBorder(cellX, y, cellX + cellW, y, cell.borderTop);
        }
        if (cell.borderLeft) {
          svg += renderSvgBorder(cellX, y, cellX, y + h, cell.borderLeft);
        }
        if (cell.borderRight) {
          svg += renderSvgBorder(cellX + cellW, y, cellX + cellW, y + h, cell.borderRight);
        }
      });
    });

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
    if (previewZoomSelect && previewZoomSelect.value === 'fit') {
      applyFitZoom();
    }
    if (statusBarRange && sheet.exportRange) {
      statusBarRange.textContent = `Export: ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}`;
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

  // 1. Copy to Clipboard (Clean, Ultra HD 600 DPI)
  btnCopyAvery.addEventListener('click', async () => {
    const sheet = getActiveSheet();
    const { svg, width, height } = generateTableSVG(sheet);

    try {
      showToast('Generujem obrázok tabuľky v Ultra HD kvalite...');
      const canvas = document.createElement('canvas');
      const scale = 8; // Ultra HD 600+ DPI maximum quality
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';

      const img = new Image();
      const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);

      img.onload = async () => {
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);

        canvas.toBlob(async (pngBlob) => {
          if (!pngBlob) return;
          try {
            if (navigator.clipboard && navigator.clipboard.write && typeof ClipboardItem !== 'undefined') {
              const item = new ClipboardItem({ 'image/png': pngBlob });
              await navigator.clipboard.write([item]);
              showToast('Skopírované v Ultra HD kvalite! Vložte cez Ctrl + V.', true);
            } else {
              throw new Error('ClipboardItem nepodporované');
            }
          } catch (clipErr) {
            // Fallback download if clipboard is restricted
            const a = document.createElement('a');
            a.href = URL.createObjectURL(pngBlob);
            a.download = sanitizeName(sheet.name) + '_UltraHD.png';
            a.click();
            showToast('Obrázok stiahnutý ako PNG v Ultra HD kvalite.', true);
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
    showToast(`Vektorový súbor "${sheet.name}.svg" bol stiahnutý.`);
  });

  // 3. Download PNG (Ultra HD 600 DPI)
  btnExportPng.addEventListener('click', () => {
    const sheet = getActiveSheet();
    const { svg, width, height } = generateTableSVG(sheet);
    const canvas = document.createElement('canvas');
    const scale = 8; // Ultra HD 600+ DPI
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
        a.download = sanitizeName(sheet.name) + '_UltraHD.png';
        a.click();
        showToast(`Tlačový obrázok "${sheet.name}_UltraHD.png" (Ultra HD 600 DPI) bol stiahnutý.`);
      }, 'image/png');
    };
    img.src = url;
  });

  // 4. Download JPG (Ultra HD 600 DPI)
  if (btnExportJpg) {
    btnExportJpg.addEventListener('click', () => {
      const sheet = getActiveSheet();
      const { svg, width, height } = generateTableSVG(sheet);
      const canvas = document.createElement('canvas');
      const scale = 8; // Ultra HD 600+ DPI
      canvas.width = width * scale;
      canvas.height = height * scale;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';

      const img = new Image();
      const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);

      img.onload = () => {
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);
        canvas.toBlob((jpgBlob) => {
          if (!jpgBlob) return;
          const a = document.createElement('a');
          a.href = URL.createObjectURL(jpgBlob);
          a.download = sanitizeName(sheet.name) + '_UltraHD.jpg';
          a.click();
          showToast(`Tlačový obrázok "${sheet.name}_UltraHD.jpg" (Ultra HD 600 DPI) bol stiahnutý.`);
        }, 'image/jpeg', 1.0);
      };
      img.src = url;
    });
  }

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
                borderBottom: key.endsWith('1') ? { width: 2, style: 'solid', color: '#000000' } : { width: 1, style: 'solid', color: '#000000' },
                borderTop: { width: 1, style: 'solid', color: '#000000' },
                borderLeft: key.startsWith('A') ? { width: 2, style: 'solid', color: '#000000' } : { width: 1, style: 'solid', color: '#000000' },
                borderRight: key.startsWith('B') ? { width: 2, style: 'solid', color: '#000000' } : { width: 1, style: 'solid', color: '#000000' }
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
          showToast(`Úspešne načítaných ${importedSheets.length} hárkov z Excelu.`);
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
