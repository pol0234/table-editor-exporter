// app.bundle.js - Standalone, zero-dependency Excel-authentic spreadsheet and Avery export engine
(function() {
  'use strict';

  const STORAGE_KEY = 'oco_excel_nutrition_sheets_v2';

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

  // --- DOM References ---
  const sheetTabsContainer = document.getElementById('sheetTabs');
  const addSheetBtn = document.getElementById('addSheetBtn');
  const gridTable = document.getElementById('gridTable');
  const formulaCellAddr = document.getElementById('formulaCellAddr');
  const formulaInput = document.getElementById('formulaInput');
  const previewSvgBox = document.getElementById('previewSvgBox');

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
  const btnCopyAvery = document.getElementById('btnCopyAvery');
  const btnExportSvg = document.getElementById('btnExportSvg');
  const btnExportPng = document.getElementById('btnExportPng');
  const btnUploadExcel = document.getElementById('btnUploadExcel');
  const excelFilePicker = document.getElementById('excelFilePicker');

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
        td.contentEditable = 'true';
        td.spellcheck = false;

        // Apply borders & formatting classes
        applyCellStyles(td, cellData);

        if (addr === appState.selectedCell) {
          td.classList.add('selected');
        }

        // Cell Selection
        td.addEventListener('focus', () => {
          selectCell(addr);
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
  }

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

  function selectCell(addr) {
    appState.selectedCell = addr;
    formulaCellAddr.textContent = addr;
    
    // Highlight in grid
    document.querySelectorAll('.grid-cell.selected').forEach(el => el.classList.remove('selected'));
    const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
    if (td) {
      td.classList.add('selected');
    }

    // Sync formula bar and ribbon
    const sheet = getActiveSheet();
    const cellData = (sheet.cells && sheet.cells[addr]) || { v: '' };
    formulaInput.value = cellData.v || '';

    // Ribbon states
    boldBtn.classList.toggle('active', !!cellData.bold);
    italicBtn.classList.toggle('active', !!cellData.italic);
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

  // --- Ribbon Button Handlers ---
  boldBtn.addEventListener('click', () => {
    const addr = appState.selectedCell;
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
    sheet.cells[addr].bold = !sheet.cells[addr].bold;
    boldBtn.classList.toggle('active', !!sheet.cells[addr].bold);
    
    const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
    if (td) applyCellStyles(td, sheet.cells[addr]);
    saveAppState();
    updatePreview();
  });

  italicBtn.addEventListener('click', () => {
    const addr = appState.selectedCell;
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
    sheet.cells[addr].italic = !sheet.cells[addr].italic;
    italicBtn.classList.toggle('active', !!sheet.cells[addr].italic);
    
    const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
    if (td) applyCellStyles(td, sheet.cells[addr]);
    saveAppState();
    updatePreview();
  });

  // Alignments
  alignLeftBtn.addEventListener('click', () => applyCellProp('align', 'left'));
  alignCenterBtn.addEventListener('click', () => applyCellProp('align', 'center'));
  alignRightBtn.addEventListener('click', () => applyCellProp('align', 'right'));

  // Indents
  indentIncBtn.addEventListener('click', () => {
    const addr = appState.selectedCell;
    const sheet = getActiveSheet();
    if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
    sheet.cells[addr].indent = (sheet.cells[addr].indent || 0) + 16;
    const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
    if (td) applyCellStyles(td, sheet.cells[addr]);
    saveAppState();
    updatePreview();
  });

  indentDecBtn.addEventListener('click', () => {
    const addr = appState.selectedCell;
    const sheet = getActiveSheet();
    if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
    sheet.cells[addr].indent = Math.max(0, (sheet.cells[addr].indent || 0) - 16);
    const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
    if (td) applyCellStyles(td, sheet.cells[addr]);
    saveAppState();
    updatePreview();
  });

  function applyCellProp(prop, value) {
    const addr = appState.selectedCell;
    const sheet = getActiveSheet();
    if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
    sheet.cells[addr][prop] = value;
    const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
    if (td) applyCellStyles(td, sheet.cells[addr]);
    saveAppState();
    updatePreview();
  }

  // --- Borders Dropdown Menu ---
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
      applyBorderToTable(type);
      borderMenu.classList.remove('open');
    });
  });

  function applyBorderToTable(type) {
    const sheet = getActiveSheet();
    const range = sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
    const cols = ['A', 'B'];

    for (let r = range.startRow; r <= range.endRow; r++) {
      cols.forEach((col, cIdx) => {
        const addr = col + r;
        if (!sheet.cells[addr]) sheet.cells[addr] = { v: '' };
        const c = sheet.cells[addr];

        if (type === 'all') {
          c.borderTop = 'thin';
          c.borderBottom = 'thin';
          c.borderLeft = 'thin';
          c.borderRight = 'thin';
        } else if (type === 'thick-box') {
          if (r === range.startRow) c.borderTop = 'thick';
          if (r === range.endRow) c.borderBottom = 'thick';
          if (cIdx === 0) c.borderLeft = 'thick';
          if (cIdx === cols.length - 1) c.borderRight = 'thick';
        } else if (type === 'thick-bottom') {
          if (r === range.endRow) c.borderBottom = 'thick';
        } else if (type === 'double-bottom') {
          if (r === range.endRow) c.borderBottom = 'double';
        } else if (type === 'header-thick') {
          if (r === range.startRow) c.borderBottom = 'thick';
        } else if (type === 'none') {
          c.borderTop = 'none';
          c.borderBottom = 'none';
          c.borderLeft = 'none';
          c.borderRight = 'none';
        }
      });
    }

    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Použité orámovanie: ${type}`);
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
      const colIndex = COLUMNS.indexOf(colName) + 1; // +1 for row number col
      const colEl = gridTable.querySelectorAll('col')[colIndex];
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

  // --- SVG Live Preview Generator ---
  function generateTableSVG(sheet) {
    const range = sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
    const colWidths = sheet.colWidths || {};
    const rowHeights = sheet.rowHeights || {};

    const wA = colWidths['A'] || 220;
    const wB = colWidths['B'] || 110;
    const totalW = wA + wB;

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

    // White background
    svg += `<rect x="${margin}" y="${margin}" width="${totalW}" height="${totalH}" fill="#ffffff" />\n`;

    // Render cells
    rowPositions.forEach(rp => {
      const r = rp.row;
      const y = rp.y + margin;
      const h = rp.height;

      const cellA = (sheet.cells && sheet.cells['A' + r]) || { v: '' };
      const cellB = (sheet.cells && sheet.cells['B' + r]) || { v: '' };

      // Text A
      const textAY = y + h / 2;
      const indentX = cellA.indent || 0;
      const textAX = margin + 8 + indentX;
      const classA = cellA.bold ? 'bold' : '';
      
      const valALines = String(cellA.v || '').split('\n');
      if (valALines.length === 1) {
        const anchorA = cellA.align === 'center' ? ' class="center ' + classA + '" x="' + (margin + wA / 2) + '"' : ' class="' + classA + '" x="' + textAX + '"';
        svg += `<text y="${textAY}"${anchorA}>${escapeXml(cellA.v)}</text>\n`;
      } else {
        svg += `<text x="${textAX}" y="${textAY - 6}" class="${classA}">`;
        valALines.forEach((line, idx) => {
          svg += `<tspan x="${textAX}" dy="${idx === 0 ? 0 : 14}">${escapeXml(line)}</tspan>`;
        });
        svg += `</text>\n`;
      }

      // Text B
      const textBY = y + h / 2;
      const classB = cellB.bold ? 'bold' : '';
      const valBLines = String(cellB.v || '').split('\n');

      if (valBLines.length === 1) {
        if (cellB.align === 'center') {
          svg += `<text x="${margin + wA + wB / 2}" y="${textBY}" class="center ${classB}">${escapeXml(cellB.v)}</text>\n`;
        } else {
          svg += `<text x="${margin + totalW - 8}" y="${textBY}" class="right ${classB}">${escapeXml(cellB.v)}</text>\n`;
        }
      } else {
        // Multi-line value (kJ / kcal)
        const startY = textBY - ((valBLines.length - 1) * 7);
        svg += `<text x="${margin + totalW - 8}" y="${startY}" class="right ${classB}">`;
        valBLines.forEach((line, idx) => {
          svg += `<tspan x="${margin + totalW - 8}" dy="${idx === 0 ? 0 : 15}">${escapeXml(line)}</tspan>`;
        });
        svg += `</text>\n`;
      }

      // Horizontal Borders
      if (cellA.borderBottom && cellA.borderBottom !== 'none') {
        const strokeW = cellA.borderBottom === 'thick' ? 2 : (cellA.borderBottom === 'double' ? 3 : 1);
        svg += `<line x1="${margin}" y1="${y + h}" x2="${margin + totalW}" y2="${y + h}" stroke="#000000" stroke-width="${strokeW}" />\n`;
      }
    });

    // Vertical Divider
    svg += `<line x1="${margin + wA}" y1="${margin}" x2="${margin + wA}" y2="${margin + totalH}" stroke="#000000" stroke-width="1" />\n`;

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
  selectCell('A1');

})();
