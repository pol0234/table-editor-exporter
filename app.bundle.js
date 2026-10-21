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

  function hasExplicitBorder(b) {
    if (!b) return false;
    const norm = normalizeBorder(b);
    return norm.width > 0 && norm.style !== 'none' && norm.color !== 'transparent';
  }

  function getBorderCss(b) {
    const norm = normalizeBorder(b);
    if (norm.width <= 0 || norm.style === 'none') {
      return '1px solid #d4d4d4'; // Default subtle Excel grid border
    }
    return `${norm.width}px ${norm.style} ${norm.color}`;
  }

  function resolveCellBorder(cIdx, r, side, sheet) {
    if (!sheet || !sheet.cells) sheet = getActiveSheet();
    if (!sheet || !sheet.cells) return '1px solid #d4d4d4';

    const colLetter = COLUMNS[cIdx];
    const addr = `${colLetter}${r}`;
    const cell = sheet.cells[addr];

    let ownBorder = null;
    let neighborBorder = null;
    let neighborCell = null;

    if (side === 'top') {
      ownBorder = cell ? cell.borderTop : null;
      if (r > 1) {
        neighborCell = sheet.cells[`${colLetter}${r - 1}`];
        neighborBorder = neighborCell ? neighborCell.borderBottom : null;
      }
    } else if (side === 'bottom') {
      ownBorder = cell ? cell.borderBottom : null;
      if (r < ROWS_COUNT) {
        neighborCell = sheet.cells[`${colLetter}${r + 1}`];
        neighborBorder = neighborCell ? neighborCell.borderTop : null;
      }
    } else if (side === 'left') {
      ownBorder = cell ? cell.borderLeft : null;
      if (cIdx > 0) {
        neighborCell = sheet.cells[`${COLUMNS[cIdx - 1]}${r}`];
        neighborBorder = neighborCell ? neighborCell.borderRight : null;
      }
    } else if (side === 'right') {
      ownBorder = cell ? cell.borderRight : null;
      if (cIdx < COLUMNS.length - 1) {
        neighborCell = sheet.cells[`${COLUMNS[cIdx + 1]}${r}`];
        neighborBorder = neighborCell ? neighborCell.borderLeft : null;
      }
    }

    const hasOwn = hasExplicitBorder(ownBorder);
    const hasNeigh = hasExplicitBorder(neighborBorder);

    // 1. Explicit border wins! Both sides of the edge agree symmetrically
    if (hasOwn && hasNeigh) {
      const nOwn = normalizeBorder(ownBorder);
      const nNeigh = normalizeBorder(neighborBorder);
      const chosen = (nOwn.width >= nNeigh.width) ? nOwn : nNeigh;
      return `${chosen.width}px ${chosen.style} ${chosen.color}`;
    }
    if (hasOwn) {
      const n = normalizeBorder(ownBorder);
      return `${n.width}px ${n.style} ${n.color}`;
    }
    if (hasNeigh) {
      const n = normalizeBorder(neighborBorder);
      return `${n.width}px ${n.style} ${n.color}`;
    }

    // 2. No explicit border: check fill behavior (White / Color Fill vs Default Gridlines)
    const cellFill = (cell && cell.fillColor && cell.fillColor !== 'transparent' && cell.fillColor !== 'none') ? cell.fillColor : null;
    const neighFill = (neighborCell && neighborCell.fillColor && neighborCell.fillColor !== 'transparent' && neighborCell.fillColor !== 'none') ? neighborCell.fillColor : null;

    if (cellFill || neighFill) {
      const fillToUse = cellFill || neighFill;
      // Seamless solid surface: hide gray gridline by matching the fill color
      return `1px solid ${fillToUse}`;
    }

    // Default subtle Excel gridline
    return '1px solid #d4d4d4';
  }

  function resolveBorderObj(cIdx, r, side, sheet) {
    if (!sheet || !sheet.cells) return null;
    const colLetter = COLUMNS[cIdx];
    const addr = `${colLetter}${r}`;
    const cell = sheet.cells[addr];

    let ownBorder = null;
    let neighborBorder = null;

    if (side === 'top') {
      ownBorder = cell ? cell.borderTop : null;
      if (r > 1) {
        const n = sheet.cells[`${colLetter}${r - 1}`];
        neighborBorder = n ? n.borderBottom : null;
      }
    } else if (side === 'bottom') {
      ownBorder = cell ? cell.borderBottom : null;
      if (r < ROWS_COUNT) {
        const n = sheet.cells[`${colLetter}${r + 1}`];
        neighborBorder = n ? n.borderTop : null;
      }
    } else if (side === 'left') {
      ownBorder = cell ? cell.borderLeft : null;
      if (cIdx > 0) {
        const n = sheet.cells[`${COLUMNS[cIdx - 1]}${r}`];
        neighborBorder = n ? n.borderRight : null;
      }
    } else if (side === 'right') {
      ownBorder = cell ? cell.borderRight : null;
      if (cIdx < COLUMNS.length - 1) {
        const n = sheet.cells[`${COLUMNS[cIdx + 1]}${r}`];
        neighborBorder = n ? n.borderLeft : null;
      }
    }

    if (hasExplicitBorder(ownBorder) && hasExplicitBorder(neighborBorder)) {
      const nOwn = normalizeBorder(ownBorder);
      const nNeigh = normalizeBorder(neighborBorder);
      return (nOwn.width >= nNeigh.width) ? nOwn : nNeigh;
    }
    if (hasExplicitBorder(ownBorder)) return normalizeBorder(ownBorder);
    if (hasExplicitBorder(neighborBorder)) return normalizeBorder(neighborBorder);
    return null;
  }

  function refreshAllGridBorders(sheet) {
    if (!sheet) sheet = getActiveSheet();
    if (!gridTable || !sheet) return;
    for (let r = 1; r <= ROWS_COUNT; r++) {
      for (let c = 0; c < COLUMNS.length; c++) {
        const col = COLUMNS[c];
        const addr = `${col}${r}`;
        const td = gridTable.querySelector(`td[data-addr="${addr}"]`);
        if (td) {
          td.style.borderTop = resolveCellBorder(c, r, 'top', sheet);
          td.style.borderBottom = resolveCellBorder(c, r, 'bottom', sheet);
          td.style.borderLeft = resolveCellBorder(c, r, 'left', sheet);
          td.style.borderRight = resolveCellBorder(c, r, 'right', sheet);
        }
      }
    }
  }

  // --- Initial Default Products Matching User's Excel Sheets ---
  function createDefaultSheets() {
    const thickBox = { width: 2, style: 'solid', color: '#000000' };
    const thinInner = { width: 1, style: 'solid', color: '#000000' };

    const sheets = [
      {
        id: 's_papaya',
        name: 'Soft-Dried-Papaya',
        colWidths: { A: 270, B: 120, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B1': { v: '100g', bold: true, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B2': { v: '1372kJ /\n328kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderTop: thickBox, borderLeft: thickBox },
          'B3': { v: '0,12g', bold: false, align: 'right', borderTop: thickBox, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B4': { v: '0,08g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderLeft: thickBox },
          'B5': { v: '80,9g', bold: false, align: 'right', borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B6': { v: '70,4g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderLeft: thickBox },
          'B7': { v: '1,54g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderLeft: thickBox },
          'B8': { v: '0,23g', bold: false, align: 'right', borderBottom: thickBox, borderRight: thickBox }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_mango',
        name: 'Dried-Mango',
        colWidths: { A: 270, B: 120, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B1': { v: '100g', bold: true, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B2': { v: '1420kJ /\n339kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderTop: thickBox, borderLeft: thickBox },
          'B3': { v: '0,50g', bold: false, align: 'right', borderTop: thickBox, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B4': { v: '0,10g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderLeft: thickBox },
          'B5': { v: '82,0g', bold: false, align: 'right', borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B6': { v: '65,0g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderLeft: thickBox },
          'B7': { v: '1,80g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderLeft: thickBox },
          'B8': { v: '0,05g', bold: false, align: 'right', borderBottom: thickBox, borderRight: thickBox }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_passion',
        name: 'Soft-Dried-PassionFruit',
        colWidths: { A: 270, B: 120, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B1': { v: '100g', bold: true, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B2': { v: '1390kJ /\n332kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderTop: thickBox, borderLeft: thickBox },
          'B3': { v: '0,20g', bold: false, align: 'right', borderTop: thickBox, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B4': { v: '0,05g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderLeft: thickBox },
          'B5': { v: '78,0g', bold: false, align: 'right', borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B6': { v: '68,0g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderLeft: thickBox },
          'B7': { v: '2,10g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderLeft: thickBox },
          'B8': { v: '0,15g', bold: false, align: 'right', borderBottom: thickBox, borderRight: thickBox }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      },
      {
        id: 's_mixed',
        name: 'Mixed-Dried-TropicalFruits',
        colWidths: { A: 270, B: 120, C: 60, D: 60, E: 60 },
        rowHeights: { 1: 26, 2: 38, 3: 24, 4: 24, 5: 24, 6: 24, 7: 24, 8: 24 },
        cells: {
          'A1': { v: 'Výživové údaje na', bold: true, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B1': { v: '100g', bold: true, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A2': { v: 'Energetická hodnota', bold: false, align: 'left', borderBottom: thickBox, borderTop: thickBox, borderLeft: thickBox },
          'B2': { v: '1410kJ /\n337kcal', bold: false, align: 'right', borderBottom: thickBox, borderTop: thickBox, borderRight: thickBox },
          'A3': { v: 'Tuky', bold: false, align: 'left', borderTop: thickBox, borderLeft: thickBox },
          'B3': { v: '0,35g', bold: false, align: 'right', borderTop: thickBox, borderRight: thickBox },
          'A4': { v: 'z toho nasýtené mastné kyseliny', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B4': { v: '0,07g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A5': { v: 'Sacharidy', bold: false, align: 'left', borderLeft: thickBox },
          'B5': { v: '81,5g', bold: false, align: 'right', borderRight: thickBox },
          'A6': { v: 'z toho cukry', bold: false, align: 'left', indent: 18, borderBottom: thinInner, borderLeft: thickBox },
          'B6': { v: '69,0g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A7': { v: 'Bielkoviny', bold: false, align: 'left', borderBottom: thinInner, borderLeft: thickBox },
          'B7': { v: '1,65g', bold: false, align: 'right', borderBottom: thinInner, borderRight: thickBox },
          'A8': { v: 'Soľ', bold: false, align: 'left', borderBottom: thickBox, borderLeft: thickBox },
          'B8': { v: '0,18g', bold: false, align: 'right', borderBottom: thickBox, borderRight: thickBox }
        },
        exportRange: { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 }
      }
    ];

    return sheets;
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
        // Ensure colWidths.A is spacious enough (at least 270px) so nutrient text doesn't wrap
        appState.sheets.forEach(sh => {
          if (!sh.colWidths) sh.colWidths = {};
          if (!sh.colWidths.A || sh.colWidths.A < 260) {
            sh.colWidths.A = 270;
          }
          // Remove legacy forced white fill so authentic gridlines show
          if (sh.cells) {
            for (const addr in sh.cells) {
              if (sh.cells[addr].fillColor === '#ffffff' && !sh.cells[addr].userFillExplicit) {
                delete sh.cells[addr].fillColor;
              }
            }
          }
        });
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

  // --- Undo / Redo History System ---
  const MAX_HISTORY = 40;
  let undoStack = [];
  let redoStack = [];

  function recordHistory() {
    try {
      const snapshot = JSON.stringify({
        sheets: appState.sheets,
        activeSheetId: appState.activeSheetId,
        selectionRange: appState.selectionRange
      });
      undoStack.push(snapshot);
      if (undoStack.length > MAX_HISTORY) undoStack.shift();
      redoStack = []; // Clear redo stack on new action
      updateUndoRedoButtons();
    } catch (e) {
      console.warn('History record error', e);
    }
  }

  function undo() {
    if (undoStack.length === 0) {
      showToast('Už nie je k dispozícii žiadny krok späť.', false);
      return;
    }
    const current = JSON.stringify({
      sheets: appState.sheets,
      activeSheetId: appState.activeSheetId,
      selectionRange: appState.selectionRange
    });
    redoStack.push(current);

    const prev = JSON.parse(undoStack.pop());
    appState.sheets = prev.sheets;
    appState.activeSheetId = prev.activeSheetId;
    appState.selectionRange = prev.selectionRange;
    saveAppState();
    renderSheetTabs();
    renderGrid();
    updatePreview();
    updateUndoRedoButtons();
    showToast('Krok späť vykonaný.');
  }

  function redo() {
    if (redoStack.length === 0) {
      showToast('Už nie je k dispozícii žiadny krok dopredu.', false);
      return;
    }
    const current = JSON.stringify({
      sheets: appState.sheets,
      activeSheetId: appState.activeSheetId,
      selectionRange: appState.selectionRange
    });
    undoStack.push(current);

    const next = JSON.parse(redoStack.pop());
    appState.sheets = next.sheets;
    appState.activeSheetId = next.activeSheetId;
    appState.selectionRange = next.selectionRange;
    saveAppState();
    renderSheetTabs();
    renderGrid();
    updatePreview();
    updateUndoRedoButtons();
    showToast('Krok dopredu vykonaný.');
  }

  function updateUndoRedoButtons() {
    const canUndo = undoStack.length > 0;
    const canRedo = redoStack.length > 0;
    if (btnUndo) btnUndo.disabled = !canUndo;
    if (btnRedo) btnRedo.disabled = !canRedo;
    if (btnTitleUndo) btnTitleUndo.disabled = !canUndo;
    if (btnTitleRedo) btnTitleRedo.disabled = !canRedo;
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
  const btnFillWhiteQuick = document.getElementById('btnFillWhiteQuick');
  const btnFillNoneQuick = document.getElementById('btnFillNoneQuick');

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
  const btnUndo = document.getElementById('btnUndo');
  const btnRedo = document.getElementById('btnRedo');
  const btnTitleUndo = document.getElementById('btnTitleUndo');
  const btnTitleRedo = document.getElementById('btnTitleRedo');

  // Export & Files
  const btnSetExportRange = document.getElementById('btnSetExportRange');
  const btnCopyAvery = document.getElementById('btnCopyAvery');
  const btnExportSvg = document.getElementById('btnExportSvg');
  const btnExportPng = document.getElementById('btnExportPng');
  const btnExportJpg = document.getElementById('btnExportJpg');
  const btnUploadExcel = document.getElementById('btnUploadExcel');
  const excelFilePicker = document.getElementById('excelFilePicker');

  // Workspace & Drawer
  const spreadsheetArea = document.getElementById('spreadsheetArea');

  // Drawer & Splitter & Toggle
  const averyPreviewDrawer = document.getElementById('averyPreviewDrawer');
  const workspaceSplitter = document.getElementById('workspaceSplitter');
  const btnTogglePreviewDrawer = document.getElementById('btnTogglePreviewDrawer');
  const btnCloseDrawer = document.getElementById('btnCloseDrawer');
  const btnToggleDrawerExpand = document.getElementById('btnToggleDrawerExpand');
  const previewZoomSelect = document.getElementById('previewZoomSelect');
  const previewViewport = document.getElementById('previewViewport');
  const btnBgWhite = document.getElementById('btnBgWhite');
  const btnBgCheck = document.getElementById('btnBgCheck');
  const btnTogglePrintGuides = document.getElementById('btnTogglePrintGuides');

  // Context Menu Elements (Mini Toolbar & Borders)
  const excelContextMenu = document.getElementById('excelContextMenu');
  const ctxFontFamily = document.getElementById('ctxFontFamily');
  const ctxFontSize = document.getElementById('ctxFontSize');
  const ctxFontInc = document.getElementById('ctxFontInc');
  const ctxFontDec = document.getElementById('ctxFontDec');
  const ctxBold = document.getElementById('ctxBold');
  const ctxItalic = document.getElementById('ctxItalic');
  const ctxUnderline = document.getElementById('ctxUnderline');
  const ctxAlignLeft = document.getElementById('ctxAlignLeft');
  const ctxAlignCenter = document.getElementById('ctxAlignCenter');
  const ctxAlignRight = document.getElementById('ctxAlignRight');

  const ctxBorderAll = document.getElementById('ctxBorderAll');
  const ctxBorderOuter = document.getElementById('ctxBorderOuter');
  const ctxBorderThickOuter = document.getElementById('ctxBorderThickOuter');
  const ctxBorderNone = document.getElementById('ctxBorderNone');
  const ctxBorderWidth = document.getElementById('ctxBorderWidth');
  const ctxBorderColor = document.getElementById('ctxBorderColor');
  const ctxFillWhite = document.getElementById('ctxFillWhite');
  const ctxFillNone = document.getElementById('ctxFillNone');
  const ctxFillYellow = document.getElementById('ctxFillYellow');
  const ctxFillColor = document.getElementById('ctxFillColor');

  const ctxShowInPreview = document.getElementById('ctxShowInPreview');
  const ctxCopyTable = document.getElementById('ctxCopyTable');
  const ctxPasteExcel = document.getElementById('ctxPasteExcel');
  const ctxSetExportRange = document.getElementById('ctxSetExportRange');
  const ctxClearContent = document.getElementById('ctxClearContent');
  const ctxClearFormat = document.getElementById('ctxClearFormat');
  const ctxAutoFit = document.getElementById('ctxAutoFit');
  const ctxInsertRow = document.getElementById('ctxInsertRow');
  const ctxDeleteRow = document.getElementById('ctxDeleteRow');
  const ctxInsertCol = document.getElementById('ctxInsertCol');
  const ctxDeleteCol = document.getElementById('ctxDeleteCol');

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
      tab.dataset.id = sheet.id;
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

  let cellEditHistoryRecorded = false;

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
    cornerTh.title = 'Vybrať celú tabuľku';
    cornerTh.addEventListener('click', () => {
      updateSelectionRange('A', 1, COLUMNS[COLUMNS.length - 1], ROWS_COUNT);
    });
    headerRow.appendChild(cornerTh);

    COLUMNS.forEach(col => {
      const th = document.createElement('th');
      th.className = 'grid-header-col';
      th.dataset.col = col;
      const w = (sheet.colWidths && sheet.colWidths[col]) || (col === 'A' ? 220 : (col === 'B' ? 110 : 70));
      th.style.width = w + 'px';
      th.innerHTML = `
        <span>${col}</span>
        <div class="col-resize-handle" data-col="${col}" title="Potiahnutím myšou zmeníte šírku stĺpca ${col}"></div>
      `;
      th.addEventListener('mousedown', (e) => {
        if (e.target.classList.contains('col-resize-handle')) return;
        updateSelectionRange(col, 1, col, ROWS_COUNT);
      });
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
      th.dataset.row = r;
      th.style.height = rowHeight + 'px';
      th.style.maxHeight = rowHeight + 'px';
      th.style.lineHeight = `${Math.max(4, rowHeight - 2)}px`;
      th.innerHTML = `
        <span>${r}</span>
        <div class="row-resize-handle" data-row="${r}" title="Potiahnutím myšou zmeníte výšku riadku ${r}"></div>
      `;
      th.addEventListener('mousedown', (e) => {
        if (e.target.classList.contains('row-resize-handle')) return;
        updateSelectionRange('A', r, COLUMNS[COLUMNS.length - 1], r);
      });
      tr.appendChild(th);

      // Data Cells
      COLUMNS.forEach((col, cIdx) => {
        const addr = col + r;
        const cellData = (sheet.cells && sheet.cells[addr]) || { v: '' };
        const td = document.createElement('td');
        td.className = 'grid-cell';
        td.dataset.addr = addr;
        td.dataset.col = col;
        td.dataset.row = r;
        td.style.height = rowHeight + 'px';
        td.style.maxHeight = rowHeight + 'px';
        td.style.lineHeight = `${Math.max(4, rowHeight - 2)}px`;
        td.contentEditable = 'true';
        td.spellcheck = false;

        // Apply borders & formatting
        applyCellStyles(td, cellData, cIdx, r, sheet);

        // Mouse Drag Selection / Table Move Handlers
        td.addEventListener('mousedown', (e) => {
          if (e.button === 2) {
            // Right-click: Keep active multi-cell selection if clicking inside it
            const cur = appState.selectionRange;
            if (cur) {
              const sC = Math.min(colIndex(cur.startCol), colIndex(cur.endCol));
              const eC = Math.max(colIndex(cur.startCol), colIndex(cur.endCol));
              const sR = Math.min(cur.startRow, cur.endRow);
              const eR = Math.max(cur.startRow, cur.endRow);
              const cIdx = colIndex(col);
              const isInside = (cIdx >= sC && cIdx <= eC && r >= sR && r <= eR);
              if (isInside) {
                return; // Keep selection!
              }
            }
            updateSelectionRange(col, r, col, r);
            return;
          }
          if (e.button !== 0) return; // Left click only

          const rect = td.getBoundingClientRect();
          const isTopBorder = td.classList.contains('range-top') && Math.abs(e.clientY - rect.top) <= 6;
          const isBottomBorder = td.classList.contains('range-bottom') && Math.abs(e.clientY - rect.bottom) <= 6;
          const isLeftBorder = td.classList.contains('range-left') && Math.abs(e.clientX - rect.left) <= 6;
          const isRightBorder = td.classList.contains('range-right') && Math.abs(e.clientX - rect.right) <= 6;
          const isCornerHandle = td.classList.contains('range-corner-handle') && (rect.right - e.clientX <= 8) && (rect.bottom - e.clientY <= 8);

          if ((isTopBorder || isBottomBorder || isLeftBorder || isRightBorder) && !isCornerHandle) {
            e.preventDefault();
            startMouseTableDrag(col, r, e);
            return;
          }

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
          cellEditHistoryRecorded = false;
          if (isDraggingSelection) return;
          const cur = appState.selectionRange;
          if (cur) {
            const sC = Math.min(colIndex(cur.startCol), colIndex(cur.endCol));
            const eC = Math.max(colIndex(cur.startCol), colIndex(cur.endCol));
            const sR = Math.min(cur.startRow, cur.endRow);
            const eR = Math.max(cur.startRow, cur.endRow);
            const cIdx = colIndex(col);
            const isInside = (cIdx >= sC && cIdx <= eC && r >= sR && r <= eR);
            const isMulti = (sC !== eC || sR !== eR);
            if (isInside && isMulti) {
              return; // Keep active multi-cell selection range intact!
            }
          }
          updateSelectionRange(col, r, col, r);
        });

        // Cell Typing / Editing
        td.addEventListener('input', () => {
          if (!cellEditHistoryRecorded) {
            recordHistory();
            cellEditHistoryRecorded = true;
          }
          const text = td.innerText;
          setCellValue(addr, text);
          formulaInput.value = text;
          updatePreview();
        });

        td.addEventListener('blur', () => {
          cellEditHistoryRecorded = false;
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

    setupGridResize();
    refreshAllGridBorders(sheet);
    refreshSelectionVisuals();
  }

  // Window mouseup ends drag selection
  window.addEventListener('mouseup', () => {
    if (isDraggingSelection) {
      isDraggingSelection = false;
    }
  });

  function applyCellStyles(td, cellData, cIdx, r, sheet) {
    if (!sheet) sheet = getActiveSheet();
    td.innerText = cellData.v || '';
    td.style.fontWeight = cellData.bold ? 'bold' : 'normal';
    td.style.fontStyle = cellData.italic ? 'italic' : 'normal';
    td.style.textDecoration = cellData.underline ? 'underline' : 'none';
    td.style.textAlign = cellData.align || 'left';
    td.style.color = cellData.textColor || '#000000';
    td.style.backgroundColor = (cellData.fillColor && cellData.fillColor !== 'transparent' && cellData.fillColor !== 'none') ? cellData.fillColor : '';
    td.style.paddingLeft = (cellData.indent ? (8 + cellData.indent) : 6) + 'px';
    const fontSize = cellData.fontSize || 11;
    td.style.fontSize = fontSize + 'pt';
    if (cellData.fontFamily) {
      td.style.fontFamily = cellData.fontFamily;
    } else {
      td.style.fontFamily = '';
    }

    // Dynamic Symmetrical Borders & Gridlines
    if (cIdx !== undefined && r !== undefined) {
      td.style.borderTop = resolveCellBorder(cIdx, r, 'top', sheet);
      td.style.borderBottom = resolveCellBorder(cIdx, r, 'bottom', sheet);
      td.style.borderLeft = resolveCellBorder(cIdx, r, 'left', sheet);
      td.style.borderRight = resolveCellBorder(cIdx, r, 'right', sheet);
    } else if (td.dataset && td.dataset.addr) {
      const parsed = parseAddr(td.dataset.addr);
      const colI = colIndex(parsed.col);
      td.style.borderTop = resolveCellBorder(colI, parsed.row, 'top', sheet);
      td.style.borderBottom = resolveCellBorder(colI, parsed.row, 'bottom', sheet);
      td.style.borderLeft = resolveCellBorder(colI, parsed.row, 'left', sheet);
      td.style.borderRight = resolveCellBorder(colI, parsed.row, 'right', sheet);
    } else {
      td.style.borderBottom = getBorderCss(cellData.borderBottom);
      td.style.borderTop = getBorderCss(cellData.borderTop);
      td.style.borderLeft = getBorderCss(cellData.borderLeft);
      td.style.borderRight = getBorderCss(cellData.borderRight);
    }
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

    // Context menu mini-toolbar sync
    if (ctxBold) ctxBold.classList.toggle('active', !!cellData.bold);
    if (ctxItalic) ctxItalic.classList.toggle('active', !!cellData.italic);
    if (ctxUnderline) ctxUnderline.classList.toggle('active', !!cellData.underline);
    if (ctxFontSize) ctxFontSize.value = String(cellData.fontSize || 11);
    if (ctxFontFamily && cellData.fontFamily) ctxFontFamily.value = cellData.fontFamily;
    if (ctxAlignLeft) ctxAlignLeft.classList.toggle('active', (cellData.align || 'left') === 'left');
    if (ctxAlignCenter) ctxAlignCenter.classList.toggle('active', cellData.align === 'center');
    if (ctxAlignRight) ctxAlignRight.classList.toggle('active', cellData.align === 'right');

    refreshSelectionVisuals();
  }

  function refreshSelectionVisuals() {
    const range = appState.selectionRange || { startCol: 'A', startRow: 1, endCol: 'A', endRow: 1 };
    const sC = colIndex(range.startCol);
    const eC = colIndex(range.endCol);
    const sR = range.startRow;
    const eR = range.endRow;

    const sheet = getActiveSheet();
    const expRange = getEffectiveExportRange(sheet);
    const expSC = colIndex(expRange.startCol);
    const expEC = colIndex(expRange.endCol);
    const expSR = expRange.startRow;
    const expER = expRange.endRow;

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

      const inExpRange = (c >= expSC && c <= expEC && r >= expSR && r <= expER);
      td.classList.toggle('in-export-range', inExpRange);
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
    recordHistory();
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
    recordHistory();
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

  function toggleUnderline() {
    recordHistory();
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const cells = getCellsInRange();
    const anchorData = sheet.cells[appState.selectedCell] || {};
    const targetUnderline = !anchorData.underline;

    cells.forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr].underline = targetUnderline;
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) applyCellStyles(td, sheet.cells[item.addr]);
    });

    if (ctxUnderline) ctxUnderline.classList.toggle('active', targetUnderline);
    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
  }

  // Font Size & Font Family
  const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48];

  function applyFontSize(size) {
    recordHistory();
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
    if (ctxFontSize) ctxFontSize.value = String(size);
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
      recordHistory();
      const fam = fontSelect.value;
      const sheet = getActiveSheet();
      if (!sheet.cells) sheet.cells = {};
      getCellsInRange().forEach(item => {
        if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
        sheet.cells[item.addr].fontFamily = fam;
        const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
        if (td) applyCellStyles(td, sheet.cells[item.addr]);
      });
      if (ctxFontFamily) ctxFontFamily.value = fam;
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
      recordHistory();
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

  function applyFillColor(color) {
    recordHistory();
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const cells = getCellsInRange();
    cells.forEach(item => {
      if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
      sheet.cells[item.addr].fillColor = color;
      const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
      if (td) {
        td.style.backgroundColor = (color && color !== 'transparent' && color !== 'none') ? color : '';
      }
    });
    if (fillColorBar) {
      fillColorBar.setAttribute('fill', color || '#ffffff');
    }
    refreshAllGridBorders(sheet);
    saveAppState();
    updatePreview();
    showToast(color ? (color.toLowerCase() === '#ffffff' ? 'Aplikovaná biela výplň (mriežka skrytá).' : `Aplikovaná farba výplne (${color}).`) : 'Výplň odstránená (mriežka obnovená).');
  }

  if (ribbonFillColor) {
    ribbonFillColor.addEventListener('input', () => {
      applyFillColor(ribbonFillColor.value);
    });
  }

  if (btnFillWhiteQuick) {
    btnFillWhiteQuick.addEventListener('click', () => {
      applyFillColor('#ffffff');
    });
  }

  if (btnFillNoneQuick) {
    btnFillNoneQuick.addEventListener('click', () => {
      applyFillColor('');
    });
  }

  if (ctxFillWhite) {
    ctxFillWhite.addEventListener('click', () => {
      applyFillColor('#ffffff');
      if (excelContextMenu) excelContextMenu.style.display = 'none';
    });
  }

  if (ctxFillNone) {
    ctxFillNone.addEventListener('click', () => {
      applyFillColor('');
      if (excelContextMenu) excelContextMenu.style.display = 'none';
    });
  }

  if (ctxFillYellow) {
    ctxFillYellow.addEventListener('click', () => {
      applyFillColor('#fff2cc');
      if (excelContextMenu) excelContextMenu.style.display = 'none';
    });
  }

  if (ctxFillColor) {
    ctxFillColor.addEventListener('input', () => {
      applyFillColor(ctxFillColor.value);
      if (excelContextMenu) excelContextMenu.style.display = 'none';
    });
  }

  // Alignments for Selection
  alignLeftBtn.addEventListener('click', () => applyRangeProp('align', 'left'));
  alignCenterBtn.addEventListener('click', () => applyRangeProp('align', 'center'));
  alignRightBtn.addEventListener('click', () => applyRangeProp('align', 'right'));

  // Indents for Selection
  indentIncBtn.addEventListener('click', () => {
    recordHistory();
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
    recordHistory();
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
    recordHistory();
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

  // --- Borders Dropdown Menu Applied to Selected Range (Floating Fixed) ---
  if (borderMenu && borderMenu.parentElement !== document.body) {
    document.body.appendChild(borderMenu);
  }

  function positionBorderMenu() {
    if (!borderMenu || !borderDropdownBtn) return;
    const rect = borderDropdownBtn.getBoundingClientRect();
    borderMenu.style.position = 'fixed';
    borderMenu.style.top = `${rect.bottom + 4}px`;
    borderMenu.style.left = `${Math.min(window.innerWidth - 260, Math.max(10, rect.left))}px`;
    borderMenu.style.zIndex = '999999';
  }

  if (borderDropdownBtn && borderMenu) {
    borderDropdownBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const isOpen = borderMenu.classList.contains('open');
      if (!isOpen) {
        positionBorderMenu();
        borderMenu.classList.add('open');
      } else {
        borderMenu.classList.remove('open');
      }
    });

    document.addEventListener('click', (e) => {
      if (!borderMenu.contains(e.target) && e.target !== borderDropdownBtn && !borderDropdownBtn.contains(e.target)) {
        borderMenu.classList.remove('open');
      }
    });

    window.addEventListener('resize', () => {
      if (borderMenu.classList.contains('open')) {
        positionBorderMenu();
      }
    });
  }

  document.querySelectorAll('.border-menu-item').forEach(item => {
    item.addEventListener('click', (e) => {
      e.stopPropagation();
      const type = item.dataset.border;
      applyBorderToSelection(type);
      borderMenu.classList.remove('open');
    });
  });

  function applyBorderToSelection(type, overrideWidth = null, overrideColor = null) {
    recordHistory();
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const range = appState.selectionRange || sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
    const sC = colIndex(range.startCol);
    const eC = colIndex(range.endCol);
    const sR = range.startRow;
    const eR = range.endRow;

    const borderWidth = overrideWidth !== null ? overrideWidth : (parseInt(borderWidthInput ? borderWidthInput.value : '1', 10) || 1);
    const borderStyle = borderStyleSelect ? borderStyleSelect.value : 'solid';
    const borderColor = overrideColor !== null ? overrideColor : (borderColorPicker ? borderColorPicker.value : '#000000');

    const userBorder = { width: borderWidth, style: borderStyle, color: borderColor };
    const thickBorder = { width: Math.max(3, borderWidth), style: 'solid', color: borderColor };
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
        } else if (type === 'top-bottom') {
          if (r === sR) cell.borderTop = Object.assign({}, userBorder);
          if (r === eR) cell.borderBottom = Object.assign({}, userBorder);
        } else if (type === 'top-thick-bottom') {
          if (r === sR) cell.borderTop = Object.assign({}, userBorder);
          if (r === eR) cell.borderBottom = Object.assign({}, thickBorder);
        } else if (type === 'top-double-bottom') {
          if (r === sR) cell.borderTop = Object.assign({}, userBorder);
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

    refreshAllGridBorders(sheet);
    refreshSelectionVisuals();
    saveAppState();
    updatePreview();
    showToast(`Orámovanie bolo aplikované na vybrané bunky (${borderWidth}px, ${borderStyle}).`);
  }

  // --- Row & Column Manipulation ---
  function insertRowAt(rowToInsert) {
    recordHistory();
    const sheet = getActiveSheet();
    const range = appState.selectionRange || { startRow: 1 };
    const targetRow = (typeof rowToInsert === 'number') ? rowToInsert : range.startRow;
    if (targetRow < 1 || targetRow >= ROWS_COUNT) return;

    // Shift cells down
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

    // Shift rowHeights down
    if (sheet.rowHeights) {
      const newHeights = {};
      for (let r = 1; r <= ROWS_COUNT; r++) {
        if (r < targetRow) {
          if (sheet.rowHeights[r]) newHeights[r] = sheet.rowHeights[r];
        } else if (r === targetRow) {
          newHeights[r] = 24;
        } else {
          if (sheet.rowHeights[r - 1]) newHeights[r] = sheet.rowHeights[r - 1];
        }
      }
      sheet.rowHeights = newHeights;
    }

    // Expand exportRange
    if (sheet.exportRange && sheet.exportRange.endRow >= targetRow) {
      sheet.exportRange.endRow = Math.min(ROWS_COUNT, sheet.exportRange.endRow + 1);
    }

    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Riadok bol vložený na pozíciu ${targetRow}.`);
  }

  function deleteRowAt(rowToDelete) {
    recordHistory();
    const sheet = getActiveSheet();
    const range = appState.selectionRange || { startRow: 1 };
    const targetRow = (typeof rowToDelete === 'number') ? rowToDelete : range.startRow;
    if (targetRow < 1) return;

    // Shift cells up
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

    // Shift rowHeights up
    if (sheet.rowHeights) {
      const newHeights = {};
      for (let r = 1; r <= ROWS_COUNT; r++) {
        if (r < targetRow) {
          if (sheet.rowHeights[r]) newHeights[r] = sheet.rowHeights[r];
        } else if (r < ROWS_COUNT) {
          if (sheet.rowHeights[r + 1]) newHeights[r] = sheet.rowHeights[r + 1];
        }
      }
      sheet.rowHeights = newHeights;
    }

    // Shrink exportRange
    if (sheet.exportRange && sheet.exportRange.endRow > sheet.exportRange.startRow) {
      if (targetRow <= sheet.exportRange.endRow) {
        sheet.exportRange.endRow--;
      }
    }

    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Riadok ${targetRow} bol zmazaný.`);
  }

  function insertColumnAt(colToInsert) {
    recordHistory();
    const sheet = getActiveSheet();
    const range = appState.selectionRange || { startCol: 'A' };
    const targetCol = colToInsert || range.startCol;
    const targetColIdx = colIndex(targetCol);
    if (targetColIdx < 0 || targetColIdx >= COLUMNS.length - 1) return;

    // Shift cells right
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

    // Shift colWidths right
    if (sheet.colWidths) {
      const newWidths = {};
      COLUMNS.forEach((c, idx) => {
        if (idx < targetColIdx) {
          if (sheet.colWidths[c]) newWidths[c] = sheet.colWidths[c];
        } else if (idx === targetColIdx) {
          newWidths[c] = 70; // default for newly inserted column
        } else if (idx > targetColIdx) {
          const prevCol = COLUMNS[idx - 1];
          if (sheet.colWidths[prevCol]) newWidths[c] = sheet.colWidths[prevCol];
        }
      });
      sheet.colWidths = newWidths;
    }

    // Adjust exportRange
    if (sheet.exportRange) {
      const sIdx = colIndex(sheet.exportRange.startCol);
      const eIdx = colIndex(sheet.exportRange.endCol);
      if (targetColIdx <= eIdx && eIdx < COLUMNS.length - 1) {
        sheet.exportRange.endCol = COLUMNS[eIdx + 1];
      }
      if (targetColIdx < sIdx && sIdx < COLUMNS.length - 1) {
        sheet.exportRange.startCol = COLUMNS[sIdx + 1];
      }
    }

    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Nový stĺpec bol vložený na pozíciu ${targetCol}.`);
  }

  function deleteColumnAt(colToDelete) {
    recordHistory();
    const sheet = getActiveSheet();
    const range = appState.selectionRange || { startCol: 'A' };
    const targetCol = colToDelete || range.startCol;
    const targetColIdx = colIndex(targetCol);
    if (targetColIdx < 0) return;

    // Shift cells left
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

    // Shift colWidths left
    if (sheet.colWidths) {
      const newWidths = {};
      COLUMNS.forEach((c, idx) => {
        if (idx < targetColIdx) {
          if (sheet.colWidths[c]) newWidths[c] = sheet.colWidths[c];
        } else if (idx < COLUMNS.length - 1) {
          const nextCol = COLUMNS[idx + 1];
          if (sheet.colWidths[nextCol]) newWidths[c] = sheet.colWidths[nextCol];
        }
      });
      sheet.colWidths = newWidths;
    }

    // Adjust exportRange
    if (sheet.exportRange) {
      const sIdx = colIndex(sheet.exportRange.startCol);
      const eIdx = colIndex(sheet.exportRange.endCol);
      if (eIdx > sIdx && targetColIdx <= eIdx) {
        sheet.exportRange.endCol = COLUMNS[Math.max(sIdx, eIdx - 1)];
      }
    }

    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Stĺpec ${targetCol} bol zmazaný.`);
  }

  if (btnInsertRow) {
    btnInsertRow.addEventListener('click', () => {
      insertRowAt();
    });
  }

  if (btnDeleteRow) {
    btnDeleteRow.addEventListener('click', () => {
      deleteRowAt();
    });
  }

  if (btnInsertCol) {
    btnInsertCol.addEventListener('click', () => {
      insertColumnAt();
    });
  }

  if (btnDeleteCol) {
    btnDeleteCol.addEventListener('click', () => {
      deleteColumnAt();
    });
  }

  // --- Clear Tools ---
  function clearContent() {
    recordHistory();
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
    recordHistory();
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
    refreshAllGridBorders(sheet);
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
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'u') {
      e.preventDefault();
      toggleUnderline();
    }
  });

  // --- Move Table Across Spreadsheet Grid ---
  function moveTable(deltaRow, deltaCol) {
    recordHistory();
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

  // --- Effective Export Range Finder ---
  function getEffectiveExportRange(sheet) {
    if (sheet.exportRange && sheet.exportRange.isManual) {
      return sheet.exportRange;
    }
    let minC = 999, maxC = -1, minR = 999, maxR = -1;
    if (sheet.cells) {
      for (const addr in sheet.cells) {
        const cell = sheet.cells[addr];
        if (cell && (
          (cell.v !== undefined && String(cell.v).trim() !== '') ||
          (cell.fillColor && cell.fillColor !== '#ffffff' && cell.fillColor !== 'transparent') ||
          (cell.borderBottom && cell.borderBottom.width > 0) ||
          (cell.borderTop && cell.borderTop.width > 0) ||
          (cell.borderLeft && cell.borderLeft.width > 0) ||
          (cell.borderRight && cell.borderRight.width > 0)
        )) {
          const { col, row } = parseAddr(addr);
          const cIdx = colIndex(col);
          if (cIdx >= 0) {
            if (cIdx < minC) minC = cIdx;
            if (cIdx > maxC) maxC = cIdx;
            if (row < minR) minR = row;
            if (row > maxR) maxR = row;
          }
        }
      }
    }
    if (minC <= maxC && minR <= maxR) {
      return {
        startCol: COLUMNS[minC],
        startRow: minR,
        endCol: COLUMNS[maxC],
        endRow: maxR
      };
    }
    return sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
  }

  // --- Button: Nastaviť výber ako oblasť pre export ---
  if (btnSetExportRange) {
    btnSetExportRange.addEventListener('click', () => {
      recordHistory();
      const sheet = getActiveSheet();
      sheet.exportRange = Object.assign({}, appState.selectionRange);
      sheet.exportRange.isManual = true;
      if (statusBarRange) {
        statusBarRange.textContent = `Export: ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}`;
      }
      saveAppState();
      refreshSelectionVisuals();
      updatePreview();
      showToast(`Oblasť exportu nastavená na ${sheet.exportRange.startCol}${sheet.exportRange.startRow}:${sheet.exportRange.endCol}${sheet.exportRange.endRow}.`);
    });
  }


  let showPrintGuides = false;
  if (btnTogglePrintGuides) {
    btnTogglePrintGuides.addEventListener('click', () => {
      showPrintGuides = !showPrintGuides;
      btnTogglePrintGuides.classList.toggle('active', showPrintGuides);
      updatePreview();
      showToast(showPrintGuides ? 'Rozmery tlače zobrazené.' : 'Rozmery tlače skryté.');
    });
  }

  // --- Column & Row Resizing Logic ---
  let activeColHandle = null;
  let activeRowHandle = null;
  let resizeStartX = 0;
  let resizeStartY = 0;
  let resizeStartDim = 0;
  let resizeTargetName = '';

  function setupGridResize() {
    // Column resize handles
    gridTable.querySelectorAll('.col-resize-handle').forEach(handle => {
      handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        activeColHandle = handle;
        resizeTargetName = handle.dataset.col;
        resizeStartX = e.clientX;
        const th = handle.closest('th');
        resizeStartDim = th.offsetWidth;
        document.body.style.cursor = 'col-resize';
      });

      handle.addEventListener('dblclick', (e) => {
        e.preventDefault();
        e.stopPropagation();
        autoFitCol(handle.dataset.col);
      });
    });

    // Row resize handles
    gridTable.querySelectorAll('.row-resize-handle').forEach(handle => {
      handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        activeRowHandle = handle;
        resizeTargetName = handle.dataset.row;
        resizeStartY = e.clientY;
        const tr = handle.closest('tr');
        resizeStartDim = tr.offsetHeight;
        document.body.style.cursor = 'row-resize';
      });

      handle.addEventListener('dblclick', (e) => {
        e.preventDefault();
        e.stopPropagation();
        autoFitRow(parseInt(handle.dataset.row, 10));
      });
    });
  }

  function autoFitCol(colLetter) {
    recordHistory();
    const sheet = getActiveSheet();
    let maxW = 50;
    for (let r = 1; r <= ROWS_COUNT; r++) {
      const td = gridTable.querySelector(`td[data-addr="${colLetter}${r}"]`);
      if (td && td.innerText.trim()) {
        const textLen = td.innerText.length;
        const estW = Math.max(60, textLen * 9 + 20);
        if (estW > maxW) maxW = estW;
      }
    }
    if (!sheet.colWidths) sheet.colWidths = {};
    sheet.colWidths[colLetter] = Math.min(400, maxW);
    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Šírka stĺpca ${colLetter} bola automaticky prispôsobená.`);
  }

  function autoFitRow(rowNum) {
    recordHistory();
    const sheet = getActiveSheet();
    let maxH = 24;
    COLUMNS.forEach(col => {
      const td = gridTable.querySelector(`td[data-addr="${col}${rowNum}"]`);
      if (td && td.innerText.trim()) {
        const lines = td.innerText.split('\n').length;
        const estH = Math.max(24, lines * 20 + 8);
        if (estH > maxH) maxH = estH;
      }
    });
    if (!sheet.rowHeights) sheet.rowHeights = {};
    sheet.rowHeights[rowNum] = maxH;
    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Výška riadku ${rowNum} bola automaticky prispôsobená.`);
  }

  window.addEventListener('mousemove', (e) => {
    if (activeColHandle) {
      const diff = e.clientX - resizeStartX;
      const newWidth = Math.max(6, resizeStartDim + diff);
      const sheet = getActiveSheet();
      if (!sheet.colWidths) sheet.colWidths = {};
      sheet.colWidths[resizeTargetName] = newWidth;

      const colIdx = COLUMNS.indexOf(resizeTargetName) + 1;
      const colEl = gridTable.querySelectorAll('col')[colIdx];
      if (colEl) colEl.style.width = newWidth + 'px';
      const thEl = gridTable.querySelector(`th.grid-header-col[data-col="${resizeTargetName}"]`);
      if (thEl) {
        thEl.style.width = newWidth + 'px';
        thEl.style.minWidth = newWidth + 'px';
        thEl.style.maxWidth = newWidth + 'px';
      }
      const tds = gridTable.querySelectorAll(`td[data-col="${resizeTargetName}"]`);
      tds.forEach(td => {
        td.style.width = newWidth + 'px';
        td.style.maxWidth = newWidth + 'px';
      });
    } else if (activeRowHandle) {
      const diff = e.clientY - resizeStartY;
      const newHeight = Math.max(4, resizeStartDim + diff);
      const sheet = getActiveSheet();
      if (!sheet.rowHeights) sheet.rowHeights = {};
      const rNum = parseInt(resizeTargetName, 10);
      sheet.rowHeights[rNum] = newHeight;

      const tr = gridTable.querySelector(`tbody tr:nth-child(${rNum})`);
      if (tr) {
        tr.style.height = newHeight + 'px';
        const th = tr.querySelector('th.grid-header-row');
        if (th) {
          th.style.height = newHeight + 'px';
          th.style.maxHeight = newHeight + 'px';
          th.style.lineHeight = `${Math.max(4, newHeight - 2)}px`;
        }
        const tds = tr.querySelectorAll('td.grid-cell');
        tds.forEach(td => {
          td.style.height = newHeight + 'px';
          td.style.maxHeight = newHeight + 'px';
          td.style.lineHeight = `${Math.max(4, newHeight - 2)}px`;
        });
      }
    }
  });

  window.addEventListener('mouseup', () => {
    if (activeColHandle || activeRowHandle) {
      recordHistory();
      activeColHandle = null;
      activeRowHandle = null;
      document.body.style.cursor = '';
      saveAppState();
      renderGrid();
      updatePreview();
    }
  });

  // --- Mouse Table Relocation (Drag & Drop) ---
  let isDraggingTableMouse = false;
  let dragTableOrigin = { col: 'A', row: 1 };
  let dragTableCurrentTarget = { col: 'A', row: 1 };
  let tableDragGhostEl = null;

  function startMouseTableDrag(col, row, e) {
    isDraggingTableMouse = true;
    dragTableOrigin = { col, row };
    dragTableCurrentTarget = { col, row };

    if (!tableDragGhostEl) {
      tableDragGhostEl = document.createElement('div');
      tableDragGhostEl.className = 'table-drag-ghost';
      document.body.appendChild(tableDragGhostEl);
    }
    updateDragGhostPosition();
    document.body.style.cursor = 'move';
    document.body.style.userSelect = 'none';

    function onMouseMove(moveEvent) {
      if (!isDraggingTableMouse) return;
      const hovered = document.elementFromPoint(moveEvent.clientX, moveEvent.clientY);
      const td = hovered && hovered.closest('td.grid-cell');
      if (td && td.dataset.col && td.dataset.row) {
        dragTableCurrentTarget = {
          col: td.dataset.col,
          row: parseInt(td.dataset.row, 10)
        };
        updateDragGhostPosition();
      }
    }

    function onMouseUp() {
      if (!isDraggingTableMouse) return;
      isDraggingTableMouse = false;
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';

      if (tableDragGhostEl && tableDragGhostEl.parentNode) {
        tableDragGhostEl.parentNode.removeChild(tableDragGhostEl);
        tableDragGhostEl = null;
      }

      const deltaRow = dragTableCurrentTarget.row - dragTableOrigin.row;
      const deltaCol = colIndex(dragTableCurrentTarget.col) - colIndex(dragTableOrigin.col);
      if (deltaRow !== 0 || deltaCol !== 0) {
        recordHistory();
        moveTable(deltaRow, deltaCol);
      }
    }

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }

  function updateDragGhostPosition() {
    if (!tableDragGhostEl) return;
    const sheet = getActiveSheet();
    const range = sheet.exportRange || { startCol: 'A', startRow: 1, endCol: 'B', endRow: 8 };
    const deltaRow = dragTableCurrentTarget.row - dragTableOrigin.row;
    const deltaCol = colIndex(dragTableCurrentTarget.col) - colIndex(dragTableOrigin.col);

    const sC = Math.max(0, Math.min(COLUMNS.length - 1, colIndex(range.startCol) + deltaCol));
    const eC = Math.max(0, Math.min(COLUMNS.length - 1, colIndex(range.endCol) + deltaCol));
    const sR = Math.max(1, Math.min(ROWS_COUNT, range.startRow + deltaRow));
    const eR = Math.max(1, Math.min(ROWS_COUNT, range.endRow + deltaRow));

    const topLeftTd = gridTable.querySelector(`td[data-addr="${COLUMNS[sC]}${sR}"]`);
    const bottomRightTd = gridTable.querySelector(`td[data-addr="${COLUMNS[eC]}${eR}"]`);

    if (topLeftTd && bottomRightTd) {
      const r1 = topLeftTd.getBoundingClientRect();
      const r2 = bottomRightTd.getBoundingClientRect();
      tableDragGhostEl.style.left = r1.left + 'px';
      tableDragGhostEl.style.top = r1.top + 'px';
      tableDragGhostEl.style.width = (r2.right - r1.left) + 'px';
      tableDragGhostEl.style.height = (r2.bottom - r1.top) + 'px';
      tableDragGhostEl.style.display = 'block';
    }
  }

  // --- Paste from Excel (HTML & TSV) with 100% Format Preservation & Dynamic Proportions ---
  function parseDimensionPx(val) {
    if (!val) return 0;
    const s = String(val).trim();
    const m = s.match(/^(\d+(?:\.\d+)?)\s*(pt|px|in|cm|mm)?$/i);
    if (!m) {
      const num = parseFloat(s);
      return isNaN(num) ? 0 : Math.round(num);
    }
    const num = parseFloat(m[1]);
    const unit = (m[2] || 'px').toLowerCase();
    if (unit === 'pt') return Math.round(num * 1.3333);
    if (unit === 'in') return Math.round(num * 96);
    if (unit === 'cm') return Math.round(num * 37.8);
    if (unit === 'mm') return Math.round(num * 3.78);
    return Math.round(num);
  }

  function autoAdjustPastedDimensions(sheet, startColIdx, maxC, startRow, maxR) {
    if (!sheet.colWidths) sheet.colWidths = {};
    if (!sheet.rowHeights) sheet.rowHeights = {};

    // 1. Calculate ideal column widths based on content so text doesn't awkwardly break
    for (let c = startColIdx; c <= maxC; c++) {
      const colLetter = COLUMNS[c];
      let maxLineChars = 0;
      let maxPt = 11;
      for (let r = startRow; r <= maxR; r++) {
        const cell = sheet.cells[`${colLetter}${r}`];
        if (cell && cell.v) {
          const lines = String(cell.v).split('\n');
          lines.forEach(l => {
            const clean = l.trim();
            if (clean.length > maxLineChars) maxLineChars = clean.length;
          });
          if (cell.fontSize && cell.fontSize > maxPt) maxPt = cell.fontSize;
        }
      }
      const charWidth = maxPt * 0.72;
      const neededWidth = Math.max(70, Math.round(maxLineChars * charWidth + 28));
      const currentW = sheet.colWidths[colLetter] || 0;
      sheet.colWidths[colLetter] = Math.max(currentW, neededWidth);
    }

    // 2. Calculate row heights based on font sizes and multiline counts
    for (let r = startRow; r <= maxR; r++) {
      let maxLines = 1;
      let maxPt = 11;
      for (let c = startColIdx; c <= maxC; c++) {
        const cell = sheet.cells[`${COLUMNS[c]}${r}`];
        if (cell && cell.v) {
          const lines = String(cell.v).split('\n').length;
          if (lines > maxLines) maxLines = lines;
          if (cell.fontSize && cell.fontSize > maxPt) maxPt = cell.fontSize;
        }
      }
      const cellPx = Math.round(maxPt * 1.33);
      const neededH = Math.round(maxLines * (cellPx * 1.25) + 10);
      const currentH = sheet.rowHeights[r] || 24;
      sheet.rowHeights[r] = Math.max(currentH, neededH);
    }
  }

  function parseExcelBorder(bStr) {
    if (!bStr) return null;
    const m = bStr.match(/([\d\.]+)\s*(pt|px)?\s+([a-z]+)\s*(.*)/i);
    if (m) {
      const rawW = parseFloat(m[1]) || 1;
      let width = 1;
      if (rawW > 1.5) width = 3;
      else if (rawW > 0.8) width = 2;
      else width = 1;
      let color = m[4] ? m[4].trim() : '#000000';
      if (color === 'windowtext' || color === 'auto' || !color) color = '#000000';
      return { width, style: m[3].toLowerCase(), color };
    }
    return null;
  }

  function pasteExcelHtml(html) {
    recordHistory();
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    if (!sheet.colWidths) sheet.colWidths = {};
    if (!sheet.rowHeights) sheet.rowHeights = {};

    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    const anchor = parseAddr(appState.selectedCell || 'A1');
    const startColIdx = colIndex(anchor.col);
    const startRow = anchor.row;

    // Check for colgroup widths
    const colEls = doc.querySelectorAll('col');
    colEls.forEach((colEl, idx) => {
      const c = startColIdx + idx;
      if (c < COLUMNS.length) {
        const rawW = parseDimensionPx(colEl.getAttribute('width') || colEl.style.width);
        if (rawW > 0) sheet.colWidths[COLUMNS[c]] = rawW;
      }
    });

    const rows = doc.querySelectorAll('tr');
    if (!rows || rows.length === 0) return;

    let maxC = startColIdx;
    let maxR = startRow;

    rows.forEach((tr, rIdx) => {
      const r = startRow + rIdx;
      if (r > ROWS_COUNT) return;
      if (r > maxR) maxR = r;

      const trH = parseDimensionPx(tr.style.height || tr.getAttribute('height'));
      if (trH > 0) {
        sheet.rowHeights[r] = trH;
      }

      const cells = tr.querySelectorAll('td, th');
      cells.forEach((td, cIdx) => {
        const c = startColIdx + cIdx;
        if (c >= COLUMNS.length) return;
        if (c > maxC) maxC = c;
        const colLetter = COLUMNS[c];
        const addr = `${colLetter}${r}`;

        const tdW = parseDimensionPx(td.style.width || td.getAttribute('width'));
        if (tdW > 0 && (!sheet.colWidths[colLetter] || tdW > sheet.colWidths[colLetter])) {
          sheet.colWidths[colLetter] = tdW;
        }

        // Convert <br> to newlines for multiline cells
        td.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
        const text = (td.textContent || td.innerText || '').trim();

        // Check bold, italic & underline
        const bold = td.style.fontWeight === 'bold' || parseInt(td.style.fontWeight, 10) >= 600 || !!td.querySelector('b, strong');
        const italic = td.style.fontStyle === 'italic' || !!td.querySelector('i, em');
        const underline = (td.style.textDecoration && td.style.textDecoration.includes('underline')) || !!td.querySelector('u');
        const align = td.style.textAlign || td.getAttribute('align') || (c > startColIdx ? 'right' : 'left');

        // Font size
        let fontSize = 11;
        const fontEl = td.querySelector('[style*="font-size"]') || td.querySelector('font');
        const fsRaw = td.style.fontSize || (fontEl && fontEl.style.fontSize) || '';
        const fsMatch = fsRaw.match(/(\d+(?:\.\d+)?)\s*(pt|px)/i);
        if (fsMatch) {
          fontSize = fsMatch[2].toLowerCase() === 'pt' ? Math.round(parseFloat(fsMatch[1])) : Math.round(parseFloat(fsMatch[1]) / 1.33);
        }

        // Font family
        let fontFamily = 'Calibri';
        const rawFam = td.style.fontFamily || (fontEl && (fontEl.style.fontFamily || fontEl.getAttribute('face')));
        if (rawFam) {
          fontFamily = rawFam.replace(/['"]/g, '').split(',')[0].trim();
        }

        // Colors
        let textColor = td.style.color || (fontEl && (fontEl.style.color || fontEl.getAttribute('color'))) || '#000000';
        if (textColor === 'windowtext' || textColor === 'auto') textColor = '#000000';
        let fillColor = td.style.backgroundColor || td.getAttribute('bgcolor') || '';

        // Borders
        const bTop = parseExcelBorder(td.style.borderTop) || parseExcelBorder(td.style.border);
        const bBottom = parseExcelBorder(td.style.borderBottom) || parseExcelBorder(td.style.border);
        const bLeft = parseExcelBorder(td.style.borderLeft) || parseExcelBorder(td.style.border);
        const bRight = parseExcelBorder(td.style.borderRight) || parseExcelBorder(td.style.border);

        const isIndent = text.toLowerCase().startsWith('z toho');

        sheet.cells[addr] = {
          v: text,
          bold,
          italic,
          underline,
          fontSize,
          fontFamily,
          textColor,
          fillColor,
          align,
          indent: isIndent ? 18 : 0,
          borderTop: bTop || { width: 1, style: 'solid', color: '#000000' },
          borderBottom: bBottom || { width: 1, style: 'solid', color: '#000000' },
          borderLeft: bLeft || { width: 1, style: 'solid', color: '#000000' },
          borderRight: bRight || { width: 1, style: 'solid', color: '#000000' }
        };
      });
    });

    // Automatically ensure proportional column widths & row heights
    autoAdjustPastedDimensions(sheet, startColIdx, maxC, startRow, maxR);

    sheet.exportRange = {
      startCol: COLUMNS[startColIdx],
      startRow,
      endCol: COLUMNS[maxC],
      endRow: maxR,
      isManual: true
    };

    updateSelectionRange(COLUMNS[startColIdx], startRow, COLUMNS[maxC], maxR);
    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Tabuľka z Excelu bola úspešne vložená so zachovaním proporcií a formátovania (${COLUMNS[startColIdx]}${startRow}:${COLUMNS[maxC]}${maxR}).`);
  }

  function pasteTsvText(text) {
    recordHistory();
    const sheet = getActiveSheet();
    if (!sheet.cells) sheet.cells = {};
    const lines = text.trim().split(/\r?\n/);
    const anchor = parseAddr(appState.selectedCell || 'A1');
    const startColIdx = colIndex(anchor.col);
    const startRow = anchor.row;

    let maxC = startColIdx;
    let maxR = startRow;

    lines.forEach((line, rIdx) => {
      const r = startRow + rIdx;
      if (r > ROWS_COUNT) return;
      if (r > maxR) maxR = r;
      const parts = line.split('\t');
      parts.forEach((val, cIdx) => {
        const c = startColIdx + cIdx;
        if (c >= COLUMNS.length) return;
        if (c > maxC) maxC = c;
        const colLetter = COLUMNS[c];
        const addr = `${colLetter}${r}`;
        const cleanVal = val.trim();
        const isIndent = cleanVal.toLowerCase().startsWith('z toho');
        sheet.cells[addr] = {
          v: cleanVal,
          bold: r === startRow,
          align: c > startColIdx ? 'right' : 'left',
          indent: isIndent ? 18 : 0,
          borderTop: { width: 1, style: 'solid', color: '#000000' },
          borderBottom: { width: 1, style: 'solid', color: '#000000' },
          borderLeft: { width: 1, style: 'solid', color: '#000000' },
          borderRight: { width: 1, style: 'solid', color: '#000000' }
        };
      });
    });

    // Automatically ensure proportional column widths & row heights
    autoAdjustPastedDimensions(sheet, startColIdx, maxC, startRow, maxR);

    sheet.exportRange = {
      startCol: COLUMNS[startColIdx],
      startRow,
      endCol: COLUMNS[maxC],
      endRow: maxR,
      isManual: true
    };

    updateSelectionRange(COLUMNS[startColIdx], startRow, COLUMNS[maxC], maxR);
    saveAppState();
    renderGrid();
    updatePreview();
    showToast(`Údaje vložené do buniek ${COLUMNS[startColIdx]}${startRow}:${COLUMNS[maxC]}${maxR}.`);
  }

  window.addEventListener('paste', (e) => {
    if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'SELECT')) {
      return;
    }
    const htmlData = e.clipboardData && e.clipboardData.getData('text/html');
    const textData = e.clipboardData && e.clipboardData.getData('text/plain');

    if (htmlData && (htmlData.includes('<table') || htmlData.includes('<tr') || htmlData.includes('<td'))) {
      e.preventDefault();
      pasteExcelHtml(htmlData);
    } else if (textData && textData.includes('\t')) {
      e.preventDefault();
      pasteTsvText(textData);
    }
  });

  // Undo & Redo button connections
  if (btnUndo) btnUndo.addEventListener('click', undo);
  if (btnRedo) btnRedo.addEventListener('click', redo);
  if (btnTitleUndo) btnTitleUndo.addEventListener('click', undo);
  if (btnTitleRedo) btnTitleRedo.addEventListener('click', redo);

  // Keyboard Shortcuts: Ctrl+Z and Ctrl+Y
  window.addEventListener('keydown', (e) => {
    if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'SELECT')) {
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
      e.preventDefault();
      undo();
    } else if (((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') || ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'z')) {
      e.preventDefault();
      redo();
    } else if (e.key === 'Delete') {
      recordHistory();
      clearContent();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b') {
      e.preventDefault();
      boldBtn.click();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i') {
      e.preventDefault();
      italicBtn.click();
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'u') {
      e.preventDefault();
      toggleUnderline();
    }
  });

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
  if (btnToggleDrawerExpand && averyPreviewDrawer) {
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
    // Calculate available width and height with safe padding
    const availW = Math.max(80, previewViewport.clientWidth - 40);
    const availH = Math.max(80, previewViewport.clientHeight - 40);
    const svgW = parseFloat(svgEl.getAttribute('width')) || 330;
    const svgH = parseFloat(svgEl.getAttribute('height')) || 300;
    
    // Scale must fit both width and height
    const scaleW = availW / svgW;
    const scaleH = availH / svgH;
    const fitScale = Math.min(1.0, scaleW, scaleH);
    const clampedScale = Math.max(0.1, fitScale);
    previewSvgBox.style.transform = `scale(${clampedScale.toFixed(2)})`;
    previewSvgBox.style.transformOrigin = 'top center';
  }

  window.addEventListener('resize', () => {
    if (previewZoomSelect && previewZoomSelect.value === 'fit') {
      applyFitZoom();
    }
  });

  // Zoom Selector
  if (previewZoomSelect && previewSvgBox) {
    previewZoomSelect.addEventListener('change', () => {
      if (previewZoomSelect.value === 'fit') {
        applyFitZoom();
      } else {
        const scale = parseFloat(previewZoomSelect.value);
        previewSvgBox.style.transform = `scale(${scale})`;
        previewSvgBox.style.transformOrigin = 'top center';
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

  // --- Toggle Preview Drawer Visibility (Persistent) ---
  const DRAWER_VISIBLE_KEY = 'oco_avery_drawer_visible_v3';

  function setDrawerVisible(visible) {
    if (!averyPreviewDrawer) return;
    if (visible) {
      averyPreviewDrawer.classList.remove('drawer-closed');
      if (workspaceSplitter) workspaceSplitter.classList.remove('drawer-closed');
      if (btnTogglePreviewDrawer) btnTogglePreviewDrawer.classList.add('active');
      localStorage.setItem(DRAWER_VISIBLE_KEY, 'true');
      updatePreview();
      if (previewZoomSelect && previewZoomSelect.value === 'fit') {
        applyFitZoom();
      }
    } else {
      averyPreviewDrawer.classList.add('drawer-closed');
      if (workspaceSplitter) workspaceSplitter.classList.add('drawer-closed');
      if (btnTogglePreviewDrawer) btnTogglePreviewDrawer.classList.remove('active');
      localStorage.setItem(DRAWER_VISIBLE_KEY, 'false');
    }
  }

  // Restore drawer visibility state
  const savedDrawerVisible = localStorage.getItem(DRAWER_VISIBLE_KEY);
  if (savedDrawerVisible === 'false') {
    setDrawerVisible(false);
  } else {
    setDrawerVisible(true);
  }

  if (btnTogglePreviewDrawer) {
    btnTogglePreviewDrawer.addEventListener('click', () => {
      const isClosed = averyPreviewDrawer && averyPreviewDrawer.classList.contains('drawer-closed');
      setDrawerVisible(isClosed);
    });
  }

  if (btnCloseDrawer) {
    btnCloseDrawer.addEventListener('click', () => {
      setDrawerVisible(false);
    });
  }

  // --- Excel Authentic Context Menu (Right-Click on Grid Cells & Headers) ---
  if (gridTable && excelContextMenu) {
    gridTable.addEventListener('contextmenu', (e) => {
      const thCol = e.target.closest('th.grid-header-col');
      const thRow = e.target.closest('th.grid-header-row');
      const td = e.target.closest('td.grid-cell');
      if (!td && !thCol && !thRow) return;
      e.preventDefault();

      let targetCellAddr = 'A1';

      if (thCol) {
        const col = thCol.dataset.col;
        if (!col) return;
        updateSelectionRange(col, 1, col, ROWS_COUNT);
        targetCellAddr = `${col}1`;
      } else if (thRow) {
        const row = parseInt(thRow.dataset.row, 10);
        if (isNaN(row)) return;
        updateSelectionRange('A', row, COLUMNS[COLUMNS.length - 1], row);
        targetCellAddr = `A${row}`;
      } else if (td) {
        const col = td.dataset.col;
        const row = parseInt(td.dataset.row, 10);
        if (!col || isNaN(row)) return;

        // Check if clicked cell is within current selection
        const cur = appState.selectionRange;
        const sC = Math.min(colIndex(cur.startCol), colIndex(cur.endCol));
        const eC = Math.max(colIndex(cur.startCol), colIndex(cur.endCol));
        const sR = Math.min(cur.startRow, cur.endRow);
        const eR = Math.max(cur.startRow, cur.endRow);
        const cIdx = colIndex(col);

        const isInside = (cIdx >= sC && cIdx <= eC && row >= sR && row <= eR);
        if (!isInside) {
          updateSelectionRange(col, row, col, row);
        }
        targetCellAddr = `${col}${row}`;
      }

      // Synchronize Mini Toolbar values with active/anchor cell
      const sheet = getActiveSheet();
      const cellData = (sheet && sheet.cells && sheet.cells[targetCellAddr]) || {};

      if (ctxFontFamily) ctxFontFamily.value = cellData.fontFamily || 'Calibri';
      if (ctxFontSize) ctxFontSize.value = String(cellData.fontSize || 11);
      if (ctxBold) ctxBold.classList.toggle('active', !!cellData.bold);
      if (ctxItalic) ctxItalic.classList.toggle('active', !!cellData.italic);
      if (ctxUnderline) ctxUnderline.classList.toggle('active', !!cellData.underline);
      if (ctxAlignLeft) ctxAlignLeft.classList.toggle('active', (cellData.align || 'left') === 'left');
      if (ctxAlignCenter) ctxAlignCenter.classList.toggle('active', cellData.align === 'center');
      if (ctxAlignRight) ctxAlignRight.classList.toggle('active', cellData.align === 'right');

      // Safe positioning within viewport
      const menuW = 270;
      const menuH = 490;
      let x = e.clientX;
      let y = e.clientY;
      if (x + menuW > window.innerWidth) x = Math.max(10, window.innerWidth - menuW - 12);
      if (y + menuH > window.innerHeight) y = Math.max(10, window.innerHeight - menuH - 12);

      excelContextMenu.style.left = `${Math.max(10, x)}px`;
      excelContextMenu.style.top = `${Math.max(10, y)}px`;
      excelContextMenu.style.display = 'block';
    });

    document.addEventListener('click', (e) => {
      if (!excelContextMenu.contains(e.target)) {
        excelContextMenu.style.display = 'none';
      }
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        excelContextMenu.style.display = 'none';
      }
    });
  }

  // --- Context Menu: Mini Toolbar Listeners ---
  if (ctxFontFamily) {
    ctxFontFamily.addEventListener('change', () => {
      recordHistory();
      const fam = ctxFontFamily.value;
      const sheet = getActiveSheet();
      if (!sheet.cells) sheet.cells = {};
      getCellsInRange().forEach(item => {
        if (!sheet.cells[item.addr]) sheet.cells[item.addr] = { v: '' };
        sheet.cells[item.addr].fontFamily = fam;
        const td = gridTable.querySelector(`td[data-addr="${item.addr}"]`);
        if (td) applyCellStyles(td, sheet.cells[item.addr]);
      });
      if (fontSelect) fontSelect.value = fam;
      saveAppState();
      updatePreview();
    });
  }

  if (ctxFontSize) {
    ctxFontSize.addEventListener('change', () => {
      const size = parseInt(ctxFontSize.value, 10) || 11;
      applyFontSize(size);
    });
  }

  if (ctxFontInc) {
    ctxFontInc.addEventListener('click', () => {
      const cur = parseInt(ctxFontSize ? ctxFontSize.value : (fontSizeSelect ? fontSizeSelect.value : '11'), 10) || 11;
      const next = FONT_SIZES.find(s => s > cur) || (cur + 2);
      applyFontSize(next);
      if (ctxFontSize) ctxFontSize.value = String(next);
    });
  }

  if (ctxFontDec) {
    ctxFontDec.addEventListener('click', () => {
      const cur = parseInt(ctxFontSize ? ctxFontSize.value : (fontSizeSelect ? fontSizeSelect.value : '11'), 10) || 11;
      const prev = [...FONT_SIZES].reverse().find(s => s < cur) || Math.max(6, cur - 2);
      applyFontSize(prev);
      if (ctxFontSize) ctxFontSize.value = String(prev);
    });
  }

  if (ctxBold) {
    ctxBold.addEventListener('click', () => {
      boldBtn.click();
      const sheet = getActiveSheet();
      const cellData = (sheet && sheet.cells && sheet.cells[appState.selectedCell]) || {};
      ctxBold.classList.toggle('active', !!cellData.bold);
    });
  }

  if (ctxItalic) {
    ctxItalic.addEventListener('click', () => {
      italicBtn.click();
      const sheet = getActiveSheet();
      const cellData = (sheet && sheet.cells && sheet.cells[appState.selectedCell]) || {};
      ctxItalic.classList.toggle('active', !!cellData.italic);
    });
  }

  if (ctxUnderline) {
    ctxUnderline.addEventListener('click', () => {
      toggleUnderline();
    });
  }

  if (ctxAlignLeft) {
    ctxAlignLeft.addEventListener('click', () => {
      applyRangeProp('align', 'left');
      ctxAlignLeft.classList.add('active');
      if (ctxAlignCenter) ctxAlignCenter.classList.remove('active');
      if (ctxAlignRight) ctxAlignRight.classList.remove('active');
    });
  }

  if (ctxAlignCenter) {
    ctxAlignCenter.addEventListener('click', () => {
      applyRangeProp('align', 'center');
      ctxAlignCenter.classList.add('active');
      if (ctxAlignLeft) ctxAlignLeft.classList.remove('active');
      if (ctxAlignRight) ctxAlignRight.classList.remove('active');
    });
  }

  if (ctxAlignRight) {
    ctxAlignRight.addEventListener('click', () => {
      applyRangeProp('align', 'right');
      ctxAlignRight.classList.add('active');
      if (ctxAlignLeft) ctxAlignLeft.classList.remove('active');
      if (ctxAlignCenter) ctxAlignCenter.classList.remove('active');
    });
  }

  // --- Context Menu: Quick Borders Listeners ---
  function getCtxBorderSettings() {
    const w = ctxBorderWidth ? (parseInt(ctxBorderWidth.value, 10) || 1) : 1;
    const c = ctxBorderColor ? ctxBorderColor.value : '#000000';
    return { width: w, color: c };
  }

  if (ctxBorderAll) {
    ctxBorderAll.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      const { width, color } = getCtxBorderSettings();
      applyBorderToSelection('all', width, color);
    });
  }

  if (ctxBorderOuter) {
    ctxBorderOuter.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      const { width, color } = getCtxBorderSettings();
      applyBorderToSelection('outer', width, color);
    });
  }

  if (ctxBorderThickOuter) {
    ctxBorderThickOuter.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      const { width, color } = getCtxBorderSettings();
      applyBorderToSelection('thick-outer', Math.max(3, width), color);
    });
  }

  if (ctxBorderNone) {
    ctxBorderNone.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      applyBorderToSelection('none');
    });
  }

  if (ctxShowInPreview) {
    ctxShowInPreview.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      recordHistory();
      const sheet = getActiveSheet();
      const cur = appState.selectionRange;
      const isMulti = (cur.startCol !== cur.endCol || cur.startRow !== cur.endRow);
      if (isMulti) {
        const sC = COLUMNS[Math.min(colIndex(cur.startCol), colIndex(cur.endCol))];
        const eC = COLUMNS[Math.max(colIndex(cur.startCol), colIndex(cur.endCol))];
        const sR = Math.min(cur.startRow, cur.endRow);
        const eR = Math.max(cur.startRow, cur.endRow);
        sheet.exportRange = { startCol: sC, startRow: sR, endCol: eC, endRow: eR, isManual: true };
      } else {
        const eff = getEffectiveExportRange(sheet);
        sheet.exportRange = Object.assign({}, eff, { isManual: true });
      }
      saveAppState();
      setDrawerVisible(true);
      updatePreview();
      renderGrid();
      showToast('Tabuľka nastavená a zobrazená v náhľade.', true);
    });
  }

  if (ctxCopyTable) {
    ctxCopyTable.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnCopyAvery.click();
    });
  }

  if (ctxPasteExcel) {
    ctxPasteExcel.addEventListener('click', async () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      try {
        if (navigator.clipboard && navigator.clipboard.readText) {
          const text = await navigator.clipboard.readText();
          if (text) {
            pasteTsvText(text);
            return;
          }
        }
      } catch (err) {
        console.warn('Clipboard read error, fallback to prompt', err);
      }
      const promptText = prompt('Vložte skopírovanú tabuľku z Excelu (Ctrl+V):');
      if (promptText) {
        pasteTsvText(promptText);
      }
    });
  }

  if (ctxSetExportRange) {
    ctxSetExportRange.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnSetExportRange.click();
    });
  }



  if (ctxClearContent) {
    ctxClearContent.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnClearContent.click();
    });
  }

  if (ctxClearFormat) {
    ctxClearFormat.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnClearFormat.click();
    });
  }

  if (ctxAutoFit) {
    ctxAutoFit.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      const cur = appState.selectionRange;
      const sC = Math.min(colIndex(cur.startCol), colIndex(cur.endCol));
      const eC = Math.max(colIndex(cur.startCol), colIndex(cur.endCol));
      for (let c = sC; c <= eC; c++) {
        autoFitCol(COLUMNS[c]);
      }
    });
  }

  if (ctxInsertRow) {
    ctxInsertRow.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnInsertRow.click();
    });
  }

  if (ctxDeleteRow) {
    ctxDeleteRow.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnDeleteRow.click();
    });
  }

  if (ctxInsertCol) {
    ctxInsertCol.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnInsertCol.click();
    });
  }

  if (ctxDeleteCol) {
    ctxDeleteCol.addEventListener('click', () => {
      if (excelContextMenu) excelContextMenu.style.display = 'none';
      btnDeleteCol.click();
    });
  }

  // --- Horizontal mouse wheel and touchpad drag scroll for ribbon toolbar ---
  const ribbonToolbar = document.querySelector('.excel-ribbon-toolbar');

  if (ribbonToolbar) {
    ribbonToolbar.addEventListener('scroll', () => {
      if (borderMenu && borderMenu.classList.contains('open')) {
        positionBorderMenu();
      }
    }, { passive: true });

    ribbonToolbar.addEventListener('wheel', (e) => {
      if (ribbonToolbar.scrollWidth > ribbonToolbar.clientWidth) {
        if (Math.abs(e.deltaY) > Math.abs(e.deltaX) && e.deltaY !== 0) {
          e.preventDefault();
          ribbonToolbar.scrollLeft += e.deltaY;
        } else if (e.deltaX !== 0) {
          ribbonToolbar.scrollLeft += e.deltaX;
        }
      }
    }, { passive: false });

    // Drag-to-scroll support for touchpads / mouse drag on ribbon background
    let isRibbonDragging = false;
    let startX = 0;
    let startScrollLeft = 0;

    ribbonToolbar.addEventListener('mousedown', (e) => {
      if (e.target.closest('button, input, select, .border-menu, .sheet-tab')) return;
      isRibbonDragging = true;
      startX = e.pageX - ribbonToolbar.offsetLeft;
      startScrollLeft = ribbonToolbar.scrollLeft;
      ribbonToolbar.style.cursor = 'grabbing';
    });

    window.addEventListener('mousemove', (e) => {
      if (!isRibbonDragging) return;
      e.preventDefault();
      const x = e.pageX - ribbonToolbar.offsetLeft;
      const walk = (x - startX) * 1.5;
      ribbonToolbar.scrollLeft = startScrollLeft - walk;
    });

    window.addEventListener('mouseup', () => {
      if (isRibbonDragging) {
        isRibbonDragging = false;
        ribbonToolbar.style.cursor = '';
      }
    });
  }

  // --- SVG Live Preview Generator With Custom Pixel Borders & Real Dimensions ---
  function generateTableSVG(sheet) {
    const range = getEffectiveExportRange(sheet);
    const colWidths = sheet.colWidths || {};
    const rowHeights = sheet.rowHeights || {};

    const sC = colIndex(range.startCol);
    const eC = colIndex(range.endCol);

    const activeCols = [];
    let totalW = 0;
    for (let c = sC; c <= eC; c++) {
      const colLetter = COLUMNS[c];
      const userW = (colWidths && colWidths[colLetter] > 0) ? colWidths[colLetter] : null;
      const thEl = gridTable ? gridTable.querySelector(`th.grid-header-col[data-col="${colLetter}"]`) : null;
      const w = userW !== null ? userW : ((thEl && thEl.offsetWidth > 0) ? thEl.offsetWidth : (colLetter === 'A' ? 270 : (colLetter === 'B' ? 120 : 70)));
      activeCols.push({ col: colLetter, width: w, x: totalW });
      totalW += w;
    }

    let totalH = 0;
    const rowPositions = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const userH = (rowHeights && rowHeights[r] > 0) ? rowHeights[r] : null;
      const trEl = gridTable ? gridTable.querySelector(`tbody tr:nth-child(${r})`) : null;
      let h;
      if (userH !== null) {
        h = userH; // Exact user height - respect completely!
      } else {
        let baseH = (trEl && trEl.offsetHeight > 0) ? trEl.offsetHeight : (r === 2 ? 38 : 24);
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
        h = Math.max(baseH, neededH);
      }
      rowPositions.push({ row: r, y: totalH, height: h });
      totalH += h;
    }

    const margin = 2;
    const svgW = totalW + margin * 2;
    const svgH = totalH + margin * 2;

    let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${svgW} ${svgH}" width="${svgW}" height="${svgH}">\n`;
    svg += `<defs>\n<style>\n`;
    svg += `  text { dominant-baseline: middle; }\n`;
    svg += `  .bold { font-weight: bold; }\n`;
    svg += `  .italic { font-style: italic; }\n`;
    svg += `  .underline { text-decoration: underline; }\n`;
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
        if (cell.fillColor && cell.fillColor !== 'transparent' && cell.fillColor !== 'none') {
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
        const textClass = (cell.bold ? 'bold ' : '') + (cell.italic ? 'italic ' : '') + (cell.underline ? 'underline ' : '');
        const textColor = cell.textColor || '#000000';
        const styleAttr = `style="font-size: ${cellFontPx}px; font-family: ${cellFontFam};"`;

        const valLines = String(cell.v || '').split('\n');
        if (valLines.length === 1) {
          if (cell.align === 'center') {
            svg += `<text x="${cellX + cellW / 2}" y="${textY}" class="center ${textClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}" ${styleAttr}>${escapeXml(cell.v)}</text>\n`;
          } else if (cell.align === 'right') {
            svg += `<text x="${cellX + cellW - 8}" y="${textY}" class="right ${textClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}" ${styleAttr}>${escapeXml(cell.v)}</text>\n`;
          } else {
            svg += `<text x="${cellX + 8 + indentX}" y="${textY}" class="${textClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}" ${styleAttr}>${escapeXml(cell.v)}</text>\n`;
          }
        } else {
          // Multiline cell (kJ / kcal)
          const startY = textY - ((valLines.length - 1) * (lineH / 2));
          const posX = cell.align === 'right' ? (cellX + cellW - 8) : (cell.align === 'center' ? (cellX + cellW / 2) : (cellX + 8 + indentX));
          const anchorClass = (cell.align === 'right' ? 'right ' : (cell.align === 'center' ? 'center ' : '')) + textClass;

          svg += `<text x="${posX}" y="${startY}" class="${anchorClass}" fill="${textColor}" font-size="${cellFontPx}px" font-family="${cellFontFam}" ${styleAttr}>`;
          valLines.forEach((line, idx) => {
            svg += `<tspan x="${posX}" dy="${idx === 0 ? 0 : lineH}" ${styleAttr}>${escapeXml(line)}</tspan>`;
          });
          svg += `</text>\n`;
        }
      });
    });

    // 3. Exact Symmetrical Cell Borders (Custom px thickness, style, color)
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

    rowPositions.forEach((rp, rIdx) => {
      const r = rp.row;
      const y = rp.y + margin;
      const h = rp.height;

      activeCols.forEach((ac, cIdxRel) => {
        const cIdx = colIndex(ac.col);
        const cellX = margin + ac.x;
        const cellW = ac.width;
        const addr = `${ac.col}${r}`;

        // Top border of each row in range
        const bTop = resolveBorderObj(cIdx, r, 'top', sheet);
        const hasTopBorder = bTop && bTop.width > 0 && bTop.style !== 'none';
        if (hasTopBorder) {
          svg += renderSvgBorder(cellX, y, cellX + cellW, y, bTop);
        } else if (rIdx > 0) {
          // Internal horizontal row divider without explicit border:
          // Check if user set a background fill (e.g. White) that covers gridlines
          const cell = (sheet.cells && sheet.cells[addr]) || {};
          const prevRowAddr = `${ac.col}${rowPositions[rIdx - 1].row}`;
          const prevCell = (sheet.cells && sheet.cells[prevRowAddr]) || {};
          const cellFill = (cell.fillColor && cell.fillColor !== 'transparent' && cell.fillColor !== 'none') ? cell.fillColor : null;
          const prevFill = (prevCell.fillColor && prevCell.fillColor !== 'transparent' && prevCell.fillColor !== 'none') ? prevCell.fillColor : null;
          if (!cellFill && !prevFill) {
            svg += `<line x1="${cellX}" y1="${y}" x2="${cellX + cellW}" y2="${y}" stroke="#d4d4d4" stroke-width="1" />\n`;
          }
        }

        // Bottom border of last row in range
        if (rIdx === rowPositions.length - 1) {
          const bBottom = resolveBorderObj(cIdx, r, 'bottom', sheet);
          if (bBottom && bBottom.width > 0 && bBottom.style !== 'none') {
            svg += renderSvgBorder(cellX, y + h, cellX + cellW, y + h, bBottom);
          }
        }

        // Left border of each col in range
        const bLeft = resolveBorderObj(cIdx, r, 'left', sheet);
        const hasLeftBorder = bLeft && bLeft.width > 0 && bLeft.style !== 'none';
        if (hasLeftBorder) {
          svg += renderSvgBorder(cellX, y, cellX, y + h, bLeft);
        } else if (cIdxRel > 0) {
          // Internal vertical column divider without explicit border:
          const cell = (sheet.cells && sheet.cells[addr]) || {};
          const prevCol = activeCols[cIdxRel - 1].col;
          const prevColAddr = `${prevCol}${r}`;
          const prevColCell = (sheet.cells && sheet.cells[prevColAddr]) || {};
          const cellFill = (cell.fillColor && cell.fillColor !== 'transparent' && cell.fillColor !== 'none') ? cell.fillColor : null;
          const prevFill = (prevColCell.fillColor && prevColCell.fillColor !== 'transparent' && prevColCell.fillColor !== 'none') ? prevColCell.fillColor : null;
          if (!cellFill && !prevFill) {
            svg += `<line x1="${cellX}" y1="${y}" x2="${cellX}" y2="${y + h}" stroke="#d4d4d4" stroke-width="1" />\n`;
          }
        }

        // Right border of last col in range
        if (cIdxRel === activeCols.length - 1) {
          const bRight = resolveBorderObj(cIdx, r, 'right', sheet);
          if (bRight && bRight.width > 0 && bRight.style !== 'none') {
            svg += renderSvgBorder(cellX + cellW, y, cellX + cellW, y + h, bRight);
          }
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
