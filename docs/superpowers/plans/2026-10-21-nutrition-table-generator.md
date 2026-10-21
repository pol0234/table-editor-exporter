# Generátor a exportér výživových tabuliek: Implementačný plán

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Vytvoriť samostatnú, offline webovú aplikáciu pre otca na tvorbu výživových tabuliek, pokročilú úpravu orámovania a riadkovania, import existujúcich Excel zošitov a export do čistého vektorového SVG, 300+ DPI PNG/JPG a priamo do schránky pre Avery.

**Architecture:** Klientska architektúra s modulárnym JavaScriptom (ES2022) a CSS3 bez nutnosti zostavovacieho nástroja (buildless), bežiaca priamo v prehliadači. Využíva SheetJS pre offline parsovanie Excelu, interný SVG DOM generátor pre 1:1 vektorovú zhodu náhľadu a exportu, HTML5 Canvas pre ultra-vysoké rozlíšenie (300 až 600 DPI) a Clipboard API pre 1-klikové vkladanie do Avery.

**Tech Stack:** HTML5, CSS3 (modern responsive layout, design tokens), Vanilla ES2022 JavaScript, SheetJS (xlsx.full.min.js), HTML5 Canvas, SVG, Node.js test runner (`node --test`).

**Spec:** [docs/superpowers/specs/2026-10-21-nutrition-table-generator-design.md](file:///c:/Users/Admin/Documents/projektiky/Oco%20-%20Vy%C5%BEivov%C3%A9%20hodnoty/docs/superpowers/specs/2026-10-21-nutrition-table-generator-design.md)

## Global Constraints

- Aplikácia musí fungovať 100 % offline bez nutnosti internetového pripojenia a bez nutnosti spúšťať backend server.
- Používateľské rozhranie a texty musia byť kompletne v zrozumiteľnej slovenčine s veľkými a jasne čitateľnými prvkami.
- Exportované SVG musí byť čistý vektorový formát bez rasterizácie textov, aby sa dalo neobmedzene zväčšovať v Avery.
- Rasterizovaný export (PNG/JPG) musí podporovať tlačové rozlíšenia 1x, 2x, 4x (300 DPI) a 6x (600 DPI) s voľbou priehľadného alebo bieleho pozadia.
- Každá zmena dát sa musí automaticky ukladať do `localStorage`.

---

### Task 1: Dátový model, správa produktov (záložiek) a perzistencia

**Files:**
- Create: `src/model.js`
- Create: `tests/model.test.js`

**Interfaces:**
- Produces:
  - `createDefaultProduct(name: string): Product`
  - `createInitialState(): AppState`
  - `saveState(state: AppState, storage?: Storage): void`
  - `loadState(storage?: Storage): AppState`
  - `addProduct(state: AppState, name: string): Product`
  - `removeProduct(state: AppState, productId: string): void`
  - `duplicateProduct(state: AppState, productId: string): Product`
  - `updateProductRow(product: Product, rowIndex: number, patch: Partial<NutritionRow>): void`
  - `addRow(product: Product, index: number, row: NutritionRow): void`
  - `removeRow(product: Product, index: number): void`
  - `reorderRows(product: Product, fromIndex: number, toIndex: number): void`

- [ ] **Step 1: Napísať zlyhávajúci test pre dátový model produktov a riadkov**

```javascript
// tests/model.test.js
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createDefaultProduct,
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
} from '../src/model.js';

test('createDefaultProduct creates standard nutrition rows matching EU packaging', () => {
  const prod = createDefaultProduct('Soft-Dried-Papaya');
  assert.equal(prod.name, 'Soft-Dried-Papaya');
  assert.equal(prod.header.title, 'Výživové údaje na');
  assert.equal(prod.header.unit, '100g');
  assert.ok(prod.rows.length >= 7);
  assert.equal(prod.rows[0].name, 'Energetická hodnota');
  assert.equal(prod.rows[1].name, 'Tuky');
  assert.equal(prod.rows[2].name, 'z toho nasýtené mastné kyseliny');
  assert.equal(prod.rows[2].indent, true);
});

test('state management: add, duplicate, remove product', () => {
  const state = createInitialState();
  assert.ok(state.products.length > 0);
  const initialCount = state.products.length;

  const added = addProduct(state, 'Dried-Mango');
  assert.equal(state.products.length, initialCount + 1);
  assert.equal(added.name, 'Dried-Mango');

  const dup = duplicateProduct(state, added.id);
  assert.equal(dup.name, 'Dried-Mango (kópia)');
  assert.equal(state.products.length, initialCount + 2);

  removeProduct(state, added.id);
  assert.equal(state.products.length, initialCount + 1);
  assert.ok(!state.products.find(p => p.id === added.id));
});

test('row manipulation: add, remove, reorder, update', () => {
  const prod = createDefaultProduct('Test');
  updateProductRow(prod, 0, { value: '1372kJ /\n328kcal' });
  assert.equal(prod.rows[0].value, '1372kJ /\n328kcal');

  addRow(prod, 1, { name: 'Vláknina', value: '2,5g', indent: false, borderBottom: 'thin' });
  assert.equal(prod.rows[1].name, 'Vláknina');

  removeRow(prod, 1);
  assert.equal(prod.rows[1].name, 'Tuky');

  reorderRows(prod, 0, 1);
  assert.equal(prod.rows[0].name, 'Tuky');
  assert.equal(prod.rows[1].name, 'Energetická hodnota');
});

test('persistence: saveState and loadState with mock storage', () => {
  const storageMap = new Map();
  const mockStorage = {
    getItem: (k) => storageMap.get(k) || null,
    setItem: (k, v) => storageMap.set(k, String(v)),
    removeItem: (k) => storageMap.delete(k)
  };

  const state = createInitialState();
  addProduct(state, 'Špeciálny produkt');
  saveState(state, mockStorage);

  const restored = loadState(mockStorage);
  assert.equal(restored.products.length, state.products.length);
  assert.ok(restored.products.some(p => p.name === 'Špeciálny produkt'));
});
```

- [ ] **Step 2: Spustiť test a overiť, že zlyhá**

Run: `node --test tests/model.test.js`  
Expected: FAIL (Cannot find module `../src/model.js`)

- [ ] **Step 3: Implementovať `src/model.js`**

```javascript
// src/model.js
export const STORAGE_KEY = 'oco_nutrition_tables_data_v1';

export function createDefaultProduct(name = 'Nový produkt') {
  return {
    id: 'prod_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
    name: name,
    header: {
      title: 'Výživové údaje na',
      unit: '100g'
    },
    styles: {
      fontFamily: 'Calibri, Arial, sans-serif',
      fontSize: 13, // pt/px base
      tableWidth: 320, // px
      colRatio: 0.68, // 68% for name column, 32% for value column
      cellPaddingY: 4,
      cellPaddingX: 8,
      outerBorderWidth: 1.5,
      outerBorderColor: '#000000',
      innerBorderColor: '#000000',
      headerBorderBottom: 'thick', // 'none' | 'thin' | 'thick' | 'double'
      verticalDivider: true,
      verticalDividerWidth: 1.0,
      backgroundColor: '#ffffff'
    },
    rows: [
      { id: 'r1', name: 'Energetická hodnota', value: '1372kJ /\n328kcal', indent: false, bold: false, borderBottom: 'thick', heightMultiplier: 1.8 },
      { id: 'r2', name: 'Tuky', value: '0,12g', indent: false, bold: false, borderBottom: 'thin', heightMultiplier: 1.0 },
      { id: 'r3', name: 'z toho nasýtené mastné kyseliny', value: '0,08g', indent: true, bold: false, borderBottom: 'thin', heightMultiplier: 1.0 },
      { id: 'r4', name: 'Sacharidy', value: '80,9g', indent: false, bold: false, borderBottom: 'thin', heightMultiplier: 1.0 },
      { id: 'r5', name: 'z toho cukry', value: '70,4g', indent: true, bold: false, borderBottom: 'thin', heightMultiplier: 1.0 },
      { id: 'r6', name: 'Bielkoviny', value: '1,54g', indent: false, bold: false, borderBottom: 'thin', heightMultiplier: 1.0 },
      { id: 'r7', name: 'Soľ', value: '0,23g', indent: false, bold: false, borderBottom: 'none', heightMultiplier: 1.0 }
    ]
  };
}

export function createInitialState() {
  const p1 = createDefaultProduct('Soft-Dried-Papaya');
  const p2 = createDefaultProduct('Dried-Mango');
  p2.rows[0].value = '1420kJ /\n339kcal';
  p2.rows[1].value = '0,50g';
  p2.rows[2].value = '0,10g';
  p2.rows[3].value = '82,0g';
  p2.rows[4].value = '65,0g';
  p2.rows[5].value = '1,80g';
  p2.rows[6].value = '0,05g';

  const p3 = createDefaultProduct('Soft-Dried-PassionFruit');
  const p4 = createDefaultProduct('Mixed-Dried-TropicalFruits');

  return {
    version: 1,
    activeProductId: p1.id,
    products: [p1, p2, p3, p4]
  };
}

export function saveState(state, storage = (typeof localStorage !== 'undefined' ? localStorage : null)) {
  if (!storage) return;
  try {
    storage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    console.error('Failed to save state to storage', e);
  }
}

export function loadState(storage = (typeof localStorage !== 'undefined' ? localStorage : null)) {
  if (!storage) return createInitialState();
  try {
    const raw = storage.getItem(STORAGE_KEY);
    if (!raw) return createInitialState();
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.products) || parsed.products.length === 0) {
      return createInitialState();
    }
    return parsed;
  } catch (e) {
    console.warn('Failed to parse stored state, using defaults', e);
    return createInitialState();
  }
}

export function addProduct(state, name = 'Nový produkt') {
  const prod = createDefaultProduct(name);
  state.products.push(prod);
  state.activeProductId = prod.id;
  return prod;
}

export function removeProduct(state, productId) {
  if (state.products.length <= 1) return; // Keep at least one
  state.products = state.products.filter(p => p.id !== productId);
  if (state.activeProductId === productId) {
    state.activeProductId = state.products[0].id;
  }
}

export function duplicateProduct(state, productId) {
  const target = state.products.find(p => p.id === productId);
  if (!target) return null;
  const copy = JSON.parse(JSON.stringify(target));
  copy.id = 'prod_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
  copy.name = target.name + ' (kópia)';
  state.products.push(copy);
  state.activeProductId = copy.id;
  return copy;
}

export function updateProductRow(product, rowIndex, patch) {
  if (product && product.rows && product.rows[rowIndex]) {
    Object.assign(product.rows[rowIndex], patch);
  }
}

export function addRow(product, index, rowData) {
  const newRow = Object.assign({
    id: 'r_' + Date.now() + '_' + Math.random().toString(36).substring(2, 5),
    name: 'Nová položka',
    value: '0,0g',
    indent: false,
    bold: false,
    borderBottom: 'thin',
    heightMultiplier: 1.0
  }, rowData);

  if (typeof index === 'number' && index >= 0 && index <= product.rows.length) {
    product.rows.splice(index, 0, newRow);
  } else {
    product.rows.push(newRow);
  }
}

export function removeRow(product, index) {
  if (product.rows.length > 1 && index >= 0 && index < product.rows.length) {
    product.rows.splice(index, 1);
  }
}

export function reorderRows(product, fromIndex, toIndex) {
  if (fromIndex < 0 || fromIndex >= product.rows.length || toIndex < 0 || toIndex >= product.rows.length) return;
  const [moved] = product.rows.splice(fromIndex, 1);
  product.rows.splice(toIndex, 0, moved);
}
```

- [ ] **Step 4: Spustiť test a overiť, že prejde**

Run: `node --test tests/model.test.js`  
Expected: PASS

- [ ] **Step 5: Commit**

Run: `git add src/model.js tests/model.test.js && git commit -m "feat: add data model, product management and persistence"`

---

### Task 2: SVG Renderovací engine (Vektorová tabuľka, orámovanie riadkov, odsadenie a typografia)

**Files:**
- Create: `src/svg-renderer.js`
- Create: `tests/svg-renderer.test.js`

**Interfaces:**
- Produces:
  - `renderTableToSVG(product: Product, options?: RenderOptions): { svgString: string, width: number, height: number }`
  - RenderOptions: `{ scale?: number, background?: 'transparent' | 'white' | string, margin?: number }`

- [ ] **Step 1: Napísať zlyhávajúci test pre generovanie SVG**

```javascript
// tests/svg-renderer.test.js
import test from 'node:test';
import assert from 'node:assert/strict';
import { createDefaultProduct } from '../src/model.js';
import { renderTableToSVG } from '../src/svg-renderer.js';

test('renderTableToSVG produces valid SVG XML containing table contents', () => {
  const prod = createDefaultProduct('Test Papaya');
  const result = renderTableToSVG(prod);

  assert.ok(result.svgString.startsWith('<svg'));
  assert.ok(result.svgString.endsWith('</svg>'));
  assert.ok(result.width > 200);
  assert.ok(result.height > 150);

  // Checks header and values
  assert.ok(result.svgString.includes('Výživové údaje na'));
  assert.ok(result.svgString.includes('100g'));
  assert.ok(result.svgString.includes('Energetická hodnota'));
  assert.ok(result.svgString.includes('1372kJ /'));
  assert.ok(result.svgString.includes('328kcal'));
});

test('renderTableToSVG renders borders according to styles', () => {
  const prod = createDefaultProduct('Test Borders');
  prod.styles.outerBorderWidth = 2.5;
  prod.styles.verticalDivider = true;
  prod.rows[0].borderBottom = 'thick';

  const result = renderTableToSVG(prod);
  // Outer rectangle border check
  assert.ok(result.svgString.includes('stroke-width="2.5"'));
  // Vertical divider line check
  assert.ok(result.svgString.includes('class="col-divider"'));
  // Multiline tspan check for kJ / kcal
  assert.ok(result.svgString.includes('<tspan'));
});

test('renderTableToSVG respects background transparency vs solid white', () => {
  const prod = createDefaultProduct('Background test');
  const transparent = renderTableToSVG(prod, { background: 'transparent' });
  assert.ok(!transparent.svgString.includes('id="bg-rect"'));

  const white = renderTableToSVG(prod, { background: 'white' });
  assert.ok(white.svgString.includes('id="bg-rect"'));
  assert.ok(white.svgString.includes('fill="#ffffff"'));
});
```

- [ ] **Step 2: Spustiť test a overiť, že zlyhá**

Run: `node --test tests/svg-renderer.test.js`  
Expected: FAIL (Cannot find module `../src/svg-renderer.js`)

- [ ] **Step 3: Implementovať `src/svg-renderer.js`**

```javascript
// src/svg-renderer.js
export function escapeXml(unsafe) {
  return String(unsafe || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function renderTableToSVG(product, options = {}) {
  const styles = product.styles || {};
  const margin = options.margin ?? 1;
  const background = options.background ?? 'transparent';
  
  const tableWidth = styles.tableWidth || 320;
  const colRatio = styles.colRatio || 0.68;
  const col1Width = Math.round(tableWidth * colRatio);
  const col2Width = tableWidth - col1Width;
  
  const fontSize = styles.fontSize || 13;
  const lineHeight = fontSize * 1.25;
  const cellPaddingY = styles.cellPaddingY ?? 4;
  const cellPaddingX = styles.cellPaddingX ?? 8;
  const outerBorderWidth = styles.outerBorderWidth ?? 1.5;
  const outerBorderColor = styles.outerBorderColor || '#000000';
  const innerBorderColor = styles.innerBorderColor || '#000000';
  const fontFamily = styles.fontFamily || 'Calibri, Arial, sans-serif';
  
  // Calculate row positions and heights
  let currentY = margin;
  
  // Header Row
  const headerLinesCount = 1;
  const headerHeight = Math.max(26, headerLinesCount * lineHeight + cellPaddingY * 2);
  const headerY = currentY;
  currentY += headerHeight;
  
  const computedRows = [];
  for (const row of product.rows) {
    const valLines = String(row.value || '').split('\n');
    const nameLines = String(row.name || '').split('\n');
    const maxLines = Math.max(valLines.length, nameLines.length, 1);
    const multiplier = row.heightMultiplier || 1.0;
    const baseH = Math.max(22, maxLines * lineHeight + cellPaddingY * 2);
    const rowH = Math.round(baseH * multiplier);
    
    computedRows.push({
      row,
      y: currentY,
      height: rowH,
      valLines,
      nameLines
    });
    currentY += rowH;
  }
  
  const tableHeight = currentY - margin;
  const totalSvgWidth = tableWidth + margin * 2;
  const totalSvgHeight = tableHeight + margin * 2;
  
  // Build SVG XML
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${totalSvgWidth} ${totalSvgHeight}" width="${totalSvgWidth}" height="${totalSvgHeight}">\n`;
  svg += `<defs>\n<style>\n`;
  svg += `  text { font-family: ${fontFamily}; font-size: ${fontSize}px; fill: #000000; }\n`;
  svg += `  .bold { font-weight: bold; }\n`;
  svg += `  .header-text { font-weight: bold; text-anchor: middle; }\n`;
  svg += `  .val-text { text-anchor: end; }\n`;
  svg += `  .indent { transform: translateX(18px); }\n`;
  svg += `</style>\n</defs>\n`;
  
  if (background && background !== 'transparent') {
    const fill = background === 'white' ? '#ffffff' : background;
    svg += `<rect id="bg-rect" x="0" y="0" width="${totalSvgWidth}" height="${totalSvgHeight}" fill="${fill}" />\n`;
  }
  
  const tableX = margin;
  const tableY = margin;
  
  // Background of table itself if transparent canvas but table has fill
  if (styles.backgroundColor && styles.backgroundColor !== 'transparent') {
    svg += `<rect x="${tableX}" y="${tableY}" width="${tableWidth}" height="${tableHeight}" fill="${styles.backgroundColor}" />\n`;
  }
  
  // Header text
  const headerMidY = headerY + headerHeight / 2 + fontSize * 0.35;
  svg += `<!-- Header Row -->\n`;
  svg += `<text x="${tableX + col1Width / 2}" y="${headerMidY}" class="header-text">${escapeXml(product.header.title)}</text>\n`;
  svg += `<text x="${tableX + col1Width + col2Width / 2}" y="${headerMidY}" class="header-text">${escapeXml(product.header.unit)}</text>\n`;
  
  // Header Bottom border
  const headerBorderThickness = styles.headerBorderBottom === 'thick' ? 2 : (styles.headerBorderBottom === 'thin' ? 1 : (styles.headerBorderBottom === 'double' ? 3 : 0));
  if (headerBorderThickness > 0) {
    svg += `<line x1="${tableX}" y1="${headerY + headerHeight}" x2="${tableX + tableWidth}" y2="${headerY + headerHeight}" stroke="${innerBorderColor}" stroke-width="${headerBorderThickness}" />\n`;
  }
  
  // Rows
  svg += `<!-- Data Rows -->\n`;
  for (let i = 0; i < computedRows.length; i++) {
    const item = computedRows[i];
    const r = item.row;
    const rY = item.y;
    const rH = item.height;
    
    // Text positioning
    const textStartY = rY + cellPaddingY + fontSize * 0.9;
    
    // Label
    const indentX = r.indent ? 16 : 0;
    const labelX = tableX + cellPaddingX + indentX;
    const labelClass = r.bold ? 'bold' : '';
    
    if (item.nameLines.length === 1) {
      svg += `<text x="${labelX}" y="${textStartY}" class="${labelClass}">${escapeXml(item.nameLines[0])}</text>\n`;
    } else {
      svg += `<text x="${labelX}" y="${textStartY}" class="${labelClass}">`;
      item.nameLines.forEach((line, lineIdx) => {
        svg += `<tspan x="${labelX}" dy="${lineIdx === 0 ? 0 : lineHeight}">${escapeXml(line)}</tspan>`;
      });
      svg += `</text>\n`;
    }
    
    // Value (Right aligned)
    const valX = tableX + tableWidth - cellPaddingX;
    const valClass = 'val-text' + (r.bold ? ' bold' : '');
    if (item.valLines.length === 1) {
      svg += `<text x="${valX}" y="${textStartY}" class="${valClass}">${escapeXml(item.valLines[0])}</text>\n`;
    } else {
      svg += `<text x="${valX}" y="${textStartY}" class="${valClass}">`;
      item.valLines.forEach((line, lineIdx) => {
        svg += `<tspan x="${valX}" dy="${lineIdx === 0 ? 0 : lineHeight}">${escapeXml(line)}</tspan>`;
      });
      svg += `</text>\n`;
    }
    
    // Horizontal row bottom border
    const borderType = r.borderBottom || 'thin';
    if (borderType !== 'none' && i < computedRows.length - 1) {
      const bStroke = borderType === 'thick' ? 2 : (borderType === 'double' ? 3 : 1);
      svg += `<line x1="${tableX}" y1="${rY + rH}" x2="${tableX + tableWidth}" y2="${rY + rH}" stroke="${innerBorderColor}" stroke-width="${bStroke}" />\n`;
    }
  }
  
  // Vertical column divider
  if (styles.verticalDivider) {
    const vStroke = styles.verticalDividerWidth || 1;
    svg += `<!-- Vertical Divider -->\n`;
    svg += `<line class="col-divider" x1="${tableX + col1Width}" y1="${tableY}" x2="${tableX + col1Width}" y2="${tableY + tableHeight}" stroke="${innerBorderColor}" stroke-width="${vStroke}" />\n`;
  }
  
  // Outer rectangle border
  if (outerBorderWidth > 0) {
    svg += `<!-- Outer Border -->\n`;
    svg += `<rect x="${tableX}" y="${tableY}" width="${tableWidth}" height="${tableHeight}" fill="none" stroke="${outerBorderColor}" stroke-width="${outerBorderWidth}" />\n`;
  }
  
  svg += `</svg>`;
  
  return {
    svgString: svg,
    width: totalSvgWidth,
    height: totalSvgHeight
  };
}
```

- [ ] **Step 4: Spustiť test a overiť, že prejde**

Run: `node --test tests/svg-renderer.test.js`  
Expected: PASS

- [ ] **Step 5: Commit**

Run: `git add src/svg-renderer.js tests/svg-renderer.test.js && git commit -m "feat: implement precise vector SVG table renderer with borders and multiline text"`

---

### Task 3: Excel Parser & Importér hárkov do záložiek (SheetJS integrácia)

**Files:**
- Create: `src/excel-parser.js`
- Create: `tests/excel-parser.test.js`
- Download/Place: `libs/xlsx.full.min.js`

**Interfaces:**
- Produces:
  - `parseExcelWorkbook(workbook: any): Product[]`
  - `parseNutritionText(text: string): { name: string, value: string, indent: boolean } | null`

- [ ] **Step 1: Napísať zlyhávajúci test pre parsovanie nutričných hodnôt z hárku**

```javascript
// tests/excel-parser.test.js
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseExcelWorkbook, parseSheetRowsToProduct } from '../src/excel-parser.js';

test('parseSheetRowsToProduct extracts nutrition table matching user screenshot', () => {
  // Simulating 2D array from sheet like the user screenshot
  const sheetData = [
    ['Výživové údaje na', '100g'],
    ['Energetická hodnota', '1372kJ /\n328kcal'],
    ['Tuky', '0,12g'],
    ['z toho nasýtené mastné kyseliny', '0,08g'],
    ['Sacharidy', '80,9g'],
    ['z toho cukry', '70,4g'],
    ['Bielkoviny', '1,54g'],
    ['Soľ', '0,23g']
  ];

  const product = parseSheetRowsToProduct('Soft-Dried-Papaya', sheetData);
  assert.equal(product.name, 'Soft-Dried-Papaya');
  assert.equal(product.header.title, 'Výživové údaje na');
  assert.equal(product.header.unit, '100g');
  assert.equal(product.rows.length, 7);
  assert.equal(product.rows[0].name, 'Energetická hodnota');
  assert.equal(product.rows[0].value, '1372kJ /\n328kcal');
  assert.equal(product.rows[2].indent, true);
  assert.equal(product.rows[4].indent, true);
});

test('parseExcelWorkbook creates multiple products from workbook sheets', () => {
  const mockWorkbook = {
    SheetNames: ['Mixed-Dried-TropicalFruits', 'Dried-Mango', 'Soft-Dried-Papaya'],
    Sheets: {
      'Mixed-Dried-TropicalFruits': {
        '!ref': 'A1:B3',
        A1: { v: 'Výživové údaje na' }, B1: { v: '100g' },
        A2: { v: 'Energetická hodnota' }, B2: { v: '1400kJ' },
        A3: { v: 'Tuky' }, B3: { v: '1g' }
      },
      'Dried-Mango': {
        '!ref': 'A1:B3',
        A1: { v: 'Výživové údaje na' }, B1: { v: '100g' },
        A2: { v: 'Energetická hodnota' }, B2: { v: '1450kJ' },
        A3: { v: 'Tuky' }, B3: { v: '0,5g' }
      },
      'Soft-Dried-Papaya': {
        '!ref': 'A1:B2',
        A1: { v: 'Výživové údaje na' }, B1: { v: '100g' },
        A2: { v: 'Energetická hodnota' }, B2: { v: '1372kJ' }
      }
    }
  };

  const products = parseExcelWorkbook(mockWorkbook);
  assert.equal(products.length, 3);
  assert.equal(products[0].name, 'Mixed-Dried-TropicalFruits');
  assert.equal(products[1].name, 'Dried-Mango');
  assert.equal(products[2].name, 'Soft-Dried-Papaya');
});
```

- [ ] **Step 2: Spustiť test a overiť, že zlyhá**

Run: `node --test tests/excel-parser.test.js`  
Expected: FAIL (Cannot find module `../src/excel-parser.js`)

- [ ] **Step 3: Implementovať `src/excel-parser.js`**

```javascript
// src/excel-parser.js
import { createDefaultProduct } from './model.js';

export function isIndentKeyword(label) {
  const l = String(label || '').trim().toLowerCase();
  return l.startsWith('z toho') || l.startsWith('- z toho') || l.startsWith('of which');
}

export function parseSheetRowsToProduct(sheetName, rawRows) {
  const product = createDefaultProduct(sheetName || 'Importovaný produkt');
  if (!Array.isArray(rawRows) || rawRows.length === 0) {
    return product;
  }

  // Find header row (e.g. contains "výživov" or "energie" or first non-empty 2-cell row)
  let headerIndex = -1;
  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    if (row && (row[0] || row[1])) {
      const c0 = String(row[0] || '').toLowerCase();
      if (c0.includes('výživov') || c0.includes('vyzivov') || c0.includes('nutrition') || c0.includes('údaje')) {
        headerIndex = i;
        break;
      }
    }
  }

  let startIndex = 0;
  if (headerIndex !== -1) {
    const hRow = rawRows[headerIndex];
    product.header.title = String(hRow[0] || 'Výživové údaje na').trim();
    product.header.unit = String(hRow[1] || '100g').trim();
    startIndex = headerIndex + 1;
  }

  const parsedRows = [];
  for (let i = startIndex; i < rawRows.length; i++) {
    const r = rawRows[i];
    if (!r) continue;
    const name = String(r[0] || '').trim();
    const val = String(r[1] || '').trim();
    if (!name && !val) continue;

    const isIndent = isIndentKeyword(name);
    const isEnergy = name.toLowerCase().includes('energetick') || name.toLowerCase().includes('energia');
    
    parsedRows.push({
      id: 'r_imp_' + i + '_' + Math.random().toString(36).substring(2, 5),
      name: name,
      value: val,
      indent: isIndent,
      bold: false,
      borderBottom: isEnergy ? 'thick' : 'thin',
      heightMultiplier: isEnergy && val.includes('\n') ? 1.8 : 1.0
    });
  }

  if (parsedRows.length > 0) {
    // Set last row border to none
    parsedRows[parsedRows.length - 1].borderBottom = 'none';
    product.rows = parsedRows;
  }

  return product;
}

export function parseExcelWorkbook(workbook) {
  if (!workbook || !workbook.SheetNames || !workbook.Sheets) return [];
  const products = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    
    // Extract rows from sheet
    const rows = [];
    if (typeof sheet['!ref'] === 'string') {
      const [startCell, endCell] = sheet['!ref'].split(':');
      const startCol = startCell.replace(/[0-9]/g, '');
      const startRow = parseInt(startCell.replace(/[^0-9]/g, ''), 10);
      const endRow = parseInt((endCell || startCell).replace(/[^0-9]/g, ''), 10);

      for (let r = startRow; r <= endRow; r++) {
        const cellA = sheet['A' + r];
        const cellB = sheet['B' + r];
        const valA = cellA ? (cellA.w || cellA.v || '') : '';
        const valB = cellB ? (cellB.w || cellB.v || '') : '';
        if (valA !== '' || valB !== '') {
          rows.push([valA, valB]);
        }
      }
    }

    products.push(parseSheetRowsToProduct(sheetName, rows));
  }

  return products;
}
```

- [ ] **Step 4: Spustiť test a overiť, že prejde**

Run: `node --test tests/excel-parser.test.js`  
Expected: PASS

- [ ] **Step 5: Zabezpečiť lokálnu knižnicu SheetJS v `libs/xlsx.full.min.js`**

Fetch/stiahnuť alebo vložiť overenú minifikovanú SheetJS knižnicu do `libs/xlsx.full.min.js` pre plné offline fungovanie.

- [ ] **Step 6: Commit**

Run: `git add src/excel-parser.js tests/excel-parser.test.js libs/xlsx.full.min.js && git commit -m "feat: add Excel workbook parser and offline SheetJS integration"`

---

### Task 4: Exportný engine (SVG, 300+ DPI PNG/JPG, Clipboard API)

**Files:**
- Create: `src/export-engine.js`
- Create: `tests/export-engine.test.js`

**Interfaces:**
- Produces:
  - `downloadSVG(svgString: string, filename: string): void`
  - `renderSVGToCanvas(svgString: string, scale: number, background: string): Promise<HTMLCanvasElement>`
  - `downloadPNG(svgString: string, filename: string, scale: number, background: string): Promise<void>`
  - `downloadJPG(svgString: string, filename: string, scale: number): Promise<void>`
  - `copyImageToClipboard(svgString: string, scale: number, background: string): Promise<boolean>`

- [ ] **Step 1: Napísať zlyhávajúci test pre kalkuláciu rozlíšení a formátovania exportu**

```javascript
// tests/export-engine.test.js
import test from 'node:test';
import assert from 'node:assert/strict';
import { getDpiDimensions, sanitizeFilename } from '../src/export-engine.js';

test('getDpiDimensions calculates scaled canvas dimensions properly', () => {
  const dims1x = getDpiDimensions(320, 200, 1);
  assert.equal(dims1x.width, 320);
  assert.equal(dims1x.height, 200);

  const dims4x = getDpiDimensions(320, 200, 4);
  assert.equal(dims4x.width, 1280);
  assert.equal(dims4x.height, 800);
  assert.equal(dims4x.dpiLabel, '300 DPI (Tlač)');
});

test('sanitizeFilename removes illegal Windows filename characters and formats extension', () => {
  assert.equal(sanitizeFilename('Soft/Dried:Papaya*?', 'png'), 'Soft-Dried-Papaya.png');
  assert.equal(sanitizeFilename('Mix ovocia 100g', 'svg'), 'Mix_ovocia_100g.svg');
});
```

- [ ] **Step 2: Spustiť test a overiť, že zlyhá**

Run: `node --test tests/export-engine.test.js`  
Expected: FAIL (Cannot find module `../src/export-engine.js`)

- [ ] **Step 3: Implementovať `src/export-engine.js`**

```javascript
// src/export-engine.js
export function sanitizeFilename(name, extension) {
  const clean = String(name || 'vyzivove-udaje')
    .trim()
    .replace(/[<>:"/\\|?*]/g, '-')
    .replace(/\s+/g, '_')
    .replace(/-+/g, '-');
  return `${clean}.${extension}`;
}

export function getDpiDimensions(width, height, scale = 4) {
  const labels = {
    1: '72 DPI (Náhľad)',
    2: '150 DPI (Stredná kvalita)',
    4: '300 DPI (Tlač)',
    6: '600 DPI (Ultra detailná tlač)'
  };
  return {
    width: Math.round(width * scale),
    height: Math.round(height * scale),
    scale,
    dpiLabel: labels[scale] || `${scale}x`
  };
}

export function downloadSVG(svgString, filename = 'tabulka.svg') {
  const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = sanitizeFilename(filename, 'svg');
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function renderSVGToCanvas(svgString, scale = 4, background = 'transparent') {
  return new Promise((resolve, reject) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(svgString, 'image/svg+xml');
    const svgElem = doc.querySelector('svg');
    const width = parseFloat(svgElem.getAttribute('width')) || 320;
    const height = parseFloat(svgElem.getAttribute('height')) || 240;

    const targetWidth = Math.round(width * scale);
    const targetHeight = Math.round(height * scale);

    const canvas = document.createElement('canvas');
    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext('2d');

    // Fill background if needed
    if (background && background !== 'transparent') {
      ctx.fillStyle = background === 'white' ? '#ffffff' : background;
      ctx.fillRect(0, 0, targetWidth, targetHeight);
    }

    const img = new Image();
    const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    img.onload = () => {
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, 0, 0, targetWidth, targetHeight);
      URL.revokeObjectURL(url);
      resolve(canvas);
    };

    img.onerror = (e) => {
      URL.revokeObjectURL(url);
      reject(e);
    };

    img.src = url;
  });
}

export async function downloadPNG(svgString, filename, scale = 4, background = 'transparent') {
  const canvas = await renderSVGToCanvas(svgString, scale, background);
  canvas.toBlob((blob) => {
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = sanitizeFilename(filename, 'png');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }, 'image/png');
}

export async function downloadJPG(svgString, filename, scale = 4) {
  // JPG always needs solid white background
  const canvas = await renderSVGToCanvas(svgString, scale, 'white');
  canvas.toBlob((blob) => {
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = sanitizeFilename(filename, 'jpg');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }, 'image/jpeg', 0.95);
}

export async function copyImageToClipboard(svgString, scale = 4, background = 'white') {
  if (!navigator.clipboard || !navigator.clipboard.write || typeof ClipboardItem === 'undefined') {
    throw new Error('Prehliadač nepodporuje priame kopírovanie obrázkov do schránky.');
  }
  const canvas = await renderSVGToCanvas(svgString, scale, background);
  return new Promise((resolve, reject) => {
    canvas.toBlob(async (blob) => {
      if (!blob) {
        reject(new Error('Chyba pri vytváraní obrázka.'));
        return;
      }
      try {
        const item = new ClipboardItem({ 'image/png': blob });
        await navigator.clipboard.write([item]);
        resolve(true);
      } catch (err) {
        reject(err);
      }
    }, 'image/png');
  });
}
```

- [ ] **Step 4: Spustiť test a overiť, že prejde**

Run: `node --test tests/export-engine.test.js`  
Expected: PASS

- [ ] **Step 5: Commit**

Run: `git add src/export-engine.js tests/export-engine.test.js && git commit -m "feat: implement high-DPI canvas rasterizer, SVG/PNG/JPG exporter and clipboard copy"`

---

### Task 5: Používateľské rozhranie (HTML & CSS) podľa dizajnového systému

**Files:**
- Create: `index.html`
- Create: `styles.css`

**Interfaces:**
- Produces:
  - Responzívny moderný layout s hornou lištou, záložkami produktov, editorom vľavo a živým náhľadom vpravo.
  - Všetky ovládacie prvky: tlačidlá importu Excelu, exportu, kopírovania do schránky, prepínania orámovania, pridávania a mazania riadkov.

- [ ] **Step 1: Vytvoriť `styles.css` s tokenmi, prehľadnou typografiou a responzívnym rozložením**
  - Implementovať farebnú paletu (príjemné sivé tóny `#0f172a`, `#f8fafc`, modrý primárny akcent `#2563eb`, zelený úspech `#16a34a`).
  - Prehľadné ovládacie prvky pre seniora (veľké tlačidlá, vysoký kontrast, jasné ikony a texty).
  - Živý náhľad s checkerboard vzorom pre priehľadnosť a bielym režimom.
  - Plávajúce notifikácie (Toast).

- [ ] **Step 2: Vytvoriť sémantický `index.html`**
  - Hlavička s logom, tlačidlom „Nahrať Excel (.xlsx)“, zálohovaním a exportnými tlačidlami.
  - Pás so záložkami produktov (`+` pre nový, záložky pre jednotlivé hárky).
  - Ľavý panel:
    - Karta 1: Hodnoty tabuľky (Hlavička, riadky, tlačidlá na pridanie riadku, odsadenie, hrúbku čiar).
    - Karta 2: Nastavenie orámovania a písma (hrúbka rámu, zvislá čiara, šírka tabuľky, font Calibri/Arial).
    - Karta 3: Nastavenia exportu (DPI 1x-6x, pozadie priehľadné/biele).
  - Pravý panel:
    - Rýchle akcie (Kopírovať do schránky, Stiahnuť SVG, Stiahnuť PNG/JPG).
    - Živý náhľad tabuľky s možnosťou priblíženia (Zoom 100%, 150%, 200%).

- [ ] **Step 3: Commit**

Run: `git add index.html styles.css && git commit -m "feat: create modern UI layout, design tokens and controls in HTML/CSS"`

---

### Task 6: Prepojenie UI radiča (`src/app.js`), interaktívny editor a end-to-end verifikácia

**Files:**
- Create: `src/app.js`

**Interfaces:**
- Produces:
  - Kompletné prepojenie DOM prvkov so stavom `model.js`, vykresľovačom `svg-renderer.js`, parsovaním `excel-parser.js` a exportom `export-engine.js`.
  - Reagovanie na kliknutia, zmeny hodnôt, nahrávanie súborov, prepínanie záložiek a export.

- [ ] **Step 1: Implementovať `src/app.js`**
  - Inicializácia stavu z `loadState()`.
  - Renderovanie záložiek produktov a prepínanie aktívneho produktu.
  - Dynamické vykreslenie zoznamu riadkov v editore s prepínačmi orámovania (spodná linka: žiadna/tenká/hrubá/dvojitá, odsadenie, viacriadková hodnota).
  - Reakcia na nahrávanie `.xlsx` súborov (drag & drop na celú obrazovku aj kliknutie na tlačidlo).
  - Generovanie živého SVG náhľadu pri každom údere do klávesnice.
  - Tlačidlá exportu a kopírovania do schránky s vizuálnym potvrdením (toast).

- [ ] **Step 2: End-to-end overenie v prehliadači pomocou subagenta alebo automatického testu**
  - Otvoriť aplikáciu a overiť:
    1. Vykreslenie všetkých záložiek.
    2. Správne orámovanie a odsadenie položiek.
    3. Reakciu na zmenu hrúbky orámovania a fontu.
    4. Funkčnosť tlačidiel exportu SVG, PNG a kopírovania do schránky.

- [ ] **Step 3: Commit**

Run: `git add src/app.js && git commit -m "feat: connect UI controller, interactive table editing, live SVG preview and export bindings"`

---

### Task 7: Používateľská príručka pre otca a finálne ladenie

**Files:**
- Create: `README.md`
- Create: `Ako-používať.html` (rýchly zástupca s inštrukciami pre otca)

- [ ] **Step 1: Vytvoriť prehľadný návod v slovenčine s krokmi ako exportovať do Avery**
- [ ] **Step 2: Spustiť finálne testy a overiť, že všetko funguje bez chýb**
- [ ] **Step 3: Commit**

Run: `git add README.md Ako-používať.html && git commit -m "docs: add user manual in Slovak for food label printing in Avery"`
