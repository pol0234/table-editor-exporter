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
