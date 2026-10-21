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
