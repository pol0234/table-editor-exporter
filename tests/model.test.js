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
