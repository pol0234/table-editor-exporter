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
