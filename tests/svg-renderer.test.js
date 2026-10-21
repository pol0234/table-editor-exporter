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
