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
