import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function runVerification() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage();

  const fileUrl = 'file://' + path.resolve(__dirname, '../index.html').replace(/\\/g, '/');
  console.log('Navigating to:', fileUrl);

  const errors = [];
  page.on('console', msg => {
    if (msg.type() === 'error') errors.push(msg.text());
  });
  page.on('pageerror', err => errors.push(err.message));

  await page.goto(fileUrl, { waitUntil: 'load' });

  // 1. Verify Clean Title without emoji and without "pre Avery"
  const title = await page.textContent('.file-title');
  console.log('Page Title:', title);
  if (title.includes('Avery') || /[\u{1F300}-\u{1F6FF}]/u.test(title)) {
    throw new Error(`Title should not contain "Avery" or emojis: ${title}`);
  }

  // 2. Verify Primary Copy Button text is clean "Kopírovať" (no "pre avery", no emojis)
  const copyBtnText = (await page.textContent('#btnCopyAvery span')).trim();
  console.log('Primary Copy Button Text:', copyBtnText);
  if (copyBtnText !== 'Kopírovať') {
    throw new Error(`Copy button text expected "Kopírovať", got "${copyBtnText}"`);
  }

  // 3. Verify Cells exist and have initial values
  const cellA1 = (await page.textContent('td[data-addr="A1"]')).trim();
  const cellB1 = (await page.textContent('td[data-addr="B1"]')).trim();
  console.log('Cell A1:', cellA1, '| Cell B1:', cellB1);

  // 4. Test Drag Selection
  console.log('Testing Drag Selection from A1 to B4...');
  const boxA1 = await page.locator('td[data-addr="A1"]').boundingBox();
  const boxB4 = await page.locator('td[data-addr="B4"]').boundingBox();

  await page.mouse.move(boxA1.x + 10, boxA1.y + 10);
  await page.mouse.down();
  await page.mouse.move(boxB4.x + 10, boxB4.y + 10);
  await page.mouse.up();

  const formulaAddr = await page.textContent('#formulaCellAddr');
  console.log('Formula Bar Address after drag:', formulaAddr);

  // 5. Test Font Size Scaling
  console.log('Testing Font Size Scaling...');
  await page.click('td[data-addr="A1"]');
  await page.selectOption('#ribbonFontSize', '16');
  
  const fontSizeA1 = await page.evaluate(() => document.querySelector('td[data-addr="A1"]').style.fontSize);
  console.log('Cell A1 fontSize after select 16:', fontSizeA1);
  if (fontSizeA1 !== '16pt') {
    throw new Error(`Expected cell A1 font-size to be 16pt, got "${fontSizeA1}"`);
  }

  // Click font increase button
  await page.click('#btnFontInc');
  const fontSizeAfterInc = await page.evaluate(() => document.querySelector('td[data-addr="A1"]').style.fontSize);
  console.log('Cell A1 fontSize after #btnFontInc:', fontSizeAfterInc);
  if (fontSizeAfterInc !== '18pt') {
    throw new Error(`Expected cell A1 font-size to be 18pt, got "${fontSizeAfterInc}"`);
  }

  // Verify font-size in rendered SVG preview
  const svgTextContent = await page.evaluate(() => document.querySelector('#previewSvgBox svg').outerHTML);
  if (!svgTextContent.includes('font-size="24px"')) {
    throw new Error('SVG preview did not scale text font-size to 24px (18pt * 1.33)!');
  }
  console.log('Font size scaling verified in both table and SVG preview.');

  // 6. Test Table Relocation (Moving Table Across Grid)
  console.log('Testing Table Moving Controls...');
  const rangeBeforeMove = await page.textContent('#statusBarRange');
  console.log('Range before move:', rangeBeforeMove);

  // Move down
  await page.click('#btnMoveDown');
  const rangeAfterDown = await page.textContent('#statusBarRange');
  console.log('Range after #btnMoveDown:', rangeAfterDown);
  if (!rangeAfterDown.includes('A2:B9')) {
    throw new Error(`Expected range to be A2:B9 after move down, got "${rangeAfterDown}"`);
  }

  // Move right
  await page.click('#btnMoveRight');
  const rangeAfterRight = await page.textContent('#statusBarRange');
  console.log('Range after #btnMoveRight:', rangeAfterRight);
  if (!rangeAfterRight.includes('B2:C9')) {
    throw new Error(`Expected range to be B2:C9 after move right, got "${rangeAfterRight}"`);
  }

  // Verify cell value moved to B2
  const movedValueB2 = (await page.textContent('td[data-addr="B2"]')).trim();
  console.log('Cell B2 value after table move:', movedValueB2);
  if (movedValueB2 !== 'Výživové údaje na') {
    throw new Error(`Expected cell B2 to have header text after move, got "${movedValueB2}"`);
  }

  // Move back up and left to A1:B8
  await page.click('#btnMoveUp');
  await page.click('#btnMoveLeft');
  const restoredRange = await page.textContent('#statusBarRange');
  console.log('Restored range:', restoredRange);
  if (!restoredRange.includes('A1:B8')) {
    throw new Error(`Expected restored range A1:B8, got "${restoredRange}"`);
  }

  // 7. Verify "Vzor EÚ údajov" has been removed completely
  const euTemplateBtn = await page.$('#btnInsertNutritionTemplate');
  if (euTemplateBtn) {
    throw new Error('Found #btnInsertNutritionTemplate in DOM, it should be removed!');
  }
  console.log('Confirmed "Vzor EU údajov" is removed.');

  // 8. Verify Export Buttons Exist (SVG, PNG, JPG)
  const btnSvg = await page.$('#btnExportSvg');
  const btnPng = await page.$('#btnExportPng');
  const btnJpg = await page.$('#btnExportJpg');
  if (!btnSvg || !btnPng || !btnJpg) {
    throw new Error('Export buttons for SVG, PNG or JPG are missing!');
  }
  console.log('Confirmed SVG, PNG, and JPG export buttons are present.');

  // 9. Verify Custom Border Pixel Controls
  console.log('Testing Border Controls...');
  await page.fill('#borderWidthInput', '3');
  await page.selectOption('#borderStyleSelect', 'solid');
  await page.click('#btnBorderDropdown');
  await page.click('.border-menu-item[data-border="all"]');

  const svgBorderContent = await page.evaluate(() => document.querySelector('#previewSvgBox svg').outerHTML);
  if (!svgBorderContent.includes('stroke-width="3"')) {
    throw new Error('SVG does not contain custom 3px border stroke-width!');
  }
  console.log('Border 3px application verified in SVG.');

  // 10. Verify NO Emojis in Entire DOM
  const hasEmoji = await page.evaluate(() => {
    const text = document.body.innerText + ' ' + document.title;
    const emojiRegex = /[\u{1F300}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1FA00}-\u{1FAFF}]/gu;
    const matches = text.match(emojiRegex);
    return matches ? matches.join(', ') : null;
  });
  if (hasEmoji) {
    throw new Error(`DOM contains unwanted emojis: ${hasEmoji}`);
  }
  console.log('Zero emojis verified across the entire DOM.');

  // 11. Test Interactive Splitter & Drawer
  console.log('Testing Drawer Expand & Splitter...');
  await page.click('#btnToggleDrawerExpand');
  let drawerWidth = await page.evaluate(() => document.getElementById('averyPreviewDrawer').offsetWidth);
  if (drawerWidth < 700) {
    throw new Error(`Drawer expected to be expanded >= 700px, got ${drawerWidth}px`);
  }
  await page.click('#btnToggleDrawerExpand');

  // 12. Capture screenshot to verification file
  await page.screenshot({ path: path.resolve(__dirname, '../verification-excel-ui.png') });
  console.log('Screenshot saved to verification-excel-ui.png');

  await browser.close();

  if (errors.length > 0) {
    console.error('Console errors encountered:', errors);
    process.exit(1);
  } else {
    console.log('All Playwright UI verifications PASSED with 0 errors!');
  }
}

runVerification().catch(err => {
  console.error('Verification failed:', err);
  process.exit(1);
});
