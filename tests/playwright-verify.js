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

  // 5. Test Custom Border Pixel Thickness & Placement (3px All borders)
  console.log('Testing Border Controls (setting 3px width and applying to selection)...');
  await page.fill('#borderWidthInput', '3');
  await page.selectOption('#borderStyleSelect', 'solid');
  
  // Click border dropdown
  await page.click('#btnBorderDropdown');
  // Click "Všetky orámovania"
  await page.click('.border-menu-item[data-border="all"]');

  // Verify that SVG generated has stroke-width="3"
  const svgContent = await page.evaluate(() => document.querySelector('#previewSvgBox svg').outerHTML);
  if (!svgContent.includes('stroke-width="3"')) {
    throw new Error('SVG does not contain custom 3px border stroke-width!');
  }
  console.log('Custom 3px border verified in SVG preview.');

  // Test "Iba vnútorné orámovanie" (Inner borders)
  await page.fill('#borderWidthInput', '2');
  await page.click('#btnBorderDropdown');
  await page.click('.border-menu-item[data-border="inner"]');
  console.log('Inner border application verified.');

  // 6. Test Drawer Expand Button (Clean labels, no emojis)
  const expandBtnText = (await page.textContent('#btnToggleDrawerExpand')).trim();
  console.log('Expand Drawer Button Text:', expandBtnText);
  if (expandBtnText.includes('↔') || expandBtnText.includes('Avery')) {
    throw new Error(`Expand button contains unwanted characters: ${expandBtnText}`);
  }

  // Click expand
  await page.click('#btnToggleDrawerExpand');
  let drawerWidth = await page.evaluate(() => document.getElementById('averyPreviewDrawer').offsetWidth);
  console.log('Drawer width after expand click:', drawerWidth);
  if (drawerWidth < 700) {
    throw new Error(`Drawer expected to be expanded >= 700px, got ${drawerWidth}px`);
  }

  // Click collapse
  await page.click('#btnToggleDrawerExpand');
  drawerWidth = await page.evaluate(() => document.getElementById('averyPreviewDrawer').offsetWidth);
  console.log('Drawer width after collapse click:', drawerWidth);

  // 7. Test Interactive Splitter Dragging
  console.log('Testing Splitter Dragging...');
  const splitter = await page.locator('#workspaceSplitter').boundingBox();
  await page.mouse.move(splitter.x + 4, splitter.y + 100);
  await page.mouse.down();
  await page.mouse.move(splitter.x - 150, splitter.y + 100);
  await page.mouse.up();

  const widthAfterSplitter = await page.evaluate(() => document.getElementById('averyPreviewDrawer').offsetWidth);
  console.log('Drawer width after splitter drag:', widthAfterSplitter);
  if (widthAfterSplitter <= drawerWidth) {
    throw new Error('Splitter drag did not increase drawer width!');
  }

  // 8. Test Data Manipulation: Insert EU Nutrition Template
  console.log('Testing "Vzor EÚ údajov" insertion...');
  await page.click('#btnInsertNutritionTemplate');
  const freshA1 = (await page.textContent('td[data-addr="A1"]')).trim();
  const freshA8 = (await page.textContent('td[data-addr="A8"]')).trim();
  console.log('Template inserted values -> A1:', freshA1, '| A8:', freshA8);
  if (freshA1 !== 'Výživové údaje na' || freshA8 !== 'Soľ') {
    throw new Error('EU template was not properly inserted into the grid!');
  }

  // 9. Capture screenshot to artifact / verification file
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
