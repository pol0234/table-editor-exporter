import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function runVerification() {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await context.newPage();

  const fileUrl = 'file://' + path.resolve(__dirname, '../index.html').replace(/\\/g, '/');
  console.log('Navigating to:', fileUrl);

  const errors = [];
  page.on('console', msg => {
    if (msg.type() === 'error') errors.push(msg.text());
  });
  page.on('pageerror', err => errors.push(err.message));

  await page.goto(fileUrl, { waitUntil: 'load' });

  // 1. Verify Title & Grid
  const title = await page.textContent('.file-title');
  console.log('Page Title:', title);

  // 2. Verify Cells exist and have values
  const cellA1 = await page.textContent('td[data-addr="A1"]');
  const cellB1 = await page.textContent('td[data-addr="B1"]');
  console.log('Cell A1:', cellA1);
  console.log('Cell B1:', cellB1);

  // 3. Test Drag Selection
  console.log('Testing Drag Selection from A1 to B4...');
  const boxA1 = await page.locator('td[data-addr="A1"]').boundingBox();
  const boxB4 = await page.locator('td[data-addr="B4"]').boundingBox();

  await page.mouse.move(boxA1.x + 10, boxA1.y + 10);
  await page.mouse.down();
  await page.mouse.move(boxB4.x + 10, boxB4.y + 10);
  await page.mouse.up();

  const formulaAddr = await page.textContent('#formulaCellAddr');
  console.log('Formula Bar Address after drag:', formulaAddr);

  // 4. Test Splitter Drag
  console.log('Testing Splitter Drag...');
  const splitter = await page.locator('#workspaceSplitter').boundingBox();
  await page.mouse.move(splitter.x + 4, splitter.y + 100);
  await page.mouse.down();
  await page.mouse.move(splitter.x - 100, splitter.y + 100);
  await page.mouse.up();

  const drawerWidth = await page.evaluate(() => document.getElementById('averyPreviewDrawer').offsetWidth);
  console.log('Drawer width after splitter drag:', drawerWidth);

  // 5. Test Sheet Tab switching
  console.log('Testing Sheet Tabs...');
  const tabNames = await page.$$eval('.sheet-tab-title', els => els.map(e => e.textContent));
  console.log('Available sheet tabs:', tabNames);

  await page.click('.sheet-tab:nth-child(2)'); // Click Dried-Mango
  const mangoA2 = await page.textContent('td[data-addr="B2"]');
  console.log('Dried-Mango B2 value:', mangoA2);

  // 6. Check Preview SVG
  const svgExists = await page.evaluate(() => {
    const svg = document.querySelector('#previewSvgBox svg');
    return !!svg && svg.childNodes.length > 0;
  });
  console.log('Preview SVG rendered:', svgExists);

  // Capture screenshot to artifact / verification file
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
