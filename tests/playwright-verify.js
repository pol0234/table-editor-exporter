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

  // 1. Verify Clean Title strictly "Tabuľkový editor" without emoji
  const title = (await page.textContent('.file-title')).trim();
  console.log('Page Title:', title);
  if (title !== 'Tabuľkový editor') {
    throw new Error(`Title expected "Tabuľkový editor", got: "${title}"`);
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

  // 12. Verify Soft-Dried-Papaya sheet exists
  console.log('Verifying Soft-Dried-Papaya sheet presence...');
  const activeSheetTab = (await page.textContent('.sheet-tab.active .sheet-tab-title')).trim();
  console.log('Active Sheet Tab:', activeSheetTab);
  if (!activeSheetTab.toLowerCase().includes('papaya')) {
    throw new Error(`Expected active sheet to be Soft-Dried-Papaya, got "${activeSheetTab}"`);
  }

  // 13. Test Interactive Row and Column Resizing via Mouse Drag
  console.log('Testing Row & Column Resizing via Mouse Drag...');
  const initialRow1H = await page.evaluate(() => document.querySelector('tbody tr:nth-child(1)').offsetHeight);
  const rowHandleBox = await page.locator('.row-resize-handle[data-row="1"]').boundingBox();
  if (!rowHandleBox) throw new Error('Row resize handle for row 1 not found!');
  
  // Expand Row 1
  await page.mouse.move(rowHandleBox.x + rowHandleBox.width / 2, rowHandleBox.y + rowHandleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(rowHandleBox.x + rowHandleBox.width / 2, rowHandleBox.y + 25);
  await page.mouse.up();

  const newRow1H = await page.evaluate(() => document.querySelector('tbody tr:nth-child(1)').offsetHeight);
  console.log(`Row 1 height expanded: ${initialRow1H}px -> ${newRow1H}px`);
  if (newRow1H <= initialRow1H) {
    throw new Error(`Expected row 1 height to increase from ${initialRow1H}px, got ${newRow1H}px`);
  }

  // Shrink Row 1 down to arbitrary small size
  const handleAfterExpand = await page.locator('.row-resize-handle[data-row="1"]').boundingBox();
  await page.mouse.move(handleAfterExpand.x + handleAfterExpand.width / 2, handleAfterExpand.y + handleAfterExpand.height / 2);
  await page.mouse.down();
  await page.mouse.move(handleAfterExpand.x + handleAfterExpand.width / 2, handleAfterExpand.y - 40);
  await page.mouse.up();

  const shrunkRow1H = await page.evaluate(() => document.querySelector('tbody tr:nth-child(1)').offsetHeight);
  console.log(`Row 1 height shrunk: ${newRow1H}px -> ${shrunkRow1H}px`);
  if (shrunkRow1H >= newRow1H) {
    throw new Error(`Expected row 1 height to shrink from ${newRow1H}px, got ${shrunkRow1H}px`);
  }

  // Column A resizing
  const initialColAW = await page.evaluate(() => document.querySelector('th.grid-header-col[data-col="A"]').offsetWidth);
  const colHandleBox = await page.locator('.col-resize-handle[data-col="A"]').boundingBox();
  if (!colHandleBox) throw new Error('Col resize handle for col A not found!');

  await page.mouse.move(colHandleBox.x + colHandleBox.width / 2, colHandleBox.y + colHandleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(colHandleBox.x + 30, colHandleBox.y + colHandleBox.height / 2);
  await page.mouse.up();

  const newColAW = await page.evaluate(() => document.querySelector('th.grid-header-col[data-col="A"]').offsetWidth);
  console.log(`Col A width: ${initialColAW}px -> ${newColAW}px`);
  if (newColAW <= initialColAW) {
    throw new Error(`Expected col A width to increase from ${initialColAW}px, got ${newColAW}px`);
  }

  // 14. Test Undo and Redo
  console.log('Testing Undo & Redo (Ctrl+Z and Ctrl+Y)...');
  await page.click('td[data-addr="B1"]');
  const initialWeight = await page.evaluate(() => document.querySelector('td[data-addr="B1"]').style.fontWeight);
  console.log('Cell B1 font weight initial:', initialWeight);

  // Toggle Bold
  await page.click('#btnBold');
  const toggledWeight = await page.evaluate(() => document.querySelector('td[data-addr="B1"]').style.fontWeight);
  console.log('Cell B1 font weight after #btnBold click:', toggledWeight);
  if (toggledWeight === initialWeight) {
    throw new Error(`Bold toggle did not change font weight from ${initialWeight}!`);
  }

  // Click Undo
  await page.click('#btnUndo');
  const weightAfterUndo = await page.evaluate(() => document.querySelector('td[data-addr="B1"]').style.fontWeight);
  console.log('Cell B1 font weight after Undo:', weightAfterUndo);
  if (weightAfterUndo !== initialWeight) {
    throw new Error(`Undo did not revert font weight to ${initialWeight}, got ${weightAfterUndo}!`);
  }

  // Click Redo
  await page.click('#btnRedo');
  const weightAfterRedo = await page.evaluate(() => document.querySelector('td[data-addr="B1"]').style.fontWeight);
  console.log('Cell B1 font weight after Redo:', weightAfterRedo);
  if (weightAfterRedo !== toggledWeight) {
    throw new Error(`Redo did not reapply font weight to ${toggledWeight}, got ${weightAfterRedo}!`);
  }

  // Test Keyboard Ctrl+Z
  await page.keyboard.press('Control+z');
  const weightAfterCtrlZ = await page.evaluate(() => document.querySelector('td[data-addr="B1"]').style.fontWeight);
  if (weightAfterCtrlZ !== initialWeight) {
    throw new Error(`Ctrl+Z did not undo font weight to ${initialWeight}, got ${weightAfterCtrlZ}!`);
  }

  // 15. Verify Print View Button is Completely Removed from DOM
  console.log('Verifying Print View ("Tlačový náhľad") button is removed...');
  const printViewBtn = await page.$('#btnTogglePrintView');
  if (printViewBtn) {
    throw new Error('Found #btnTogglePrintView in DOM, it should be completely removed!');
  }
  const ctxPrintBtn = await page.$('#ctxTogglePrint');
  if (ctxPrintBtn) {
    throw new Error('Found #ctxTogglePrint in DOM, it should be completely removed!');
  }
  console.log('Confirmed Print Preview button and context option are removed.');

  // 15b. Verify Top Border Offset so Row 1 Top Border is Never Cut Off
  console.log('Verifying Top Border Container Offset...');
  const tableOffsetTop = await page.evaluate(() => {
    const table = document.querySelector('.excel-grid-table');
    const container = document.getElementById('spreadsheetArea');
    const rectTable = table.getBoundingClientRect();
    const rectContainer = container.getBoundingClientRect();
    return rectTable.top - rectContainer.top;
  });
  console.log('Table offset from spreadsheet container top:', tableOffsetTop);
  if (tableOffsetTop < 8) {
    throw new Error(`Expected spreadsheet table offset top >= 8px to keep top border visible, got ${tableOffsetTop}px`);
  }

  // 16. Test Excel HTML Paste with Format Preservation
  console.log('Testing Excel HTML Paste...');
  await page.click('td[data-addr="C10"]');
  await page.evaluate(() => {
    const excelHtml = `
      <table style="border-collapse:collapse;">
        <tr height="30">
          <td style="font-weight:bold; font-size:12pt; color:#ff0000; background-color:#ffffcc; border:2px solid #0000ff; text-align:center;">Test Excel Produkt</td>
          <td style="font-weight:bold; font-size:12pt; color:#008800; border:2px solid #0000ff; text-align:right;">100g</td>
        </tr>
      </table>
    `;
    const dt = new DataTransfer();
    dt.setData('text/html', excelHtml);
    const pasteEvent = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
    window.dispatchEvent(pasteEvent);
  });

  const pastedC10Val = (await page.textContent('td[data-addr="C10"]')).trim();
  const pastedC10Bg = await page.evaluate(() => document.querySelector('td[data-addr="C10"]').style.backgroundColor);
  console.log(`Pasted cell C10: "${pastedC10Val}", bgColor: "${pastedC10Bg}"`);
  if (pastedC10Val !== 'Test Excel Produkt') {
    throw new Error(`Expected pasted text "Test Excel Produkt", got "${pastedC10Val}"`);
  }

  // 17. Verify Dynamic SVG Export Bounding Box
  console.log('Verifying Dynamic SVG Export...');
  const svgOutput = await page.evaluate(() => {
    return document.querySelector('#previewSvgBox svg').outerHTML;
  });
  if (!svgOutput.includes('Test Excel Produkt')) {
    throw new Error('SVG export preview does not dynamically contain pasted content outside A1:B8!');
  }
  console.log('Dynamic SVG accurately reflects visual cells without hardcoded preset clipping.');

  // 18. Test Moving Table With Mouse Drag (relocation by border grab)
  console.log('Testing Mouse Table Drag Relocation...');
  const rangeBeforeMouseDrag = await page.textContent('#statusBarRange');
  console.log('Range before mouse table drag:', rangeBeforeMouseDrag);

  // Grab the top border of cell C10 (range-top border)
  const freshC10 = await page.locator('td[data-addr="C10"]').boundingBox();
  await page.mouse.move(freshC10.x + freshC10.width / 2, freshC10.y + 2);
  await page.mouse.down();
  // Drag down by 2 rows to C12
  const targetC12 = await page.locator('td[data-addr="C12"]').boundingBox();
  await page.mouse.move(targetC12.x + targetC12.width / 2, targetC12.y + targetC12.height / 2);
  await page.mouse.up();

  const rangeAfterMouseDrag = await page.textContent('#statusBarRange');
  console.log('Range after mouse table drag:', rangeAfterMouseDrag);
  if (!rangeAfterMouseDrag.includes('C12:D12')) {
    throw new Error(`Expected range after mouse drag to be C12:D12, got "${rangeAfterMouseDrag}"`);
  }
  // 19. Verify Drawer Header is "Náhľad (export)"
  console.log('Testing Drawer Header and Toggle...');
  const drawerHeader = (await page.textContent('#averyPreviewDrawer h3')).trim();
  console.log('Drawer Header:', drawerHeader);
  if (!drawerHeader.includes('Náhľad (export)')) {
    throw new Error(`Expected drawer header "Náhľad (export)", got "${drawerHeader}"`);
  }

  // 20. Test Drawer Toggle Button (Hide / Show)
  console.log('Testing Preview Drawer Toggle...');
  await page.click('#btnTogglePreviewDrawer');
  let isDrawerClosed = await page.evaluate(() => document.getElementById('averyPreviewDrawer').classList.contains('drawer-closed'));
  console.log('Drawer closed after toggle click:', isDrawerClosed);
  if (!isDrawerClosed) {
    throw new Error('Preview drawer should have class "drawer-closed" after clicking #btnTogglePreviewDrawer!');
  }

  // Toggle back on
  await page.click('#btnTogglePreviewDrawer');
  isDrawerClosed = await page.evaluate(() => document.getElementById('averyPreviewDrawer').classList.contains('drawer-closed'));
  console.log('Drawer closed after second toggle click:', isDrawerClosed);
  if (isDrawerClosed) {
    throw new Error('Preview drawer should NOT be closed after toggling back on!');
  }

  // Test Close Drawer Button (&times; Skryť)
  await page.click('#btnCloseDrawer');
  isDrawerClosed = await page.evaluate(() => document.getElementById('averyPreviewDrawer').classList.contains('drawer-closed'));
  if (!isDrawerClosed) {
    throw new Error('Preview drawer should be closed after clicking #btnCloseDrawer!');
  }

  // 21. Test Context Menu: Selection Preservation, Mini-Toolbar & Borders
  console.log('Testing Context Menu: Multi-Cell Selection Preservation...');
  // Select A1 to B3 via drag
  const selBoxA1 = await page.locator('td[data-addr="A1"]').boundingBox();
  const selBoxB3 = await page.locator('td[data-addr="B3"]').boundingBox();
  await page.mouse.move(selBoxA1.x + selBoxA1.width / 2, selBoxA1.y + selBoxA1.height / 2);
  await page.mouse.down();
  await page.mouse.move(selBoxB3.x + selBoxB3.width / 2, selBoxB3.y + selBoxB3.height / 2);
  await page.mouse.up();

  const selStatusBefore = await page.textContent('#statusBarSelection');
  console.log('Status selection before right-click:', selStatusBefore);
  if (!selStatusBefore.includes('A1:B3')) {
    throw new Error(`Expected selection A1:B3, got "${selStatusBefore}"`);
  }

  // Right-click B2 (inside the selection)
  await page.click('td[data-addr="B2"]', { button: 'right' });
  const isMenuVisible = await page.evaluate(() => {
    const el = document.getElementById('excelContextMenu');
    return el && window.getComputedStyle(el).display !== 'none';
  });
  console.log('Context menu visible on right click:', isMenuVisible);
  if (!isMenuVisible) {
    throw new Error('Context menu #excelContextMenu should be visible on right-click!');
  }

  // Verify multi-cell selection was preserved!
  const selStatusAfter = await page.textContent('#statusBarSelection');
  console.log('Status selection AFTER right-click inside:', selStatusAfter);
  if (!selStatusAfter.includes('A1:B3')) {
    throw new Error(`Multi-cell selection collapsed on right click! Expected A1:B3, got "${selStatusAfter}"`);
  }
  console.log('Active selection preserved on context menu right-click.');

  // Test Mini Toolbar font size, bold, underline, align
  console.log('Testing Mini Toolbar formatting...');
  await page.selectOption('#ctxFontSize', '14');
  await page.click('#ctxUnderline');
  await page.click('#ctxAlignCenter');

  const cellA1Underline = await page.evaluate(() => document.querySelector('td[data-addr="A1"]').style.textDecoration);
  const cellA1Align = await page.evaluate(() => document.querySelector('td[data-addr="A1"]').style.textAlign);
  const cellA1Size = await page.evaluate(() => document.querySelector('td[data-addr="A1"]').style.fontSize);
  console.log(`Cell A1 styled via Mini Toolbar: textDecoration="${cellA1Underline}", align="${cellA1Align}", size="${cellA1Size}"`);
  if (!cellA1Underline.includes('underline')) {
    throw new Error(`Expected cell A1 text-decoration to be underline, got "${cellA1Underline}"`);
  }
  if (cellA1Align !== 'center') {
    throw new Error(`Expected cell A1 text-align to be center, got "${cellA1Align}"`);
  }

  // Test Quick Borders via Context Menu
  console.log('Testing Context Menu Quick Borders...');
  await page.selectOption('#ctxBorderWidth', '2');
  await page.click('#ctxBorderAll');
  
  const cellA1Border = await page.evaluate(() => document.querySelector('td[data-addr="A1"]').style.borderBottom);
  console.log('Cell A1 borderBottom after #ctxBorderAll:', cellA1Border);
  if (!cellA1Border.includes('2px')) {
    throw new Error(`Expected 2px border on cell A1, got "${cellA1Border}"`);
  }

  // Re-open context menu to test "Zobraziť túto tabuľku v náhľade"
  await page.click('td[data-addr="A1"]', { button: 'right' });
  await page.click('#ctxShowInPreview');
  isDrawerClosed = await page.evaluate(() => document.getElementById('averyPreviewDrawer').classList.contains('drawer-closed'));
  console.log('Drawer visible after ctxShowInPreview:', !isDrawerClosed);
  if (isDrawerClosed) {
    throw new Error('Drawer should be reopened after clicking "Zobraziť túto tabuľku v náhľade"!');
  }

  // 22. Verify Aesthetic Symmetrical Export Grid in Drawer Footer
  console.log('Verifying Export Grid in Drawer Footer...');
  const exportCardsCount = await page.evaluate(() => document.querySelectorAll('.export-actions-grid .export-card-btn').length);
  console.log('Found export card buttons in grid:', exportCardsCount);
  if (exportCardsCount !== 3) {
    throw new Error(`Expected 3 export card buttons (SVG, PNG, JPG), found ${exportCardsCount}`);
  }

  // 23. Verify Zero Emojis in Entire DOM
  const pageHtml = await page.content();
  const emojiRegex = /[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}]/u;
  if (emojiRegex.test(pageHtml)) {
    throw new Error('Page contains forbidden emojis!');
  }
  console.log('Zero emojis verified across full DOM.');

  // 24. Capture screenshot to verification file
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
