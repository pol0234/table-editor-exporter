# Table Editor & Exporter

A professional, client-side spreadsheet editor and high-fidelity table exporter. Built for generating clean, publication-ready nutritional tables, product specifications, label assets, and technical documentation with zero backend dependencies.

---

## Overview

Table Editor & Exporter combines the familiar layout and ergonomics of Microsoft Excel with a dedicated vector and raster export engine. It runs entirely inside any modern web browser, operates 100% offline without local server dependencies, and can be hosted statically on GitHub Pages with native Progressive Web App (PWA) installation.

---

## Key Features

- **Authentic Excel Ergonomics:**
  - Standard green ribbon toolbar with categorized command groups (Písmo, Zarovnanie, Orámovanie, Tabuľka, Náhľad & Export).
  - Formula bar with live address coordinates and cell formula input.
  - 16 authentic Microsoft Excel border styles (all, outer, thick outer, top-and-bottom, top-thick-bottom, top-double-bottom, inner horizontal, inner vertical, none).
  - Typography options: Font families (Calibri, Arial, Segoe UI, Roboto), font sizes (8pt to 36pt), bold, italic, underline, alignment, and indentation.
  - Cell fill color and text color pickers with real-time grid and export synchronization.
  - Subtle gray gridlines (`#d4d4d4`) visible during editing and accurately reproduced in vector/raster exports. Solid white fill seamlessly masks underlying gridlines.

- **Grid Interaction & Editing:**
  - Multi-cell rectangular drag selection with active border outline and corner fill handle.
  - Full column and row selection by clicking header cells (`A`, `B`, `C`, row numbers `1`, `2`, `3`, or top-left corner).
  - Column and row insertion and deletion with automatic cell coordinate, column width, and export range recalculation.
  - Drag-to-resize column widths and row heights with visual guidelines.
  - Excel clipboard fidelity: Copy and paste directly between Microsoft Excel and the web grid while preserving formatting, cell boundaries, and line breaks.
  - Full Undo / Redo history (`Ctrl + Z` / `Ctrl + Y`).
  - Right-click context menu with mini-toolbar and quick formatting shortcuts.

- **Export Engine:**
  - **Direct Clipboard Copy:** Copies a crisp, 300 DPI high-resolution raster image directly into the Windows clipboard for instant pasting into labeling software (such as Avery Design & Print), Word, PowerPoint, or graphic design tools.
  - **Vector SVG:** Infinitely scalable vector export with exact typography and border stroke styles.
  - **High-Resolution PNG / JPG:** Crisp 300 DPI raster exports with user-configurable solid or transparent background.
  - **Live Preview Drawer:** Side drawer displaying a live SVG render with 2D auto-fit zoom, ensuring large tables adapt without clipping.

- **Multi-Sheet Management:**
  - Bottom sheet tab bar supporting multiple tables, sheet renaming, duplication, and addition.
  - Automatic persistent local storage saving across browser reloads.

---

## Local Usage

### Quick Start
1. Double-click `Spustiť-Aplikáciu.bat` or open `index.html` directly in Google Chrome, Microsoft Edge, or Mozilla Firefox.
2. The application opens immediately. No Node.js runtime, Python server, or internet connection is required.

### Desktop Shortcut
To create a clean desktop shortcut with custom icon on Windows:
1. Right-click `Vytvoriť-Odkaz-Na-Plochu.ps1` and select **Run with PowerShell**.
2. A shortcut named `Table Editor & Exporter` will be placed on your Desktop.

---

## GitHub Pages Deployment

The repository is structured as a static web application and is ready for 1-click deployment to GitHub Pages.

### Automatic Deployment (GitHub Actions)
1. Push this repository to GitHub:
   ```bash
   git init
   git remote add origin https://github.com/<your-username>/table-editor-exporter.git
   git add .
   git commit -m "Initial release of Table Editor & Exporter"
   git push -u origin master
   ```
2. In your repository on GitHub, go to **Settings** > **Pages**.
3. Under **Build and deployment** > **Source**, select **GitHub Actions**.
4. The included workflow (`.github/workflows/deploy.yml`) will automatically build and deploy the app to `https://<your-username>.github.io/table-editor-exporter/`.
5. Once hosted on GitHub Pages over HTTPS, users will see the native browser **Install App** button in Chrome/Edge to install Table Editor & Exporter as a desktop PWA with automatic updates.

---

## Security & Architecture

- **Zero Server Attack Surface:** 100% client-side application. No backend processes, APIs, or database connections.
- **Privacy First:** All spreadsheet data remains strictly local in the user's browser (`localStorage`). No telemetry, analytics, or external data transmission.
- **Dependency Minimalist:** Core grid rendering and SVG vector generation are implemented with zero runtime dependencies.
- **Automated Verification:** Fully covered by Playwright end-to-end automated UI test suite and unit tests.

---

## Development & Testing

```bash
# Install development dependencies (Playwright for verification)
npm install

# Run unit tests
npm test

# Run full Playwright automated UI verification
npm run test:ui
```

---

## License

MIT License. Open for personal, commercial, and open-source use.
