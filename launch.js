// launch.js - Launches Tabuľkový editor as a true standalone PWA window
import { spawn } from 'child_process';
import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PORT = 4321;
const APP_URL = `http://localhost:${PORT}`;

// Check if server is already running
function checkServer(cb) {
  const req = http.get(APP_URL, (res) => {
    cb(true);
  });
  req.on('error', () => {
    cb(false);
  });
  req.setTimeout(500, () => {
    req.destroy();
    cb(false);
  });
}

function findBrowser() {
  const candidates = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

function launchBrowser() {
  const browser = findBrowser();
  const profileDir = path.join(process.env.LOCALAPPDATA || path.join(process.env.USERPROFILE, 'AppData', 'Local'), 'TabulkovyEditorProfile');

  if (browser) {
    const args = [
      `--app=${APP_URL}`,
      `--user-data-dir=${profileDir}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--window-size=1400,900'
    ];
    const child = spawn(browser, args, {
      detached: true,
      stdio: 'ignore'
    });
    child.unref();
  } else {
    // Fallback to default browser
    spawn('cmd.exe', ['/c', 'start', APP_URL], { detached: true, stdio: 'ignore' }).unref();
  }
}

// Check server status, start if needed, then launch browser window
checkServer((running) => {
  if (running) {
    launchBrowser();
    process.exit(0);
  } else {
    // Start server.js in background
    const serverScript = path.join(__dirname, 'server.js');
    const srv = spawn(process.execPath, [serverScript], {
      detached: true,
      stdio: 'ignore'
    });
    srv.unref();

    // Wait 300ms for server to bind then launch browser
    setTimeout(() => {
      launchBrowser();
      process.exit(0);
    }, 350);
  }
});
