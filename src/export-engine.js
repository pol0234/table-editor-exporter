export function sanitizeFilename(name, extension) {
  const clean = String(name || 'vyzivove-udaje')
    .trim()
    .replace(/[<>:"/\\|?*]/g, '-')
    .replace(/\s+/g, '_')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '');
  return `${clean || 'vyzivove-udaje'}.${extension}`;
}

export function getDpiDimensions(width, height, scale = 4) {
  const labels = {
    1: '72 DPI (Náhľad)',
    2: '150 DPI (Stredná kvalita)',
    4: '300 DPI (Tlač)',
    6: '600 DPI (Ultra detailná tlač)'
  };
  return {
    width: Math.round(width * scale),
    height: Math.round(height * scale),
    scale,
    dpiLabel: labels[scale] || `${scale}x`
  };
}

export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function downloadSVG(svgString, filename = 'tabulka.svg') {
  const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
  downloadBlob(blob, sanitizeFilename(filename, 'svg'));
}

export function renderSVGToCanvas(svgString, scale = 4, background = 'transparent') {
  return new Promise((resolve, reject) => {
    // If running in browser with DOMParser
    if (typeof DOMParser === 'undefined' || typeof Image === 'undefined') {
      return reject(new Error('Canvas rendering requires browser environment'));
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(svgString, 'image/svg+xml');
    const svgElem = doc.querySelector('svg');
    const width = parseFloat(svgElem ? svgElem.getAttribute('width') : '320') || 320;
    const height = parseFloat(svgElem ? svgElem.getAttribute('height') : '240') || 240;

    const targetWidth = Math.round(width * scale);
    const targetHeight = Math.round(height * scale);

    const canvas = document.createElement('canvas');
    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext('2d');

    // Fill background if needed
    if (background && background !== 'transparent') {
      ctx.fillStyle = background === 'white' ? '#ffffff' : background;
      ctx.fillRect(0, 0, targetWidth, targetHeight);
    }

    const img = new Image();
    const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    img.onload = () => {
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, 0, 0, targetWidth, targetHeight);
      URL.revokeObjectURL(url);
      resolve(canvas);
    };

    img.onerror = (e) => {
      URL.revokeObjectURL(url);
      reject(e);
    };

    img.src = url;
  });
}

export async function downloadPNG(svgString, filename, scale = 4, background = 'transparent') {
  const canvas = await renderSVGToCanvas(svgString, scale, background);
  return new Promise((resolve) => {
    canvas.toBlob((blob) => {
      if (blob) {
        downloadBlob(blob, sanitizeFilename(filename, 'png'));
      }
      resolve();
    }, 'image/png');
  });
}

export async function downloadJPG(svgString, filename, scale = 4) {
  // JPG always needs solid white background
  const canvas = await renderSVGToCanvas(svgString, scale, 'white');
  return new Promise((resolve) => {
    canvas.toBlob((blob) => {
      if (blob) {
        downloadBlob(blob, sanitizeFilename(filename, 'jpg'));
      }
      resolve();
    }, 'image/jpeg', 0.95);
  });
}

export async function copyImageToClipboard(svgString, scale = 4, background = 'white') {
  if (typeof navigator === 'undefined' || !navigator.clipboard || !navigator.clipboard.write || typeof ClipboardItem === 'undefined') {
    throw new Error('Prehliadač nepodporuje priame kopírovanie obrázkov do schránky.');
  }
  const canvas = await renderSVGToCanvas(svgString, scale, background);
  return new Promise((resolve, reject) => {
    canvas.toBlob(async (blob) => {
      if (!blob) {
        reject(new Error('Chyba pri generovaní obrázka pre schránku.'));
        return;
      }
      try {
        const item = new ClipboardItem({ 'image/png': blob });
        await navigator.clipboard.write([item]);
        resolve(true);
      } catch (err) {
        reject(err);
      }
    }, 'image/png');
  });
}
