export function escapeXml(unsafe) {
  return String(unsafe || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export function renderTableToSVG(product, options = {}) {
  const styles = product.styles || {};
  const margin = options.margin ?? 1;
  const background = options.background ?? 'transparent';
  
  const tableWidth = styles.tableWidth || 320;
  const colRatio = styles.colRatio || 0.68;
  const col1Width = Math.round(tableWidth * colRatio);
  const col2Width = tableWidth - col1Width;
  
  const fontSize = styles.fontSize || 13;
  const lineHeight = fontSize * 1.25;
  const cellPaddingY = styles.cellPaddingY ?? 4;
  const cellPaddingX = styles.cellPaddingX ?? 8;
  const outerBorderWidth = styles.outerBorderWidth ?? 1.5;
  const outerBorderColor = styles.outerBorderColor || '#000000';
  const innerBorderColor = styles.innerBorderColor || '#000000';
  const fontFamily = styles.fontFamily || 'Calibri, Arial, sans-serif';
  
  // Calculate row positions and heights
  let currentY = margin;
  
  // Header Row
  const headerLinesCount = 1;
  const headerHeight = Math.max(26, headerLinesCount * lineHeight + cellPaddingY * 2);
  const headerY = currentY;
  currentY += headerHeight;
  
  const computedRows = [];
  for (const row of product.rows) {
    const valLines = String(row.value || '').split('\n');
    const nameLines = String(row.name || '').split('\n');
    const maxLines = Math.max(valLines.length, nameLines.length, 1);
    const multiplier = row.heightMultiplier || 1.0;
    const baseH = Math.max(22, maxLines * lineHeight + cellPaddingY * 2);
    const rowH = Math.round(baseH * multiplier);
    
    computedRows.push({
      row,
      y: currentY,
      height: rowH,
      valLines,
      nameLines
    });
    currentY += rowH;
  }
  
  const tableHeight = currentY - margin;
  const totalSvgWidth = tableWidth + margin * 2;
  const totalSvgHeight = tableHeight + margin * 2;
  
  // Build SVG XML
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${totalSvgWidth} ${totalSvgHeight}" width="${totalSvgWidth}" height="${totalSvgHeight}">\n`;
  svg += `<defs>\n<style>\n`;
  svg += `  text { font-family: ${fontFamily}; font-size: ${fontSize}px; fill: #000000; }\n`;
  svg += `  .bold { font-weight: bold; }\n`;
  svg += `  .header-text { font-weight: bold; text-anchor: middle; }\n`;
  svg += `  .val-text { text-anchor: end; }\n`;
  svg += `  .indent { transform: translateX(18px); }\n`;
  svg += `</style>\n</defs>\n`;
  
  if (background && background !== 'transparent') {
    const fill = background === 'white' ? '#ffffff' : background;
    svg += `<rect id="bg-rect" x="0" y="0" width="${totalSvgWidth}" height="${totalSvgHeight}" fill="${fill}" />\n`;
  }
  
  const tableX = margin;
  const tableY = margin;
  
  // Background of table itself if transparent canvas but table has fill
  if (styles.backgroundColor && styles.backgroundColor !== 'transparent') {
    svg += `<rect x="${tableX}" y="${tableY}" width="${tableWidth}" height="${tableHeight}" fill="${styles.backgroundColor}" />\n`;
  }
  
  // Header text
  const headerMidY = headerY + headerHeight / 2 + fontSize * 0.35;
  svg += `<!-- Header Row -->\n`;
  svg += `<text x="${tableX + col1Width / 2}" y="${headerMidY}" class="header-text">${escapeXml(product.header.title)}</text>\n`;
  svg += `<text x="${tableX + col1Width + col2Width / 2}" y="${headerMidY}" class="header-text">${escapeXml(product.header.unit)}</text>\n`;
  
  // Header Bottom border
  const headerBorderThickness = styles.headerBorderBottom === 'thick' ? 2 : (styles.headerBorderBottom === 'thin' ? 1 : (styles.headerBorderBottom === 'double' ? 3 : 0));
  if (headerBorderThickness > 0) {
    svg += `<line x1="${tableX}" y1="${headerY + headerHeight}" x2="${tableX + tableWidth}" y2="${headerY + headerHeight}" stroke="${innerBorderColor}" stroke-width="${headerBorderThickness}" />\n`;
  }
  
  // Rows
  svg += `<!-- Data Rows -->\n`;
  for (let i = 0; i < computedRows.length; i++) {
    const item = computedRows[i];
    const r = item.row;
    const rY = item.y;
    const rH = item.height;
    
    // Text positioning (vertically centered in row or aligned with padding)
    const textStartY = rY + cellPaddingY + fontSize * 0.9;
    
    // Label
    const indentX = r.indent ? 16 : 0;
    const labelX = tableX + cellPaddingX + indentX;
    const labelClass = r.bold ? 'bold' : '';
    
    if (item.nameLines.length === 1) {
      svg += `<text x="${labelX}" y="${textStartY}" class="${labelClass}">${escapeXml(item.nameLines[0])}</text>\n`;
    } else {
      svg += `<text x="${labelX}" y="${textStartY}" class="${labelClass}">`;
      item.nameLines.forEach((line, lineIdx) => {
        svg += `<tspan x="${labelX}" dy="${lineIdx === 0 ? 0 : lineHeight}">${escapeXml(line)}</tspan>`;
      });
      svg += `</text>\n`;
    }
    
    // Value (Right aligned)
    const valX = tableX + tableWidth - cellPaddingX;
    const valClass = 'val-text' + (r.bold ? ' bold' : '');
    if (item.valLines.length === 1) {
      svg += `<text x="${valX}" y="${textStartY}" class="${valClass}">${escapeXml(item.valLines[0])}</text>\n`;
    } else {
      svg += `<text x="${valX}" y="${textStartY}" class="${valClass}">`;
      item.valLines.forEach((line, lineIdx) => {
        svg += `<tspan x="${valX}" dy="${lineIdx === 0 ? 0 : lineHeight}">${escapeXml(line)}</tspan>`;
      });
      svg += `</text>\n`;
    }
    
    // Horizontal row bottom border
    const borderType = r.borderBottom || 'thin';
    if (borderType !== 'none' && i < computedRows.length - 1) {
      const bStroke = borderType === 'thick' ? 2 : (borderType === 'double' ? 3 : 1);
      svg += `<line x1="${tableX}" y1="${rY + rH}" x2="${tableX + tableWidth}" y2="${rY + rH}" stroke="${innerBorderColor}" stroke-width="${bStroke}" />\n`;
    }
  }
  
  // Vertical column divider
  if (styles.verticalDivider) {
    const vStroke = styles.verticalDividerWidth || 1;
    svg += `<!-- Vertical Divider -->\n`;
    svg += `<line class="col-divider" x1="${tableX + col1Width}" y1="${tableY}" x2="${tableX + col1Width}" y2="${tableY + tableHeight}" stroke="${innerBorderColor}" stroke-width="${vStroke}" />\n`;
  }
  
  // Outer rectangle border
  if (outerBorderWidth > 0) {
    svg += `<!-- Outer Border -->\n`;
    svg += `<rect x="${tableX}" y="${tableY}" width="${tableWidth}" height="${tableHeight}" fill="none" stroke="${outerBorderColor}" stroke-width="${outerBorderWidth}" />\n`;
  }
  
  svg += `</svg>`;
  
  return {
    svgString: svg,
    width: totalSvgWidth,
    height: totalSvgHeight
  };
}
