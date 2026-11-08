export function buildTestPdf({ text = true } = {}) {
  const stream = text ? [
    'BT /F2 20 Tf 1 0 0 1 203 730 Tm (OpenThesis Template) Tj ET',
    'BT /F1 12 Tf 1 0 0 1 72 680 Tm (1 Introduction) Tj ET',
    'BT /F1 12 Tf 1 0 0 1 72 660 Tm (This reference line describes the template layout and typography.) Tj ET',
    'BT /F1 12 Tf 1 0 0 1 72 640 Tm (A second line makes body text the dominant inferred style.) Tj ET',
    'BT /F1 12 Tf 1 0 0 1 72 620 Tm (A third line provides stable margins and line spacing.) Tj ET',
  ].join('\n') : '';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${Buffer.byteLength(stream)} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>',
  ];
  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  pdf += offsets.slice(1).map(offset => `${String(offset).padStart(10, '0')} 00000 n \n`).join('');
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(pdf);
}

/**
 * A PDF laid out line by line, so a test can put text exactly where the
 * inference should measure it — an indented first line, a running head in the
 * header band, a table-of-contents line with dot leaders.
 *
 * `font` picks one of the base-14 faces. They are the faces PDF.js reports by
 * their PostScript names (`Helvetica-Bold`), which is what the parser reads to
 * recover bold and italic; an embedded font reports `sans-serif` instead.
 */
export function buildLinedPdf({ pages = [[]], shapes = [], width = 612, height = 792 } = {}) {
  const baseFonts = {
    regular: ['F1', 'Helvetica'],
    bold: ['F2', 'Helvetica-Bold'],
    serif: ['F3', 'Times-Roman'],
    italic: ['F4', 'Helvetica-Oblique'],
  };
  const fontKeys = Object.keys(baseFonts);
  const pageCount = pages.length || 1;
  const pageNumbers = [];
  const contentNumbers = [];
  for (let index = 0; index < pageCount; index += 1) {
    pageNumbers.push(3 + index * 2);
    contentNumbers.push(4 + index * 2);
  }
  const firstFont = 3 + pageCount * 2;
  const fontNumbers = {};
  fontKeys.forEach((key, index) => { fontNumbers[key] = firstFont + index; });

  const objects = [];
  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  objects[2] = `<< /Type /Pages /Kids [${pageNumbers.map(number => `${number} 0 R`).join(' ')}] /Count ${pageCount} >>`;

  const fontResources = fontKeys
    .map(key => `/${baseFonts[key][0]} ${fontNumbers[key]} 0 R`)
    .join(' ');
  pages.forEach((lines, index) => {
    const drawing = (shapes[index] ?? []).map(shapeStream).join('\n');
    const text = lines.map(line => {
      const [id] = baseFonts[line.font || 'regular'] || baseFonts.regular;
      return `BT /${id} ${line.size ?? 12} Tf 1 0 0 1 ${line.x ?? 72} ${line.y ?? 700} Tm (${escapePdfText(line.text)}) Tj ET`;
    }).join('\n');
    const stream = [drawing, text].filter(Boolean).join('\n');
    objects[pageNumbers[index]] = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${width} ${height}] /Resources << /Font << ${fontResources} >> >> /Contents ${contentNumbers[index]} 0 R >>`;
    objects[contentNumbers[index]] = `<< /Length ${Buffer.byteLength(stream)} >>\nstream\n${stream}\nendstream`;
  });

  fontKeys.forEach((key, index) => {
    objects[fontNumbers[key]] = `<< /Type /Font /Subtype /Type1 /BaseFont /${baseFonts[key][1]} >>`;
  });

  const total = objects.length - 1;
  let pdf = '%PDF-1.4\n';
  const offsets = [];
  for (let number = 1; number <= total; number += 1) {
    offsets[number] = Buffer.byteLength(pdf);
    pdf += `${number} 0 obj\n${objects[number]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${total + 1}\n0000000000 65535 f \n`;
  pdf += offsets.slice(1).map(offset => `${String(offset).padStart(10, '0')} 00000 n \n`).join('');
  pdf += `trailer\n<< /Size ${total + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(pdf);
}

function escapePdfText(text) {
  return text.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
}

/**
 * A drawn shape for `buildLinedPdf` — the table rules a PDF-inférence test
 * needs. `line` strokes a segment; `rect` fills (or strokes) a rectangle, which
 * is how Word draws a shaded header band.
 */
function shapeStream(shape) {
  const color = pdfColor(shape.color ?? '#000000');
  if (shape.type === 'line') {
    const width = shape.width ?? 1;
    return `${width} w ${color} RG ${shape.x1} ${shape.y1} m ${shape.x2} ${shape.y2} l S`;
  }
  const fill = shape.fill !== false;
  const paint = fill ? 'f' : 'S';
  return `${color} ${fill ? 'rg' : 'RG'} ${shape.x} ${shape.y} ${shape.w} ${shape.h} re ${paint}`;
}

function pdfColor(value) {
  const hex = value.replace('#', '');
  const full = hex.length === 3 ? hex.split('').map(char => char + char).join('') : hex;
  const [red, green, blue] = [0, 2, 4].map(offset => (parseInt(full.slice(offset, offset + 2), 16) / 255).toFixed(3));
  return `${red} ${green} ${blue}`;
}