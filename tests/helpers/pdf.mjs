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
