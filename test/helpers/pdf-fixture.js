'use strict';

// Writes a small, valid PDF: one Helvetica text line per page, with a correct cross-reference table.
function makePdf(lines) {
    const objects = [];
    const add = body => objects.push(body) && objects.length;
    const catalog = add('');
    const pagesId = add('');
    const font = add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    const kids = [];
    for (const line of lines) {
        const text = line.replace(/[\\()]/g, m => '\\' + m);
        const stream = `BT /F1 24 Tf 72 720 Td (${text}) Tj ET`;
        const content = add(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
        kids.push(
            add(
                `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 ${font} 0 R >> >> /Contents ${content} 0 R >>`
            )
        );
    }
    objects[catalog - 1] = `<< /Type /Catalog /Pages ${pagesId} 0 R >>`;
    objects[pagesId - 1] = `<< /Type /Pages /Kids [${kids.map(k => k + ' 0 R').join(' ')}] /Count ${kids.length} >>`;
    let out = '%PDF-1.4\n';
    const offsets = [];
    objects.forEach((body, i) => {
        offsets.push(out.length);
        out += `${i + 1} 0 obj\n${body}\nendobj\n`;
    });
    const xref = out.length;
    out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (const offset of offsets) {
        out += String(offset).padStart(10, '0') + ' 00000 n \n';
    }
    out += `trailer\n<< /Size ${objects.length + 1} /Root ${catalog} 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
    return Buffer.from(out, 'latin1');
}

module.exports = { makePdf };
