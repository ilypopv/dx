'use strict';

// Builds small .docx packages from XML snippets, so the tests cover exactly the markup they talk about.
const fflate = require('../../media/vendor/fflate.js');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const enc = s => new TextEncoder().encode(s);

// A one-pixel PNG.
const PNG = Uint8Array.from(
    Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
        'base64'
    )
);

function makeDocx({ body, styles, numbering, footnotes, rels = [], sectPr } = {}) {
    const docRels = [
        ...(styles ? [`<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>`] : []),
        ...(numbering ? [`<Relationship Id="rIdNum" Type="${REL}/numbering" Target="numbering.xml"/>`] : []),
        ...(footnotes ? [`<Relationship Id="rIdFn" Type="${REL}/footnotes" Target="footnotes.xml"/>`] : []),
        ...rels
    ];
    const files = {
        '[Content_Types].xml': enc(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        ),
        '_rels/.rels': enc(
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
        ),
        'word/document.xml': enc(
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${NS}><w:body>${body || ''}${sectPr || ''}</w:body></w:document>`
        ),
        'word/_rels/document.xml.rels': enc(
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${docRels.join('')}</Relationships>`
        ),
        'word/media/image1.png': PNG
    };
    if (styles) {
        files['word/styles.xml'] = enc(`<?xml version="1.0" encoding="UTF-8"?><w:styles ${NS}>${styles}</w:styles>`);
    }
    if (numbering) {
        files['word/numbering.xml'] = enc(
            `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${NS}>${numbering}</w:numbering>`
        );
    }
    if (footnotes) {
        files['word/footnotes.xml'] = enc(
            `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${NS}>${footnotes}</w:footnotes>`
        );
    }
    return fflate.zipSync(files);
}

function documentXml(bytes) {
    return new TextDecoder().decode(fflate.unzipSync(bytes)['word/document.xml']);
}

// The plain text of a model paragraph.
const textOf = block => block.runs.map(r => r.text || '').join('');

module.exports = { makeDocx, documentXml, textOf, PNG };
