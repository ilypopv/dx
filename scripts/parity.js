'use strict';

// Compares the pages of dx's layout engine with Word's: node scripts/parity.js <file.docx> <file.pdf> [-v]
// The PDF is the document exported by Word (File > Save As > PDF); the check finds where every page of the PDF
// starts in the text of the document and compares it with where the engine starts its pages. The engine measures
// with the font files installed on the system, as the webview does: the fonts that only Word ships are not there,
// a missing font falls back to Times or Helvetica like the browser's serif and sans-serif.
const fs = require('fs');
const path = require('path');
const { readFont } = require('./ttf');
const { readDocx } = require('../src/core/docx');
const { layoutDocument } = require('../src/core/layout');

const [file, pdf] = process.argv.slice(2).filter(a => !a.startsWith('-'));
const verbose = process.argv.includes('-v');
if (!file || !pdf) {
    console.error('usage: node scripts/parity.js <file.docx> <exported-by-word.pdf> [-v]');
    process.exit(2);
}

const FONT_DIRS = [
    '/System/Library/Fonts/Supplemental',
    '/System/Library/Fonts',
    '/Library/Fonts',
    path.join(require('os').homedir(), 'Library/Fonts'),
    'C:/Windows/Fonts',
    '/usr/share/fonts/truetype/msttcorefonts'
];
const fontFiles = new Map();
for (const dir of FONT_DIRS) {
    try {
        for (const name of fs.readdirSync(dir)) {
            if (/\.(ttf|ttc|otf)$/i.test(name)) {
                const key = name.toLowerCase().replace(/\.(ttf|ttc|otf)$/, '');
                if (!fontFiles.has(key)) {
                    fontFiles.set(key, path.join(dir, name));
                }
            }
        }
    } catch {
        // Not on this system.
    }
}
const fonts = new Map();
function fontOf(font) {
    const family = (font.family || 'Times New Roman').toLowerCase();
    const style = font.b && font.i ? ' bold italic' : font.b ? ' bold' : font.i ? ' italic' : '';
    const key = family + style;
    if (!fonts.has(key)) {
        const candidates = [family + style, family.replace(/ /g, '') + style.replace(/ /g, ''), family];
        // The browser's fallback for a family it does not have (the generic family of fontStack in the webview).
        const generic = /times|georgia|cambria|garamond|serif|book|minion|palatino/i.test(family)
            ? 'times'
            : 'helvetica';
        const found =
            candidates.map(c => fontFiles.get(c)).find(Boolean) ||
            fontFiles.get(generic) ||
            fontFiles.get('times new roman');
        fonts.set(key, readFont(found));
    }
    return fonts.get(key);
}
const measure = {
    width(text, font) {
        const f = fontOf(font);
        let w = 0;
        for (const ch of text) {
            w += f.width(ch.codePointAt(0));
        }
        return w * font.size;
    },
    metrics(font) {
        const { upm, metrics } = fontOf({ family: font.family });
        const k = font.size / upm;
        return {
            ascent: (metrics.ascender + metrics.lineGap) * k,
            descent: -metrics.descender * k,
            natural: (metrics.ascender - metrics.descender + metrics.lineGap) * k
        };
    }
};

async function wordPages() {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const doc = await pdfjs.getDocument({
        data: new Uint8Array(fs.readFileSync(pdf)),
        disableFontFace: true,
        verbosity: 0
    }).promise;
    const pages = [];
    for (let i = 1; i <= doc.numPages; i++) {
        const content = await (await doc.getPage(i)).getTextContent();
        pages.push(content.items.map(it => it.str).join(''));
    }
    return pages;
}

const squash = s => s.replace(/\s+/g, '');
const charsOf = runs => runs.flatMap(r => (r.br || r.image ? [null] : [...(r.text || '')]));

(async () => {
    const { model } = readDocx(new Uint8Array(fs.readFileSync(file)));
    const t0 = Date.now();
    const { pages } = layoutDocument(model, measure);
    const ms = Date.now() - t0;
    const word = await wordPages();

    // The text of the body without spaces, and where every paragraph character and table row is in it.
    let text = '';
    const paraPos = [];
    const rowPos = [];
    model.blocks.forEach((b, i) => {
        if (b.type === 'p') {
            paraPos[i] = [];
            for (const ch of charsOf(b.runs)) {
                paraPos[i].push(text.length);
                if (ch && !/\s/.test(ch)) {
                    text += ch;
                }
            }
            paraPos[i].push(text.length);
        } else if (b.type === 'table') {
            rowPos[i] = [];
            for (const row of b.rows) {
                rowPos[i].push(text.length);
                for (const c of row) {
                    for (const x of c.blocks) {
                        if (x.runs) {
                            text += squash(charsOf(x.runs).filter(Boolean).join(''));
                        }
                    }
                }
            }
            rowPos[i].push(text.length);
        }
    });

    let same = 0;
    let from = 0;
    const report = [];
    const count = Math.max(pages.length, word.length);
    for (let p = 0; p < count; p++) {
        const page = pages[p];
        const first = page && page.parts.find(q => !q.empty);
        let dx = -1;
        let head = '';
        if (first) {
            const b = model.blocks[first.block];
            dx = first.rows
                ? rowPos[first.block][first.rows[0]]
                : b.type === 'p'
                  ? paraPos[first.block][first.from]
                  : -1;
            if (first.rows && first.repeat) {
                // Word repeats the header rows of a table on the pages it continues on.
                head = text.slice(rowPos[first.block][0], rowPos[first.block][first.repeat]);
            }
        }
        let body = squash(word[p] || '');
        // The page number of the header comes first in Word's text.
        const number = page && String(page.number);
        if (number && body.startsWith(number)) {
            body = body.slice(number.length);
        }
        if (head && body.startsWith(head)) {
            body = body.slice(head.length);
        }
        let ok;
        if (!word[p] || !page) {
            ok = false;
        } else if (!body) {
            // A page with a picture only.
            ok = !!first && page.parts.filter(q => !q.empty).length === 1;
        } else {
            let at = -1;
            for (let len = 80; len >= 10 && at < 0; len -= 5) {
                const probe = body.slice(0, len);
                for (const variant of [probe, probe.replace(/^(\d+(\.\d+)*[.)]|[•▪◦–-])/, '')]) {
                    const q = text.indexOf(variant, from);
                    if (q >= 0 && at < 0) {
                        at = q;
                    }
                }
            }
            ok = at >= 0 && Math.abs(at - dx) <= 1;
            if (at >= 0) {
                from = at + 1;
            }
            if (!ok) {
                report.push(
                    `  page ${p + 1}: Word ${JSON.stringify(text.slice(at, at + 40))}, dx ${JSON.stringify(text.slice(dx, dx + 40))}`
                );
            }
        }
        same += ok ? 1 : 0;
    }
    console.log(
        `dx ${pages.length} pages, Word ${word.length}; ${same} of ${count} pages start where Word's do (layout ${ms} ms)`
    );
    if (report.length) {
        console.log((verbose ? report : report.slice(0, 15)).join('\n'));
    }
    process.exitCode = same === count ? 0 : 1;
})();
