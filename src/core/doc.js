'use strict';

// A reader of the Word 97-2003 binary format (MS-DOC).
// The text comes from the piece table, the paragraph and character properties from the PAPX / CHPX
// formatted disk pages. The result is the same block model that the DOCX reader produces:
//   { blocks: [ { type: 'p', style, align, indent, list, runs: [ { text, b, i, u, s, sup, sub, size, color, href, br } ] }
//               | { type: 'table', rows: [ [ blocks ] ] }  (row -> cell -> blocks)
//               | { type: 'pagebreak' } ] }

const { isCfb, readCfb } = require('./cfb');

// Windows-1252 differs from Latin-1 in 0x80-0x9F; compressed pieces are stored in it.
const CP1252 = {
    0x80: 0x20ac,
    0x82: 0x201a,
    0x83: 0x0192,
    0x84: 0x201e,
    0x85: 0x2026,
    0x86: 0x2020,
    0x87: 0x2021,
    0x88: 0x02c6,
    0x89: 0x2030,
    0x8a: 0x0160,
    0x8b: 0x2039,
    0x8c: 0x0152,
    0x8e: 0x017d,
    0x91: 0x2018,
    0x92: 0x2019,
    0x93: 0x201c,
    0x94: 0x201d,
    0x95: 0x2022,
    0x96: 0x2013,
    0x97: 0x2014,
    0x98: 0x02dc,
    0x99: 0x2122,
    0x9a: 0x0161,
    0x9b: 0x203a,
    0x9c: 0x0153,
    0x9e: 0x017e,
    0x9f: 0x0178
};

// The colour palette of sprmCIco.
const ICO = [
    undefined,
    '#000000',
    '#0000ff',
    '#00ffff',
    '#00ff00',
    '#ff00ff',
    '#ff0000',
    '#ffff00',
    '#ffffff',
    '#000080',
    '#008080',
    '#008000',
    '#800080',
    '#800000',
    '#808000',
    '#808080',
    '#c0c0c0'
];

const ALIGN = ['left', 'center', 'right', 'justify'];

class DocError extends Error {}

function operandSize(sprm, bytes, offset) {
    switch (sprm >> 13) {
        case 0:
        case 1:
            return 1;
        case 2:
        case 4:
        case 5:
            return 2;
        case 3:
            return 4;
        case 7:
            return 3;
        default:
            // Variable length. sprmTDefTable and sprmPChgTabs carry a 16 bit / special size.
            if (sprm === 0xd608 || sprm === 0xd606) {
                return 2 + (bytes[offset] | (bytes[offset + 1] << 8)) - 1;
            }
            if (sprm === 0xc615) {
                const cb = bytes[offset];
                return cb === 255 ? 1 + bytes.length - offset : 1 + cb;
            }
            return 1 + bytes[offset];
    }
}

// Walks a grpprl and calls visit(sprm, offset) for every property modifier.
function eachSprm(bytes, start, end, visit) {
    let offset = start;
    while (offset + 2 <= end) {
        const sprm = bytes[offset] | (bytes[offset + 1] << 8);
        offset += 2;
        const size = operandSize(sprm, bytes, offset);
        if (offset + size > end) {
            break;
        }
        visit(sprm, offset, size);
        offset += size;
    }
}

function toggle(value) {
    // 0 off, 1 on, 0x80 keep the style value, 0x81 invert it: the styles are not resolved, so it means on.
    return value === 1 || value === 0x81 ? true : value === 0 ? false : undefined;
}

function readFib(word) {
    const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
    if (word.length < 0x200 || view.getUint16(0, true) !== 0xa5ec) {
        throw new DocError('The WordDocument stream has no valid file information block.');
    }
    const nFib = view.getUint16(2, true);
    const flags = view.getUint16(0x0a, true);
    if (flags & 0x0100) {
        throw new DocError('The document is encrypted or password protected.');
    }
    if (nFib < 101) {
        throw new DocError('Word 6.0/95 documents are not supported, only Word 97-2003.');
    }
    const csw = view.getUint16(32, true);
    const lwStart = 34 + csw * 2 + 2;
    const cslw = view.getUint16(lwStart - 2, true);
    const fcStart = lwStart + cslw * 4 + 2;
    const pair = index => ({
        fc: view.getUint32(fcStart + index * 8, true),
        lcb: view.getUint32(fcStart + index * 8 + 4, true)
    });
    return {
        table: flags & 0x0200 ? '1Table' : '0Table',
        ccpText: view.getUint32(lwStart + 3 * 4, true),
        stshf: pair(1),
        chpx: pair(12),
        papx: pair(13),
        clx: pair(33)
    };
}

function readPieces(table, clx) {
    const view = new DataView(table.buffer, table.byteOffset, table.byteLength);
    let offset = clx.fc;
    const end = clx.fc + clx.lcb;
    while (offset < end && table[offset] === 0x01) {
        offset += 3 + view.getInt16(offset + 1, true);
    }
    if (offset >= end || table[offset] !== 0x02) {
        throw new DocError('The document has no piece table.');
    }
    const lcb = view.getUint32(offset + 1, true);
    const base = offset + 5;
    const count = (lcb - 4) / 12;
    if (!Number.isInteger(count) || base + lcb > table.length) {
        throw new DocError('The piece table is damaged.');
    }
    const pieces = [];
    for (let i = 0; i < count; i++) {
        const raw = view.getUint32(base + (count + 1) * 4 + i * 8 + 2, true);
        const compressed = (raw & 0x40000000) !== 0;
        pieces.push({
            cpStart: view.getUint32(base + i * 4, true),
            cpEnd: view.getUint32(base + (i + 1) * 4, true),
            compressed,
            fc: compressed ? (raw & 0x3fffffff) / 2 : raw & 0x3fffffff
        });
    }
    return pieces;
}

// Decodes the main document text and remembers the file offset of every character,
// the formatting of a character is looked up by that offset.
function readText(word, pieces, ccpText) {
    const chars = [];
    const fcs = [];
    for (const piece of pieces) {
        if (piece.cpStart >= ccpText) {
            break;
        }
        const count = Math.min(piece.cpEnd, ccpText) - piece.cpStart;
        for (let i = 0; i < count; i++) {
            if (piece.compressed) {
                const fc = piece.fc + i;
                const byte = word[fc];
                chars.push(CP1252[byte] || byte);
                fcs.push(fc);
            } else {
                const fc = piece.fc + i * 2;
                chars.push(word[fc] | (word[fc + 1] << 8));
                fcs.push(fc);
            }
        }
    }
    return { chars, fcs };
}

// Reads every FKP of a PlcBte into a sorted list of [fcStart, fcEnd) ranges with their parsed properties.
function readFkps(word, table, plc, parseEntry) {
    const runs = [];
    if (!plc.lcb) {
        return runs;
    }
    const view = new DataView(table.buffer, table.byteOffset, table.byteLength);
    const count = (plc.lcb - 4) / 8;
    for (let i = 0; i < count; i++) {
        const pn = view.getUint32(plc.fc + (count + 1) * 4 + i * 4, true) & 0x3fffff;
        const page = pn * 512;
        if (page + 512 > word.length) {
            continue;
        }
        const fkp = word.subarray(page, page + 512);
        const fkpView = new DataView(fkp.buffer, fkp.byteOffset, 512);
        const crun = fkp[511];
        for (let r = 0; r < crun; r++) {
            runs.push({
                start: fkpView.getUint32(r * 4, true),
                end: fkpView.getUint32((r + 1) * 4, true),
                props: parseEntry(fkp, crun, r)
            });
        }
    }
    runs.sort((a, b) => a.start - b.start);
    return runs;
}

function findRun(runs, fc) {
    let low = 0;
    let high = runs.length - 1;
    while (low <= high) {
        const mid = (low + high) >> 1;
        if (fc < runs[mid].start) {
            high = mid - 1;
        } else if (fc >= runs[mid].end) {
            low = mid + 1;
        } else {
            return runs[mid].props;
        }
    }
    return undefined;
}

function parsePapx(fkp, crun, r) {
    const props = { istd: 0 };
    const offset = fkp[(crun + 1) * 4 + r * 13] * 2;
    if (!offset) {
        return props;
    }
    let cb = fkp[offset];
    let start = offset + 1;
    let size = cb * 2 - 1;
    if (cb === 0) {
        cb = fkp[offset + 1];
        start = offset + 2;
        size = cb * 2;
    }
    const end = Math.min(start + size, 511);
    props.istd = fkp[start] | (fkp[start + 1] << 8);
    eachSprm(fkp, start + 2, end, (sprm, at) => {
        switch (sprm) {
            case 0x2416:
                props.inTable = fkp[at] !== 0;
                break;
            case 0x2417:
                props.ttp = fkp[at] !== 0;
                break;
            case 0x6649:
                props.itap = fkp[at] | (fkp[at + 1] << 8);
                break;
            case 0x2403:
            case 0x2461:
                props.align = ALIGN[fkp[at]];
                break;
            case 0x460b:
                props.ilfo = fkp[at] | (fkp[at + 1] << 8);
                break;
            case 0x260a:
                props.ilvl = fkp[at];
                break;
            case 0x840f:
            case 0x845e: {
                const v = fkp[at] | (fkp[at + 1] << 8);
                props.indent = v > 0x7fff ? v - 0x10000 : v;
                break;
            }
        }
    });
    if (props.itap > 0) {
        props.inTable = true;
    }
    return props;
}

function parseChpx(fkp, crun, r) {
    const props = {};
    const offset = fkp[(crun + 1) * 4 + r] * 2;
    if (!offset) {
        return props;
    }
    const cb = fkp[offset];
    eachSprm(fkp, offset + 1, Math.min(offset + 1 + cb, 511), (sprm, at) => {
        const v = fkp[at];
        switch (sprm) {
            case 0x0835:
                props.b = toggle(v);
                break;
            case 0x0836:
                props.i = toggle(v);
                break;
            case 0x0837:
            case 0x2a53:
                props.s = toggle(v);
                break;
            case 0x083c:
                props.hidden = toggle(v);
                break;
            case 0x2a3e:
                props.u = v !== 0;
                break;
            case 0x2a48:
                props.sup = v === 1;
                props.sub = v === 2;
                break;
            case 0x4a43:
                props.size = (fkp[at] | (fkp[at + 1] << 8)) / 2;
                break;
            case 0x2a42:
                props.color = ICO[v];
                break;
            case 0x6870:
                if (!(fkp[at + 3] & 0xff)) {
                    props.color =
                        '#' + [fkp[at], fkp[at + 1], fkp[at + 2]].map(x => x.toString(16).padStart(2, '0')).join('');
                }
                break;
        }
    });
    return props;
}

// The sti of every style: 1-9 are the headings, 62 the title, 74 the subtitle.
function readStyles(table, stshf) {
    const styles = [];
    if (!stshf.lcb || stshf.fc + stshf.lcb > table.length) {
        return styles;
    }
    const view = new DataView(table.buffer, table.byteOffset, table.byteLength);
    const cbStshi = view.getUint16(stshf.fc, true);
    const cstd = view.getUint16(stshf.fc + 2, true);
    let offset = stshf.fc + 2 + cbStshi;
    const end = stshf.fc + stshf.lcb;
    for (let i = 0; i < cstd && offset + 2 <= end; i++) {
        const cbStd = view.getUint16(offset, true);
        styles.push(cbStd >= 2 ? view.getUint16(offset + 2, true) & 0x0fff : undefined);
        offset += 2 + cbStd;
    }
    return styles;
}

function styleName(sti) {
    if (sti >= 1 && sti <= 9) {
        return 'h' + Math.min(sti, 6);
    }
    if (sti === 62) {
        return 'title';
    }
    if (sti === 74) {
        return 'subtitle';
    }
    return undefined;
}

function sameFormat(a, b) {
    return (
        a.b === b.b &&
        a.i === b.i &&
        a.u === b.u &&
        a.s === b.s &&
        a.sup === b.sup &&
        a.sub === b.sub &&
        a.size === b.size &&
        a.color === b.color &&
        a.href === b.href
    );
}

function buildRuns(chars, fcs, start, end, chpx, fields) {
    const runs = [];
    let current;
    const push = (text, format) => {
        if (current && !current.br && sameFormat(current, format)) {
            current.text += text;
        } else {
            current = Object.assign({ text }, format);
            runs.push(current);
        }
    };
    for (let i = start; i < end; i++) {
        const c = chars[i];
        if (c === 0x13) {
            fields.push({ code: '', result: false, href: undefined });
            continue;
        }
        if (c === 0x14 && fields.length) {
            const field = fields[fields.length - 1];
            field.result = true;
            const link = /HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/i.exec(field.code);
            if (link && !/\\l/.test(field.code)) {
                field.href = link[1];
            }
            continue;
        }
        if (c === 0x15 && fields.length) {
            fields.pop();
            continue;
        }
        if (fields.length && !fields[fields.length - 1].result) {
            fields[fields.length - 1].code += String.fromCharCode(c);
            continue;
        }
        const props = findRun(chpx, fcs[i]) || {};
        if (props.hidden) {
            continue;
        }
        let text;
        if (c === 0x0b) {
            runs.push({ text: '', br: true });
            current = undefined;
            continue;
        } else if (c === 0x1e) {
            text = '‑';
        } else if (c === 0x09 || c >= 0x20) {
            text = String.fromCharCode(c);
        } else {
            // Pictures (0x01, 0x08), footnote and annotation references and other control characters.
            continue;
        }
        if (c === 0x1f || c === 0xad) {
            continue;
        }
        let href;
        for (let f = fields.length - 1; f >= 0 && !href; f--) {
            href = fields[f].href;
        }
        const format = {};
        for (const key of ['b', 'i', 'u', 's', 'sup', 'sub', 'size', 'color']) {
            if (props[key]) {
                format[key] = props[key];
            }
        }
        if (href) {
            format.href = href;
        }
        push(text, format);
    }
    return runs;
}

function readDoc(bytes) {
    if (!isCfb(bytes)) {
        throw new DocError('The file is not a Word 97-2003 document.');
    }
    const cfb = readCfb(bytes);
    const word = cfb.read('WordDocument');
    if (!word) {
        throw new DocError('The compound document has no WordDocument stream, it is not a Word document.');
    }
    const fib = readFib(word);
    const table = cfb.read(fib.table);
    if (!table) {
        throw new DocError(`The document has no ${fib.table} stream.`);
    }
    const pieces = readPieces(table, fib.clx);
    const { chars, fcs } = readText(word, pieces, fib.ccpText);
    const papx = readFkps(word, table, fib.papx, parsePapx);
    const chpx = readFkps(word, table, fib.chpx, parseChpx);
    const styles = readStyles(table, fib.stshf);

    const blocks = [];
    const fields = [];
    let row;
    let cell;
    let tableBlock;
    let start = 0;

    const closeTable = () => {
        if (tableBlock && row && (row.length || (cell && cell.length))) {
            if (cell && cell.length) {
                row.push(cell);
            }
            tableBlock.rows.push(row);
        }
        tableBlock = row = cell = undefined;
    };

    for (let i = 0; i < chars.length; i++) {
        const c = chars[i];
        if (c !== 0x0d && c !== 0x07 && c !== 0x0c) {
            continue;
        }
        const pap = findRun(papx, fcs[i]) || { istd: 0 };
        const runs = buildRuns(chars, fcs, start, i, chpx, fields);
        start = i + 1;

        if (pap.inTable) {
            if (!tableBlock) {
                tableBlock = { type: 'table', rows: [] };
                blocks.push(tableBlock);
                row = [];
                cell = [];
            }
            if (pap.ttp) {
                // The row end mark: a paragraph of its own after the last cell.
                tableBlock.rows.push(row);
                row = [];
                cell = [];
                continue;
            }
            cell.push(paragraph(runs, pap, styles));
            if (c === 0x07) {
                row.push(cell);
                cell = [];
            }
            continue;
        }

        closeTable();
        if (c === 0x0c && !runs.length) {
            blocks.push({ type: 'pagebreak' });
            continue;
        }
        blocks.push(paragraph(runs, pap, styles));
        if (c === 0x0c) {
            blocks.push({ type: 'pagebreak' });
        }
    }
    closeTable();
    if (start < chars.length) {
        blocks.push(paragraph(buildRuns(chars, fcs, start, chars.length, chpx, fields), { istd: 0 }, styles));
    }
    return { blocks };
}

function paragraph(runs, pap, styles) {
    const block = { type: 'p', runs };
    const style = styleName(styles[pap.istd]);
    if (style) {
        block.style = style;
    }
    if (pap.align && pap.align !== 'left') {
        block.align = pap.align;
    }
    if (pap.ilfo) {
        block.list = { ordered: false, level: pap.ilvl || 0 };
    } else if (pap.indent > 0) {
        block.indent = Math.round(pap.indent / 15);
    }
    return block;
}

module.exports = { readDoc, DocError };
