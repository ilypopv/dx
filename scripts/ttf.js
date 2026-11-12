'use strict';

// A minimal TrueType reader for the parity check: units per em, the hhea line metrics and the advance width of a
// code point (cmap formats 4 and 12, hmtx). Collections (.ttc) are read at their first font.
const fs = require('fs');

function readFont(file) {
    const buf = fs.readFileSync(file);
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    const base = dv.getUint32(0) === 0x74746366 ? dv.getUint32(12) : 0;
    const tables = {};
    for (let i = 0; i < dv.getUint16(base + 4); i++) {
        const o = base + 12 + i * 16;
        const tag = String.fromCharCode(dv.getUint8(o), dv.getUint8(o + 1), dv.getUint8(o + 2), dv.getUint8(o + 3));
        tables[tag] = dv.getUint32(o + 8);
    }
    const upm = dv.getUint16(tables.head + 18);
    const hhea = tables.hhea;
    const metrics = {
        ascender: dv.getInt16(hhea + 4),
        descender: dv.getInt16(hhea + 6),
        lineGap: dv.getInt16(hhea + 8)
    };
    const numH = dv.getUint16(hhea + 34);
    const advance = g => dv.getUint16(tables.hmtx + 4 * Math.min(g, numH - 1));

    let sub = -1;
    let format = 0;
    const cmap = tables.cmap;
    for (let i = 0; i < dv.getUint16(cmap + 2); i++) {
        const platform = dv.getUint16(cmap + 4 + i * 8);
        const encoding = dv.getUint16(cmap + 6 + i * 8);
        const offset = cmap + dv.getUint32(cmap + 8 + i * 8);
        const f = dv.getUint16(offset);
        if (platform === 3 && (encoding === 10 || encoding === 1) && (f === 12 || (f === 4 && format !== 12))) {
            sub = offset;
            format = f;
        }
    }
    const glyph = cp => {
        if (format === 12) {
            for (let i = 0; i < dv.getUint32(sub + 12); i++) {
                const o = sub + 16 + i * 12;
                if (cp >= dv.getUint32(o) && cp <= dv.getUint32(o + 4)) {
                    return dv.getUint32(o + 8) + (cp - dv.getUint32(o));
                }
            }
            return 0;
        }
        const segX2 = dv.getUint16(sub + 6);
        const ends = sub + 14;
        const starts = ends + segX2 + 2;
        const deltas = starts + segX2;
        const ranges = deltas + segX2;
        for (let i = 0; i < segX2 / 2; i++) {
            if (cp > dv.getUint16(ends + i * 2)) {
                continue;
            }
            const start = dv.getUint16(starts + i * 2);
            if (cp < start) {
                return 0;
            }
            const delta = dv.getInt16(deltas + i * 2);
            const range = dv.getUint16(ranges + i * 2);
            if (!range) {
                return (cp + delta) & 0xffff;
            }
            const g = dv.getUint16(ranges + i * 2 + range + (cp - start) * 2);
            return g ? (g + delta) & 0xffff : 0;
        }
        return 0;
    };
    const widths = new Map();
    return {
        upm,
        metrics,
        // The advance of a code point, in ems.
        width(cp) {
            let w = widths.get(cp);
            if (w === undefined) {
                w = advance(glyph(cp)) / upm;
                widths.set(cp, w);
            }
            return w;
        }
    };
}

module.exports = { readFont };
