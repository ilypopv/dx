'use strict';

// A reader and a minimal writer of Office Open XML word processing documents (.docx).
// readDocx() unzips the package, resolves styles, numbering, relationships and the theme fonts, and turns
// word/document.xml into the block model shared with the .doc reader (see doc.js). Every rendered piece of a
// paragraph keeps a slot that points back to its XML node, so text edits made in the viewer can be written
// into the original tree: everything that was not edited is serialized exactly as it was read.

const fflate = require('../../media/vendor/fflate.js');
const { parseXml, serialize, element, setAttr, child, children, find, textOf } = require('./xml');

const LIMITS = { maxEntries: 20000, maxPart: 256 * 1024 * 1024, maxTotal: 1024 * 1024 * 1024 };
const EMU_PER_PT = 12700;

const HIGHLIGHT = {
    yellow: '#ffff00',
    green: '#00ff00',
    cyan: '#00ffff',
    magenta: '#ff00ff',
    blue: '#0000ff',
    red: '#ff0000',
    darkBlue: '#000080',
    darkCyan: '#008080',
    darkGreen: '#008000',
    darkMagenta: '#800080',
    darkRed: '#800000',
    darkYellow: '#808000',
    darkGray: '#808080',
    lightGray: '#c0c0c0',
    black: '#000000',
    white: '#ffffff'
};

const IMAGE_TYPES = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    svg: 'image/svg+xml',
    webp: 'image/webp',
    tif: 'image/tiff',
    tiff: 'image/tiff'
};

class DocxError extends Error {}

function unzip(bytes) {
    let entries = 0;
    let total = 0;
    const files = fflate.unzipSync(bytes, {
        filter(file) {
            entries++;
            total += file.originalSize;
            if (entries > LIMITS.maxEntries || file.originalSize > LIMITS.maxPart || total > LIMITS.maxTotal) {
                throw new DocxError('The file is too large or looks like a zip bomb.');
            }
            return true;
        }
    });
    return files;
}

const utf8 = new TextDecoder('utf-8');
const readPart = (files, name) => (files[name] ? parseXml(utf8.decode(files[name])) : undefined);

function onOff(node) {
    if (!node) {
        return undefined;
    }
    const v = node.attrs['w:val'];
    return v === undefined || !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

const num = (node, name) => {
    const v = node && node.attrs[name];
    if (v === undefined) {
        return undefined;
    }
    const n = parseFloat(v);
    return Number.isFinite(n) ? n : undefined;
};

// ---------- relationships ----------

function partDir(part) {
    const slash = part.lastIndexOf('/');
    return slash < 0 ? '' : part.slice(0, slash + 1);
}

function resolvePath(base, target) {
    if (target.startsWith('/')) {
        return target.slice(1);
    }
    const parts = (base + target).split('/');
    const out = [];
    for (const p of parts) {
        if (p === '..') {
            out.pop();
        } else if (p !== '.' && p !== '') {
            out.push(p);
        }
    }
    return out.join('/');
}

function readRels(files, part) {
    const rels = new Map();
    const name = partDir(part) + '_rels/' + part.slice(partDir(part).length) + '.rels';
    const tree = readPart(files, name);
    const root = tree && child(tree, 'Relationships');
    for (const rel of children(root, 'Relationship')) {
        const external = rel.attrs.TargetMode === 'External';
        rels.set(rel.attrs.Id, {
            type: (rel.attrs.Type || '').replace(/^.*\//, ''),
            target: external ? rel.attrs.Target : resolvePath(partDir(part), rel.attrs.Target || ''),
            external
        });
    }
    return rels;
}

function mainPart(files) {
    const tree = readPart(files, '_rels/.rels');
    const root = tree && child(tree, 'Relationships');
    for (const rel of children(root, 'Relationship')) {
        if (/\/officeDocument$/.test(rel.attrs.Type || '')) {
            return resolvePath('', rel.attrs.Target);
        }
    }
    return files['word/document.xml'] ? 'word/document.xml' : undefined;
}

// ---------- theme, styles, numbering ----------

function readTheme(files, rels) {
    const fonts = {};
    for (const rel of rels.values()) {
        if (rel.type === 'theme') {
            const tree = readPart(files, rel.target);
            const major = find(tree, 'a:majorFont');
            const minor = find(tree, 'a:minorFont');
            fonts.major = major && child(major, 'a:latin') && child(major, 'a:latin').attrs.typeface;
            fonts.minor = minor && child(minor, 'a:latin') && child(minor, 'a:latin').attrs.typeface;
        }
    }
    return fonts;
}

// Run properties as a flat object; undefined keys are inherited.
function runProps(rPr, theme) {
    const p = {};
    if (!rPr) {
        return p;
    }
    for (const c of rPr.children) {
        switch (c.name) {
            case 'w:b':
                p.b = onOff(c);
                break;
            case 'w:i':
                p.i = onOff(c);
                break;
            case 'w:strike':
            case 'w:dstrike':
                p.s = onOff(c);
                break;
            case 'w:u':
                p.u = c.attrs['w:val'] !== 'none';
                break;
            case 'w:caps':
                p.caps = onOff(c);
                break;
            case 'w:smallCaps':
                p.smallCaps = onOff(c);
                break;
            case 'w:vanish':
                p.hidden = onOff(c);
                break;
            case 'w:sz':
                p.size = num(c, 'w:val') / 2;
                break;
            case 'w:color': {
                const v = c.attrs['w:val'];
                if (v && /^[0-9a-fA-F]{6}$/.test(v)) {
                    p.color = '#' + v.toLowerCase();
                } else if (v === 'auto') {
                    p.color = null;
                }
                break;
            }
            case 'w:highlight':
                p.highlight = HIGHLIGHT[c.attrs['w:val']] || null;
                break;
            case 'w:shd': {
                const fill = c.attrs['w:fill'];
                if (fill && /^[0-9a-fA-F]{6}$/.test(fill) && !p.highlight) {
                    p.highlight = '#' + fill.toLowerCase();
                }
                break;
            }
            case 'w:vertAlign': {
                const v = c.attrs['w:val'];
                p.sup = v === 'superscript';
                p.sub = v === 'subscript';
                break;
            }
            case 'w:rFonts': {
                const a = c.attrs;
                const themed = a['w:asciiTheme'];
                if (a['w:ascii']) {
                    p.font = a['w:ascii'];
                } else if (themed) {
                    p.font = /^major/.test(themed) ? theme.major : theme.minor;
                } else if (a['w:hAnsi']) {
                    p.font = a['w:hAnsi'];
                }
                break;
            }
        }
    }
    return p;
}

function paraProps(pPr) {
    const p = {};
    if (!pPr) {
        return p;
    }
    for (const c of pPr.children) {
        switch (c.name) {
            case 'w:jc': {
                const v = c.attrs['w:val'];
                p.align =
                    v === 'center'
                        ? 'center'
                        : v === 'right' || v === 'end'
                          ? 'right'
                          : v === 'both' || v === 'distribute'
                            ? 'justify'
                            : 'left';
                break;
            }
            case 'w:ind': {
                const left = num(c, 'w:left') ?? num(c, 'w:start');
                if (left !== undefined) {
                    p.left = left / 20;
                }
                const right = num(c, 'w:right') ?? num(c, 'w:end');
                if (right !== undefined) {
                    p.right = right / 20;
                }
                if (c.attrs['w:hanging'] !== undefined) {
                    p.first = -num(c, 'w:hanging') / 20;
                } else if (c.attrs['w:firstLine'] !== undefined) {
                    p.first = num(c, 'w:firstLine') / 20;
                }
                break;
            }
            case 'w:spacing': {
                if (c.attrs['w:before'] !== undefined) {
                    p.before = num(c, 'w:before') / 20;
                }
                if (c.attrs['w:after'] !== undefined) {
                    p.after = num(c, 'w:after') / 20;
                }
                const line = num(c, 'w:line');
                if (line !== undefined) {
                    const rule = c.attrs['w:lineRule'] || 'auto';
                    p.line = rule === 'auto' ? { factor: line / 240 } : { pt: line / 20 };
                }
                break;
            }
            case 'w:numPr': {
                const id = child(c, 'w:numId');
                const lvl = child(c, 'w:ilvl');
                if (id) {
                    p.numId = id.attrs['w:val'];
                }
                if (lvl) {
                    p.ilvl = num(lvl, 'w:val');
                }
                break;
            }
            case 'w:outlineLvl':
                p.outline = num(c, 'w:val');
                break;
            case 'w:pageBreakBefore':
                p.pageBreakBefore = onOff(c);
                break;
            case 'w:shd': {
                const fill = c.attrs['w:fill'];
                if (fill && /^[0-9a-fA-F]{6}$/.test(fill)) {
                    p.shade = '#' + fill.toLowerCase();
                }
                break;
            }
        }
    }
    return p;
}

const merge = (...items) => {
    const out = {};
    for (const item of items) {
        if (item) {
            for (const key of Object.keys(item)) {
                if (item[key] !== undefined) {
                    out[key] = item[key];
                }
            }
        }
    }
    return out;
};

function readStyles(files, rels, theme) {
    const styles = { byId: new Map(), defaults: { r: {}, p: {} }, defaultPara: undefined, defaultTable: undefined };
    let tree;
    for (const rel of rels.values()) {
        if (rel.type === 'styles') {
            tree = readPart(files, rel.target);
        }
    }
    const root = tree && child(tree, 'w:styles');
    if (!root) {
        return styles;
    }
    const defaults = child(root, 'w:docDefaults');
    styles.defaults.r = runProps(child(child(defaults, 'w:rPrDefault'), 'w:rPr'), theme);
    styles.defaults.p = paraProps(child(child(defaults, 'w:pPrDefault'), 'w:pPr'));
    for (const s of children(root, 'w:style')) {
        const id = s.attrs['w:styleId'];
        const nameNode = child(s, 'w:name');
        const basedOn = child(s, 'w:basedOn');
        const entry = {
            id,
            type: s.attrs['w:type'],
            name: nameNode ? nameNode.attrs['w:val'] || '' : '',
            basedOn: basedOn ? basedOn.attrs['w:val'] : undefined,
            r: runProps(child(s, 'w:rPr'), theme),
            p: paraProps(child(s, 'w:pPr')),
            borders: hasBorders(child(child(s, 'w:tblPr'), 'w:tblBorders'))
        };
        styles.byId.set(id, entry);
        if (s.attrs['w:default'] === '1' || s.attrs['w:default'] === 'true') {
            if (entry.type === 'paragraph') {
                styles.defaultPara = id;
            } else if (entry.type === 'table') {
                styles.defaultTable = id;
            }
        }
    }
    return styles;
}

// The chain of a style from the root to the style itself.
function styleChain(styles, id) {
    const chain = [];
    const seen = new Set();
    let current = id && styles.byId.get(id);
    while (current && !seen.has(current.id) && chain.length < 32) {
        seen.add(current.id);
        chain.unshift(current);
        current = current.basedOn ? styles.byId.get(current.basedOn) : undefined;
    }
    return chain;
}

function headingOf(chain, outline) {
    for (let i = chain.length - 1; i >= 0; i--) {
        const name = chain[i].name.toLowerCase();
        const m = /^heading\s*(\d)$/.exec(name);
        if (m) {
            return 'h' + Math.min(6, Math.max(1, +m[1]));
        }
        if (name === 'title') {
            return 'title';
        }
        if (name === 'subtitle') {
            return 'subtitle';
        }
    }
    if (outline !== undefined && outline >= 0 && outline < 9) {
        return 'h' + Math.min(6, outline + 1);
    }
    return undefined;
}

function readNumbering(files, rels) {
    const abstract = new Map();
    const nums = new Map();
    let tree;
    for (const rel of rels.values()) {
        if (rel.type === 'numbering') {
            tree = readPart(files, rel.target);
        }
    }
    const root = tree && child(tree, 'w:numbering');
    if (!root) {
        return { abstract, nums };
    }
    const readLevels = node => {
        const levels = [];
        for (const lvl of children(node, 'w:lvl')) {
            const fmt = child(lvl, 'w:numFmt');
            const text = child(lvl, 'w:lvlText');
            const start = child(lvl, 'w:start');
            const font = child(child(lvl, 'w:rPr'), 'w:rFonts');
            levels[num(lvl, 'w:ilvl') || 0] = {
                fmt: fmt ? fmt.attrs['w:val'] : 'decimal',
                text: text ? text.attrs['w:val'] || '' : '',
                start: start ? num(start, 'w:val') : 1,
                font: font ? font.attrs['w:ascii'] : undefined,
                p: paraProps(child(lvl, 'w:pPr'))
            };
        }
        return levels;
    };
    for (const a of children(root, 'w:abstractNum')) {
        abstract.set(a.attrs['w:abstractNumId'], readLevels(a));
    }
    for (const n of children(root, 'w:num')) {
        const ref = child(n, 'w:abstractNumId');
        const overrides = {};
        for (const o of children(n, 'w:lvlOverride')) {
            const start = child(o, 'w:startOverride');
            const lvl = child(o, 'w:lvl');
            overrides[num(o, 'w:ilvl') || 0] = {
                start: start ? num(start, 'w:val') : undefined,
                level: lvl ? readLevels({ children: [lvl] })[num(lvl, 'w:ilvl') || 0] : undefined
            };
        }
        nums.set(n.attrs['w:numId'], { abstractId: ref ? ref.attrs['w:val'] : undefined, overrides });
    }
    return { abstract, nums };
}

function roman(n) {
    const table = [
        [1000, 'm'],
        [900, 'cm'],
        [500, 'd'],
        [400, 'cd'],
        [100, 'c'],
        [90, 'xc'],
        [50, 'l'],
        [40, 'xl'],
        [10, 'x'],
        [9, 'ix'],
        [5, 'v'],
        [4, 'iv'],
        [1, 'i']
    ];
    let out = '';
    for (const [value, sym] of table) {
        while (n >= value) {
            out += sym;
            n -= value;
        }
    }
    return out;
}

function letters(n) {
    let out = '';
    while (n > 0) {
        const m = (n - 1) % 26;
        out = String.fromCharCode(97 + m) + out;
        n = (n - 1 - m) / 26;
    }
    return out;
}

function formatNumber(n, fmt) {
    switch (fmt) {
        case 'lowerLetter':
            return letters(n);
        case 'upperLetter':
            return letters(n).toUpperCase();
        case 'lowerRoman':
            return roman(n);
        case 'upperRoman':
            return roman(n).toUpperCase();
        case 'decimalZero':
            return n < 10 ? '0' + n : String(n);
        case 'none':
            return '';
        default:
            return String(n);
    }
}

// Symbol and Wingdings bullets live in the private use area; show their usual Unicode look.
function bulletText(text, font) {
    const c = text.charCodeAt(0);
    if (!text) {
        return '';
    }
    if (c === 0xf0a7 || (font === 'Wingdings' && text === '§')) {
        return '▪';
    }
    if (c === 0xf0d8 || c === 0xf0e8) {
        return '➢';
    }
    if (c === 0xf0fc) {
        return '✓';
    }
    if ((c >= 0xf000 && c <= 0xf0ff) || (font === 'Symbol' && text === '·')) {
        return '•';
    }
    if (text === 'o' && /courier/i.test(font || '')) {
        return '◦';
    }
    return text;
}

class Numbering {
    constructor(def) {
        this.def = def;
        this.counters = new Map();
    }

    level(numId, ilvl) {
        const n = this.def.nums.get(numId);
        if (!n) {
            return undefined;
        }
        const levels = this.def.abstract.get(n.abstractId) || [];
        const override = n.overrides[ilvl];
        return (override && override.level) || levels[ilvl];
    }

    // The marker text of the next item of a list level; deeper levels restart.
    next(numId, ilvl) {
        const n = this.def.nums.get(numId);
        if (!n) {
            return undefined;
        }
        const levels = this.def.abstract.get(n.abstractId) || [];
        let counters = this.counters.get(numId);
        if (!counters) {
            counters = [];
            this.counters.set(numId, counters);
        }
        const startOf = l => {
            const o = n.overrides[l];
            const def = (o && o.level) || levels[l];
            return o && o.start !== undefined ? o.start : def ? def.start : 1;
        };
        counters[ilvl] = counters[ilvl] === undefined ? startOf(ilvl) : counters[ilvl] + 1;
        counters.length = ilvl + 1;
        const lvl = this.level(numId, ilvl);
        if (!lvl) {
            return '•';
        }
        if (lvl.fmt === 'bullet') {
            return bulletText(lvl.text, lvl.font);
        }
        return lvl.text.replace(/%(\d)/g, (m, k) => {
            const index = +k - 1;
            const value = counters[index] === undefined ? startOf(index) : counters[index];
            const def = (n.overrides[index] && n.overrides[index].level) || levels[index];
            return formatNumber(value, index === ilvl ? lvl.fmt : def ? def.fmt : 'decimal');
        });
    }
}

// ---------- body ----------

function hasBorders(node) {
    if (!node) {
        return undefined;
    }
    return node.children.some(
        c => c.attrs && c.attrs['w:val'] && c.attrs['w:val'] !== 'none' && c.attrs['w:val'] !== 'nil'
    );
}

class BodyReader {
    constructor(ctx, editable) {
        this.ctx = ctx;
        this.editable = editable;
    }

    blocks(container, out) {
        for (const node of container.children) {
            switch (node.name) {
                case 'w:p':
                    this.paragraph(node, out);
                    break;
                case 'w:tbl':
                    out.push(this.table(node));
                    break;
                case 'w:sdt':
                    this.blocks(child(node, 'w:sdtContent') || { children: [] }, out);
                    break;
                case 'w:customXml':
                case 'w:ins':
                case 'w:moveTo':
                    this.blocks(node, out);
                    break;
                case 'mc:AlternateContent':
                    this.blocks(child(node, 'mc:Fallback') || child(node, 'mc:Choice') || { children: [] }, out);
                    break;
            }
        }
        return out;
    }

    paragraph(node, out) {
        const ctx = this.ctx;
        const pPr = child(node, 'w:pPr');
        const direct = paraProps(pPr);
        const styleNode = child(pPr, 'w:pStyle');
        const styleId = styleNode ? styleNode.attrs['w:val'] : ctx.styles.defaultPara;
        const chain = styleChain(ctx.styles, styleId);
        let props = merge(ctx.styles.defaults.p, ...chain.map(s => s.p), direct);
        const baseRun = merge(ctx.styles.defaults.r, ...chain.map(s => s.r));

        const block = { type: 'p', runs: [] };
        if (this.editable) {
            block.id = ctx.paragraphs.length;
            ctx.paragraphs.push({ node, slots: [], baseRun: child(pPr, 'w:rPr') });
        }
        const heading = headingOf(chain, props.outline);
        if (heading) {
            block.style = heading;
        }

        if (props.numId && props.numId !== '0') {
            const ilvl = props.ilvl || 0;
            const lvl = ctx.numbering.level(props.numId, ilvl);
            const marker = ctx.numbering.next(props.numId, ilvl);
            if (lvl) {
                props = merge(props, lvl.p, direct);
            }
            if (marker !== undefined) {
                block.list = { marker, level: ilvl };
            }
        }

        if (props.align && props.align !== 'left') {
            block.align = props.align;
        }
        const box = {};
        for (const key of ['left', 'right', 'first', 'before', 'after', 'shade']) {
            if (props[key]) {
                box[key] = props[key];
            }
        }
        if (props.line) {
            box.line = props.line;
        }
        if (Object.keys(box).length) {
            block.layout = box;
        }
        if (props.pageBreakBefore) {
            out.push({ type: 'pagebreak' });
        }

        this.state = { fields: [], baseRun, block, after: [], seenText: false };
        this.inline(node, block.runs, {});
        out.push(block);
        for (const extra of this.state.after) {
            out.push(extra);
        }
        // A section break inside the paragraph properties starts a new page unless it is continuous.
        const sect = child(pPr, 'w:sectPr');
        if (sect) {
            const type = child(sect, 'w:type');
            if (!type || type.attrs['w:val'] !== 'continuous') {
                out.push({ type: 'pagebreak' });
            }
        }
    }

    inline(container, runs, link) {
        for (const node of container.children) {
            switch (node.name) {
                case 'w:r':
                    this.run(node, runs, link);
                    break;
                case 'w:hyperlink': {
                    const rel = node.attrs['r:id'] && this.ctx.rels.get(node.attrs['r:id']);
                    const href = rel && rel.external ? rel.target : undefined;
                    this.inline(node, runs, href ? { href } : link);
                    break;
                }
                case 'w:fldSimple': {
                    const m = /HYPERLINK\s+"([^"]+)"/i.exec(node.attrs['w:instr'] || '');
                    this.inline(node, runs, m ? { href: m[1] } : link);
                    break;
                }
                case 'w:sdt':
                    this.inline(child(node, 'w:sdtContent') || { children: [] }, runs, link);
                    break;
                case 'w:smartTag':
                case 'w:customXml':
                case 'w:ins':
                case 'w:moveTo':
                case 'w:dir':
                case 'w:bdo':
                    this.inline(node, runs, link);
                    break;
                case 'mc:AlternateContent':
                    this.inline(child(node, 'mc:Choice') || child(node, 'mc:Fallback') || { children: [] }, runs, link);
                    break;
            }
        }
    }

    slot(node, parent, kind) {
        const ctx = this.ctx;
        if (!this.editable) {
            return undefined;
        }
        const index = ctx.slots.length;
        ctx.slots.push({ node, parent, kind });
        ctx.paragraphs[this.state.block.id].slots.push(index);
        return index;
    }

    run(node, runs, link) {
        const ctx = this.ctx;
        const rPr = child(node, 'w:rPr');
        const rStyle = child(rPr, 'w:rStyle');
        const charChain = rStyle ? styleChain(ctx.styles, rStyle.attrs['w:val']) : [];
        const props = merge(this.state.baseRun, ...charChain.map(s => s.r), runProps(rPr, ctx.theme));
        const format = {};
        for (const key of [
            'b',
            'i',
            'u',
            's',
            'sup',
            'sub',
            'caps',
            'smallCaps',
            'size',
            'color',
            'highlight',
            'font'
        ]) {
            if (props[key]) {
                format[key] = props[key];
            }
        }
        for (const c of node.children) {
            this.runChild(c, node, format, runs, props, link);
        }
    }

    // The format of a piece: the run format plus the link of the enclosing hyperlink or HYPERLINK field.
    pieceFormat(format, link) {
        let href = link.href;
        const fields = this.state.fields;
        for (let f = fields.length - 1; f >= 0 && !href; f--) {
            href = fields[f].href;
        }
        return href ? Object.assign({}, format, { href }) : format;
    }

    runChild(c, parent, base, runs, props, link) {
        const state = this.state;
        const fields = state.fields;
        if (c.name === 'w:fldChar') {
            const type = c.attrs['w:fldCharType'];
            if (type === 'begin') {
                fields.push({ code: '', result: false });
            } else if (type === 'separate' && fields.length) {
                const field = fields[fields.length - 1];
                field.result = true;
                const m = /HYPERLINK\s+"([^"]+)"/i.exec(field.code);
                if (m && !/\\l/.test(field.code)) {
                    field.href = m[1];
                }
            } else if (type === 'end') {
                fields.pop();
            }
            return;
        }
        if (fields.length && !fields[fields.length - 1].result) {
            if (c.name === 'w:instrText') {
                fields[fields.length - 1].code += textOf(c);
            }
            return;
        }
        if (props.hidden) {
            return;
        }
        const format = this.pieceFormat(base, link);
        switch (c.name) {
            case 'w:t': {
                const text = textOf(c);
                if (text) {
                    state.seenText = true;
                }
                runs.push(Object.assign({ text, slot: this.slot(c, parent, 't') }, format));
                return;
            }
            case 'w:tab':
            case 'w:ptab':
                runs.push(Object.assign({ text: '\t', slot: this.slot(c, parent, 'x'), atom: true }, format));
                return;
            case 'w:noBreakHyphen':
                runs.push(Object.assign({ text: '‑', slot: this.slot(c, parent, 'x'), atom: true }, format));
                return;
            case 'w:sym': {
                const code = parseInt(c.attrs['w:char'] || '', 16);
                const text = Number.isFinite(code) ? bulletText(String.fromCharCode(code), c.attrs['w:font']) : '';
                runs.push(Object.assign({ text, slot: this.slot(c, parent, 'x'), atom: true }, format));
                return;
            }
            case 'w:br':
            case 'w:cr': {
                const type = c.attrs['w:type'];
                if (type === 'page' || type === 'column') {
                    state.after.push({ type: 'pagebreak' });
                    return;
                }
                runs.push({ text: '', br: true, slot: this.slot(c, parent, 'x') });
                return;
            }
            case 'w:lastRenderedPageBreak':
                // Where Word broke the page when it last laid the document out; used at paragraph granularity.
                if (!state.seenText && !runs.length) {
                    state.block.pageStart = true;
                }
                return;
            case 'w:footnoteReference':
            case 'w:endnoteReference': {
                const id = c.attrs['w:id'];
                const kind = c.name === 'w:footnoteReference' ? 'footnotes' : 'endnotes';
                const mark = this.ctx.noteMark(kind, id);
                runs.push(
                    Object.assign({}, format, {
                        text: mark,
                        sup: true,
                        atom: true,
                        note: true,
                        slot: this.slot(c, parent, 'x')
                    })
                );
                return;
            }
            case 'w:drawing':
            case 'w:pict':
            case 'w:object':
                this.drawing(c, parent, runs);
                return;
            case 'mc:AlternateContent': {
                const choice = child(c, 'mc:Choice');
                const pick = choice && find(choice, 'w:drawing') ? choice : child(c, 'mc:Fallback') || choice;
                if (pick) {
                    for (const inner of pick.children) {
                        this.runChild(inner, parent, base, runs, props, link);
                    }
                }
                return;
            }
        }
        return;
    }

    drawing(node, parent, runs) {
        const ctx = this.ctx;
        const blip = find(node, 'a:blip');
        const imageData = find(node, 'v:imagedata');
        const id = blip
            ? blip.attrs['r:embed'] || blip.attrs['r:link']
            : imageData
              ? imageData.attrs['r:id']
              : undefined;
        let width;
        let height;
        const extent = find(node, 'wp:extent');
        if (extent) {
            width = num(extent, 'cx') / EMU_PER_PT;
            height = num(extent, 'cy') / EMU_PER_PT;
        } else {
            const shape = find(node, 'v:shape');
            const style = (shape && shape.attrs.style) || '';
            const w = /width:\s*([\d.]+)pt/.exec(style);
            const h = /height:\s*([\d.]+)pt/.exec(style);
            width = w ? +w[1] : undefined;
            height = h ? +h[1] : undefined;
        }
        const rel = id && ctx.rels.get(id);
        const box = find(node, 'w:txbxContent');
        if (rel && !rel.external) {
            const ext = rel.target.replace(/^.*\./, '').toLowerCase();
            runs.push({
                image: { part: rel.target, type: IMAGE_TYPES[ext], width, height },
                atom: true,
                slot: this.slot(node, parent, 'x')
            });
        } else if (box) {
            // A text box: its paragraphs are shown read-only next to the anchor paragraph.
            const blocks = new BodyReader(ctx, false).blocks(box, []);
            if (blocks.length) {
                this.state.after.push({ type: 'box', blocks, width });
            }
        }
    }

    table(node) {
        const ctx = this.ctx;
        const tblPr = child(node, 'w:tblPr');
        const styleNode = child(tblPr, 'w:tblStyle');
        const styleId = styleNode ? styleNode.attrs['w:val'] : ctx.styles.defaultTable;
        let borders = hasBorders(child(tblPr, 'w:tblBorders'));
        if (borders === undefined) {
            for (const s of styleChain(ctx.styles, styleId)) {
                if (s.borders !== undefined) {
                    borders = s.borders;
                }
            }
        }
        const grid = children(child(node, 'w:tblGrid'), 'w:gridCol').map(c => (num(c, 'w:w') || 0) / 20);
        const rows = [];
        const collectRows = container => {
            for (const tr of container.children) {
                if (tr.name === 'w:tr') {
                    rows.push(tr);
                } else if (tr.name === 'w:sdt') {
                    collectRows(child(tr, 'w:sdtContent') || { children: [] });
                } else if (tr.name === 'w:ins' || tr.name === 'w:customXml') {
                    collectRows(tr);
                }
            }
        };
        collectRows(node);

        const out = { type: 'table', rows: [], grid, borders: borders !== false };
        // Vertical merges: the cell that restarts a merge grows a rowspan for every continuing cell below it.
        const open = new Map();
        for (const tr of rows) {
            const row = [];
            let col = 0;
            const cells = [];
            const collectCells = container => {
                for (const tc of container.children) {
                    if (tc.name === 'w:tc') {
                        cells.push(tc);
                    } else if (tc.name === 'w:sdt') {
                        collectCells(child(tc, 'w:sdtContent') || { children: [] });
                    } else if (tc.name === 'w:customXml') {
                        collectCells(tc);
                    }
                }
            };
            collectCells(tr);
            for (const tc of cells) {
                const tcPr = child(tc, 'w:tcPr');
                const span = num(child(tcPr, 'w:gridSpan'), 'w:val') || 1;
                const vMerge = child(tcPr, 'w:vMerge');
                const shd = child(tcPr, 'w:shd');
                if (vMerge && vMerge.attrs['w:val'] !== 'restart' && open.has(col)) {
                    open.get(col).rowspan = (open.get(col).rowspan || 1) + 1;
                    // The continuing cell is hidden, but its paragraphs stay editable through the same ids.
                    this.blocks(tc, []);
                    col += span;
                    continue;
                }
                const cell = { blocks: this.blocks(tc, []) };
                if (span > 1) {
                    cell.colspan = span;
                }
                const fill = shd && shd.attrs['w:fill'];
                if (fill && /^[0-9a-fA-F]{6}$/.test(fill)) {
                    cell.fill = '#' + fill.toLowerCase();
                }
                if (vMerge) {
                    open.set(col, cell);
                } else {
                    open.delete(col);
                }
                row.push(cell);
                col += span;
            }
            out.rows.push(row);
        }
        return out;
    }
}

function readNotes(files, rels, kind) {
    const notes = new Map();
    for (const rel of rels.values()) {
        if (rel.type === kind) {
            const tree = readPart(files, rel.target);
            const root = tree && child(tree, kind === 'footnotes' ? 'w:footnotes' : 'w:endnotes');
            for (const note of children(root, kind === 'footnotes' ? 'w:footnote' : 'w:endnote')) {
                const type = note.attrs['w:type'];
                if (!type || type === 'normal') {
                    notes.set(note.attrs['w:id'], note);
                }
            }
        }
    }
    return notes;
}

function readDocx(bytes) {
    if (!bytes || bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
        throw new DocxError('The file is not a ZIP package, so it is not a .docx document.');
    }
    let files;
    try {
        files = unzip(bytes);
    } catch (error) {
        throw error instanceof DocxError ? error : new DocxError('The .docx package is damaged: ' + error.message);
    }
    const part = mainPart(files);
    if (!part || !files[part]) {
        throw new DocxError('The package has no main document part, it is not a .docx document.');
    }
    const tree = readPart(files, part);
    const documentNode = child(tree, 'w:document');
    const body = child(documentNode, 'w:body');
    if (!body) {
        throw new DocxError('The main document part has no body.');
    }
    const rels = readRels(files, part);
    const theme = readTheme(files, rels);
    const styles = readStyles(files, rels, theme);
    const numbering = new Numbering(readNumbering(files, rels));

    const notes = { footnotes: readNotes(files, rels, 'footnotes'), endnotes: readNotes(files, rels, 'endnotes') };
    const noteOrder = { footnotes: [], endnotes: [] };
    const ctx = {
        files,
        rels,
        theme,
        styles,
        numbering,
        slots: [],
        paragraphs: [],
        noteMark(kind, id) {
            let index = noteOrder[kind].indexOf(id);
            if (index < 0) {
                noteOrder[kind].push(id);
                index = noteOrder[kind].length - 1;
            }
            return kind === 'footnotes' ? String(index + 1) : roman(index + 1);
        }
    };

    const blocks = new BodyReader(ctx, true).blocks(body, []);

    // Footnotes and endnotes follow the body, read-only.
    const noteBlocks = [];
    for (const kind of ['footnotes', 'endnotes']) {
        for (const id of noteOrder[kind]) {
            const note = notes[kind].get(id);
            if (!note) {
                continue;
            }
            const reader = new BodyReader(ctx, false);
            const content = reader.blocks(note, []);
            const mark =
                kind === 'footnotes' ? String(noteOrder[kind].indexOf(id) + 1) : roman(noteOrder[kind].indexOf(id) + 1);
            if (content.length && content[0].type === 'p') {
                content[0].runs = content[0].runs.filter(r => !r.note);
                content[0].runs.unshift({ text: mark + ' ', sup: true });
            }
            noteBlocks.push(...content);
        }
    }

    const sect = child(body, 'w:sectPr');
    const pgSz = child(sect, 'w:pgSz');
    const pgMar = child(sect, 'w:pgMar');
    const page = {
        width: (num(pgSz, 'w:w') || 11906) / 20,
        height: (num(pgSz, 'w:h') || 16838) / 20,
        top: (num(pgMar, 'w:top') ?? 1440) / 20,
        right: (num(pgMar, 'w:right') ?? 1440) / 20,
        bottom: (num(pgMar, 'w:bottom') ?? 1440) / 20,
        left: (num(pgMar, 'w:left') ?? 1440) / 20
    };
    const defaults = styles.defaults.r;
    const font = { family: defaults.font || theme.minor, size: defaults.size || 11 };

    return {
        model: { blocks, notes: noteBlocks, page, font },
        files,
        part,
        tree,
        slots: ctx.slots,
        paragraphs: ctx.paragraphs
    };
}

// ---------- writing ----------

function setText(t, text) {
    t.children = [{ text }];
    if (/^\s|\s$/.test(text) || /\s\s/.test(text)) {
        setAttr(t, 'xml:space', 'preserve');
    }
}

function removeNode(parent, node) {
    const index = parent.children.indexOf(node);
    if (index >= 0) {
        parent.children.splice(index, 1);
    }
}

// edits: [{ p: paragraph id, segments: [{ slot?: number, text: string }] }] in document order.
// A segment with a slot is a rendered piece that still exists; loose text (no slot) was typed outside them.
function applyEdits(doc, edits) {
    for (const edit of edits) {
        const para = doc.paragraphs[edit.p];
        if (!para) {
            continue;
        }
        const present = new Set();
        let lastText;
        let pending = '';
        for (const segment of edit.segments) {
            const slot = segment.slot !== undefined ? doc.slots[segment.slot] : undefined;
            if (!slot || para.slots.indexOf(segment.slot) < 0) {
                // Loose text joins the previous text piece, or the next one when it comes first.
                if (lastText) {
                    lastText.value += segment.text;
                } else {
                    pending += segment.text;
                }
                continue;
            }
            present.add(segment.slot);
            if (slot.kind === 't') {
                lastText = { slot, value: pending + segment.text };
                pending = '';
                slot.value = lastText;
            }
        }
        for (const index of para.slots) {
            const slot = doc.slots[index];
            if (!present.has(index)) {
                if (slot.kind === 't') {
                    setText(slot.node, '');
                } else {
                    removeNode(slot.parent, slot.node);
                }
            } else if (slot.value) {
                setText(slot.node, slot.value.value);
                slot.value = undefined;
            }
        }
        if (para.extra) {
            // The run an earlier save created holds the loose text now.
            setText(para.extra, pending);
        } else if (pending) {
            // Text typed into a paragraph that had no text piece: a new run with the paragraph mark formatting.
            const run = element('w:r');
            if (para.baseRun) {
                // The paragraph mark properties may carry revision marks that a run must not have.
                const keep = para.baseRun.children.filter(c => !/^w:(ins|del|moveFrom|moveTo)$/.test(c.name));
                run.children.push(element('w:rPr', {}, keep));
            }
            const t = element('w:t');
            setText(t, pending);
            run.children.push(t);
            para.node.children.push(run);
            para.extra = t;
        }
    }
}

function writeDocx(doc, edits) {
    applyEdits(doc, edits);
    const files = Object.assign({}, doc.files);
    files[doc.part] = new TextEncoder().encode(serialize(doc.tree));
    return fflate.zipSync(files, { level: 6 });
}

module.exports = { readDocx, writeDocx, applyEdits, DocxError, formatNumber, bulletText };
