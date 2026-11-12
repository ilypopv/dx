'use strict';
// The page layout of a Word document, computed the way Word computes it, without a DOM.
//
// Word does not keep its pages in the file: the lastRenderedPageBreak marks are whatever machine saved the file last
// laid out, and Word ignores them, it lays the document out again on every open. So does this module:
//   - lines are broken greedily at Word's break opportunities, with the advance widths of the fonts and no kerning
//     (Word's default), measured by an injected `measure` so that the same code runs in the webview (canvas) and in
//     the tests (font files);
//   - with East Asian layout on (useFELayout) closing punctuation may hang past the right margin, and the kinsoku
//     rules keep some characters from starting or ending a line;
//   - a line is as tall as the tallest ascent plus the tallest descent of what it holds (its text, the list marker,
//     a picture standing on the baseline); multiple line spacing adds a share of a line of its text below it, and
//     that share may run into the bottom margin of a page;
//   - the space between two paragraphs is the larger of the space after and the space before, contextual spacing
//     drops it between paragraphs of one style, a page that a manual break starts drops the space before;
//   - pages are filled to the bottom margin with Word's rules: keep with next, keep lines together, page break
//     before, widow and orphan control (on unless the document switches it off), manual and section breaks, tables
//     broken between rows with their header rows repeated and never left alone at the bottom of a page.
// The result names, for every page, the parts of the blocks it shows: a paragraph from one character to another,
// a table from one row to another.

const EPSILON = 0.01;

// ---------- characters ----------

// How many characters of a paragraph a run stands for; the renderer cuts runs by the same count.
function runLength(run) {
    if (run.br || run.image) {
        return 1;
    }
    return run.text ? run.text.length : 0;
}

// Breaking spaces (the no-break space and the figure space are not), and the hyphens and dashes a line may break
// after (the non-breaking hyphen is not).
const SPACE = /[\u0020\u2000-\u2006\u2008-\u200A\u205F\u3000]/;
const BREAK_AFTER = /[-\u2010\u2012-\u2015]/;
// With East Asian layout on, these may hang past the right margin at the end of a line...
const HANGING = /[.,:;!?)\]}%»”’、。，．：；！？）」』]/;
// ...these never start a line, and these never end one (kinsoku).
const NO_LINE_START = /^[%)\]}»”’!?,.:;>、。，．：；！？）」』]/;
const NO_LINE_END = /[([{«“‘$（「『]$/;

function fontOf(run, block, model) {
    const base = (model && model.font) || {};
    const size = run.size || block.markSize || base.size || 11;
    return {
        family: run.font || block.markFont || base.family || 'Calibri',
        size: run.sup || run.sub ? size * 0.65 : size,
        b: !!run.b,
        i: !!run.i
    };
}

// The pieces a line can break between: words with the spaces after them, tabs, line breaks and pictures. Every
// token keeps its character range, so that lines and pages can be cut there.
function tokenize(block, runs, model, measure) {
    const fe = !!(model && model.feLayout);
    const tokens = [];
    let pos = 0;
    let word = null;
    const flush = () => {
        if (word) {
            tokens.push(word);
            word = null;
        }
    };
    const start = () => {
        word = {
            kind: 'word',
            start: pos,
            end: pos,
            width: 0,
            tail: 0,
            spaces: 0,
            inner: 0,
            hang: 0,
            fonts: [],
            text: ''
        };
    };
    for (const run of runs) {
        if (run.br) {
            flush();
            tokens.push({ kind: 'br', start: pos, end: pos + 1, width: 0, tail: 0, font: fontOf(run, block, model) });
            pos += 1;
            continue;
        }
        if (run.image) {
            flush();
            const { width = 0, height = 0 } = run.image;
            tokens.push({ kind: 'image', start: pos, end: pos + 1, width, tail: 0, height });
            pos += 1;
            continue;
        }
        const text = run.text || '';
        const font = fontOf(run, block, model);
        for (let k = 0; k < text.length; k++) {
            const ch = text[k];
            if (ch === '\t') {
                flush();
                tokens.push({ kind: 'tab', start: pos, end: pos + 1, width: 0, tail: 0, font });
                pos += 1;
                continue;
            }
            if (SPACE.test(ch)) {
                if (!word) {
                    start();
                }
                word.tail += measure.width(ch, font);
                if (ch === ' ') {
                    word.spaces++;
                }
                word.end = pos + 1;
                word.fonts.push(font);
                pos += 1;
                continue;
            }
            if (word && word.tail) {
                flush();
            }
            if (!word) {
                start();
            }
            const w = measure.width(ch, font);
            word.width += w;
            word.text += ch;
            word.hang = fe && HANGING.test(ch) ? w : 0;
            word.end = pos + 1;
            word.fonts.push(font);
            pos += 1;
            // A line may break after a hyphen or a dash, also where the next run begins.
            if (BREAK_AFTER.test(ch) && !(k + 1 < text.length && SPACE.test(text[k + 1]))) {
                flush();
            }
        }
    }
    flush();
    return { tokens: fe ? kinsoku(tokens) : tokens, length: pos };
}

// No break before a word that may not start a line, nor after one that may not end it: they are glued.
function kinsoku(tokens) {
    const out = [];
    for (const t of tokens) {
        const prev = out[out.length - 1];
        if (
            prev &&
            prev.kind === 'word' &&
            t.kind === 'word' &&
            (NO_LINE_START.test(t.text) || NO_LINE_END.test(prev.text))
        ) {
            prev.width += prev.tail + t.width;
            prev.inner += prev.spaces + t.inner;
            prev.tail = t.tail;
            prev.spaces = t.spaces;
            prev.hang = t.hang;
            prev.end = t.end;
            prev.text += t.text;
            prev.fonts.push(...t.fonts);
        } else {
            out.push(t);
        }
    }
    return out;
}

// ---------- lines ----------

// Symbol fonts of Word that the browser cannot measure as Word does (the system fonts of those names differ):
// ascent, descent and line gap in 2048ths of an em, from the fonts Word ships.
const WORD_FONT_METRICS = {
    symbol: [2059, 450, 0],
    wingdings: [1841, 432, 0],
    'wingdings 2': [1727, 432, 0],
    'wingdings 3': [1900, 432, 0]
};

// One line of a font, in points: ascent (with the line gap above the text), descent and their sum.
function fontLine(measure, font) {
    const known = WORD_FONT_METRICS[(font.family || '').toLowerCase()];
    if (known) {
        const k = font.size / 2048;
        return {
            ascent: (known[0] + known[2]) * k,
            descent: known[1] * k,
            natural: (known[0] + known[1] + known[2]) * k
        };
    }
    return measure.metrics(font);
}

// Left and right indent and the indent of the first line, in points.
function indents(block) {
    const layout = block.layout || {};
    return {
        left: block.list && !layout.left ? 18 * (block.list.level + 1) : layout.left || 0,
        right: layout.right || 0,
        // A list item starts its text at the indent; its marker hangs in front of it.
        first: block.list ? 0 : layout.first || 0
    };
}

// Breaks a paragraph into lines for a column `width` points wide and gives every line its height.
// Lines: { start, end (characters), height, fit (what has to fit on the page), single (the line without spacing) }.
function paragraphLines(block, model, measure, width, runs = block.runs) {
    const layout = block.layout || {};
    const { left, right, first } = indents(block);
    const tabStop = (model && model.tabStop) || 36;
    const { tokens, length } = tokenize(block, runs, model, measure);
    const mark = fontOf({}, block, model);
    const markLine = fontLine(measure, mark);
    const end = width - right;
    const lines = [];

    let line = null;
    const grow = font => {
        const m = fontLine(measure, font);
        line.ascent = Math.max(line.ascent, m.ascent);
        line.descent = Math.max(line.descent, m.descent);
        line.natural = Math.max(line.natural, m.natural);
        line.text = true;
    };
    const open = at => {
        const x = lines.length ? left : left + first;
        line = {
            start: at,
            end: at,
            x,
            offset: x,
            width: 0,
            spaces: 0,
            pending: 0,
            tabs: [],
            ascent: 0,
            descent: 0,
            natural: 0,
            image: 0,
            text: false
        };
        if (!lines.length && block.list) {
            // The marker is part of the first line: its font can make the line taller.
            const family = block.list.metricFont || (block.list.format && block.list.format.font);
            if (family) {
                const m = fontLine(measure, {
                    family,
                    size: (block.list.format && block.list.format.size) || mark.size
                });
                line.ascent = m.ascent;
                line.descent = m.descent;
            }
        }
    };
    const close = () => {
        lines.push(line);
        line = null;
    };
    open(0);
    for (const token of tokens) {
        if (token.kind === 'br') {
            grow(token.font);
            line.end = token.end;
            close();
            open(token.end);
            continue;
        }
        if (token.kind === 'tab') {
            // The next default tab stop, measured from the margin.
            const next = (Math.floor((line.x + EPSILON) / tabStop) + 1) * tabStop;
            if (next > end && line.end > line.start) {
                close();
                open(token.start);
            }
            const to = Math.max(line.x, Math.min(next, end));
            line.tabs.push({ at: token.start, width: to - line.x });
            line.x = to;
            line.width = line.x - line.offset;
            // Justification stretches the spaces after the last tab only.
            line.spaces = 0;
            line.pending = 0;
            line.end = token.end;
            grow(token.font);
            continue;
        }
        if (line.x + token.width - (token.hang || 0) > end + EPSILON && line.end > line.start) {
            close();
            open(token.start);
        }
        if (token.kind === 'image') {
            line.image = Math.max(line.image, token.height);
        } else {
            for (const font of token.fonts) {
                grow(font);
            }
        }
        line.spaces += line.pending + (token.inner || 0);
        line.pending = token.spaces || 0;
        line.width = line.x + token.width - line.offset;
        line.x += token.width + token.tail;
        line.end = token.end;
    }
    if (line.end > line.start || !lines.length) {
        line.end = length;
        close();
    }

    const spacing = layout.line;
    for (const l of lines) {
        if (!l.text && !l.image) {
            l.ascent = Math.max(l.ascent, markLine.ascent);
            l.descent = Math.max(l.descent, markLine.descent);
        }
        const natural = l.natural || markLine.natural;
        // Nothing hangs below a picture alone on its line.
        const single = l.image
            ? Math.max(l.image, l.text ? l.ascent : 0) + (l.text ? l.descent : 0)
            : l.ascent + l.descent;
        if (spacing && spacing.pt && !spacing.atLeast) {
            l.height = l.fit = spacing.pt;
        } else if (spacing && spacing.atLeast) {
            l.height = l.fit = Math.max(spacing.pt, single);
        } else {
            const factor = spacing ? Math.max(0.06, spacing.factor) : 1;
            l.height = single + (factor - 1) * natural;
            // What multiple spacing adds hangs below the text and may run into the bottom margin.
            l.fit = Math.min(single, l.height);
        }
        l.single = single;
        // For the alignment: the text starts at `offset` and is `width` long; justification spreads `room - width`
        // over its `spaces`.
        l.room = end - l.offset;
        delete l.pending;
    }
    return lines;
}

// ---------- blocks ----------

const CELL_MARGIN = 5.4;
const BORDER = 0.5;

function spacingOf(block) {
    const layout = block.layout || {};
    return { before: layout.before || 0, after: layout.after || 0 };
}

function sameStyle(a, b) {
    return !!(a && b && a.sid && a.sid === b.sid);
}

// The gap between two paragraphs: the larger of the space after and the space before, without the sides that
// contextual spacing drops between paragraphs of one style.
function gapBetween(prev, block) {
    const { before } = spacingOf(block);
    if (!prev) {
        return before;
    }
    const same = sameStyle(prev, block);
    const after = same && prev.layout && prev.layout.contextual ? 0 : spacingOf(prev).after;
    const own = same && block.layout && block.layout.contextual ? 0 : before;
    return Math.max(after, own);
}

// The widths of the cells of every row, inside their margins.
function cellWidths(block) {
    const grid = block.grid || [];
    return block.rows.map(row => {
        let col = 0;
        return row.map(cell => {
            const span = cell.colspan || 1;
            let w = 0;
            for (let k = col; k < col + span; k++) {
                w += grid[k] || 0;
            }
            col += span;
            return Math.max(1, w - 2 * CELL_MARGIN);
        });
    });
}

// The height of a stack of blocks (a cell, a text box, a header) laid out in `width` points. Framed paragraphs
// (a page number in a header) are placed by their frame and take no room.
function stackHeight(blocks, model, measure, width) {
    let height = 0;
    let prev = null;
    for (const b of blocks) {
        if (b.type === 'pagebreak' || b.type === 'section' || (b.layout && b.layout.frame)) {
            continue;
        }
        if (b.type === 'table') {
            height += tableRows(b, model, measure).reduce((s, r) => s + r.height, 0);
            prev = null;
            continue;
        }
        if (b.type === 'box') {
            height += stackHeight(b.blocks, model, measure, Math.min(width, b.width || width));
            prev = null;
            continue;
        }
        height += gapBetween(prev, b) - (prev ? spacingOf(prev).after : 0);
        height += paragraphLines(b, model, measure, width).reduce((s, l) => s + l.height, 0) + spacingOf(b).after;
        prev = b;
    }
    return height;
}

function tableRows(block, model, measure) {
    const widths = cellWidths(block);
    return block.rows.map((row, r) => {
        let height = 0;
        row.forEach((cell, c) => {
            if (!cell.rowspan || cell.rowspan === 1) {
                const blocks = Array.isArray(cell) ? cell : cell.blocks;
                height = Math.max(height, stackHeight(blocks, model, measure, widths[r][c]));
            }
        });
        return { height: height + (block.borders === false ? 0 : BORDER) };
    });
}

// ---------- pages ----------

const DEFAULT_PAGE = { width: 595.3, height: 841.9, top: 56.7, right: 56.7, bottom: 56.7, left: 56.7 };

// Lays the body out page by page. `layout.steps()` does one block per step (the renderer shows the pages that are
// complete while the rest is still being laid out); `layout.pages` grows as it goes.
function createLayout(model, measure) {
    const blocks = model.blocks;
    const sections = model.sections && model.sections.length ? model.sections : [{ header: {}, footer: {} }];
    const pageOf = s => (sections[s] && sections[s].page) || model.page || DEFAULT_PAGE;
    const pages = [];
    const lineCache = new Map();
    const hfCache = new Map();

    let section = 0;
    let page = null;
    let number = 0;
    let y = 0;
    let room = 0;
    let width = 0;
    let prev = null; // the paragraph before, on this page
    let afterHard = false;

    // A header (or footer) pushes the body when it is taller than the margin.
    const hfHeight = (s, kind, which) => {
        const list = sections[s] && sections[s][which] && sections[s][which][kind];
        if (!list || !list.length) {
            return 0;
        }
        const key = s + which + kind;
        if (!hfCache.has(key)) {
            const p = pageOf(s);
            hfCache.set(key, stackHeight(list, model, measure, p.width - p.left - p.right));
        }
        return hfCache.get(key);
    };

    const newPage = (first, hard) => {
        const sec = sections[section] || {};
        number = first && sec.start !== undefined ? sec.start : number + 1;
        const kind = first && sec.titlePg ? 'first' : model.evenOdd && number % 2 === 0 ? 'even' : 'default';
        const p = pageOf(section);
        const top = Math.max(p.top, (p.header ?? 36) + hfHeight(section, kind, 'header'));
        const bottom = Math.max(p.bottom, (p.footer ?? 36) + hfHeight(section, kind, 'footer'));
        page = { section, first: !!first, hard: !!hard, number, kind, parts: [] };
        pages.push(page);
        room = p.height - top - bottom;
        width = p.width - p.left - p.right;
        y = 0;
        prev = null;
        afterHard = !!hard;
    };

    const linesOf = i => {
        let lines = lineCache.get(i);
        if (!lines) {
            lines = paragraphLines(blocks[i], model, measure, width);
            lineCache.set(i, lines);
        }
        return lines;
    };
    const heightOf = lines => lines.reduce((s, l) => s + l.height, 0);
    // The height a run of lines needs on a page: the last one only needs its text to fit.
    const fitOf = lines => (lines.length ? heightOf(lines.slice(0, -1)) + lines[lines.length - 1].fit : 0);

    // The gap above a paragraph that starts here.
    const gapHere = block => {
        if (y === 0) {
            return afterHard ? 0 : spacingOf(block).before;
        }
        // The space after the paragraph before is already counted.
        return gapBetween(prev, block) - (prev ? spacingOf(prev).after : 0);
    };

    const placeParagraph = i => {
        const block = blocks[i];
        const lines = linesOf(i);
        const layout = block.layout || {};
        const { after } = spacingOf(block);
        let k = 0;
        while (k < lines.length) {
            let used = k === 0 ? gapHere(block) : 0;
            let n = 0;
            while (k + n < lines.length && y + used + lines[k + n].fit <= room + EPSILON) {
                used += lines[k + n].height;
                n++;
            }
            if (n === lines.length - k) {
                page.parts.push({
                    block: i,
                    from: lines[k].start,
                    to: lines[lines.length - 1].end,
                    lines: [k, lines.length],
                    first: k === 0,
                    last: true
                });
                y += used + after;
                prev = block;
                afterHard = false;
                return;
            }
            if (y > 0) {
                const rest = lines.length - k;
                if (layout.keepLines && k === 0) {
                    n = 0;
                } else if (!layout.noWidow) {
                    // No last line alone at the top of the next page: one more line goes along.
                    if (n > 0 && rest - n === 1) {
                        n -= 1;
                    }
                    // No first line alone at the bottom of this page: the paragraph starts on the next one.
                    if (k === 0 && n === 1) {
                        n = 0;
                    }
                }
            } else if (n === 0) {
                // Taller than a page: it takes the page anyway.
                n = 1;
            }
            if (n > 0) {
                page.parts.push({
                    block: i,
                    from: lines[k].start,
                    to: lines[k + n - 1].end,
                    lines: [k, k + n],
                    first: k === 0,
                    last: false
                });
                afterHard = false;
            }
            k += n;
            newPage(false, false);
        }
    };

    const placeTable = i => {
        const block = blocks[i];
        const rows = tableRows(block, model, measure);
        const header = block.headerRows || 0;
        const headerHeight = rows.slice(0, header).reduce((s, r) => s + r.height, 0);
        let r = 0;
        while (r < rows.length) {
            const repeat = r > 0 ? header : 0;
            let used = repeat ? headerHeight : 0;
            let n = 0;
            while (r + n < rows.length && y + used + rows[r + n].height <= room + EPSILON) {
                used += rows[r + n].height;
                n++;
            }
            // The header rows are never left alone at the bottom of a page.
            if (r === 0 && header && n <= header && n < rows.length && y > 0) {
                n = 0;
            }
            if (n === 0 && y === 0) {
                n = 1;
                used += rows[r].height;
            }
            if (n > 0) {
                page.parts.push({ block: i, rows: [r, r + n], repeat, last: r + n === rows.length });
                y += used;
                prev = null;
                afterHard = false;
            }
            r += n;
            if (r < rows.length) {
                newPage(false, false);
            }
        }
    };

    // What has to come along with a block that keeps with it: the first lines of a paragraph (all of them if it keeps
    // its lines together, two with widow control), the first rows of a table.
    const headHeight = i => {
        const b = blocks[i];
        if (!b) {
            return 0;
        }
        if (b.type === 'table') {
            const rows = tableRows(b, model, measure);
            return rows.slice(0, (b.headerRows || 0) + 1).reduce((s, r) => s + r.height, 0);
        }
        if (b.type !== 'p') {
            return 0;
        }
        const lines = linesOf(i);
        // With widow control a paragraph of up to three lines cannot be split at all (2 + 1 leaves a widow).
        const count =
            b.layout && b.layout.keepLines
                ? lines.length
                : b.layout && b.layout.noWidow
                  ? 1
                  : lines.length <= 3
                    ? lines.length
                    : 2;
        return spacingOf(b).before + fitOf(lines.slice(0, count));
    };

    // Keep with next: a paragraph (and the chain of those that keep with it) stays on the page of the first lines of
    // the block after the chain. Keep lines together: the paragraph starts a page rather than being split. Neither
    // moves anything when what has to stay together would not fit on an empty page either.
    const keepTogether = i => {
        const block = blocks[i];
        const layout = block.layout || {};
        let need = 0;
        let j = i;
        while (
            blocks[j] &&
            blocks[j].type === 'p' &&
            blocks[j].layout &&
            blocks[j].layout.keepNext &&
            !blocks[j].pageBreakOnly
        ) {
            need +=
                (j === i ? gapHere(blocks[j]) : spacingOf(blocks[j]).before) +
                heightOf(linesOf(j)) +
                spacingOf(blocks[j]).after;
            j++;
        }
        if (j > i && blocks[j] && blocks[j].type !== 'pagebreak' && blocks[j].type !== 'section') {
            need += headHeight(j);
        } else if (layout.keepLines) {
            need = gapHere(block) + fitOf(linesOf(i));
        } else {
            return;
        }
        if (y + need > room + EPSILON && need <= room + EPSILON) {
            newPage(false, false);
        }
    };

    function* steps() {
        newPage(true, false);
        let pendingBreak = false;
        let pendingFirst = false;
        for (let i = 0; i < blocks.length; i++) {
            const block = blocks[i];
            if (block.type === 'pagebreak') {
                pendingBreak = true;
                continue;
            }
            if (block.type === 'section') {
                section++;
                pendingFirst = pendingFirst || !!block.newPage;
                continue;
            }
            if (pendingBreak) {
                newPage(pendingFirst, true);
                pendingBreak = false;
                pendingFirst = false;
            }
            if (block.type === 'table') {
                placeTable(i);
            } else if (block.type === 'box') {
                const h = stackHeight(block.blocks, model, measure, Math.min(width, block.width || width));
                if (y + h > room + EPSILON && y > 0) {
                    newPage(false, false);
                }
                page.parts.push({ block: i, whole: true });
                y += h;
                prev = null;
            } else if (block.type === 'p') {
                if (block.pageBreakOnly) {
                    // A paragraph that only carries a manual page break ends the page it is on.
                    page.parts.push({ block: i, from: 0, to: 0, first: true, last: true, empty: true });
                } else {
                    if (block.layout && block.layout.pageBreakBefore && y > 0) {
                        newPage(false, false);
                    }
                    if (y > 0) {
                        keepTogether(i);
                    }
                    placeParagraph(i);
                }
            }
            yield;
        }
    }

    return {
        pages,
        steps,
        // The lines of a paragraph for a column `w` points wide (the renderer draws them line by line).
        linesOf: (block, w) => paragraphLines(block, model, measure, w)
    };
}

function layoutDocument(model, measure) {
    const layout = createLayout(model, measure);
    const steps = layout.steps();
    while (!steps.next().done) {
        // drained
    }
    return { pages: layout.pages };
}

module.exports = { createLayout, layoutDocument, paragraphLines, cellWidths, runLength, tokenize, CELL_MARGIN };
