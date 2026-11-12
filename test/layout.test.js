'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { layoutDocument, paragraphLines } = require('../src/core/layout.js');

// A monospaced font: every character is half its size wide, a space a quarter; a line of it is 1.2 sizes tall.
const measure = {
    width: (text, font) => [...text].reduce((w, ch) => w + (ch === ' ' ? 0.25 : 0.5) * font.size, 0),
    metrics: font => ({ ascent: font.size, descent: font.size * 0.2, natural: font.size * 1.2 })
};
const run = (text, extra) => Object.assign({ text, size: 10, font: 'Mono' }, extra);
const para = (text, layout, extra) => Object.assign({ type: 'p', runs: [run(text)], layout }, extra);
// Pages of a 100pt column and 13 lines of 12pt (156pt of room).
const PAGE = { width: 140, height: 196, top: 20, bottom: 20, left: 20, right: 20, header: 10, footer: 10 };
const model = (blocks, extra) =>
    Object.assign({ blocks, page: PAGE, sections: [{ header: {}, footer: {}, page: PAGE }] }, extra);
const lines = (block, m = model([block])) => paragraphLines(block, m, measure, 100);
const texts = block => lines(block).map(l => block.runs[0].text.slice(l.start, l.end));
// A paragraph of `n` lines of exactly one word each (a word of 20 characters fills the column).
const tall = (n, layout, extra) => para(Array.from({ length: n }, () => 'w'.repeat(20)).join(' '), layout, extra);
const pagesOf = m =>
    layoutDocument(m, measure).pages.map(p =>
        p.parts.map(q => (q.rows ? `t${q.block}:${q.rows}` : `${q.block}:${q.from}-${q.to}`))
    );

test('lines break greedily between words, the spaces at the end of a line do not count', () => {
    // A word of 9 characters is 45pt, its space 2.5pt: two words fit in 100pt, three do not.
    assert.deepEqual(texts(para('aaaaaaaaa aaaaaaaaa aaaaaaaaa')), ['aaaaaaaaa aaaaaaaaa ', 'aaaaaaaaa']);
    // A line may break after a hyphen, also at the end of a run.
    const block = { type: 'p', runs: [run('aaaaaaaaaaaaaaaaaa-'), run('bbbbbbbbbb')] };
    assert.deepEqual(
        lines(block).map(l => [l.start, l.end]),
        [
            [0, 19],
            [19, 29]
        ]
    );
});

test('with East Asian layout punctuation hangs past the margin and kinsoku keeps characters off line ends', () => {
    // 45 + 2.5 + 50 + 5 (the full stop) is 102.5pt.
    const block = para('aaaaaaaaa aaaaaaaaaa.');
    assert.equal(lines(block).length, 2, 'without it the full stop does not fit');
    assert.equal(lines(block, model([block], { feLayout: true })).length, 1, 'the full stop hangs');
    const percent = para('aaaaaaaaaaaaaaaaaa 1 %%');
    const fe = model([percent], { feLayout: true });
    assert.deepEqual(
        lines(percent, fe).map(l => percent.runs[0].text.slice(l.start, l.end)),
        ['aaaaaaaaaaaaaaaaaa ', '1 %%'],
        'a line never starts with %: the number goes along'
    );
});

test('a line is as tall as its ascent and descent; multiple spacing adds a share of a text line below it', () => {
    const [single] = lines(para('a'));
    assert.equal(single.height, 12);
    const [spaced] = lines(para('a', { line: { factor: 1.5 } }));
    assert.equal(spaced.height, 18);
    assert.equal(spaced.fit, 12, 'only the text has to fit at the bottom of a page');
    const [exact] = lines(para('a', { line: { pt: 30 } }));
    assert.equal(exact.height, 30);
    const [atLeast] = lines(para('a', { line: { pt: 5, atLeast: true } }));
    assert.equal(atLeast.height, 12);
});

test('a picture stands on the baseline, nothing hangs below it, spacing still adds its share', () => {
    const block = {
        type: 'p',
        runs: [{ image: { width: 50, height: 100 } }],
        layout: { line: { factor: 1.5 } },
        markSize: 10
    };
    const [line] = paragraphLines(block, model([block]), measure, 100);
    assert.equal(line.height, 100 + 6);
});

test('the font of a list marker makes the first line taller', () => {
    const block = para('a', {}, { list: { marker: '•', level: 0, metricFont: 'Symbol', format: { size: 10 } } });
    const [line] = lines(block);
    // Word's Symbol: ascent 2059 and descent 450 per 2048.
    assert.ok(Math.abs(line.height - (2059 + 450) / 204.8) < 1e-9);
});

test('the space between paragraphs is the larger of after and before; contextual spacing drops it in a style', () => {
    const m = model([
        para('a', { after: 10 }),
        para('b', { before: 4 }),
        para('c', { before: 30 }, { sid: 'S' }),
        para('d', { before: 30, contextual: true }, { sid: 'S' })
    ]);
    const pages = layoutDocument(m, measure).pages;
    assert.equal(pages.length, 1);
    // Fill a page exactly: 13 lines fit, the 14th does not.
    const fill = n => model(Array.from({ length: n }, () => para('a')));
    assert.equal(layoutDocument(fill(13), measure).pages.length, 1);
    assert.equal(layoutDocument(fill(14), measure).pages.length, 2);
});

test('widow and orphan control moves lines, unless the document switches it off', () => {
    // 11 lines fill the page but two; a 3-line paragraph would leave one line on the next page.
    const blocks = [tall(11), tall(3)];
    assert.deepEqual(pagesOf(model(blocks)), [['0:0-230'], ['1:0-62']], 'the paragraph moves whole');
    const off = [tall(11), tall(3, { noWidow: true })];
    assert.deepEqual(pagesOf(model(off)), [['0:0-230', '1:0-42'], ['1:42-62']]);
    // A 4-line paragraph gives two lines to each page.
    assert.deepEqual(pagesOf(model([tall(11), tall(4)])), [['0:0-230', '1:0-42'], ['1:42-83']]);
});

test('keep with next keeps a heading with the first lines of what follows; keep lines keeps a paragraph whole', () => {
    const heading = para('Heading', { keepNext: true });
    assert.deepEqual(pagesOf(model([tall(11), heading, tall(3)])), [['0:0-230'], ['1:0-7', '2:0-62']]);
    assert.deepEqual(pagesOf(model([tall(10), tall(4, { keepLines: true })])), [['0:0-209'], ['1:0-83']]);
    assert.deepEqual(pagesOf(model([tall(1), para('x', { pageBreakBefore: true })])), [['0:0-20'], ['1:0-1']]);
});

test('a manual page break starts a page without the space before', () => {
    const blocks = [para('a'), { type: 'pagebreak' }, para('b', { before: 50 })];
    const pages = layoutDocument(model(blocks), measure).pages;
    assert.equal(pages.length, 2);
    const fill = Array.from({ length: 12 }, () => para('x'));
    // After the break the page holds 13 lines: the 50pt before is dropped.
    const after = layoutDocument(
        model([para('a'), { type: 'pagebreak' }, para('b', { before: 50 }), ...fill]),
        measure
    );
    assert.equal(after.pages.length, 2);
});

test('tables break between rows, repeat their header rows and never leave them alone at the bottom', () => {
    const cell = text => ({ blocks: [para(text)] });
    const table = {
        type: 'table',
        grid: [110.8],
        borders: false,
        headerRows: 1,
        rows: Array.from({ length: 20 }, (_, i) => [cell('r' + i)])
    };
    const pages = layoutDocument(model([table]), measure).pages;
    assert.deepEqual(
        pages.map(p => p.parts.map(q => [q.rows, q.repeat])),
        [[[[0, 13], 0]], [[[13, 20], 1]]],
        'the second page starts with the header row again'
    );
    const late = layoutDocument(model([tall(12), table]), measure).pages;
    assert.deepEqual(
        late[0].parts.map(q => q.block),
        [0],
        'the header row does not stay alone under the text'
    );
});

test('a header taller than the top margin pushes the body down; a framed paragraph takes no room', () => {
    const header = [para('h1'), para('h2'), para('h3')];
    const sections = [{ header: { default: header }, footer: {}, page: PAGE }];
    const blocks = Array.from({ length: 13 }, () => para('x'));
    assert.equal(
        layoutDocument(model(blocks, { sections }), measure).pages.length,
        2,
        'the header (10 + 36) takes 26pt of room'
    );
    const framed = [para('1', { frame: true })];
    const quiet = [{ header: { default: framed }, footer: {}, page: PAGE }];
    assert.equal(layoutDocument(model(blocks, { sections: quiet }), measure).pages.length, 1);
});

test('every line tells where its text starts, how wide it is and which spaces justification stretches', () => {
    // First line indent 10pt; the words are 20pt, the spaces 2.5pt.
    const block = para('aaaa aaaa aaaa aaaa aaaa', { first: 10 });
    const [one, two] = lines(block);
    assert.equal(one.offset, 10);
    assert.equal(one.width, 4 * 20 + 3 * 2.5, 'four words fit in 90pt, the space after the last one is not counted');
    assert.equal(one.spaces, 3);
    assert.equal(one.room, 90);
    assert.equal(two.offset, 0);
    assert.equal(two.spaces, 0);
    // A tab jumps to the next default stop (36pt); spaces before it are not stretched.
    const tabbed = { type: 'p', runs: [run('a a'), run('\t', { atom: true }), run('b b')] };
    const [line] = lines(tabbed, model([tabbed], { tabStop: 36 }));
    assert.deepEqual(line.tabs, [{ at: 3, width: 36 - 12.5 }]);
    assert.equal(line.spaces, 1, 'only the space after the tab');
});

test('keep with next needs room for a whole paragraph of three lines: it cannot be split 2 + 1', () => {
    const picture = {
        type: 'p',
        runs: [{ image: { width: 50, height: 100 } }],
        layout: { keepNext: true },
        markSize: 10
    };
    // 3 lines of text leave 120pt: the picture (100 + 0) and two lines fit, three do not.
    const blocks = [tall(3), picture, tall(3)];
    const pages = layoutDocument(model(blocks), measure).pages;
    assert.deepEqual(
        pages.map(p => p.parts.map(q => q.block)),
        [[0], [1, 2]],
        'the picture goes along with its three-line caption'
    );
});
