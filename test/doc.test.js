'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { readDoc } = require('../src/core/doc.js');
const { isCfb, readCfb } = require('../src/core/cfb.js');

// Made by the macOS textutil from an HTML page: a real Word 97-2003 binary.
const sample = new Uint8Array(fs.readFileSync(path.join(__dirname, 'fixtures', 'sample.doc')));
const textOf = block => block.runs.map(r => r.text || '').join('');

test('the compound file directory lists the Word streams', () => {
    assert.ok(isCfb(sample));
    const cfb = readCfb(sample);
    assert.ok(cfb.has('WordDocument'));
    assert.ok(cfb.has('1Table') || cfb.has('0Table'));
    assert.equal(cfb.read('WordDocument')[0], 0xec);
});

test('the text, including Cyrillic and cp1252 punctuation, is decoded', () => {
    const { blocks } = readDoc(sample);
    const paragraphs = blocks.filter(b => b.type === 'p').map(textOf);
    assert.equal(paragraphs[0], 'Заголовок документа');
    assert.match(paragraphs[1], /^Обычный абзац с жирным, курсивом и подчёркнутым текстом\./);
    assert.ok(paragraphs.includes('Centered paragraph'));
    assert.ok(
        paragraphs.some(t => t.includes('“quotes” — and a euro sign €')),
        paragraphs.join('|')
    );
});

test('character and paragraph formatting', () => {
    const { blocks } = readDoc(sample);
    const body = blocks[1];
    const bold = body.runs.find(r => r.text === 'жирным');
    const italic = body.runs.find(r => r.text === 'курсивом');
    const underline = body.runs.find(r => r.text === 'подчёркнутым');
    assert.equal(bold.b, true);
    assert.equal(italic.i, true);
    assert.equal(underline.u, true);
    assert.equal(blocks[0].runs[0].size, 24);
    assert.equal(blocks.find(b => b.type === 'p' && textOf(b) === 'Centered paragraph').align, 'center');
});

test('tables come out as rows of cells', () => {
    const table = readDoc(sample).blocks.find(b => b.type === 'table');
    assert.ok(table);
    const cells = table.rows.map(row => row.map(cell => cell.map(textOf).join(' ')));
    assert.deepEqual(cells, [
        ['A1', 'B1'],
        ['A2', 'B2 bold']
    ]);
});

test('files that are not Word 97-2003 documents are refused', () => {
    assert.throws(() => readDoc(new Uint8Array(600)), /not a Word 97-2003/);
    // A compound file without a WordDocument stream (for example an .xls).
    const fake = sample.slice();
    const cfb = readCfb(fake);
    assert.ok(cfb.names().length > 0);
});
