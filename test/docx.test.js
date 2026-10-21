'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fflate = require('../media/vendor/fflate.js');
const { readDocx, writeDocx, formatNumber, bulletText } = require('../src/core/docx.js');
const { makeDocx, documentXml, textOf } = require('./helpers/docx-fixtures.js');

const p = (text, pPr = '', rPr = '') =>
    `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const STYLES =
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="2F5496"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="MyHeading"><w:name w:val="My Heading"/><w:basedOn w:val="Heading1"/></w:style>' +
    '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>' +
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:style>';

const NUMBERING =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>';

const numPr = (id, lvl = 0) => `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${id}"/></w:numPr>`;

test('paragraphs, headings and run formatting come through the style chain', () => {
    const bytes = makeDocx({
        styles: STYLES,
        body:
            p('Title text', '<w:pStyle w:val="MyHeading"/>') +
            p('Body', '<w:jc w:val="center"/>', '<w:i/><w:color w:val="FF0000"/>') +
            '<w:p><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>strong</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:u w:val="single"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r></w:p>'
    });
    const { model } = readDocx(bytes);
    const [h, body, mixed] = model.blocks;
    assert.equal(h.style, 'h1', 'a style based on heading 1 is a heading');
    assert.equal(h.runs[0].b, true);
    assert.equal(h.runs[0].size, 16);
    assert.equal(h.runs[0].color, '#2f5496');
    assert.equal(h.layout.before, 12);
    assert.equal(body.align, 'center');
    assert.equal(body.runs[0].i, true);
    assert.equal(body.runs[0].color, '#ff0000');
    assert.equal(body.runs[0].font, 'Calibri');
    assert.equal(body.layout.after, 8);
    assert.deepEqual(body.layout.line, { factor: 259 / 240 });
    assert.equal(mixed.runs[0].b, true, 'character style');
    assert.equal(mixed.runs[1].b, undefined, 'w:val="0" switches bold off');
    assert.equal(mixed.runs[1].u, true);
    assert.equal(mixed.runs[1].sup, true);
    assert.equal(model.font.family, 'Calibri');
    assert.equal(model.font.size, 11);
});

test('numbered and bulleted lists get their markers', () => {
    const bytes = makeDocx({
        numbering: NUMBERING,
        body:
            p('one', numPr(1)) +
            p('one a', numPr(1, 1)) +
            p('one b', numPr(1, 1)) +
            p('two', numPr(1)) +
            p('two a', numPr(1, 1)) +
            p('bullet', numPr(2)) +
            p('five', numPr(3)) +
            p('none', numPr(0))
    });
    const markers = readDocx(bytes).model.blocks.map(b => b.list && b.list.marker);
    assert.deepEqual(markers, ['1.', '1.a)', '1.b)', '2.', '2.a)', '•', '5.', undefined]);
});

test('number formats and private-use bullets', () => {
    assert.equal(formatNumber(4, 'upperRoman'), 'IV');
    assert.equal(formatNumber(28, 'lowerLetter'), 'ab');
    assert.equal(formatNumber(3, 'decimalZero'), '03');
    assert.equal(bulletText('', 'Wingdings'), '▪');
    assert.equal(bulletText('o', 'Courier New'), '◦');
    assert.equal(bulletText('–', undefined), '–');
});

test('tables keep spans, vertical merges, shading and borders', () => {
    const tc = (text, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${p(text)}</w:tc>`;
    const bytes = makeDocx({
        styles: STYLES,
        body:
            '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
            `<w:tr>${tc('wide', '<w:gridSpan w:val="2"/>')}${tc('tall', '<w:vMerge w:val="restart"/><w:shd w:fill="FFFF00"/>')}</w:tr>` +
            `<w:tr>${tc('a')}${tc('b')}${tc('', '<w:vMerge/>')}</w:tr>` +
            `<w:tr>${tc('c')}${tc('d')}${tc('e')}</w:tr>` +
            '</w:tbl>' +
            '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="nil"/></w:tblBorders></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    });
    const [grid, plain] = readDocx(bytes).model.blocks;
    assert.equal(grid.type, 'table');
    assert.equal(grid.borders, true);
    assert.deepEqual(grid.grid, [100, 100, 100]);
    assert.equal(grid.rows[0][0].colspan, 2);
    assert.equal(grid.rows[0][1].rowspan, 2);
    assert.equal(grid.rows[0][1].fill, '#ffff00');
    assert.equal(grid.rows[1].length, 2, 'the continued cell is not rendered');
    assert.equal(grid.rows[2].length, 3);
    assert.equal(textOf(grid.rows[2][2].blocks[0]), 'e');
    assert.equal(plain.borders, false);
});

test('hyperlinks, HYPERLINK fields and their codes', () => {
    const bytes = makeDocx({
        rels: [
            `<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a" TargetMode="External"/>`
        ],
        body:
            '<w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>site</w:t></w:r></w:hyperlink>' +
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.org/b" </w:instrText></w:r>' +
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>field</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
            '<w:r><w:t> after</w:t></w:r></w:p>'
    });
    const [para] = readDocx(bytes).model.blocks;
    assert.equal(textOf(para), 'sitefield after', 'field codes are not shown');
    assert.equal(para.runs[0].href, 'https://example.com/a');
    assert.equal(para.runs[1].href, 'https://example.org/b');
    assert.equal(para.runs[2].href, undefined);
});

test('breaks, tabs, deleted text, images, footnotes and the page setup', () => {
    const bytes = makeDocx({
        rels: [
            `<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`
        ],
        footnotes:
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>The note.</w:t></w:r></w:p></w:footnote>',
        body:
            '<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r><w:del><w:r><w:delText>gone</w:delText></w:r></w:del><w:ins><w:r><w:t>new</w:t></w:r></w:ins></w:p>' +
            '<w:p><w:r><w:t>before break</w:t><w:br w:type="page"/></w:r></w:p>' +
            '<w:p><w:r><w:lastRenderedPageBreak/><w:t>next page</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
            '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
        sectPr: '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>'
    });
    const { model, files } = readDocx(bytes);
    const [first, second, brk, third, picture] = model.blocks;
    assert.equal(textOf(first), 'a\tb' + 'c' + 'new');
    assert.ok(first.runs.some(r => r.br));
    assert.equal(textOf(second), 'before break');
    assert.equal(brk.type, 'pagebreak');
    assert.equal(third.pageStart, true);
    assert.equal(third.runs[1].text, '1');
    assert.equal(third.runs[1].sup, true);
    assert.deepEqual(picture.runs[0].image, {
        part: 'word/media/image1.png',
        type: 'image/png',
        width: 100,
        height: 50
    });
    assert.ok(files['word/media/image1.png'].length > 0);
    assert.equal(textOf(model.notes[0]), '1 The note.');
    assert.deepEqual(model.page, { width: 612, height: 792, top: 36, right: 36, bottom: 36, left: 36 });
});

test('text edits are written back into the original runs', () => {
    const bytes = makeDocx({
        body:
            '<w:p w:rsidR="1"><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r><w:r><w:t xml:space="preserve"> world</w:t></w:r></w:p>' +
            '<w:p><w:pPr><w:rPr><w:i/></w:rPr></w:pPr></w:p>' +
            '<w:p><w:r><w:t>keep &amp; me</w:t></w:r><w:r><w:tab/><w:t>tail</w:t></w:r></w:p>'
    });
    const doc = readDocx(bytes);
    const [first, empty, third] = doc.model.blocks;
    const slots = first.runs.map(r => r.slot);
    const out = writeDocx(doc, [
        // The bold run was retyped, a loose text node was typed after the second run.
        { p: first.id, segments: [{ slot: slots[0], text: 'Howdy' }, { slot: slots[1], text: ' all' }, { text: '!' }] },
        // Typing into an empty paragraph creates a run with the paragraph mark formatting.
        { p: empty.id, segments: [{ text: 'fresh text' }] },
        // The tab was deleted, the text piece after it was kept.
        {
            p: third.id,
            segments: [
                { slot: third.runs[0].slot, text: 'keep & me' },
                { slot: third.runs[2].slot, text: 'tail' }
            ]
        }
    ]);
    const xml = documentXml(out);
    assert.ok(
        xml.includes(
            '<w:p w:rsidR="1"><w:r><w:rPr><w:b/></w:rPr><w:t>Howdy</w:t></w:r><w:r><w:t xml:space="preserve"> all!</w:t></w:r></w:p>'
        ),
        xml
    );
    assert.ok(xml.includes('<w:r><w:rPr><w:i/></w:rPr><w:t>fresh text</w:t></w:r>'), xml);
    assert.ok(xml.includes('<w:t>keep &amp; me</w:t></w:r><w:r><w:t>tail</w:t>'), xml);
    assert.ok(!xml.includes('<w:tab/>'));

    // Everything that was not edited is untouched: the other parts byte for byte.
    const before = fflate.unzipSync(bytes);
    const after = fflate.unzipSync(out);
    assert.deepEqual(Object.keys(after), Object.keys(before));
    for (const name of Object.keys(before)) {
        if (name !== 'word/document.xml') {
            assert.deepEqual(after[name], before[name], name);
        }
    }

    // The written file reads back with the new text.
    const again = readDocx(out).model.blocks.map(textOf);
    assert.deepEqual(again, ['Howdy all!', 'fresh text', 'keep & metail']);
});

test('saving twice does not duplicate typed text', () => {
    const bytes = makeDocx({ body: '<w:p/>' });
    const doc = readDocx(bytes);
    const id = doc.model.blocks[0].id;
    writeDocx(doc, [{ p: id, segments: [{ text: 'one' }] }]);
    const out = writeDocx(doc, [{ p: id, segments: [{ text: 'one two' }] }]);
    assert.deepEqual(readDocx(out).model.blocks.map(textOf), ['one two']);
});

test('a deleted text piece becomes empty, untouched paragraphs stay as they were', () => {
    const body = '<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p><w:p><w:r><w:t>same</w:t></w:r></w:p>';
    const doc = readDocx(makeDocx({ body }));
    const first = doc.model.blocks[0];
    const xml = documentXml(writeDocx(doc, [{ p: first.id, segments: [{ slot: first.runs[1].slot, text: 'b' }] }]));
    assert.ok(xml.includes('<w:r><w:t></w:t></w:r><w:r><w:t>b</w:t></w:r>'), xml);
    assert.ok(xml.includes('<w:p><w:r><w:t>same</w:t></w:r></w:p>'));
});

test('damaged and foreign files are refused with a reason', () => {
    assert.throws(() => readDocx(new Uint8Array([1, 2, 3, 4])), /not a ZIP/);
    assert.throws(() => readDocx(fflate.zipSync({ 'a.txt': new Uint8Array([1]) })), /no main document/);
    const truncated = makeDocx({ body: p('x') }).slice(0, 60);
    assert.throws(() => readDocx(truncated), /damaged|ZIP/);
});
