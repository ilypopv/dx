'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { parseXml, serialize, child, find, textOf, setAttr } = require('../src/core/xml.js');

const SOURCE =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<!-- a comment --><w:document xmlns:w="urn:w" mc:Ignorable="w14 wp14">' +
    '<w:body><w:p w:rsidR="00AB"><w:r><w:t xml:space="preserve"> a &amp; b &lt;c&gt; </w:t></w:r>' +
    '<w:r><w:t a=\'single "quoted"\'>x</w:t></w:r><w:empty></w:empty><w:br/>' +
    '<![CDATA[raw <text>]]></w:p></w:body></w:document>';

test('a parsed part serializes back to the same text', () => {
    assert.equal(serialize(parseXml(SOURCE)), SOURCE);
});

test('text and attributes are decoded', () => {
    const doc = parseXml(SOURCE);
    const t = find(doc, 'w:t');
    assert.equal(textOf(t), ' a & b <c> ');
    assert.equal(t.attrs['xml:space'], 'preserve');
    const p = find(doc, 'w:p');
    assert.equal(p.attrs['w:rsidR'], '00AB');
    assert.equal(child(child(p, 'w:r'), 'w:t').attrs['xml:space'], 'preserve');
});

test('an edited node is re-encoded, the rest stays verbatim', () => {
    const doc = parseXml(SOURCE);
    const t = find(doc, 'w:t');
    t.children = [{ text: 'new <&> text' }];
    setAttr(t, 'xml:space', 'preserve');
    const out = serialize(doc);
    assert.ok(out.includes('<w:t xml:space="preserve">new &lt;&amp;&gt; text</w:t>'));
    assert.ok(out.includes(`<w:t a='single "quoted"'>x</w:t>`));
    assert.ok(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<!-- a comment -->'));
});

test('numeric character references are decoded', () => {
    const doc = parseXml('<a>&#x41;&#66;&#x1F600;</a>');
    assert.equal(textOf(doc), 'AB😀');
});

test('unterminated markup is reported', () => {
    assert.throws(() => parseXml('<a><b'), /Unterminated/);
});
