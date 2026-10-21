'use strict';

// Executed by VS Code in the extension host (see run.js).
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vscode = require('vscode');

const workspace = process.env.DX_E2E_WORKSPACE;
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

async function statusOf(file, done, timeout = 20000) {
    const uri = vscode.Uri.file(path.join(workspace, file)).toString();
    const end = Date.now() + timeout;
    let last;
    while (Date.now() < end) {
        const views = await vscode.commands.executeCommand('dx._test.views');
        const view = views.find(v => v.uri === uri);
        last = view && view.status;
        if (last && done(last)) {
            return last;
        }
        await sleep(200);
    }
    throw new Error(`${file}: no status in time, last: ${JSON.stringify(last)}`);
}

async function open(file, viewType) {
    await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(path.join(workspace, file)), viewType);
}

async function saveActive(file) {
    const target = path.join(workspace, file);
    const before = fs.statSync(target).mtimeMs;
    await vscode.commands.executeCommand('workbench.action.files.save');
    const end = Date.now() + 10000;
    while (Date.now() < end && fs.statSync(target).mtimeMs === before) {
        await sleep(100);
    }
    return fs.readFileSync(target);
}

const cases = [
    [
        'PDF renders all pages with a text layer',
        async () => {
            await open('sample.pdf', 'dx-pdf-editor');
            const s = await statusOf('sample.pdf', s => s.text || s.error);
            assert.equal(s.error, undefined, s.error);
            assert.equal(s.pages, 3);
            assert.equal(s.text, true, 'the text layer is there, so text can be selected and found');
        }
    ],
    [
        'PDF annotations are saved into the file',
        async () => {
            await vscode.commands.executeCommand(
                'dx._test.edit',
                vscode.Uri.file(path.join(workspace, 'sample.pdf')).toString()
            );
            await sleep(500);
            const bytes = await saveActive('sample.pdf');
            const text = bytes.toString('latin1');
            assert.ok(text.startsWith('%PDF-'));
            assert.ok(/\/FreeText/.test(text), 'the saved PDF has the free text annotation');
            assert.ok(text.includes('dx test note'));
        }
    ],
    [
        'a damaged PDF shows an error instead of a blank page',
        async () => {
            await open('broken.pdf', 'dx-pdf-editor');
            const s = await statusOf('broken.pdf', s => s.error);
            assert.match(s.error, /could not open|damaged|not a PDF/i);
        }
    ],
    [
        'DOCX renders pages and words',
        async () => {
            await open('sample.docx', 'dx-docx-editor');
            const s = await statusOf('sample.docx', s => s.kind || s.error);
            assert.equal(s.error, undefined, s.error);
            assert.equal(s.kind, 'docx');
            assert.equal(s.pages, 2, 'the page break starts a second sheet');
            assert.equal(s.editable, true);
            assert.ok(s.words >= 7);
        }
    ],
    [
        'DOCX text edits are saved',
        async () => {
            await vscode.commands.executeCommand(
                'dx._test.edit',
                vscode.Uri.file(path.join(workspace, 'sample.docx')).toString()
            );
            await sleep(500);
            const bytes = await saveActive('sample.docx');
            const fflate = require(path.join(__dirname, '../../media/vendor/fflate.js'));
            const xml = new TextDecoder().decode(fflate.unzipSync(new Uint8Array(bytes))['word/document.xml']);
            assert.ok(xml.includes('Edited by dx'), xml);
            assert.ok(xml.includes('<w:b/></w:rPr><w:t>Bold</w:t>'), 'other paragraphs keep their markup');
        }
    ],
    [
        'DOC renders read-only',
        async () => {
            await open('sample.doc', 'dx-doc-editor');
            const s = await statusOf('sample.doc', s => s.kind || s.error);
            assert.equal(s.error, undefined, s.error);
            assert.equal(s.kind, 'doc');
            assert.equal(s.editable, false);
            assert.ok(s.words > 20);
        }
    ]
];

exports.run = async function () {
    let failed = 0;
    for (const [name, fn] of cases) {
        try {
            await fn();
            console.log('ok   ' + name);
        } catch (error) {
            failed++;
            console.log('FAIL ' + name + '\n     ' + (error && error.stack ? error.stack : error));
        }
    }
    if (failed) {
        throw new Error(`${failed} of ${cases.length} integration checks failed`);
    }
};
