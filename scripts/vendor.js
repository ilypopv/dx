'use strict';

// Copies the runtime assets of pdf.js out of node_modules into media/pdfjs/.
// The webview loads them through asWebviewUri(), so they must live inside the extension.
// Only what keeps the package small is copied:
//   - the viewer stylesheet and the icons it references;
//   - the CMaps of the Unicode encodings and the Adobe character collections, which cover the
//     CJK documents with non-embedded fonts (the legacy national encodings are left out);
//   - the Symbol and Dingbats fonts: the other standard fonts resolve to the system fonts.
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const source = path.join(root, 'node_modules', 'pdfjs-dist');
const target = path.join(root, 'media', 'pdfjs');

if (!fs.existsSync(source)) {
    console.error('pdfjs-dist is not installed; run npm install first.');
    process.exit(1);
}

fs.rmSync(target, { recursive: true, force: true });

function copy(from, to, filter) {
    fs.mkdirSync(to, { recursive: true });
    let count = 0;
    for (const name of fs.readdirSync(from)) {
        if (filter(name)) {
            fs.copyFileSync(path.join(from, name), path.join(to, name));
            count++;
        }
    }
    return count;
}

const CMAPS = /^(Uni(CNS|GB|JIS|KS)-(UCS2|UTF16)-[HV]|Adobe-(CNS1|GB1|Japan1|Korea1)-UCS2)\.bcmap$/;
const FONTS = /^(FoxitSymbol\.pfb|FoxitDingbats\.pfb|LICENSE_FOXIT)$/;

fs.mkdirSync(target, { recursive: true });
fs.copyFileSync(path.join(source, 'web', 'pdf_viewer.css'), path.join(target, 'pdf_viewer.css'));
fs.copyFileSync(path.join(source, 'LICENSE'), path.join(target, 'LICENSE'));
const images = copy(path.join(source, 'web', 'images'), path.join(target, 'images'), name => /\.(svg|gif)$/.test(name));
const cmaps = copy(path.join(source, 'cmaps'), path.join(target, 'cmaps'), name => CMAPS.test(name));
const fonts = copy(path.join(source, 'standard_fonts'), path.join(target, 'standard_fonts'), name => FONTS.test(name));
console.log(`pdf.js assets: ${images} images, ${cmaps} cmaps, ${fonts} font files`);
