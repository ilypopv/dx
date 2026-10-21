'use strict';

// Runs the integration suite inside a real VS Code: node test/e2e/run.js [path to the VS Code executable].
// The suite opens a PDF, a .docx and a .doc in the dx editors, waits for the webviews to report what they
// rendered, makes an edit through a test hook and saves it. Needs a desktop session (it opens a window).
const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { makePdf } = require('../helpers/pdf-fixture');
const { makeDocx } = require('../helpers/docx-fixtures');

const root = path.resolve(__dirname, '../..');
const code =
    process.argv[2] ||
    process.env.VSCODE_EXECUTABLE ||
    (process.platform === 'darwin' ? '/Applications/Visual Studio Code.app/Contents/MacOS/Code' : 'code');

const work = fs.mkdtempSync(path.join(os.tmpdir(), 'dx-e2e-'));
const workspace = path.join(work, 'workspace');
fs.mkdirSync(workspace);
fs.writeFileSync(path.join(workspace, 'sample.pdf'), makePdf(['First page of dx', 'Second page', 'Third page']));
fs.writeFileSync(
    path.join(workspace, 'sample.docx'),
    makeDocx({
        body:
            '<w:p><w:r><w:t>Original first paragraph</w:t></w:r></w:p>' +
            '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> and plain</w:t></w:r></w:p>' +
            '<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>Second page</w:t></w:r></w:p>'
    })
);
fs.copyFileSync(path.join(root, 'test', 'fixtures', 'sample.doc'), path.join(workspace, 'sample.doc'));
fs.writeFileSync(path.join(workspace, 'broken.pdf'), 'this is not a pdf');

const result = spawnSync(
    code,
    [
        workspace,
        '--extensionDevelopmentPath=' + root,
        '--extensionTestsPath=' + path.join(__dirname, 'suite.js'),
        '--user-data-dir=' + path.join(work, 'user'),
        '--extensions-dir=' + path.join(work, 'extensions'),
        '--disable-extensions',
        '--disable-workspace-trust',
        '--skip-welcome',
        '--skip-release-notes',
        '--disable-gpu-sandbox'
    ],
    { stdio: 'inherit', env: { ...process.env, ELECTRON_RUN_AS_NODE: undefined, DX_E2E_WORKSPACE: workspace } }
);
fs.rmSync(work, { recursive: true, force: true });
process.exit(result.status === null ? 1 : result.status);
