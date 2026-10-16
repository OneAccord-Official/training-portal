'use strict';
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
fs.mkdirSync(path.join(root, 'vendor'), {recursive:true});
fs.copyFileSync(path.join(root, 'node_modules/@azure/msal-browser/lib/msal-browser.min.js'), path.join(root, 'vendor/msal-browser.min.js'));
// The locked, locally served library avoids an unversioned third-party script.
fs.copyFileSync(path.join(root, 'node_modules/@azure/msal-browser/LICENSE'), path.join(root, 'vendor/MSAL-LICENSE'));

// Publish only browser assets. Never publish server source, tests or review evidence.
const output = path.join(root, 'dist');
fs.rmSync(output, {recursive:true, force:true});
fs.mkdirSync(output, {recursive:true});
for (const entry of ['index.html', 'auth.html', 'css', 'js', 'data', 'vendor']) {
  fs.cpSync(path.join(root, entry), path.join(output, entry), {recursive:true});
}
