const fs = require('node:fs');
const path = require('node:path');

const versionFile = path.join(__dirname, 'src', 'common', 'version.ts');
const timestamp = new Date().toLocaleString();
const buildId = `Build-${Math.floor(Math.random() * 10000)}-${Date.now().toString().slice(-4)}`;

const content = `// This file is auto-generated. Do not edit.
export const BUILD_ID = "${buildId}";
export const BUILD_TIMESTAMP = "${timestamp}";
`;

fs.writeFileSync(versionFile, content);
console.log(`Version updated: ${buildId} (${timestamp})`);
