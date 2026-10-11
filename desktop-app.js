import { spawn } from 'child_process';
import path from 'path';
import fs from 'fs';
import http from 'http';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Check if server is already running, if not start it
function checkServer(callback) {
    const req = http.get('http://localhost:3000/index.html', (res) => {
        callback(true);
    });
    req.on('error', () => {
        // Start server in background
        import('./server.js');
        setTimeout(() => callback(true), 500);
    });
}

const chromePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';

const browserExe = fs.existsSync(chromePath) ? chromePath : (fs.existsSync(edgePath) ? edgePath : null);

if (!browserExe) {
    console.error('No compatible desktop runtime found.');
    process.exit(1);
}

const userProfileDir = path.join(__dirname, '.desktop_profile');

checkServer(() => {
    console.log(`Launching ApexSpreadsheet in dedicated desktop window via: ${browserExe}`);

    const appProcess = spawn(browserExe, [
        `--app=http://localhost:3000`,
        `--user-data-dir=${userProfileDir}`,
        '--window-size=1360,860',
        '--disable-extensions',
        '--no-first-run',
        '--no-default-browser-check'
    ], {
        detached: true,
        stdio: 'ignore'
    });

    appProcess.unref();
    console.log('Desktop application window launched successfully!');
    process.exit(0);
});
