const http = require('http');
const fs = require('fs');
const path = require('path');
const previewDir = __dirname;
const sourceDir = process.env.PORTAL_SOURCE_DIR;
if (!sourceDir) throw new Error('Set PORTAL_SOURCE_DIR to the existing portal directory.');
const board = JSON.parse(fs.readFileSync(path.join(sourceDir, 'data', 'jobs.json'), 'utf8'));
const installers = JSON.parse(fs.readFileSync(path.join(sourceDir, 'data', 'installers-live.json'), 'utf8'));
const snapshot = {label: 'Local snapshot', board: {jobs: board.jobs || [], designBoard: board.designBoard || {cards: []}, holidayRequests: board.holidayRequests || [], attendanceEntries: board.attendanceEntries || [], notifications: board.notifications || []}, installers};
http.createServer((req, res) => {
  if (req.method !== 'GET') {res.writeHead(405);res.end();return;}
  res.setHeader('Cache-Control', 'no-store');
  if (req.url === '/snapshot') {res.setHeader('Content-Type','application/json');res.end(JSON.stringify(snapshot));return;}
  const script = req.url === '/snapshot.js';
  res.setHeader('Content-Type', script ? 'text/javascript; charset=utf-8' : 'text/html; charset=utf-8');
  res.end(fs.readFileSync(path.join(previewDir,script ? 'snapshot.js' : 'index.html')));
}).listen(5185,'127.0.0.1',()=>console.log('Read-only preview: http://127.0.0.1:5185'));

