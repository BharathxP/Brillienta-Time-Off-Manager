// Marketplace screenshots (1366 × 768, as Microsoft Marketplace asks) from the local preview with the sample
// organisation: npm run preview, then node tools/screenshots/capture.cjs → sharepoint/images/screenshot-1..5.png
const fs = require('fs');
const http = require('http');
const path = require('path');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..', '..');
const dist = path.join(root, 'preview-dist');
const out = path.join(root, 'sharepoint', 'images');
const TYPES = { '.html': 'text/html', '.js': 'application/javascript', '.woff': 'font/woff' };

const server = http.createServer((req, res) => {
  const file = path.join(dist, decodeURIComponent(req.url.split('?')[0]) === '/' ? 'index.html' : decodeURIComponent(req.url.split('?')[0]));
  if (file.indexOf(dist) !== 0 || !fs.existsSync(file)) {
    res.writeHead(404);
    return res.end();
  }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream' });
  res.end(fs.readFileSync(file));
});

const SHOTS = [
  { name: 'screenshot-1', hash: 'home' },
  { name: 'screenshot-2', hash: 'home', act: async (p) => p.click('.to-nav-cta button').then(() => p.waitForSelector('.ms-Panel .to-request-form')) },
  { name: 'screenshot-3', hash: 'calendar', act: async (p) => p.click('.ms-Dropdown[aria-label="Whose time off"]').then(() => p.click('.ms-Dropdown-item:has-text("Everyone")')) },
  { name: 'screenshot-4', hash: 'approvals', act: async (p) => p.click('.to-table tbody tr').then(() => p.waitForSelector('.ms-Panel')) },
  { name: 'screenshot-5', hash: 'reports' }
];

server.listen(0, async () => {
  const base = `http://localhost:${server.address().port}/`;
  const browser = await chromium.launch();
  for (const s of SHOTS) {
    const page = await browser.newPage({ viewport: { width: 1366, height: 768 } });
    await page.goto(`${base}?role=admin#/${s.hash}`);
    await page.addStyleTag({ content: 'body{padding:0!important;background:#fff!important} .to-app{border:0!important;border-radius:0!important;box-shadow:none!important;min-height:768px!important}' });
    await page.waitForSelector('.to-app');
    await page.evaluate(() => document.fonts.ready);
    await page.waitForTimeout(1200);
    if (s.act) await s.act(page);
    await page.mouse.move(1, 760); // no hover tooltips in the picture
    await page.waitForTimeout(700);
    await page.screenshot({ path: path.join(out, `${s.name}.png`) });
    console.log(`wrote sharepoint/images/${s.name}.png`);
    await page.close();
  }
  await browser.close();
  server.close();
});
