// Renders the app's icons from the logo SVG (src/.../assets/defaultLogo.ts) with the pre-installed Chromium:
//   sharepoint/images/app-icon.png         96 px  (app catalog, marketplace, web part picker)
//   teams/<web part id>_color.png         192 px  (Teams colour icon)
//   teams/<web part id>_outline.png        32 px  (Teams outline icon: white on transparent)
//   docs/brand/time-off-manager-logo-512.png
// Run: node tools/icons/make-icons.cjs   (then node tools/teams/make-app-package.cjs)
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '..', '..');
const src = fs.readFileSync(path.join(root, 'src/webparts/timeOffManager/assets/defaultLogo.ts'), 'utf8');
const LOGO = eval(/LOGO_SVG =\s*([\s\S]*?);\n/.exec(src)[1]);
const id = JSON.parse(fs.readFileSync(path.join(root, 'src/webparts/timeOffManager/TimeOffManagerWebPart.manifest.json'), 'utf8')).id;
// The outline: the same calendar-with-check, drawn as white strokes on transparent.
const OUTLINE =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">' +
  '<rect x="4" y="6.5" width="24" height="22" rx="4" fill="none" stroke="#fff" stroke-width="2.4"/>' +
  '<line x1="4" y1="12.5" x2="28" y2="12.5" stroke="#fff" stroke-width="2.4"/>' +
  '<line x1="10.5" y1="3" x2="10.5" y2="8.5" stroke="#fff" stroke-width="2.6" stroke-linecap="round"/>' +
  '<line x1="21.5" y1="3" x2="21.5" y2="8.5" stroke="#fff" stroke-width="2.6" stroke-linecap="round"/>' +
  '<path d="M10.5 20.5 l3.8 3.8 l7.4 -7.4" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/>' +
  '</svg>';

(async () => {
  const browser = await chromium.launch();
  const render = async (svg, size, out) => {
    const page = await browser.newPage({ viewport: { width: size, height: size } });
    await page.setContent(`<html><body style="margin:0;background:transparent">${svg.replace('<svg ', `<svg width="${size}" height="${size}" `)}</body></html>`);
    fs.mkdirSync(path.dirname(path.join(root, out)), { recursive: true });
    await page.screenshot({ path: path.join(root, out), omitBackground: true });
    await page.close();
    console.log('wrote ' + out);
  };
  await render(LOGO, 96, 'sharepoint/images/app-icon.png');
  await render(LOGO, 192, `teams/${id}_color.png`);
  await render(OUTLINE, 32, `teams/${id}_outline.png`);
  await render(LOGO, 512, 'docs/brand/time-off-manager-logo-512.png');
  await browser.close();
})();
