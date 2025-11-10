const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  page.on('console', (msg) => {
    const type = msg.type();
    console.log('[console]', type, msg.text());
  });
  page.on('pageerror', (error) => {
    console.log('[pageerror]', error.message);
  });
  await page.goto('https://mapaindicador-engeman.vercel.app/analise.html', { waitUntil: 'networkidle' });
  await page.waitForTimeout(5000);
  await browser.close();
})();
