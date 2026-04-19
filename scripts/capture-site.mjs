#!/usr/bin/env node

import { chromium } from 'playwright';
import { resolve } from 'path';

const [,, url, outputPath] = process.argv;

if (!url || !outputPath) {
  console.error('Usage: capture-site.mjs <url> <output.png>');
  process.exit(1);
}

let browser;
try {
  browser = await chromium.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  // Fixed viewport; screenshot is viewport-only (not full-page).
  // deviceScaleFactor: 2 yields a 2560x1600 physical pixel image, which is
  // sufficient for the downstream resize step that targets 4000px wide.
  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
    deviceScaleFactor: 2,
  });
  const page = await context.newPage();

  await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });

  await page.screenshot({
    path: resolve(outputPath),
    fullPage: false,
  });
} catch (err) {
  console.error(`capture-site: failed to capture '${url}': ${err.message}`);
  process.exit(1);
} finally {
  if (browser) await browser.close();
}
