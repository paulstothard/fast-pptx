import { chromium } from 'playwright';
process.stdout.write(chromium.executablePath());
