// Turns the HTML pictures from `RENDER=<dir> node harness.mjs ...` into PNGs. Usage: node screenshot.mjs <dir>
// Needs Playwright (dev tool only; the scripts themselves never use it).
import { createRequire } from "node:module";
import { execSync } from "node:child_process";
import fs from "node:fs";
const require = createRequire(execSync("npm root -g").toString().trim() + "/");
const { chromium } = require("playwright");
const dir = process.argv[2];
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1800, height: 900 } });
for (const f of fs.readdirSync(dir).filter(x => x.endsWith(".html"))) {
  await page.goto("file://" + fs.realpathSync(`${dir}/${f}`));
  await page.screenshot({ path: `${dir}/${f.replace(/\.html$/, ".png")}`, fullPage: true });
}
await browser.close();
console.log("screenshots in", dir);
