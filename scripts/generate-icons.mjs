import { mkdir } from "node:fs/promises";
import { createRequire } from "node:module";
import { h } from "vue";
import { renderToString } from "vue/server-renderer";
import { Gauge } from "@lucide/vue";

const require = createRequire(import.meta.url);
const sharp = require(process.env.SHARP_PATH || "sharp");
const icon = await renderToString(
  h(Gauge, { size: 280, color: "white", strokeWidth: 1.8 }),
);
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="512" height="512"><rect width="512" height="512" fill="#087f5b"/><g transform="translate(116 116)">${icon}</g></svg>`;
await mkdir(new URL("../public/", import.meta.url), { recursive: true });
for (const size of [192, 512]) {
  await sharp(Buffer.from(svg))
    .resize(size, size)
    .png()
    .toFile(new URL(`../public/icon-${size}.png`, import.meta.url).pathname);
}
