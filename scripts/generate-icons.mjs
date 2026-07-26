import { PNG } from "pngjs";
import fs from "node:fs";

function roundedRect(x, y, left, top, right, bottom, radius) {
  const clampedX = Math.max(left + radius, Math.min(x, right - radius));
  const clampedY = Math.max(top + radius, Math.min(y, bottom - radius));
  return Math.hypot(x - clampedX, y - clampedY) <= radius;
}

function colorAt(x, y, size) {
  const scale = size / 512;
  const px = x / scale;
  const py = y / scale;
  const outer = roundedRect(px, py, 0, 0, 512, 512, 116);
  if (!outer) return [13, 15, 20, 255];
  const tile = roundedRect(px, py, 48, 48, 464, 464, 100);
  if (!tile) return [13, 15, 20, 255];

  const t = Math.max(0, Math.min(1, (px + py - 96) / 736));
  let color = [
    Math.round(91 + (36 - 91) * t),
    Math.round(150 + (74 - 150) * t),
    Math.round(240 + (139 - 240) * t),
    255,
  ];

  const vertical = px >= 173 && px <= 226 && py >= 144 && py <= 368;
  const bowlOuter = px >= 200 && px <= 368 && py >= 144 && py <= 323;
  const bowlInner = px >= 226 && px <= 315 && py >= 191 && py <= 276;
  if (vertical || (bowlOuter && !bowlInner)) color = [255, 255, 255, 255];

  if (Math.hypot(px - 388, py - 126) <= 25) color = [82, 210, 115, 255];
  return color;
}

function generate(size, filename) {
  const png = new PNG({ width: size, height: size });
  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      const [red, green, blue, alpha] = colorAt(x + 0.5, y + 0.5, size);
      const index = (size * y + x) * 4;
      png.data[index] = red;
      png.data[index + 1] = green;
      png.data[index + 2] = blue;
      png.data[index + 3] = alpha;
    }
  }
  fs.writeFileSync(filename, PNG.sync.write(png));
}

fs.mkdirSync("public/icons", { recursive: true });
generate(192, "public/icons/icon-192.png");
generate(512, "public/icons/icon-512.png");
generate(512, "public/icons/icon-maskable-512.png");
