/**
 * Generates assets/icon.png — the Atlyn Calendar Slicer mark.
 *
 * Draws a small calendar page: a rounded card, a header band with two binder
 * tabs, and a 3-row grid of day dots with one highlighted (selected) day.
 * Rendered at 4x and downscaled for crisp antialiasing. Run: npm run icon.
 *
 * @napi-rs/canvas is a devDependency only; it never ships inside the .pbiviz.
 */
import { createCanvas } from "@napi-rs/canvas";
import { writeFileSync } from "node:fs";
import { resolve } from "node:path";

const SIZE = 20;
const SCALE = 4;
const W = SIZE * SCALE;

const CARD = "#118DFF";
const HEADER = "#12239E";
const PAGE = "#FFFFFF";
const DOT = "#B9C6E8";
const SELECTED = "#E66C37";

function roundedRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
}

const canvas = createCanvas(W, W);
const ctx = canvas.getContext("2d");
ctx.clearRect(0, 0, W, W);
ctx.scale(SCALE, SCALE);

// Card + header band.
ctx.fillStyle = CARD;
roundedRect(ctx, 2, 3, 16, 15, 2.5);
ctx.fill();
ctx.fillStyle = HEADER;
roundedRect(ctx, 2, 3, 16, 5, 2.5);
ctx.fill();
ctx.fillRect(2, 6, 16, 2);

// Binder tabs.
ctx.fillStyle = HEADER;
roundedRect(ctx, 6, 1.5, 2, 3, 1);
ctx.fill();
roundedRect(ctx, 12, 1.5, 2, 3, 1);
ctx.fill();

// Page area.
ctx.fillStyle = PAGE;
roundedRect(ctx, 3.5, 9, 13, 7.5, 1);
ctx.fill();

// Day dots (3 x 4 grid), one selected.
const cols = 4;
const rows = 3;
const x0 = 5;
const y0 = 10.5;
const dx = 2.6;
const dy = 2.1;
for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
        const selected = r === 1 && c === 2;
        ctx.fillStyle = selected ? SELECTED : DOT;
        ctx.beginPath();
        ctx.arc(x0 + c * dx, y0 + r * dy, selected ? 0.95 : 0.7, 0, Math.PI * 2);
        ctx.fill();
    }
}

// Downscale to the final 20x20 for antialiasing.
const out = createCanvas(SIZE, SIZE);
const octx = out.getContext("2d");
octx.imageSmoothingEnabled = true;
octx.drawImage(canvas, 0, 0, SIZE, SIZE);

const target = resolve(process.cwd(), "assets", "icon.png");
writeFileSync(target, out.toBuffer("image/png"));
console.log(`Wrote ${target} (${SIZE}x${SIZE})`);
