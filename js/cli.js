#!/usr/bin/env node
// cli.js — Node CLI for .doc text extraction and PNG rendering
import { readFileSync, writeFileSync } from 'fs';

const args = process.argv.slice(2);
if (!args.length) {
  console.error('Usage: node cli.js <file.doc> [--render output.png] [--text]');
  process.exit(1);
}

let filePath = null;
let renderPath = null;
let textMode = false;

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--render' && args[i + 1]) {
    renderPath = args[++i];
  } else if (args[i] === '--text') {
    textMode = true;
  } else {
    filePath = args[i];
  }
}

if (!filePath) {
  console.error('No input file specified');
  process.exit(1);
}

// Default to text mode if no --render
if (!renderPath) textMode = true;

const fileBuffer = readFileSync(filePath);
const wasmPath = new URL('../doc.wasm', import.meta.url);
const wasmBytes = readFileSync(wasmPath);

// Try to load canvas for rendering
let createCanvas = null;
if (renderPath) {
  try {
    const canvasModule = await import('canvas');
    createCanvas = canvasModule.createCanvas;
  } catch {
    console.error('node-canvas not installed. Run: npm install canvas');
    console.error('Falling back to text-only mode.');
    textMode = true;
    renderPath = null;
  }
}

let canvas = null;
let ctx = null;
let currentFont = '12pt serif';

function decodeText(ptr, len) {
  return new TextDecoder('utf-16le').decode(
    new Uint8Array(memory.buffer, ptr, len)
  );
}

function updateFont(size, bold, italic) {
  const pt = size / 2;
  const style = (italic ? 'italic ' : '') + (bold ? 'bold ' : '');
  currentFont = `${style}${pt}pt serif`;
  if (ctx) ctx.font = currentFont;
}

const imports = {
  canvas: {
    measureText(ptr, len) {
      if (!ctx) return 0;
      const text = decodeText(ptr, len);
      return ctx.measureText(text).width;
    },
    setFont(size, bold, italic) {
      updateFont(size, bold, italic);
    },
    setColor(rgb) {
      if (ctx) ctx.fillStyle = '#' + (rgb & 0xFFFFFF).toString(16).padStart(6, '0');
    },
    fillText(ptr, len, x, y) {
      if (!ctx) return;
      const text = decodeText(ptr, len);
      ctx.fillText(text, x, y);
    },
    fillRect(x, y, w, h) {
      if (ctx) ctx.fillRect(x, y, w, h);
    },
    setPage(pageNum, widthPx, heightPx) {
      if (!createCanvas) return;
      canvas = createCanvas(widthPx, heightPx);
      ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, widthPx, heightPx);
      ctx.fillStyle = '#000000';
      ctx.font = currentFont;
    },
    drawImage(ptr, len, x, y, w, h) {
      if (!ctx || !canvas) return;
      // Queue image for async loading after render
      const imgBuf = Buffer.from(new Uint8Array(memory.buffer, ptr, len));
      pendingImages.push({ imgBuf, x, y, w, h, canvas, ctx });
    },
  },
  env: {
    log(val) { console.error('[wasm]', val); },
  },
};

const pendingImages = [];

const { instance } = await WebAssembly.instantiate(wasmBytes, imports);
const memory = instance.exports.memory;

// If rendering, create an initial canvas for measureText during layout
if (createCanvas) {
  canvas = createCanvas(816, 1056);
  ctx = canvas.getContext('2d');
  ctx.font = currentFont;
}

const inputPtr = 0x004C0000;
const neededBytes = inputPtr + fileBuffer.length + fileBuffer.length * 3;
const neededPages = Math.ceil(neededBytes / 65536);
const currentPages = memory.buffer.byteLength / 65536;
if (neededPages > currentPages) {
  memory.grow(neededPages - currentPages);
}

new Uint8Array(memory.buffer).set(
  new Uint8Array(fileBuffer.buffer, fileBuffer.byteOffset, fileBuffer.length),
  inputPtr
);
instance.exports.set_input(inputPtr, fileBuffer.length);

const err = instance.exports.parse();
if (err) {
  console.error(`Parse error: code ${err}`);
  process.exit(1);
}

// Text extraction
if (textMode) {
  const textPtr = instance.exports.get_text_ptr();
  const textLen = instance.exports.get_text_len();
  if (textPtr && textLen) {
    const text = decodeText(textPtr, textLen);
    process.stdout.write(text);
    if (!renderPath) process.stdout.write('\n');
  }
}

// Render to PNG
if (renderPath && createCanvas) {
  const pageCount = instance.exports.get_page_count();
  console.error(`Rendering ${pageCount} page(s)`);

  const { loadImage } = await import('canvas');

  async function renderPage(p) {
    pendingImages.length = 0;
    instance.exports.render(p);
    // Process pending images
    for (const img of pendingImages) {
      try {
        const loaded = await loadImage(img.imgBuf);
        img.ctx.drawImage(loaded, img.x, img.y, img.w, img.h);
      } catch (e) {
        img.ctx.strokeStyle = '#cccccc';
        img.ctx.strokeRect(img.x, img.y, img.w, img.h);
        img.ctx.fillStyle = '#000000';
      }
    }
  }

  if (pageCount === 1) {
    await renderPage(0);
    if (canvas) {
      const pngBuffer = canvas.toBuffer('image/png');
      writeFileSync(renderPath, pngBuffer);
      console.error(`Wrote ${renderPath} (${pngBuffer.length} bytes)`);
    }
  } else {
    const base = renderPath.replace(/\.png$/i, '');
    for (let p = 0; p < pageCount; p++) {
      await renderPage(p);
      if (canvas) {
        const outPath = `${base}-${p}.png`;
        const pngBuffer = canvas.toBuffer('image/png');
        writeFileSync(outPath, pngBuffer);
        console.error(`Wrote ${outPath} (${pngBuffer.length} bytes)`);
      }
    }
  }
}
