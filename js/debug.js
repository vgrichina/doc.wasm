#!/usr/bin/env node
// debug.js — Dump parsed CHP/PAP runs from wasm memory
import { readFileSync } from 'fs';

const filePath = process.argv[2];
if (!filePath) { console.error('Usage: node debug.js <file.doc>'); process.exit(1); }

const fileBuffer = readFileSync(filePath);
const wasmBytes = readFileSync(new URL('../doc.wasm', import.meta.url));

let canvas, ctx, currentFont = '12pt serif';
try {
  const cm = await import('canvas');
  canvas = cm.createCanvas(816, 1056);
  ctx = canvas.getContext('2d');
  ctx.font = currentFont;
} catch {}

function decodeText(mem, ptr, len) {
  return new TextDecoder('utf-16le').decode(new Uint8Array(mem.buffer, ptr, len));
}

const imports = {
  canvas: {
    measureText(ptr, len) { if (!ctx) return 0; return ctx.measureText(decodeText(memory, ptr, len)).width; },
    setFont(size, bold, italic) {
      const pt = size / 2;
      currentFont = `${italic?'italic ':''}${bold?'bold ':''}${pt}pt serif`;
      if (ctx) ctx.font = currentFont;
    },
    setColor() {}, fillText() {}, fillRect() {}, setPage() {},
  },
  env: { log(v) { console.error('[wasm]', v); } },
};

const { instance } = await WebAssembly.instantiate(wasmBytes, imports);
const memory = instance.exports.memory;

const inputPtr = 0x004C0000;
const needed = Math.ceil((inputPtr + fileBuffer.length * 4) / 65536);
const cur = memory.buffer.byteLength / 65536;
if (needed > cur) memory.grow(needed - cur);

new Uint8Array(memory.buffer).set(new Uint8Array(fileBuffer.buffer, fileBuffer.byteOffset, fileBuffer.length), inputPtr);
instance.exports.set_input(inputPtr, fileBuffer.length);

const err = instance.exports.parse();
if (err) { console.error('Parse error:', err); process.exit(1); }

const view = new DataView(memory.buffer);

// CHP runs at 0x00094000, 28 bytes each
const CHP_BASE = 0x00094000;
// Read chp_run_count — it's a global but not exported. Let's scan for non-zero runs.
console.log('=== CHP Runs ===');
for (let i = 0; i < 200; i++) {
  const ptr = CHP_BASE + i * 28;
  const cpStart = view.getInt32(ptr, true);
  const cpEnd = view.getInt32(ptr + 4, true);
  const flags = view.getUint32(ptr + 8, true);
  const fontSize = view.getUint32(ptr + 12, true);
  const color = view.getUint32(ptr + 16, true);
  if (cpStart === 0 && cpEnd === 0 && i > 0) break;
  const bold = flags & 1 ? 'B' : '.';
  const italic = flags & 2 ? 'I' : '.';
  const ul = flags & 4 ? 'U' : '.';
  const strike = flags & 8 ? 'S' : '.';
  console.log(`  [${i}] cp=${cpStart}-${cpEnd} ${bold}${italic}${ul}${strike} size=${fontSize/2}pt color=#${color.toString(16).padStart(6,'0')}`);
}

// PAP runs at 0x00194000, 28 bytes each
const PAP_BASE = 0x00194000;
console.log('\n=== PAP Runs ===');
for (let i = 0; i < 200; i++) {
  const ptr = PAP_BASE + i * 28;
  const cpStart = view.getInt32(ptr, true);
  const cpEnd = view.getInt32(ptr + 4, true);
  const align = view.getUint32(ptr + 8, true);
  const spaceBefore = view.getUint32(ptr + 12, true);
  const spaceAfter = view.getUint32(ptr + 16, true);
  if (cpStart === 0 && cpEnd === 0 && i > 0) break;
  const alignNames = ['left', 'center', 'right', 'justify'];
  console.log(`  [${i}] cp=${cpStart}-${cpEnd} align=${alignNames[align]||align} spaceBefore=${spaceBefore} spaceAfter=${spaceAfter}`);
}

// Page count
console.log(`\nPages: ${instance.exports.get_page_count()}`);
console.log(`Text length: ${instance.exports.get_text_len()} bytes`);
