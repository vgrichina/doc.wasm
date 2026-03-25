#!/usr/bin/env node
// cli.js — Node CLI for text extraction from .doc files
import { readFileSync } from 'fs';
import { readFile } from 'fs/promises';

const args = process.argv.slice(2);
if (!args.length) {
  console.error('Usage: node cli.js <file.doc>');
  process.exit(1);
}

const filePath = args[0];
const fileBuffer = readFileSync(filePath);

// Load wasm
const wasmPath = new URL('../doc.wasm', import.meta.url);
const wasmBytes = readFileSync(wasmPath);

const imports = {
  canvas: {
    measureText() { return 0; },
    setFont() {},
    setColor() {},
    fillText() {},
    fillRect() {},
    setPage() {},
  },
  env: {
    log(val) { console.error('[wasm]', val); },
  },
};

const { instance } = await WebAssembly.instantiate(wasmBytes, imports);
const memory = instance.exports.memory;

// Place input at 0x004C0000
const inputPtr = 0x004C0000;
const neededBytes = inputPtr + fileBuffer.length + fileBuffer.length * 3;
const neededPages = Math.ceil(neededBytes / 65536);
const currentPages = memory.buffer.byteLength / 65536;
if (neededPages > currentPages) {
  memory.grow(neededPages - currentPages);
}

new Uint8Array(memory.buffer).set(new Uint8Array(fileBuffer.buffer, fileBuffer.byteOffset, fileBuffer.length), inputPtr);
instance.exports.set_input(inputPtr, fileBuffer.length);

const err = instance.exports.parse();
if (err) {
  console.error(`Parse error: code ${err}`);
  process.exit(1);
}

const textPtr = instance.exports.get_text_ptr();
const textLen = instance.exports.get_text_len();
if (textPtr && textLen) {
  const text = new TextDecoder('utf-16le').decode(
    new Uint8Array(memory.buffer, textPtr, textLen)
  );
  process.stdout.write(text);
}
