#!/usr/bin/env node
// test.js — Automated test runner for doc.wasm
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const wasmBytes = readFileSync(join(__dirname, '..', 'doc.wasm'));
const fixturesDir = join(__dirname, 'fixtures');

let passed = 0, failed = 0;

function decodeText(mem, ptr, len) {
  return new TextDecoder('utf-16le').decode(new Uint8Array(mem.buffer, ptr, len));
}

async function loadAndParse(filename) {
  const fileBuffer = readFileSync(join(fixturesDir, filename));

  const imports = {
    canvas: {
      measureText() { return 0; },
      setFont() {}, setColor() {}, fillText() {}, fillRect() {}, setPage() {}, drawImage() {},
    },
    env: { log() {} },
  };

  const { instance } = await WebAssembly.instantiate(wasmBytes, imports);
  const memory = instance.exports.memory;

  const inputPtr = 0x004C0000;
  const needed = Math.ceil((inputPtr + fileBuffer.length * 4) / 65536);
  const cur = memory.buffer.byteLength / 65536;
  if (needed > cur) memory.grow(needed - cur);

  new Uint8Array(memory.buffer).set(
    new Uint8Array(fileBuffer.buffer, fileBuffer.byteOffset, fileBuffer.length),
    inputPtr
  );
  instance.exports.set_input(inputPtr, fileBuffer.length);

  const err = instance.exports.parse();
  const textPtr = instance.exports.get_text_ptr();
  const textLen = instance.exports.get_text_len();
  const text = textPtr && textLen ? decodeText(memory, textPtr, textLen) : '';
  const view = new DataView(memory.buffer);

  return { err, text, view, instance, memory };
}

function getCHPRun(view, index) {
  const CHP_BASE = 0x00094000;
  const ptr = CHP_BASE + index * 28;
  return {
    cpStart: view.getInt32(ptr, true),
    cpEnd: view.getInt32(ptr + 4, true),
    flags: view.getUint32(ptr + 8, true),
    fontSize: view.getUint32(ptr + 12, true),
    color: view.getUint32(ptr + 16, true),
  };
}

function getPAPRun(view, index) {
  const PAP_BASE = 0x00194000;
  const ptr = PAP_BASE + index * 28;
  return {
    cpStart: view.getInt32(ptr, true),
    cpEnd: view.getInt32(ptr + 4, true),
    align: view.getUint32(ptr + 8, true),
  };
}

function assert(name, condition, detail) {
  if (condition) {
    console.log(`  PASS: ${name}`);
    passed++;
  } else {
    console.log(`  FAIL: ${name}${detail ? ' — ' + detail : ''}`);
    failed++;
  }
}

// --- Test cases ---

async function testTdf116194() {
  console.log('\ntdf116194.doc — colored text (sprmCCv)');
  const { err, view } = await loadAndParse('tdf116194.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  // Find any CHP run with color 0xc00000
  let foundColor = false;
  let colors = [];
  for (let i = 0; i < 200; i++) {
    const run = getCHPRun(view, i);
    if (run.cpStart === 0 && run.cpEnd === 0 && i > 0) break;
    if (run.color !== 0) colors.push(`0x${run.color.toString(16).padStart(6, '0')}`);
    if (run.color === 0xc00000) foundColor = true;
  }
  assert('has CHP run with color #c00000', foundColor,
    `non-black colors found: ${colors.join(', ')}`);
}

async function testTdf38778() {
  console.log('\ntdf38778.doc — center alignment');
  const { err, view } = await loadAndParse('tdf38778.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  const pap = getPAPRun(view, 0);
  assert('first PAP run has center alignment (1)', pap.align === 1,
    `align=${pap.align}`);
}

async function testTestDoc() {
  console.log('\ntest.doc — basic text');
  const { err, text } = await loadAndParse('test.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text starts with "My name is Ryan"', text.startsWith('My name is Ryan'),
    `text starts with: "${text.slice(0, 30)}"`);
}

async function testPoiTest() {
  console.log('\npoi-test.doc — field code filtering');
  const { err, text } = await loadAndParse('poi-test.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  const lower = text.toLowerCase();
  assert('text contains "microsoft word document"',
    lower.includes('microsoft word document'),
    `text: "${text.slice(0, 80)}"`);
}

async function testBold() {
  console.log('\ntestBold.doc — mini-stream, bold text');
  const { err, text } = await loadAndParse('testBold.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text contains "Introduction"', text.includes('Introduction'),
    `text: "${text.slice(0, 80)}"`);
}

async function testEmpty() {
  console.log('\nempty.doc — empty document');
  const { err, instance } = await loadAndParse('empty.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  const textLen = instance.exports.get_text_len();
  assert('text_len is small (< 100)', textLen < 100, `text_len=${textLen}`);
}

async function test47304() {
  console.log('\n47304.doc — basic text');
  const { err, text } = await loadAndParse('47304.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text contains "test"', text.toLowerCase().includes('test'),
    `text: "${text.slice(0, 80)}"`);
}

// Run all tests
console.log('doc.wasm test suite');
console.log('===================');

await testTdf116194();
await testTdf38778();
await testTestDoc();
await testPoiTest();
await testBold();
await testEmpty();
await test47304();

console.log(`\n===================`);
console.log(`Results: ${passed} passed, ${failed} failed`);
process.exit(failed > 0 ? 1 : 0);
