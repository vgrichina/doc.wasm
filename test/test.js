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
  const { err, text, view } = await loadAndParse('tdf116194.doc');
  assert('parse succeeds', err === 0, `err=${err}`);

  // Validate full text content — all 13 color labels with RGB values
  assert('text contains "Dark Red rgb(192,0,0)"', text.includes('Dark Red rgb(192,0,0)'),
    `text: "${text.slice(0, 60)}"`);
  assert('text contains "CornFlowerBlue Accent 1"', text.includes('CornFlowerBlue Accent 1'),
    `missing CornFlowerBlue`);
  assert('text contains "SteelBlue Accent 5 rgb(180,198,231)"',
    text.includes('SteelBlue Accent 5 rgb(180,198,231)'), `missing last entry`);

  // Validate specific CHP color values (sprmCCv 24-bit RGB)
  const expectedColors = [0xc00000, 0x92d050, 0xffff00, 0x00b0f0, 0x0070c0];
  const foundColors = new Set();
  for (let i = 0; i < 200; i++) {
    const run = getCHPRun(view, i);
    if (run.cpStart === 0 && run.cpEnd === 0 && i > 0) break;
    foundColors.add(run.color);
  }
  for (const c of expectedColors) {
    const hex = '#' + c.toString(16).padStart(6, '0');
    assert(`has CHP color ${hex}`, foundColors.has(c),
      `colors found: ${[...foundColors].map(v => '#' + v.toString(16).padStart(6, '0')).join(', ')}`);
  }
}

async function testTdf38778() {
  console.log('\ntdf38778.doc — center alignment');
  const { err, text, view } = await loadAndParse('tdf38778.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  const clean = text.replace(/[\x00-\x20\r\n]+$/g, '').replace(/[\x00-\x1f]/g, '');
  assert('text is "1"', clean === '1', `text="${clean}"`);
  const pap = getPAPRun(view, 0);
  assert('first PAP run has center alignment (1)', pap.align === 1,
    `align=${pap.align}`);
}

async function testTestDoc() {
  console.log('\ntest.doc — basic text');
  const { err, text } = await loadAndParse('test.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text starts with "My name is Ryan"', text.startsWith('My name is Ryan'),
    `starts: "${text.slice(0, 30)}"`);
  assert('text contains "This is a test blahblahblayh"', text.includes('This is a test blahblahblayh'),
    `missing phrase`);
  assert('text contains "several FKPs for testing purposes"',
    text.includes('several FKPs for testing purposes'), `missing FKP phrase`);
  // Raw text has \r paragraph breaks between chars near end; verify last visible content
  const visible = text.replace(/\r/g, '');
  assert('visible text ends with "sds dsd"', visible.trimEnd().endsWith('sds dsd'),
    `ends: "${visible.slice(-30)}"`);
  assert('text length > 300 chars', text.length > 300, `len=${text.length}`);
}

async function testPoiTest() {
  console.log('\npoi-test.doc — field code filtering');
  const { err, text } = await loadAndParse('poi-test.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text contains "Microsoft word document with 4 embedded images"',
    text.includes('Microsoft word document with 4 embedded images'),
    `text: "${text.slice(0, 80)}"`);
  assert('text contains "Image 1"', text.includes('Image 1'), `missing Image 1`);
  assert('text contains "Image 4"', text.includes('Image 4'), `missing Image 4`);
  assert('text contains "MS Word Document"', text.includes('MS Word Document'),
    `missing MS Word Document`);
  // Field codes should be filtered — no raw field instruction markers
  assert('no field begin marker (0x13)', !text.includes('\x13'), `found field begin`);
  assert('no field sep marker (0x14)', !text.includes('\x14'), `found field sep`);
}

async function testBold() {
  console.log('\ntestBold.doc — mini-stream, unicode, text replacement');
  const { err, text } = await loadAndParse('testBold.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text starts with "Introduction"', text.startsWith('Introduction'),
    `starts: "${text.slice(0, 30)}"`);
  assert('text contains "MS-Word 97 formatted document"',
    text.includes('MS-Word 97 formatted document'), `missing format description`);
  assert('text contains "NeoOffice"', text.includes('NeoOffice'), `missing NeoOffice`);
  assert('text contains Unicode em-dash U+2014', text.includes('\u2014'),
    `missing em-dash`);
  assert('text contains Unicode check mark U+2714', text.includes('\u2714'),
    `missing check mark`);
  assert('text contains "${organization}"', text.includes('${organization}'),
    `missing template variable`);
  assert('text ends with "${organization}!"', text.trimEnd().endsWith('${organization}!'),
    `ends: "${text.slice(-40)}"`);
}

async function testEmpty() {
  console.log('\nempty.doc — empty document');
  const { err, text, instance } = await loadAndParse('empty.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  const textLen = instance.exports.get_text_len();
  assert('text_len is small (< 100)', textLen < 100, `text_len=${textLen}`);
  assert('text is blank after trim', text.trim().length === 0,
    `trimmed text: "${text.trim()}", len=${text.trim().length}`);
}

async function test47304() {
  console.log('\n47304.doc — quoted text');
  const { err, text } = await loadAndParse('47304.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  const clean = text.replace(/[\x00-\x20\r\n]+$/g, '').replace(/[\x00-\x1f]/g, '');
  assert('exact text is: Just  a \u201Ctest\u201D', clean === 'Just  a \u201Ctest\u201D',
    `text="${clean}"`);
}

async function testTdf118412() {
  console.log('\ntdf118412.doc — long doc with field codes, headings');
  const { err, text } = await loadAndParse('tdf118412.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text contains "Xar Format Specification"', text.includes('Xar Format Specification'),
    `missing title`);
  assert('text contains "ultra-compact, open, vector graphic format"',
    text.includes('ultra-compact, open, vector graphic format'), `missing abstract phrase`);
  assert('text contains "Adobe Postscript rendering model"',
    text.includes('Adobe Postscript rendering model'), `missing background phrase`);
  // Field codes like INCLUDEPICTURE and HYPERLINK should be filtered
  assert('no INCLUDEPICTURE field instruction', !text.includes('INCLUDEPICTURE'),
    `found raw field code`);
  assert('no HYPERLINK field instruction', !text.includes('HYPERLINK'),
    `found raw field code`);
  assert('text length > 1500', text.length > 1500, `len=${text.length}`);
}

async function testTdf138345() {
  console.log('\ntdf138345.doc — paragraph style shading');
  const { err, text } = await loadAndParse('tdf138345.doc');
  assert('parse succeeds', err === 0, `err=${err}`);
  assert('text contains "Paragraph styles  can define shading"',
    text.includes('Paragraph styles  can define shading'), `missing first para`);
  assert('text contains "overridden or cancelled by a character style"',
    text.includes('overridden or cancelled by a character style'), `missing second para`);
  assert('text contains "direct formatting"', text.includes('direct formatting'),
    `missing third para concept`);
  assert('text contains "cancel a character style background"',
    text.includes('cancel a character style background'), `missing last para`);
}

async function testTdf59896() {
  console.log('\ntdf59896.doc — Word 6.0/95 format (unsupported)');
  const { err } = await loadAndParse('tdf59896.doc');
  // wIdent=0xA5DC (Word 6/95), not 0xA5EC (Word 97+) — should fail gracefully
  assert('parse returns ERR_BAD_FIB (5)', err === 5, `err=${err}`);
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
await testTdf118412();
await testTdf138345();
await testTdf59896();

console.log(`\n===================`);
console.log(`Results: ${passed} passed, ${failed} failed`);
process.exit(failed > 0 ? 1 : 0);
