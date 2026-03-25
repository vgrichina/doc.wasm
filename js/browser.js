// browser.js — Drag-drop harness
import { createDocParser } from './loader.js';

const dropZone = document.getElementById('drop-zone');
const canvas = document.getElementById('doc-canvas');
const status = document.getElementById('status');
const textOutput = document.getElementById('text-output');

let parser;

async function init() {
  parser = await createDocParser(canvas);
  status.textContent = 'Ready — drop a .doc file';
}

function handleFile(file) {
  status.textContent = `Loading ${file.name} (${file.size} bytes)...`;

  const reader = new FileReader();
  reader.onload = () => {
    const err = parser.parse(reader.result);
    if (err) {
      status.textContent = `Parse error: code ${err}`;
      return;
    }

    status.textContent = `Parsed OK — ${parser.getPageCount()} page(s)`;
    parser.render(0);

    const text = parser.getText();
    if (textOutput) {
      textOutput.textContent = text;
    }
  };
  reader.readAsArrayBuffer(file);
}

// Drag & drop
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

// Also allow click to select file
dropZone.addEventListener('click', () => {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.doc';
  input.onchange = () => {
    if (input.files[0]) handleFile(input.files[0]);
  };
  input.click();
});

init();
