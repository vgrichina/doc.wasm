// browser.js — Drag-drop harness with multi-page navigation
import { createDocParser } from './loader.js';

const dropZone = document.getElementById('drop-zone');
const canvas = document.getElementById('doc-canvas');
const status = document.getElementById('status');
const textOutput = document.getElementById('text-output');
const prevBtn = document.getElementById('prev-page');
const nextBtn = document.getElementById('next-page');
const pageInfo = document.getElementById('page-info');

let parser;
let currentPage = 0;
let totalPages = 0;

async function init() {
  parser = await createDocParser(canvas);
  status.textContent = 'Ready — drop a .doc file';
}

function updatePageControls() {
  if (pageInfo) pageInfo.textContent = `Page ${currentPage + 1} / ${totalPages}`;
  if (prevBtn) prevBtn.disabled = currentPage <= 0;
  if (nextBtn) nextBtn.disabled = currentPage >= totalPages - 1;
}

function renderCurrentPage() {
  parser.render(currentPage);
  updatePageControls();
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

    totalPages = parser.getPageCount();
    currentPage = 0;
    status.textContent = `Parsed OK — ${totalPages} page(s)`;
    renderCurrentPage();

    const text = parser.getText();
    if (textOutput) {
      textOutput.textContent = text;
    }
  };
  reader.readAsArrayBuffer(file);
}

if (prevBtn) prevBtn.addEventListener('click', () => {
  if (currentPage > 0) { currentPage--; renderCurrentPage(); }
});
if (nextBtn) nextBtn.addEventListener('click', () => {
  if (currentPage < totalPages - 1) { currentPage++; renderCurrentPage(); }
});

// Keyboard nav
document.addEventListener('keydown', (e) => {
  if (e.key === 'ArrowLeft' && currentPage > 0) { currentPage--; renderCurrentPage(); }
  if (e.key === 'ArrowRight' && currentPage < totalPages - 1) { currentPage++; renderCurrentPage(); }
});

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

// Click to browse
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
