// loader.js — Load doc.wasm with canvas imports, copy file, parse + render

export async function createDocParser(canvas) {
  const ctx = canvas ? canvas.getContext('2d') : null;

  // Offscreen canvas for measureText when no visible canvas
  const measureCanvas = new OffscreenCanvas(1, 1);
  const measureCtx = measureCanvas.getContext('2d');

  let currentFont = '12pt serif';
  let memory;

  const fontMap = {
    'Times New Roman': 'Times New Roman, Times, serif',
    'Arial': 'Arial, Helvetica, sans-serif',
    'Calibri': 'Calibri, Arial, sans-serif',
    'Cambria': 'Cambria, Georgia, serif',
    'Courier New': 'Courier New, Courier, monospace',
    'Verdana': 'Verdana, Geneva, sans-serif',
    'Tahoma': 'Tahoma, Geneva, sans-serif',
    'Georgia': 'Georgia, Times, serif',
    'Trebuchet MS': 'Trebuchet MS, sans-serif',
    'Comic Sans MS': 'Comic Sans MS, cursive',
    'Impact': 'Impact, sans-serif',
    'Symbol': 'Symbol, serif',
    'Wingdings': 'Wingdings, serif',
  };

  function updateFont(size, bold, italic, namePtr, nameLen) {
    // size is in half-points, convert to pt
    const pt = size / 2;
    const style = (italic ? 'italic ' : '') + (bold ? 'bold ' : '');
    let family = 'serif';
    if (namePtr && nameLen && memory) {
      const name = new TextDecoder('utf-16le').decode(
        new Uint8Array(memory.buffer, namePtr, nameLen)
      );
      family = fontMap[name] || `"${name}", serif`;
    }
    currentFont = `${style}${pt}pt ${family}`;
    measureCtx.font = currentFont;
    if (ctx) ctx.font = currentFont;
  }

  function decodeText(ptr, len) {
    return new TextDecoder('utf-16le').decode(
      new Uint8Array(memory.buffer, ptr, len)
    );
  }

  const imports = {
    canvas: {
      measureText(ptr, len) {
        const text = decodeText(ptr, len);
        return measureCtx.measureText(text).width;
      },
      setFont(size, bold, italic, namePtr, nameLen) {
        updateFont(size, bold, italic, namePtr, nameLen);
      },
      setColor(rgb) {
        const color = '#' + (rgb & 0xFFFFFF).toString(16).padStart(6, '0');
        if (ctx) ctx.fillStyle = color;
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
        if (!canvas) return;
        canvas.width = widthPx;
        canvas.height = heightPx;
        if (ctx) {
          ctx.clearRect(0, 0, widthPx, heightPx);
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, widthPx, heightPx);
          ctx.fillStyle = '#000000';
          ctx.font = currentFont;
        }
      },
      drawImage(ptr, len, x, y, w, h) {
        if (!ctx) return;
        try {
          const imgData = new Uint8Array(memory.buffer, ptr, len);
          const blob = new Blob([imgData]);
          const url = URL.createObjectURL(blob);
          const img = new Image();
          img.onload = () => {
            ctx.drawImage(img, x, y, w, h);
            URL.revokeObjectURL(url);
          };
          img.src = url;
        } catch {
          ctx.strokeStyle = '#cccccc';
          ctx.strokeRect(x, y, w, h);
        }
      },
    },
    env: {
      log(val) {
        console.log('[wasm]', val);
      },
    },
  };

  const wasmPath = new URL('../doc.wasm', import.meta.url);
  const { instance } = await WebAssembly.instantiateStreaming(
    fetch(wasmPath),
    imports
  );

  memory = instance.exports.memory;
  updateFont(24, 0, 0); // default 12pt

  return {
    /**
     * Load and parse a .doc file from an ArrayBuffer.
     * Returns 0 on success, error code on failure.
     */
    parse(arrayBuffer) {
      const fileBytes = new Uint8Array(arrayBuffer);
      const fileLen = fileBytes.length;

      // Place input after the fixed region (page-aligned at 0x004C0000)
      const inputPtr = 0x004C0000;
      const neededBytes = inputPtr + fileLen + fileLen * 3; // file + working space
      const neededPages = Math.ceil(neededBytes / 65536);
      const currentPages = memory.buffer.byteLength / 65536;
      if (neededPages > currentPages) {
        memory.grow(neededPages - currentPages);
      }

      // Copy file into wasm memory
      new Uint8Array(memory.buffer).set(fileBytes, inputPtr);
      instance.exports.set_input(inputPtr, fileLen);

      return instance.exports.parse();
    },

    /**
     * Render a specific page onto the canvas.
     */
    render(page = 0) {
      instance.exports.render(page);
    },

    /**
     * Extract plain text as a JS string.
     */
    getText() {
      const ptr = instance.exports.get_text_ptr();
      const len = instance.exports.get_text_len();
      if (!ptr || !len) return '';
      return decodeText(ptr, len);
    },

    getPageCount() {
      return instance.exports.get_page_count();
    },

    getErrorCode() {
      return instance.exports.get_error_code();
    },
  };
}
