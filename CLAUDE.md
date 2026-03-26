# doc.wasm

A Microsoft .doc (OLE2/Word 97-2003) parser and renderer written entirely in raw WebAssembly Text format (WAT).

## Build

```bash
bash build.sh
```

Requires `wat2wasm` from the [WebAssembly Binary Toolkit (WABT)](https://github.com/WebAssembly/wabt).

## Test

```bash
node test/test.js
```

## Render fixtures to PNG

```bash
node js/cli.js test/fixtures/<file>.doc --render /tmp/output.png
```

Multi-page docs produce numbered outputs: `output-0.png`, `output-1.png`, etc.

## Architecture

Single WAT file (`wat/main.wat`) containing all parsing and layout logic:

- **OLE2/CFBF container** — FAT/MiniFAT chain walking, directory parsing, stream extraction
- **FIB** — File Information Block parsing for stream offsets
- **Piece Table (CLX)** — maps character positions (CPs) to file character offsets (FCs)
- **CHP (Character Properties)** — FKP-based parsing of character formatting runs (bold, italic, font size, color, font index, highlight)
- **PAP (Paragraph Properties)** — FKP-based parsing of paragraph runs (alignment, spacing, indent, dxaLeft)
- **SEP (Section Properties)** — page dimensions and margins
- **STSH (Stylesheet)** — style inheritance chain for CHP defaults (flags, font_size, font_index)
- **Font Table (SttbfFfn)** — font name extraction
- **Layout engine** — word-wrapping, pagination, alignment, indent
- **Renderer** — emits draw calls via imported canvas API

### Memory layout

Fixed regions at known offsets (see top of `wat/main.wat`):
- `0x00094000` — CHP runs (28-byte records)
- `0x00194000` — PAP runs (32-byte records)
- `0x002A4000` — Style table (20-byte records per istd)
- `0x002B4000` — Layout segments (28-byte records)

### PAP run format (32 bytes)

| Offset | Field | Description |
|--------|-------|-------------|
| 0 | cp_start | Start character position |
| 4 | cp_end | End character position |
| 8 | alignment | 0=left, 1=center, 2=right, 3=justify |
| 12 | space_before | Twips |
| 16 | space_after | Twips |
| 20 | first_indent | First-line indent (twips, signed) |
| 24 | istd | Style index |
| 28 | dxaLeft | Left indent (twips, signed) |

## Key design decisions

- **Font size bleed fix**: CHP runs can span multiple PAP paragraphs. If a CHP run's font_size matches its start paragraph's style default, the layout engine re-resolves the font_size from the current paragraph's style. This prevents heading sizes from bleeding into body text.
- **Indent via dxaLeft**: Parsed from `sprmPDxaLeft` (0x840F) and `sprmPDxaLeft80` (0x845E). Applied in layout as pixel offset from the page left margin.
