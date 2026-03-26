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
- `0x00194000` — PAP runs (56-byte records)
- `0x002A4000` — Style table (28-byte records per istd)
- `0x002B4000` — Layout segments (28-byte records)

### PAP run format (56 bytes)

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
| 32 | ilvl | List indent level |
| 36 | ilfo | List format override index (>0 = list paragraph) |
| 40 | dyaLine | Line spacing value (twips or 240ths) |
| 44 | fMultLinespace | 0=absolute twips, 1=proportional (240=single) |
| 48 | fInTable | 1 if paragraph is in a table |
| 52 | fTtp | 1 if paragraph is table row end mark |

### Style table format (28 bytes per entry)

| Offset | Field | Description |
|--------|-------|-------------|
| 0 | flags | CHP flags (bold, italic, etc.) |
| 4 | font_size | Half-points, 0=not set |
| 8 | color | RGB, 0xFFFFFFFF=not set |
| 12 | istdBase | Base style index, 0xFFF=no base |
| 16 | font_index | Font table index, 0xFFFF=not set |
| 20 | alignment | PAP alignment, 0xFF=not set |
| 24 | dxaLeft | PAP left indent twips, 0x80000000=not set |

## Key design decisions

- **Font size bleed fix**: CHP runs can span multiple PAP paragraphs. If a CHP run's font_size matches its start paragraph's style default, the layout engine re-resolves the font_size from the current paragraph's style. This prevents heading sizes from bleeding into body text.
- **Indent via dxaLeft**: Parsed from `sprmPDxaLeft80` (0x845E). Applied in layout as pixel offset from the page left margin.
- **Style-level PAP properties**: Paragraph styles' papx UPX is parsed for alignment and dxaLeft, which serve as defaults before direct formatting overrides.
- **List bullet synthesis**: List paragraphs (ilfo > 0) get a bullet character (U+2022 "•") prepended during layout. Parsed from `sprmPIlvl` (0x260A) and `sprmPIlfo` (0x460B).
- **sprm_size overrides**: Some sprms (0x845E, 0x8460) have misleading spra bits suggesting variable-length, but are actually fixed 2-byte operands. These are special-cased in `$sprm_size`.
- **Line spacing**: Parsed from `sprmPDyaLine` (0x6412), a 4-byte LSPD structure (dyaLine + fMultLinespace). When fMultLinespace=1, dyaLine is in 240ths of a line (240=single, 480=double). When fMultLinespace=0, dyaLine is absolute twips.
- **Table layout**: Table paragraphs detected via `sprmPFInTable` (0x2416) and `sprmPFTtp` (0x2417). Cells in a row are laid out side-by-side with equal-width columns (content width / cell count). Column count is determined by scanning ahead for 0x07 marks until the row-end (fTtp) paragraph.
