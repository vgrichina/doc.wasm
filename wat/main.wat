(module
  ;; ============================================================
  ;; doc.wasm — Microsoft .doc (OLE2/CFBF) parser + renderer
  ;; Written in raw WAT
  ;; ============================================================

  ;; ── Imports (canvas-like API provided by JS) ────────────────

  (import "canvas" "measureText" (func $measureText (param $ptr i32) (param $len i32) (result f32)))
  (import "canvas" "setFont"     (func $setFont (param $size i32) (param $bold i32) (param $italic i32) (param $namePtr i32) (param $nameLen i32)))
  (import "canvas" "setColor"    (func $setColor (param $rgb i32)))
  (import "canvas" "fillText"    (func $fillText (param $ptr i32) (param $len i32) (param $x f32) (param $y f32)))
  (import "canvas" "fillRect"    (func $fillRect (param $x f32) (param $y f32) (param $w f32) (param $h f32)))
  (import "canvas" "setPage"     (func $setPage (param $pageNum i32) (param $widthPx f32) (param $heightPx f32)))
  (import "canvas" "drawImage"   (func $drawImage (param $ptr i32) (param $len i32) (param $x f32) (param $y f32) (param $w f32) (param $h f32)))
  (import "env"    "log"         (func $log (param i32)))

  ;; ── Memory ──────────────────────────────────────────────────
  ;; Initial 16 MiB (256 pages of 64 KiB each)
  (memory (export "memory") 256)

  ;; ── Fixed region addresses ──────────────────────────────────
  ;; 0x00000000  256 KB  FAT array
  ;; 0x00040000  256 KB  Mini-FAT array
  ;; 0x00080000  64 KB   Directory entries
  ;; 0x00090000  16 KB   FIB parsed fields
  ;; 0x00094000  1 MiB   CHP runs
  ;; 0x00194000  1 MiB   PAP runs
  ;; 0x00294000  64 KB   SEP runs
  ;; 0x002A4000  64 KB   Stylesheet
  ;; 0x002B4000  2 MiB   LAYOUT
  ;; 0x004B4000          end of fixed region (~4.7 MiB)

  ;; ── Globals ─────────────────────────────────────────────────

  ;; Input location (set by JS via set_input)
  (global $input_ptr (mut i32) (i32.const 0))
  (global $input_len (mut i32) (i32.const 0))

  ;; Arena bump allocator (starts after fixed region)
  (global $arena_base (mut i32) (i32.const 0x004C0000))  ;; page-aligned after fixed
  (global $arena_ptr  (mut i32) (i32.const 0x004C0000))

  ;; Mini-stream container (root entry's data read via FAT)
  (global $mini_stream_ptr (mut i32) (i32.const 0))
  (global $mini_stream_len (mut i32) (i32.const 0))

  ;; Parsed stream pointers (set during parse)
  (global $worddoc_ptr (mut i32) (i32.const 0))
  (global $worddoc_len (mut i32) (i32.const 0))
  (global $table_ptr   (mut i32) (i32.const 0))
  (global $table_len   (mut i32) (i32.const 0))
  (global $data_ptr    (mut i32) (i32.const 0))
  (global $data_len    (mut i32) (i32.const 0))
  (global $text_ptr    (mut i32) (i32.const 0))
  (global $text_len    (mut i32) (i32.const 0))

  ;; Layout results
  (global $page_count  (mut i32) (i32.const 0))

  ;; Error code
  (global $error_code  (mut i32) (i32.const 0))

  ;; ── Constants ───────────────────────────────────────────────

  ;; Fixed region pointers
  (global $FAT_BASE      i32 (i32.const 0x00000000))
  (global $MINIFAT_BASE  i32 (i32.const 0x00040000))
  (global $DIR_BASE      i32 (i32.const 0x00080000))
  (global $FIB_BASE      i32 (i32.const 0x00090000))
  (global $CHP_BASE      i32 (i32.const 0x00094000))
  (global $PAP_BASE      i32 (i32.const 0x00194000))
  (global $SEP_BASE      i32 (i32.const 0x00294000))
  (global $STYLE_BASE    i32 (i32.const 0x002A4000))
  (global $LAYOUT_BASE   i32 (i32.const 0x002B4000))
  (global $FONT_TABLE    i32 (i32.const 0x002A8000))  ;; 256 entries × 8 bytes = 2KB within STYLE region

  ;; Font table count
  (global $font_count (mut i32) (i32.const 0))

  ;; Error codes
  (global $ERR_NONE          i32 (i32.const 0))
  (global $ERR_TOO_SMALL     i32 (i32.const 1))
  (global $ERR_BAD_MAGIC     i32 (i32.const 2))
  (global $ERR_BAD_SECTOR_SZ i32 (i32.const 3))
  (global $ERR_NO_WORDDOC    i32 (i32.const 4))
  (global $ERR_BAD_FIB       i32 (i32.const 5))
  (global $ERR_NO_TABLE      i32 (i32.const 6))
  (global $ERR_BAD_CLX       i32 (i32.const 7))

  ;; CFBF constants
  (global $SECTOR_SIZE     i32 (i32.const 512))
  (global $DIR_ENTRY_SIZE  i32 (i32.const 128))
  (global $MINI_SECTOR_SIZE i32 (i32.const 64))
  (global $MINI_STREAM_CUTOFF i32 (i32.const 0x1000))  ;; 4096

  ;; ── Data segments ───────────────────────────────────────────
  ;; OLE2 magic: D0 CF 11 E0 A1 B1 1A E1
  (data (i32.const 0x004B4000) "\D0\CF\11\E0\A1\B1\1A\E1")

  ;; Stream names in UTF-16LE for comparison (stored at 0x004B4010+)
  ;; "WordDocument" (24 bytes)
  (data (i32.const 0x004B4010) "W\00o\00r\00d\00D\00o\00c\00u\00m\00e\00n\00t\00")
  ;; "0Table" (12 bytes)
  (data (i32.const 0x004B4030) "0\00T\00a\00b\00l\00e\00")
  ;; "1Table" (12 bytes)
  (data (i32.const 0x004B4040) "1\00T\00a\00b\00l\00e\00")
  ;; "Data" (8 bytes)
  (data (i32.const 0x004B4050) "D\00a\00t\00a\00")

  ;; Windows-1252 to Unicode lookup for 0x80-0x9F (32 entries × 4 bytes)
  ;; Stored at 0x004B4060
  (data (i32.const 0x004B4060)
    "\AC\20\00\00"  ;; 0x80 → U+20AC €
    "\81\00\00\00"  ;; 0x81 → U+0081 (undefined, keep as-is)
    "\1A\20\00\00"  ;; 0x82 → U+201A ‚
    "\92\01\00\00"  ;; 0x83 → U+0192 ƒ
    "\1E\20\00\00"  ;; 0x84 → U+201E „
    "\26\20\00\00"  ;; 0x85 → U+2026 …
    "\20\20\00\00"  ;; 0x86 → U+2020 †
    "\21\20\00\00"  ;; 0x87 → U+2021 ‡
    "\C6\02\00\00"  ;; 0x88 → U+02C6 ˆ
    "\30\20\00\00"  ;; 0x89 → U+2030 ‰
    "\60\01\00\00"  ;; 0x8A → U+0160 Š
    "\39\20\00\00"  ;; 0x8B → U+2039 ‹
    "\52\01\00\00"  ;; 0x8C → U+0152 Œ
    "\8D\00\00\00"  ;; 0x8D → U+008D
    "\7D\01\00\00"  ;; 0x8E → U+017D Ž
    "\8F\00\00\00"  ;; 0x8F → U+008F
    "\90\00\00\00"  ;; 0x90 → U+0090
    "\18\20\00\00"  ;; 0x91 → U+2018 '
    "\19\20\00\00"  ;; 0x92 → U+2019 '
    "\1C\20\00\00"  ;; 0x93 → U+201C "
    "\1D\20\00\00"  ;; 0x94 → U+201D "
    "\22\20\00\00"  ;; 0x95 → U+2022 •
    "\13\20\00\00"  ;; 0x96 → U+2013 –
    "\14\20\00\00"  ;; 0x97 → U+2014 —
    "\DC\02\00\00"  ;; 0x98 → U+02DC ˜
    "\22\21\00\00"  ;; 0x99 → U+2122 ™
    "\61\01\00\00"  ;; 0x9A → U+0161 š
    "\3A\20\00\00"  ;; 0x9B → U+203A ›
    "\53\01\00\00"  ;; 0x9C → U+0153 œ
    "\9D\00\00\00"  ;; 0x9D → U+009D
    "\7E\01\00\00"  ;; 0x9E → U+017E ž
    "\78\01\00\00"  ;; 0x9F → U+0178 Ÿ
  )

  ;; ── Memory helpers ──────────────────────────────────────────

  (func $read_u8 (param $off i32) (result i32)
    (i32.load8_u (local.get $off))
  )

  (func $read_u16_le (param $off i32) (result i32)
    (i32.load16_u (local.get $off))
  )

  (func $read_u32_le (param $off i32) (result i32)
    (i32.load (local.get $off))
  )

  ;; ── Arena allocator ─────────────────────────────────────────

  (func $arena_alloc (param $size i32) (result i32)
    (local $ptr i32)
    ;; Align size to 4 bytes
    (local.set $size
      (i32.and
        (i32.add (local.get $size) (i32.const 3))
        (i32.const -4)
      )
    )
    (local.set $ptr (global.get $arena_ptr))
    (global.set $arena_ptr
      (i32.add (global.get $arena_ptr) (local.get $size))
    )
    (local.get $ptr)
  )

  (func $arena_reset
    (global.set $arena_ptr (global.get $arena_base))
  )

  ;; ── Byte comparison helper ──────────────────────────────────

  (func $memcmp (param $a i32) (param $b i32) (param $len i32) (result i32)
    ;; Returns 0 if equal, 1 if not
    (local $i i32)
    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (local.get $len)))
        (if (i32.ne
              (i32.load8_u (i32.add (local.get $a) (local.get $i)))
              (i32.load8_u (i32.add (local.get $b) (local.get $i)))
            )
          (then (return (i32.const 1)))
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )
    (i32.const 0)
  )

  ;; ── Memory copy ─────────────────────────────────────────────

  (func $memcpy (param $dst i32) (param $src i32) (param $len i32)
    (local $i i32)
    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (local.get $len)))
        (i32.store8
          (i32.add (local.get $dst) (local.get $i))
          (i32.load8_u (i32.add (local.get $src) (local.get $i)))
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )
  )

  ;; ── CFBF Header Parser ──────────────────────────────────────

  ;; CFBF header fields (stored as globals after parsing)
  (global $cfbf_sector_size   (mut i32) (i32.const 512))
  (global $cfbf_mini_sector_size (mut i32) (i32.const 64))
  (global $cfbf_fat_sectors   (mut i32) (i32.const 0))  ;; count
  (global $cfbf_dir_start     (mut i32) (i32.const 0))  ;; first directory sector
  (global $cfbf_minifat_start (mut i32) (i32.const 0))  ;; first mini-FAT sector
  (global $cfbf_minifat_count (mut i32) (i32.const 0))
  (global $cfbf_difat_start   (mut i32) (i32.const 0))  ;; first DIFAT sector
  (global $cfbf_difat_count   (mut i32) (i32.const 0))
  (global $cfbf_mini_stream_start (mut i32) (i32.const 0)) ;; root entry start sector
  (global $cfbf_mini_stream_size  (mut i32) (i32.const 0))

  ;; Number of directory entries found
  (global $dir_count (mut i32) (i32.const 0))

  ;; Stream locations found in directory
  (global $stream_worddoc_start (mut i32) (i32.const -1))
  (global $stream_worddoc_size  (mut i32) (i32.const 0))
  (global $stream_0table_start  (mut i32) (i32.const -1))
  (global $stream_0table_size   (mut i32) (i32.const 0))
  (global $stream_1table_start  (mut i32) (i32.const -1))
  (global $stream_1table_size   (mut i32) (i32.const 0))
  (global $stream_data_start    (mut i32) (i32.const -1))
  (global $stream_data_size     (mut i32) (i32.const 0))

  ;; Convert sector number to byte offset in file
  (func $sector_offset (param $sector i32) (result i32)
    ;; offset = (sector + 1) * 512  (sector 0 starts at byte 512, after header)
    (i32.mul
      (i32.add (local.get $sector) (i32.const 1))
      (i32.const 512)
    )
  )

  ;; Get next sector from FAT
  (func $fat_next (param $sector i32) (result i32)
    ;; FAT is array of i32 at FAT_BASE
    (i32.load
      (i32.add
        (global.get $FAT_BASE)
        (i32.mul (local.get $sector) (i32.const 4))
      )
    )
  )

  ;; Get next mini-sector from mini-FAT
  (func $minifat_next (param $sector i32) (result i32)
    (i32.load
      (i32.add
        (global.get $MINIFAT_BASE)
        (i32.mul (local.get $sector) (i32.const 4))
      )
    )
  )

  ;; Read a stream via FAT chain into dest. Returns bytes copied.
  (func $read_stream (param $start_sector i32) (param $size i32) (param $dest i32) (result i32)
    (local $sector i32)
    (local $remaining i32)
    (local $chunk i32)
    (local $offset i32)

    (local.set $sector (local.get $start_sector))
    (local.set $remaining (local.get $size))
    (local.set $offset (i32.const 0))

    (block $done
      (loop $loop
        ;; Stop if no more data or end-of-chain
        (br_if $done (i32.le_s (local.get $remaining) (i32.const 0)))
        (br_if $done (i32.ge_u (local.get $sector) (i32.const 0xFFFFFFFE)))

        ;; Chunk = min(512, remaining)
        (local.set $chunk (i32.const 512))
        (if (i32.lt_u (local.get $remaining) (i32.const 512))
          (then (local.set $chunk (local.get $remaining)))
        )

        ;; Copy from file
        (call $memcpy
          (i32.add (local.get $dest) (local.get $offset))
          (i32.add (global.get $input_ptr) (call $sector_offset (local.get $sector)))
          (local.get $chunk)
        )

        (local.set $offset (i32.add (local.get $offset) (local.get $chunk)))
        (local.set $remaining (i32.sub (local.get $remaining) (local.get $chunk)))
        (local.set $sector (call $fat_next (local.get $sector)))
        (br $loop)
      )
    )

    (local.get $offset)
  )

  ;; Read a mini-stream via mini-FAT chain into dest.
  ;; Mini-stream sectors are 64 bytes each, stored inside the root entry's stream
  ;; which has been read into $mini_stream_ptr via FAT.
  (func $read_mini_stream (param $start_sector i32) (param $size i32) (param $dest i32) (result i32)
    (local $sector i32)
    (local $remaining i32)
    (local $chunk i32)
    (local $offset i32)
    (local $mini_offset i32)

    (local.set $sector (local.get $start_sector))
    (local.set $remaining (local.get $size))
    (local.set $offset (i32.const 0))

    (block $done
      (loop $loop
        (br_if $done (i32.le_s (local.get $remaining) (i32.const 0)))
        (br_if $done (i32.ge_u (local.get $sector) (i32.const 0xFFFFFFFE)))

        ;; Byte offset within mini-stream container
        (local.set $mini_offset (i32.mul (local.get $sector) (i32.const 64)))

        ;; Chunk = min(64, remaining)
        (local.set $chunk (i32.const 64))
        (if (i32.lt_u (local.get $remaining) (i32.const 64))
          (then (local.set $chunk (local.get $remaining)))
        )

        ;; Copy from mini-stream container
        (call $memcpy
          (i32.add (local.get $dest) (local.get $offset))
          (i32.add (global.get $mini_stream_ptr) (local.get $mini_offset))
          (local.get $chunk)
        )

        (local.set $offset (i32.add (local.get $offset) (local.get $chunk)))
        (local.set $remaining (i32.sub (local.get $remaining) (local.get $chunk)))
        (local.set $sector (call $minifat_next (local.get $sector)))
        (br $loop)
      )
    )

    (local.get $offset)
  )

  ;; Read a stream, choosing FAT or mini-FAT based on size
  (func $read_stream_auto (param $start_sector i32) (param $size i32) (param $dest i32) (result i32)
    (if (result i32) (i32.lt_u (local.get $size) (global.get $MINI_STREAM_CUTOFF))
      (then
        (call $read_mini_stream (local.get $start_sector) (local.get $size) (local.get $dest))
      )
      (else
        (call $read_stream (local.get $start_sector) (local.get $size) (local.get $dest))
      )
    )
  )

  ;; Parse CFBF header
  (func $parse_cfbf_header (result i32)
    (local $base i32)
    (local $i i32)
    (local $sector i32)
    (local $fat_idx i32)

    (local.set $base (global.get $input_ptr))

    ;; Check minimum file size (at least 1 sector + header = 1024 bytes)
    (if (i32.lt_u (global.get $input_len) (i32.const 1024))
      (then (return (global.get $ERR_TOO_SMALL)))
    )

    ;; Verify magic bytes
    (if (call $memcmp
          (local.get $base)
          (i32.const 0x004B4000)  ;; our stored magic
          (i32.const 8)
        )
      (then (return (global.get $ERR_BAD_MAGIC)))
    )

    ;; Read sector size power (offset 0x1E) — should be 9 (=512) for .doc
    ;; Sector size = 1 << power
    (if (i32.ne (call $read_u16_le (i32.add (local.get $base) (i32.const 0x1E))) (i32.const 9))
      (then (return (global.get $ERR_BAD_SECTOR_SZ)))
    )

    ;; Read header fields
    (global.set $cfbf_fat_sectors
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x2C))))
    (global.set $cfbf_dir_start
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x30))))
    (global.set $cfbf_minifat_start
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x3C))))
    (global.set $cfbf_minifat_count
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x40))))
    (global.set $cfbf_difat_start
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x44))))
    (global.set $cfbf_difat_count
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x48))))

    ;; Build FAT from DIFAT entries in header (offsets 0x4C..0x01FC, 109 entries)
    (local.set $fat_idx (i32.const 0))
    (local.set $i (i32.const 0))
    (block $fat_done
      (loop $fat_loop
        (br_if $fat_done (i32.ge_u (local.get $i) (i32.const 109)))
        (local.set $sector
          (call $read_u32_le
            (i32.add (local.get $base)
              (i32.add (i32.const 0x4C) (i32.mul (local.get $i) (i32.const 4)))
            )
          )
        )
        ;; ENDOFCHAIN = 0xFFFFFFFE, FREE = 0xFFFFFFFF — skip
        (if (i32.lt_u (local.get $sector) (i32.const 0xFFFFFFFE))
          (then
            ;; Copy this FAT sector (512 bytes = 128 i32 entries) into FAT_BASE
            (call $memcpy
              (i32.add (global.get $FAT_BASE) (i32.mul (local.get $fat_idx) (i32.const 512)))
              (i32.add (global.get $input_ptr) (call $sector_offset (local.get $sector)))
              (i32.const 512)
            )
            (local.set $fat_idx (i32.add (local.get $fat_idx) (i32.const 1)))
          )
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $fat_loop)
      )
    )

    ;; TODO: Follow DIFAT chain for files with >109 FAT sectors

    ;; Parse mini-FAT chain
    (if (i32.and
          (i32.gt_u (global.get $cfbf_minifat_count) (i32.const 0))
          (i32.lt_u (global.get $cfbf_minifat_start) (i32.const 0xFFFFFFFE))
        )
      (then
        (local.set $sector (global.get $cfbf_minifat_start))
        (local.set $i (i32.const 0))
        (block $mf_done
          (loop $mf_loop
            (br_if $mf_done (i32.ge_u (local.get $sector) (i32.const 0xFFFFFFFE)))
            (call $memcpy
              (i32.add (global.get $MINIFAT_BASE) (i32.mul (local.get $i) (i32.const 512)))
              (i32.add (global.get $input_ptr) (call $sector_offset (local.get $sector)))
              (i32.const 512)
            )
            (local.set $i (i32.add (local.get $i) (i32.const 1)))
            (local.set $sector (call $fat_next (local.get $sector)))
            (br $mf_loop)
          )
        )
      )
    )

    (global.get $ERR_NONE)
  )

  ;; ── Directory Parser ────────────────────────────────────────

  ;; Compare directory entry name (UTF-16LE) at entry_ptr with reference at ref_ptr
  ;; name_size is byte count of the name in the dir entry (including null terminator)
  ;; ref_len is byte count of the reference name (without null terminator)
  (func $dir_name_eq (param $entry_ptr i32) (param $ref_ptr i32) (param $ref_len i32) (result i32)
    (local $name_size i32)
    ;; Name size is at offset 64 in the entry (u16, in bytes including null terminator)
    (local.set $name_size
      (call $read_u16_le (i32.add (local.get $entry_ptr) (i32.const 64)))
    )
    ;; name_size includes 2-byte null terminator, so compare name_size-2 bytes
    (if (i32.ne (i32.sub (local.get $name_size) (i32.const 2)) (local.get $ref_len))
      (then (return (i32.const 0)))
    )
    ;; Compare the name bytes
    (i32.eqz (call $memcmp (local.get $entry_ptr) (local.get $ref_ptr) (local.get $ref_len)))
  )

  (func $parse_directory (result i32)
    (local $sector i32)
    (local $offset i32)
    (local $entry_ptr i32)
    (local $entry_type i32)
    (local $dir_offset i32)
    (local $entries_in_sector i32)
    (local $j i32)
    (local $start_sect i32)
    (local $stream_size i32)

    ;; 512 / 128 = 4 entries per sector
    (local.set $entries_in_sector (i32.const 4))
    (local.set $sector (global.get $cfbf_dir_start))
    (global.set $dir_count (i32.const 0))

    (block $done
      (loop $sector_loop
        (br_if $done (i32.ge_u (local.get $sector) (i32.const 0xFFFFFFFE)))

        (local.set $offset
          (i32.add (global.get $input_ptr) (call $sector_offset (local.get $sector)))
        )

        ;; Process each entry in this sector
        (local.set $j (i32.const 0))
        (block $entry_done
          (loop $entry_loop
            (br_if $entry_done (i32.ge_u (local.get $j) (local.get $entries_in_sector)))

            (local.set $entry_ptr
              (i32.add (local.get $offset)
                (i32.mul (local.get $j) (i32.const 128))
              )
            )

            ;; Entry type at offset 66
            (local.set $entry_type
              (call $read_u8 (i32.add (local.get $entry_ptr) (i32.const 66)))
            )

            ;; Skip empty entries (type 0)
            (if (i32.ne (local.get $entry_type) (i32.const 0))
              (then
                ;; Copy entry to DIR_BASE for reference
                (call $memcpy
                  (i32.add (global.get $DIR_BASE)
                    (i32.mul (global.get $dir_count) (i32.const 128))
                  )
                  (local.get $entry_ptr)
                  (i32.const 128)
                )

                ;; Get start sector (offset 116) and size (offset 120)
                (local.set $start_sect
                  (call $read_u32_le (i32.add (local.get $entry_ptr) (i32.const 116)))
                )
                (local.set $stream_size
                  (call $read_u32_le (i32.add (local.get $entry_ptr) (i32.const 120)))
                )

                ;; Root entry (type 5) — save mini-stream info
                (if (i32.eq (local.get $entry_type) (i32.const 5))
                  (then
                    (global.set $cfbf_mini_stream_start (local.get $start_sect))
                    (global.set $cfbf_mini_stream_size (local.get $stream_size))
                  )
                )

                ;; Check for known stream names
                ;; "WordDocument" = 24 bytes UTF-16LE
                (if (call $dir_name_eq (local.get $entry_ptr) (i32.const 0x004B4010) (i32.const 24))
                  (then
                    (global.set $stream_worddoc_start (local.get $start_sect))
                    (global.set $stream_worddoc_size (local.get $stream_size))
                  )
                )
                ;; "0Table" = 12 bytes
                (if (call $dir_name_eq (local.get $entry_ptr) (i32.const 0x004B4030) (i32.const 12))
                  (then
                    (global.set $stream_0table_start (local.get $start_sect))
                    (global.set $stream_0table_size (local.get $stream_size))
                  )
                )
                ;; "1Table" = 12 bytes
                (if (call $dir_name_eq (local.get $entry_ptr) (i32.const 0x004B4040) (i32.const 12))
                  (then
                    (global.set $stream_1table_start (local.get $start_sect))
                    (global.set $stream_1table_size (local.get $stream_size))
                  )
                )
                ;; "Data" = 8 bytes
                (if (call $dir_name_eq (local.get $entry_ptr) (i32.const 0x004B4050) (i32.const 8))
                  (then
                    (global.set $stream_data_start (local.get $start_sect))
                    (global.set $stream_data_size (local.get $stream_size))
                  )
                )

                (global.set $dir_count (i32.add (global.get $dir_count) (i32.const 1)))
              )
            )

            (local.set $j (i32.add (local.get $j) (i32.const 1)))
            (br $entry_loop)
          )
        )

        (local.set $sector (call $fat_next (local.get $sector)))
        (br $sector_loop)
      )
    )

    ;; Verify we found WordDocument
    (if (i32.eq (global.get $stream_worddoc_start) (i32.const -1))
      (then (return (global.get $ERR_NO_WORDDOC)))
    )

    (global.get $ERR_NONE)
  )

  ;; ── FIB Parser ──────────────────────────────────────────────

  ;; FIB fields stored at FIB_BASE as flat struct:
  ;; offset 0:  wIdent (u16)
  ;; offset 2:  nFib (u16)
  ;; offset 4:  fWhichTblStm (u8, 0 or 1)
  ;; offset 8:  fcClx (u32)
  ;; offset 12: lcbClx (u32)
  ;; offset 16: fcPlcfBteChpx (u32)
  ;; offset 20: lcbPlcfBteChpx (u32)
  ;; offset 24: fcPlcfBtePapx (u32)
  ;; offset 28: lcbPlcfBtePapx (u32)
  ;; offset 32: fcStshf (u32)
  ;; offset 36: lcbStshf (u32)
  ;; offset 40: fcPlcfSed (u32)
  ;; offset 44: lcbPlcfSed (u32)
  ;; offset 48: ccpText (u32) — character count of main text
  ;; offset 52: ccpFtn (u32)
  ;; offset 56: ccpHdd (u32)

  (func $parse_fib (result i32)
    (local $base i32)  ;; base of WordDocument stream in memory
    (local $wIdent i32)
    (local $nFib i32)
    (local $flags i32)
    (local $fWhichTblStm i32)
    (local $fcLcb_base i32) ;; base offset of fibRgFcLcbBlob

    (local.set $base (global.get $worddoc_ptr))

    ;; Validate wIdent
    (local.set $wIdent (call $read_u16_le (local.get $base)))
    (i32.store16 (global.get $FIB_BASE) (local.get $wIdent))
    (if (i32.ne (local.get $wIdent) (i32.const 0xA5EC))
      (then (return (global.get $ERR_BAD_FIB)))
    )

    ;; nFib at offset 2
    (local.set $nFib (call $read_u16_le (i32.add (local.get $base) (i32.const 2))))
    (i32.store16 (i32.add (global.get $FIB_BASE) (i32.const 2)) (local.get $nFib))

    ;; flags at offset 0x000A — bit 9 is fWhichTblStm
    (local.set $flags (call $read_u16_le (i32.add (local.get $base) (i32.const 0x0A))))
    (local.set $fWhichTblStm
      (i32.and (i32.shr_u (local.get $flags) (i32.const 9)) (i32.const 1))
    )
    (i32.store8 (i32.add (global.get $FIB_BASE) (i32.const 4)) (local.get $fWhichTblStm))

    ;; ccpText at FIB offset 0x004C
    (i32.store
      (i32.add (global.get $FIB_BASE) (i32.const 48))
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x004C)))
    )
    ;; ccpFtn at 0x0050
    (i32.store
      (i32.add (global.get $FIB_BASE) (i32.const 52))
      (call $read_u32_le (i32.add (local.get $base) (i32.const 0x0050)))
    )

    ;; fibRgFcLcbBlob starts at offset 0x0062 in the FIB for nFib >= 0x00C1 (Word 97+)
    ;; The fcLcb pairs are at known indices:
    ;; Each pair is 8 bytes (fc:u32, lcb:u32)
    ;; Index 0 = fcStshfOrig, 1 = fcStshf, etc. (see [MS-DOC] 2.5.10)
    ;;
    ;; For Word 97 (nFib=0x00C1), fibRgFcLcb97 starts at base+0x0062:
    ;; The actual layout depends on nFib version, but critical offsets:

    ;; For nFib=0x00C1 (Word 97):
    ;; fibRgFcLcb starts at 0x0062 in the FIB
    ;; Offsets within fibRgFcLcb (each is fc:u32 + lcb:u32 = 8 bytes):
    ;;   fcStshf      = index 1  → 0x0062 + 1*8 = 0x006A
    ;;   fcPlcfSed    = index 3  → 0x0062 + 3*8 = 0x007A
    ;;   fcPlcfBteChpx= index 10 → 0x0062 + 10*8 = 0x00B2
    ;;   fcPlcfBtePapx= index 11 → 0x0062 + 11*8 = 0x00BA
    ;;   fcClx        = index 17 → 0x0062 + 17*8 = 0x00CA

    ;; Actually, the FIB layout is more complex. Let me use the standard offsets
    ;; from [MS-DOC] §2.5.10 FibRgFcLcb97:
    ;; The blob starts at FIB offset 0x009A for nFib=0x00C1
    ;; But there's also fibRgW97 and fibRgLw97 before it.
    ;;
    ;; FIB structure:
    ;; 0x0000: fibBase (32 bytes)
    ;; 0x0020: csw (u16) = count of shorts in fibRgW
    ;; 0x0022: fibRgW97 (28 bytes = 14 u16)
    ;; 0x003E: cslw (u16) = count of longs in fibRgLw
    ;; 0x0040: fibRgLw97 (88 bytes = 22 u32)
    ;;   - fibRgLw97.ccpText at offset 0x004C (index 3 in the array starting at 0x0040+4=0x0044... wait)
    ;;   Actually fibRgLw97 starts at 0x0040, and ccpText is at index 0 relative to fibRgLw97
    ;;   Offset 0x004C = 0x0040 + 0x0C → so ccpText is the 3rd u32 in fibRgLw97? No...
    ;;   Let me recalculate:
    ;;   0x003E: cslw (u16)
    ;;   0x0040: fibRgLw97[0] = cbMac (u32)
    ;;   0x0044: fibRgLw97[1] = reserved
    ;;   0x0048: fibRgLw97[2] = ccpText... or 0x004C?
    ;;   From MS-DOC: ccpText is at fibRgLw97 offset 0x000C → absolute 0x0040+0x000C = 0x004C ✓
    ;;
    ;; 0x0098: cbRgFcLcb (u16)
    ;; 0x009A: fibRgFcLcbBlob starts here
    ;;
    ;; fibRgFcLcb97 field offsets (relative to 0x009A):
    ;;   fcStshf      = 0x0A (fc) / 0x0E (lcb)  → absolute 0x00A4 / 0x00A8
    ;;   fcPlcfSed    = 0x62 (fc) / 0x66 (lcb)
    ;;   fcPlcfBteChpx= 0xFA (fc) / 0xFE (lcb)
    ;;   fcPlcfBtePapx= 0x102 (fc)/ 0x106 (lcb)
    ;;   fcClx        = 0x11A (fc)/ 0x11E (lcb)
    ;;
    ;; Actually I should use the well-known absolute offsets from the FIB:

    (local.set $fcLcb_base (i32.add (local.get $base) (i32.const 0x009A)))

    ;; fibRgFcLcb97 field indices (each entry = 8 bytes: fc u32 + lcb u32)
    ;; Indices verified against actual Word 97 files:
    ;; idx  1 (0x08): fcStshf / lcbStshf
    ;; idx  6 (0x30): fcPlcfSed / lcbPlcfSed
    ;; idx 12 (0x60): fcPlcfBteChpx / lcbPlcfBteChpx
    ;; idx 13 (0x68): fcPlcfBtePapx / lcbPlcfBtePapx
    ;; idx 33 (0x108): fcClx / lcbClx

    ;; fcStshf / lcbStshf (index 1, offset 0x08)
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 32))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x08))))
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 36))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x0C))))

    ;; fcPlcfSed / lcbPlcfSed (index 6, offset 0x30)
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 40))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x30))))
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 44))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x34))))

    ;; fcPlcfBteChpx / lcbPlcfBteChpx (index 12, offset 0x60)
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 16))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x60))))
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 20))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x64))))

    ;; fcPlcfBtePapx / lcbPlcfBtePapx (index 13, offset 0x68)
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 24))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x68))))
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 28))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x6C))))

    ;; fcClx / lcbClx (index 33, offset 0x108)
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 8))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x0108))))
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 12))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x010C))))

    ;; fcSttbfFfn / lcbSttbfFfn (index 15, offset 0x78)
    ;; Store at FIB_BASE+48 / FIB_BASE+52
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 48))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x0078))))
    (i32.store (i32.add (global.get $FIB_BASE) (i32.const 52))
      (call $read_u32_le (i32.add (local.get $fcLcb_base) (i32.const 0x007C))))

    (global.get $ERR_NONE)
  )

  ;; ── Piece Table / CLX Parser ────────────────────────────────

  ;; Piece descriptor count
  (global $piece_count (mut i32) (i32.const 0))
  ;; Pointer to piece CPs array (n+1 i32s) in arena
  (global $piece_cps_ptr (mut i32) (i32.const 0))
  ;; Pointer to piece PCDs array (n * 8 bytes) in arena
  (global $piece_pcds_ptr (mut i32) (i32.const 0))

  (func $parse_clx (result i32)
    (local $clx_offset i32)  ;; offset within table stream
    (local $clx_end i32)
    (local $pos i32)
    (local $type i32)
    (local $size i32)
    (local $n i32)
    (local $cps_size i32)
    (local $pcds_size i32)

    ;; fcClx is offset into the Table stream
    (local.set $clx_offset
      (i32.load (i32.add (global.get $FIB_BASE) (i32.const 8)))
    )
    (local.set $clx_end
      (i32.add (local.get $clx_offset)
        (i32.load (i32.add (global.get $FIB_BASE) (i32.const 12)))
      )
    )

    (local.set $pos (i32.add (global.get $table_ptr) (local.get $clx_offset)))

    ;; Skip Prc records (type 0x01)
    (block $found_pcdt
      (loop $skip_prc
        (local.set $type (call $read_u8 (local.get $pos)))
        ;; Pcdt has type 0x02
        (br_if $found_pcdt (i32.eq (local.get $type) (i32.const 0x02)))
        ;; Prc has type 0x01
        (if (i32.ne (local.get $type) (i32.const 0x01))
          (then (return (global.get $ERR_BAD_CLX)))
        )
        ;; Skip: 1 byte type + 2 byte cbGrpprl + cbGrpprl bytes
        (local.set $size
          (call $read_u16_le (i32.add (local.get $pos) (i32.const 1)))
        )
        (local.set $pos
          (i32.add (local.get $pos) (i32.add (i32.const 3) (local.get $size)))
        )
        (br $skip_prc)
      )
    )

    ;; Now at Pcdt: type=0x02, then u32 lcb, then PlcPcd data
    (local.set $pos (i32.add (local.get $pos) (i32.const 1))) ;; skip type byte
    (local.set $size (call $read_u32_le (local.get $pos)))
    (local.set $pos (i32.add (local.get $pos) (i32.const 4))) ;; now at PlcPcd

    ;; PlcPcd: (n+1) CPs as i32, then n PCDs of 8 bytes each
    ;; size = (n+1)*4 + n*8 = 4 + n*12
    ;; n = (size - 4) / 12
    (local.set $n (i32.div_u (i32.sub (local.get $size) (i32.const 4)) (i32.const 12)))
    (global.set $piece_count (local.get $n))

    ;; Allocate and copy CPs
    (local.set $cps_size (i32.mul (i32.add (local.get $n) (i32.const 1)) (i32.const 4)))
    (global.set $piece_cps_ptr (call $arena_alloc (local.get $cps_size)))
    (call $memcpy (global.get $piece_cps_ptr) (local.get $pos) (local.get $cps_size))

    ;; Allocate and copy PCDs
    (local.set $pcds_size (i32.mul (local.get $n) (i32.const 8)))
    (global.set $piece_pcds_ptr (call $arena_alloc (local.get $pcds_size)))
    (call $memcpy
      (global.get $piece_pcds_ptr)
      (i32.add (local.get $pos) (local.get $cps_size))
      (local.get $pcds_size)
    )

    (global.get $ERR_NONE)
  )

  ;; ── Text Extraction ─────────────────────────────────────────

  (func $extract_text (result i32)
    (local $i i32)
    (local $cp_start i32)
    (local $cp_end i32)
    (local $char_count i32)
    (local $pcd_ptr i32)
    (local $fc_raw i32)
    (local $fc i32)
    (local $compressed i32)
    (local $src i32)
    (local $dst i32)
    (local $j i32)
    (local $byte_val i32)
    (local $codepoint i32)

    ;; Allocate text buffer from arena
    ;; Max text size: sum of all piece char counts * 2 (UTF-16)
    ;; We'll compute the total CPs first
    (local.set $cp_end
      (call $read_u32_le
        (i32.add (global.get $piece_cps_ptr)
          (i32.mul (global.get $piece_count) (i32.const 4))
        )
      )
    )
    ;; Allocate cp_end * 2 bytes for UTF-16
    (global.set $text_ptr (call $arena_alloc (i32.mul (local.get $cp_end) (i32.const 2))))
    (local.set $dst (global.get $text_ptr))

    (local.set $i (i32.const 0))
    (block $done
      (loop $piece_loop
        (br_if $done (i32.ge_u (local.get $i) (global.get $piece_count)))

        ;; CP range for this piece
        (local.set $cp_start
          (call $read_u32_le
            (i32.add (global.get $piece_cps_ptr)
              (i32.mul (local.get $i) (i32.const 4))
            )
          )
        )
        (local.set $cp_end
          (call $read_u32_le
            (i32.add (global.get $piece_cps_ptr)
              (i32.mul (i32.add (local.get $i) (i32.const 1)) (i32.const 4))
            )
          )
        )
        (local.set $char_count (i32.sub (local.get $cp_end) (local.get $cp_start)))

        ;; PCD for this piece (8 bytes: 2 flags + 4 fc + 2 prm)
        (local.set $pcd_ptr
          (i32.add (global.get $piece_pcds_ptr)
            (i32.mul (local.get $i) (i32.const 8))
          )
        )

        ;; fc field at PCD offset 2 (4 bytes)
        (local.set $fc_raw
          (call $read_u32_le (i32.add (local.get $pcd_ptr) (i32.const 2)))
        )

        ;; Bit 30: fCompressed
        (local.set $compressed
          (i32.and (i32.shr_u (local.get $fc_raw) (i32.const 30)) (i32.const 1))
        )

        ;; fc = fc_raw & 0x3FFFFFFF
        (local.set $fc
          (i32.and (local.get $fc_raw) (i32.const 0x3FFFFFFF))
        )

        (if (local.get $compressed)
          (then
            ;; Compressed: Windows-1252, real offset = fc / 2
            (local.set $src
              (i32.add (global.get $worddoc_ptr) (i32.div_u (local.get $fc) (i32.const 2)))
            )

            ;; Convert each byte to UTF-16LE
            (local.set $j (i32.const 0))
            (block $cp_done
              (loop $cp_loop
                (br_if $cp_done (i32.ge_u (local.get $j) (local.get $char_count)))
                (local.set $byte_val
                  (call $read_u8 (i32.add (local.get $src) (local.get $j)))
                )

                ;; Map 0x80-0x9F via Windows-1252 lookup table
                (if (i32.and
                      (i32.ge_u (local.get $byte_val) (i32.const 0x80))
                      (i32.le_u (local.get $byte_val) (i32.const 0x9F))
                    )
                  (then
                    (local.set $codepoint
                      (i32.load
                        (i32.add (i32.const 0x004B4060)
                          (i32.mul (i32.sub (local.get $byte_val) (i32.const 0x80)) (i32.const 4))
                        )
                      )
                    )
                  )
                  (else
                    (local.set $codepoint (local.get $byte_val))
                  )
                )

                ;; Write as UTF-16LE (assuming BMP only)
                (i32.store16 (local.get $dst) (local.get $codepoint))
                (local.set $dst (i32.add (local.get $dst) (i32.const 2)))

                (local.set $j (i32.add (local.get $j) (i32.const 1)))
                (br $cp_loop)
              )
            )
          )
          (else
            ;; Uncompressed: UTF-16LE at fc offset in WordDocument stream
            (local.set $src
              (i32.add (global.get $worddoc_ptr) (local.get $fc))
            )
            (call $memcpy
              (local.get $dst)
              (local.get $src)
              (i32.mul (local.get $char_count) (i32.const 2))
            )
            (local.set $dst
              (i32.add (local.get $dst) (i32.mul (local.get $char_count) (i32.const 2)))
            )
          )
        )

        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $piece_loop)
      )
    )

    ;; Set text length
    (global.set $text_len (i32.sub (local.get $dst) (global.get $text_ptr)))

    ;; Filter out field codes from extracted text
    (call $filter_field_codes)

    (global.get $ERR_NONE)
  )

  ;; ── Field code filter ───────────────────────────────────────
  ;; Remove field instruction text (between 0x13 and 0x14) from TEXT_BUFFER
  ;; Keep field result text (between 0x14 and 0x15)
  ;; Remove the 0x13, 0x14, 0x15 marker chars themselves
  (func $filter_field_codes
    (local $src i32)
    (local $dst i32)
    (local $end i32)
    (local $ch i32)
    (local $in_field i32)  ;; 1 = inside field instruction (skip)

    (local.set $src (global.get $text_ptr))
    (local.set $dst (global.get $text_ptr))
    (local.set $end (i32.add (global.get $text_ptr) (global.get $text_len)))

    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $src) (local.get $end)))

        (local.set $ch (call $read_u16_le (local.get $src)))

        ;; 0x13 = field begin — start skipping
        (if (i32.eq (local.get $ch) (i32.const 0x13))
          (then
            (local.set $in_field (i32.const 1))
            (local.set $src (i32.add (local.get $src) (i32.const 2)))
            (br $loop)
          )
        )
        ;; 0x14 = field separator — stop skipping (show result)
        (if (i32.eq (local.get $ch) (i32.const 0x14))
          (then
            (local.set $in_field (i32.const 0))
            (local.set $src (i32.add (local.get $src) (i32.const 2)))
            (br $loop)
          )
        )
        ;; 0x15 = field end — stop skipping
        (if (i32.eq (local.get $ch) (i32.const 0x15))
          (then
            (local.set $in_field (i32.const 0))
            (local.set $src (i32.add (local.get $src) (i32.const 2)))
            (br $loop)
          )
        )

        ;; Skip if inside field instruction
        (if (local.get $in_field)
          (then
            (local.set $src (i32.add (local.get $src) (i32.const 2)))
            (br $loop)
          )
        )

        ;; Copy character
        (i32.store16 (local.get $dst) (local.get $ch))
        (local.set $dst (i32.add (local.get $dst) (i32.const 2)))
        (local.set $src (i32.add (local.get $src) (i32.const 2)))
        (br $loop)
      )
    )

    ;; Update text length
    (global.set $text_len (i32.sub (local.get $dst) (global.get $text_ptr)))
  )

  ;; ── Main parse orchestrator ─────────────────────────────────

  (func $parse (result i32)
    (local $err i32)
    (local $fWhichTblStm i32)
    (local $tbl_start i32)
    (local $tbl_size i32)

    ;; Reset arena
    (call $arena_reset)

    ;; Phase 1: CFBF header + FAT
    (local.set $err (call $parse_cfbf_header))
    (if (local.get $err) (then
      (global.set $error_code (local.get $err))
      (return (local.get $err))
    ))

    ;; Phase 2: Directory
    (local.set $err (call $parse_directory))
    (if (local.get $err) (then
      (global.set $error_code (local.get $err))
      (return (local.get $err))
    ))

    ;; Phase 2.5: Read mini-stream container (root entry's data via FAT)
    ;; This is needed to read streams < 4096 bytes
    (if (i32.and
          (i32.gt_u (global.get $cfbf_mini_stream_size) (i32.const 0))
          (i32.lt_u (global.get $cfbf_mini_stream_start) (i32.const 0xFFFFFFFE))
        )
      (then
        (global.set $mini_stream_ptr (call $arena_alloc (global.get $cfbf_mini_stream_size)))
        (global.set $mini_stream_len (global.get $cfbf_mini_stream_size))
        (drop (call $read_stream
          (global.get $cfbf_mini_stream_start)
          (global.get $cfbf_mini_stream_size)
          (global.get $mini_stream_ptr)
        ))
      )
    )

    ;; Phase 3: Extract WordDocument stream
    (global.set $worddoc_ptr (call $arena_alloc (global.get $stream_worddoc_size)))
    (global.set $worddoc_len (global.get $stream_worddoc_size))
    (drop (call $read_stream_auto
      (global.get $stream_worddoc_start)
      (global.get $stream_worddoc_size)
      (global.get $worddoc_ptr)
    ))

    ;; Phase 4: Parse FIB
    (local.set $err (call $parse_fib))
    (if (local.get $err) (then
      (global.set $error_code (local.get $err))
      (return (local.get $err))
    ))

    ;; Phase 5: Extract Table stream (0Table or 1Table based on FIB flag)
    (local.set $fWhichTblStm
      (i32.load8_u (i32.add (global.get $FIB_BASE) (i32.const 4)))
    )
    (if (local.get $fWhichTblStm)
      (then
        ;; Use 1Table
        (if (i32.eq (global.get $stream_1table_start) (i32.const -1))
          (then
            (global.set $error_code (global.get $ERR_NO_TABLE))
            (return (global.get $ERR_NO_TABLE))
          )
        )
        (local.set $tbl_start (global.get $stream_1table_start))
        (local.set $tbl_size (global.get $stream_1table_size))
      )
      (else
        ;; Use 0Table
        (if (i32.eq (global.get $stream_0table_start) (i32.const -1))
          (then
            (global.set $error_code (global.get $ERR_NO_TABLE))
            (return (global.get $ERR_NO_TABLE))
          )
        )
        (local.set $tbl_start (global.get $stream_0table_start))
        (local.set $tbl_size (global.get $stream_0table_size))
      )
    )
    (global.set $table_ptr (call $arena_alloc (local.get $tbl_size)))
    (global.set $table_len (local.get $tbl_size))
    (drop (call $read_stream_auto
      (local.get $tbl_start)
      (local.get $tbl_size)
      (global.get $table_ptr)
    ))

    ;; Phase 6: Extract Data stream (optional)
    (if (i32.ne (global.get $stream_data_start) (i32.const -1))
      (then
        (global.set $data_ptr (call $arena_alloc (global.get $stream_data_size)))
        (global.set $data_len (global.get $stream_data_size))
        (drop (call $read_stream_auto
          (global.get $stream_data_start)
          (global.get $stream_data_size)
          (global.get $data_ptr)
        ))
      )
    )

    ;; Phase 7: Parse piece table (CLX)
    (local.set $err (call $parse_clx))
    (if (local.get $err) (then
      (global.set $error_code (local.get $err))
      (return (local.get $err))
    ))

    ;; Phase 8: Extract text
    (local.set $err (call $extract_text))
    (if (local.get $err) (then
      (global.set $error_code (local.get $err))
      (return (local.get $err))
    ))

    ;; Phase 8.5: Parse STSH (stylesheet — style defaults)
    (call $parse_stsh)

    ;; Phase 8.6: Parse font table (SttbfFfn)
    (call $parse_font_table)

    ;; Phase 9: Parse PAP (paragraph properties) — before CHP so style istd is available
    (call $parse_pap)

    ;; Phase 10: Parse CHP (character properties) — uses PAP istd for style defaults
    (call $parse_chp)

    ;; Phase 11: Parse SEP (section properties — page size, margins)
    (call $parse_sep)

    ;; Phase 12: Scan for embedded images in Data stream
    (call $scan_images)

    ;; Phase 13: Layout
    (call $do_layout)

    (global.set $error_code (global.get $ERR_NONE))
    (global.get $ERR_NONE)
  )

  ;; ── CHP (Character Properties) Parser ───────────────────────

  ;; Style defaults from STSH (Normal style) — fallback if no style table
  (global $style_default_font_size  (mut i32) (i32.const 24))  ;; default 12pt
  (global $style_default_flags     (mut i32) (i32.const 0))
  (global $style_default_font_index (mut i32) (i32.const 0))
  (global $style_count             (mut i32) (i32.const 0))

  ;; ── STSH (Stylesheet) Parser ──────────────────────────────────
  ;; Style table at STYLE_BASE: up to 256 styles, 20 bytes each:
  ;; [0..3]   flags (i32): bit0=bold, bit1=italic, etc. 0xFFFFFFFF = not set
  ;; [4..7]   font_size (i32, half-points). 0 = not set
  ;; [8..11]  color (i32). 0xFFFFFFFF = not set
  ;; [12..15] istdBase (i32): base style index. 0xFFF = no base
  ;; [16..19] font_index (i32): index into FONT_TABLE. 0xFFFF = not set

  ;; Read style entry field
  (func $style_ptr (param $istd i32) (result i32)
    (i32.add (global.get $STYLE_BASE) (i32.mul (local.get $istd) (i32.const 20)))
  )

  ;; Get resolved font_size for a style, walking the base chain (max 10 deep)
  (func $style_get_font_size (param $istd i32) (result i32)
    (local $ptr i32)
    (local $val i32)
    (local $base i32)
    (local $depth i32)

    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $istd) (global.get $style_count)))
        (br_if $done (i32.ge_u (local.get $depth) (i32.const 10)))

        (local.set $ptr (call $style_ptr (local.get $istd)))
        (local.set $val (i32.load (i32.add (local.get $ptr) (i32.const 4))))
        (if (local.get $val) (then (return (local.get $val))))

        ;; Walk to base style
        (local.set $base (i32.load (i32.add (local.get $ptr) (i32.const 12))))
        (br_if $done (i32.eq (local.get $base) (i32.const 0x0FFF)))
        (local.set $istd (local.get $base))
        (local.set $depth (i32.add (local.get $depth) (i32.const 1)))
        (br $loop)
      )
    )
    (i32.const 24) ;; fallback 12pt
  )

  ;; Get resolved flags for a style — return this style's flags directly
  ;; (flags are always explicitly stored, 0 = no bold/italic)
  (func $style_get_flags (param $istd i32) (result i32)
    (if (result i32) (i32.lt_u (local.get $istd) (global.get $style_count))
      (then (i32.load (call $style_ptr (local.get $istd))))
      (else (i32.const 0))
    )
  )

  ;; Get resolved font_index for a style, walking the base chain
  (func $style_get_font_index (param $istd i32) (result i32)
    (local $ptr i32)
    (local $val i32)
    (local $base i32)
    (local $depth i32)

    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $istd) (global.get $style_count)))
        (br_if $done (i32.ge_u (local.get $depth) (i32.const 10)))

        (local.set $ptr (call $style_ptr (local.get $istd)))
        (local.set $val (i32.load (i32.add (local.get $ptr) (i32.const 16))))
        (if (i32.ne (local.get $val) (i32.const 0xFFFF)) (then (return (local.get $val))))

        ;; Walk to base style
        (local.set $base (i32.load (i32.add (local.get $ptr) (i32.const 12))))
        (br_if $done (i32.eq (local.get $base) (i32.const 0x0FFF)))
        (local.set $istd (local.get $base))
        (local.set $depth (i32.add (local.get $depth) (i32.const 1)))
        (br $loop)
      )
    )
    (i32.const 0) ;; fallback: font index 0
  )

  ;; Parse one STD entry's chpx UPX and store results at STYLE_BASE + istd*20
  ;; $upx_ptr points to start of the UPX region, $sgc is style group code
  (func $parse_style_chpx (param $istd i32) (param $upx_ptr i32) (param $end i32) (param $sgc i32)
                           (param $std_base i32)
    (local $cbUpx i32)
    (local $pos i32)
    (local $flags i32)
    (local $font_size i32)
    (local $color i32)
    (local $font_index i32)
    (local $sptr i32)

    (local.set $pos (local.get $upx_ptr))

    ;; For paragraph style (sgc=1): skip papx UPX first, then chpx UPX
    (if (i32.eq (local.get $sgc) (i32.const 1))
      (then
        (if (i32.lt_u (i32.add (local.get $pos) (i32.const 2)) (local.get $end))
          (then
            (local.set $cbUpx (call $read_u16_le (local.get $pos)))
            (local.set $pos (i32.add (local.get $pos) (i32.add (i32.const 2) (local.get $cbUpx))))
            (if (i32.and (i32.sub (local.get $pos) (local.get $std_base)) (i32.const 1))
              (then (local.set $pos (i32.add (local.get $pos) (i32.const 1))))
            )
          )
        )
      )
    )

    ;; Read chpx UPX
    (if (i32.ge_u (i32.add (local.get $pos) (i32.const 2)) (local.get $end))
      (then (return))
    )
    (local.set $cbUpx (call $read_u16_le (local.get $pos)))
    (local.set $pos (i32.add (local.get $pos) (i32.const 2)))

    (if (i32.eqz (local.get $cbUpx)) (then (return)))
    (if (i32.gt_u (i32.add (local.get $pos) (local.get $cbUpx)) (local.get $end))
      (then (return))
    )

    ;; Parse sprms — start with defaults
    (local.set $flags (i32.const 0))
    (local.set $font_size (i32.const 0)) ;; 0 = not set (inherit from base)
    (local.set $color (i32.const 0))

    (call $parse_chp_sprms
      (local.get $pos) (local.get $cbUpx)
      (local.get $flags) (local.get $font_size) (local.get $color) (i32.const 0xFFFF)
    )
    ;; Returns (flags, font_size, color, font_index) — pop in reverse order
    (local.set $font_index)
    (local.set $color)
    (local.set $font_size)
    (local.set $flags)

    ;; Store in style table
    (local.set $sptr (call $style_ptr (local.get $istd)))
    (i32.store (local.get $sptr) (local.get $flags))
    (i32.store (i32.add (local.get $sptr) (i32.const 4)) (local.get $font_size))
    (i32.store (i32.add (local.get $sptr) (i32.const 8)) (local.get $color))
    (i32.store (i32.add (local.get $sptr) (i32.const 16)) (local.get $font_index))
  )

  ;; ── Font Table (SttbfFfn) Parser ──────────────────────────────
  ;; Parses font names from the SttbfFfn structure in the table stream.
  ;; Each entry stored at FONT_TABLE + i*8: name_ptr(4) + name_len(4)
  (func $parse_font_table
    (local $fc i32)
    (local $lcb i32)
    (local $pos i32)
    (local $end i32)
    (local $cData i32)
    (local $i i32)
    (local $cbFfn i32)
    (local $entry_end i32)
    (local $name_start i32)
    (local $name_len i32)
    (local $dst i32)
    (local $j i32)

    (local.set $fc (i32.load (i32.add (global.get $FIB_BASE) (i32.const 48))))
    (local.set $lcb (i32.load (i32.add (global.get $FIB_BASE) (i32.const 52))))

    (if (i32.eqz (local.get $lcb)) (then (return)))
    (if (i32.gt_u (local.get $lcb) (i32.const 0x00100000)) (then (return)))
    (if (i32.eq (local.get $fc) (i32.const 0xFFFFFFFF)) (then (return)))

    (local.set $pos (i32.add (global.get $table_ptr) (local.get $fc)))
    (local.set $end (i32.add (local.get $pos) (local.get $lcb)))

    ;; SttbfFfn header: optional 0xFFFF (extended), cData (u16), cbExtra (u16)
    (if (i32.eq (call $read_u16_le (local.get $pos)) (i32.const 0xFFFF))
      (then
        ;; Extended format (UTF-16 strings)
        (local.set $pos (i32.add (local.get $pos) (i32.const 2)))
      )
    )
    (local.set $cData (call $read_u16_le (local.get $pos)))
    (local.set $pos (i32.add (local.get $pos) (i32.const 4)))  ;; skip cData + cbExtra

    (if (i32.gt_u (local.get $cData) (i32.const 256))
      (then (local.set $cData (i32.const 256)))
    )

    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (local.get $cData)))
        (br_if $done (i32.ge_u (i32.add (local.get $pos) (i32.const 1)) (local.get $end)))

        ;; Each STTB data item: cbFfn (1 byte) + cbFfn bytes of FFN data
        (local.set $cbFfn (call $read_u8 (local.get $pos)))
        (local.set $entry_end (i32.add (local.get $pos) (i32.add (i32.const 1) (local.get $cbFfn))))
        (local.set $pos (i32.add (local.get $pos) (i32.const 1)))

        ;; FFN structure within the cbFfn bytes:
        ;; Fixed fields: prq/ff(1) + wWeight(2) + chs(1) + ixchSzAlt(1) = 5 bytes min
        ;; Optional: panose(10) at offset 5, FONTSIGNATURE(24) at offset 15
        ;; Name offset depends on total size: 39 if cbFfn >= 40, else 15 if >= 16, else 5
        (if (i32.gt_u (local.get $cbFfn) (i32.const 5))
          (then
            (local.set $name_start
              (i32.add (local.get $pos)
                (if (result i32) (i32.ge_u (local.get $cbFfn) (i32.const 40))
                  (then (i32.const 39))
                  (else (if (result i32) (i32.ge_u (local.get $cbFfn) (i32.const 16))
                    (then (i32.const 15))
                    (else (i32.const 5))
                  ))
                )
              )
            )
            ;; Find name length: scan for null terminator (0x0000) in UTF-16
            (local.set $name_len (i32.const 0))
            (block $name_done
              (loop $name_loop
                (br_if $name_done (i32.ge_u (i32.add (local.get $name_start) (i32.add (local.get $name_len) (i32.const 2))) (local.get $entry_end)))
                (br_if $name_done (i32.eqz (call $read_u16_le (i32.add (local.get $name_start) (local.get $name_len)))))
                (local.set $name_len (i32.add (local.get $name_len) (i32.const 2)))
                (br $name_loop)
              )
            )

            ;; Copy name to arena
            (if (i32.gt_u (local.get $name_len) (i32.const 0))
              (then
                (local.set $dst (call $arena_alloc (local.get $name_len)))
                (call $memcpy (local.get $dst) (local.get $name_start) (local.get $name_len))
                ;; Store ptr+len at FONT_TABLE + i*8
                (i32.store (i32.add (global.get $FONT_TABLE) (i32.mul (local.get $i) (i32.const 8)))
                  (local.get $dst))
                (i32.store (i32.add (i32.add (global.get $FONT_TABLE) (i32.mul (local.get $i) (i32.const 8))) (i32.const 4))
                  (local.get $name_len))
              )
            )
          )
        )

        (local.set $pos (local.get $entry_end))
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )

    (global.set $font_count (local.get $i))
  )

  (func $parse_stsh
    (local $fc i32)
    (local $lcb i32)
    (local $stsh_ptr i32)
    (local $cbStshi i32)
    (local $cbSTDBaseInFile i32)
    (local $pos i32)
    (local $end i32)
    (local $istd i32)
    (local $std_base i32)
    (local $cbStd i32)
    (local $w0 i32)
    (local $w1 i32)
    (local $w2 i32)
    (local $sgc i32)
    (local $istdBase i32)
    (local $cupx i32)
    (local $name_len i32)
    (local $sptr i32)

    (local.set $fc (i32.load (i32.add (global.get $FIB_BASE) (i32.const 32))))
    (local.set $lcb (i32.load (i32.add (global.get $FIB_BASE) (i32.const 36))))

    (if (i32.eqz (local.get $lcb)) (then (return)))
    ;; Guard against 0xFFFFFFFF or out-of-bounds
    (if (i32.gt_u (local.get $lcb) (i32.const 0x00100000)) (then (return)))
    (if (i32.eq (local.get $fc) (i32.const 0xFFFFFFFF)) (then (return)))
    (if (i32.ge_u (local.get $fc) (global.get $table_len)) (then (return)))

    (local.set $stsh_ptr (i32.add (global.get $table_ptr) (local.get $fc)))
    (local.set $end (i32.add (local.get $stsh_ptr) (local.get $lcb)))

    (local.set $cbStshi (call $read_u16_le (local.get $stsh_ptr)))

    ;; Read cbSTDBaseInFile from Stshi header (offset 2 within Stshi = stsh_ptr + 2 + 2)
    ;; Stshi fields: cstd(2), cbSTDBaseInFile(2), ...
    (local.set $cbSTDBaseInFile (i32.const 10)) ;; fallback
    (if (i32.ge_u (local.get $cbStshi) (i32.const 4))
      (then
        (local.set $cbSTDBaseInFile
          (call $read_u16_le (i32.add (local.get $stsh_ptr) (i32.const 4)))
        )
        ;; Sanity check
        (if (i32.lt_u (local.get $cbSTDBaseInFile) (i32.const 10))
          (then (local.set $cbSTDBaseInFile (i32.const 10)))
        )
      )
    )

    (local.set $pos (i32.add (local.get $stsh_ptr) (i32.add (i32.const 2) (local.get $cbStshi))))

    ;; Initialize all style slots to "not set"
    (local.set $istd (i32.const 0))
    (block $init_done
      (loop $init_loop
        (br_if $init_done (i32.ge_u (local.get $istd) (i32.const 256)))
        (local.set $sptr (call $style_ptr (local.get $istd)))
        (i32.store (local.get $sptr) (i32.const 0))             ;; flags = 0
        (i32.store (i32.add (local.get $sptr) (i32.const 4)) (i32.const 0))  ;; font_size = 0 (not set)
        (i32.store (i32.add (local.get $sptr) (i32.const 8)) (i32.const 0))  ;; color = 0
        (i32.store (i32.add (local.get $sptr) (i32.const 12)) (i32.const 0x0FFF)) ;; no base
        (i32.store (i32.add (local.get $sptr) (i32.const 16)) (i32.const 0xFFFF)) ;; font_index not set
        (local.set $istd (i32.add (local.get $istd) (i32.const 1)))
        (br $init_loop)
      )
    )

    ;; Parse each STD entry
    (local.set $istd (i32.const 0))
    (block $all_done
      (loop $style_loop
        (br_if $all_done (i32.ge_u (local.get $istd) (i32.const 256)))
        (br_if $all_done (i32.ge_u (i32.add (local.get $pos) (i32.const 2)) (local.get $end)))

        (local.set $cbStd (call $read_u16_le (local.get $pos)))
        (local.set $pos (i32.add (local.get $pos) (i32.const 2)))

        ;; Empty slot — skip
        (if (i32.eqz (local.get $cbStd))
          (then
            (local.set $istd (i32.add (local.get $istd) (i32.const 1)))
            (br $style_loop)
          )
        )

        (local.set $std_base (local.get $pos))

        ;; StdfBase: word0 bits 0-11 = sti
        ;; word1 bits 0-3 = sgc, bits 4-15 = istdBase
        ;; word2 bits 0-3 = cupx
        (local.set $w0 (call $read_u16_le (local.get $pos)))
        (local.set $w1 (call $read_u16_le (i32.add (local.get $pos) (i32.const 2))))
        (local.set $w2 (call $read_u16_le (i32.add (local.get $pos) (i32.const 4))))
        (local.set $sgc (i32.and (local.get $w1) (i32.const 0x0F)))
        (local.set $istdBase (i32.and (i32.shr_u (local.get $w1) (i32.const 4)) (i32.const 0x0FFF)))
        (local.set $cupx (i32.and (local.get $w2) (i32.const 0x0F)))

        ;; Store base style index
        (local.set $sptr (call $style_ptr (local.get $istd)))
        (i32.store (i32.add (local.get $sptr) (i32.const 12)) (local.get $istdBase))

        ;; Skip to UPX data: StdfBase(cbSTDBaseInFile) + xstzName
        (local.set $pos (i32.add (local.get $std_base) (local.get $cbSTDBaseInFile)))

        ;; Skip name: 2-byte char count + (count+1)*2 bytes + padding
        (if (i32.lt_u (i32.add (local.get $pos) (i32.const 2)) (i32.add (local.get $std_base) (local.get $cbStd)))
          (then
            (local.set $name_len (call $read_u16_le (local.get $pos)))
            ;; Sanity check name length
            (if (i32.lt_u (local.get $name_len) (i32.const 200))
              (then
                (local.set $pos (i32.add (local.get $pos)
                  (i32.add (i32.const 2)
                    (i32.mul (i32.add (local.get $name_len) (i32.const 1)) (i32.const 2))
                  )
                ))
                (if (i32.and (i32.sub (local.get $pos) (local.get $std_base)) (i32.const 1))
                  (then (local.set $pos (i32.add (local.get $pos) (i32.const 1))))
                )
              )
            )
          )
        )

        ;; Parse chpx UPX if this is paragraph (sgc=1) or character (sgc=2) style
        (if (i32.or
              (i32.eq (local.get $sgc) (i32.const 1))
              (i32.eq (local.get $sgc) (i32.const 2))
            )
          (then
            (call $parse_style_chpx
              (local.get $istd)
              (local.get $pos)
              (i32.add (local.get $std_base) (local.get $cbStd))
              (local.get $sgc)
              (local.get $std_base)
            )
          )
        )

        ;; Advance to next STD
        (local.set $pos (i32.add (local.get $std_base) (local.get $cbStd)))
        (local.set $istd (i32.add (local.get $istd) (i32.const 1)))
        (br $style_loop)
      )
    )

    (global.set $style_count (local.get $istd))

    ;; Set global defaults from style 0 (Normal) via inheritance
    (global.set $style_default_font_size (call $style_get_font_size (i32.const 0)))
    (global.set $style_default_flags (call $style_get_flags (i32.const 0)))
    (global.set $style_default_font_index (call $style_get_font_index (i32.const 0)))
  )

  ;; CHP run format at CHP_BASE: array of 28-byte records
  ;; [0..3]  cp_start (i32)
  ;; [4..7]  cp_end (i32)
  ;; [8..11] flags: bit0=bold, bit1=italic, bit2=underline, bit3=strike
  ;; [12..15] font_size (half-points, default 24 = 12pt)
  ;; [16..19] color (0x00RRGGBB, default 0)
  ;; [20..23] font_index (index into FONT_TABLE)
  ;; [24..27] reserved
  (global $chp_run_count (mut i32) (i32.const 0))

  ;; Get operand size for a sprm opcode
  ;; Bits 13-15 of opcode encode the type:
  ;; 0=toggle(1), 1=byte(1), 2=word(2), 3=dword(4), 4=variable, 5=variable, 6=variable, 7=triple(3)
  (func $sprm_size (param $opcode i32) (result i32)
    (local $spra i32)
    (local.set $spra (i32.and (i32.shr_u (local.get $opcode) (i32.const 13)) (i32.const 7)))
    (if (i32.eqz (local.get $spra)) (then (return (i32.const 1))))           ;; toggle
    (if (i32.eq (local.get $spra) (i32.const 1)) (then (return (i32.const 1)))) ;; byte
    (if (i32.eq (local.get $spra) (i32.const 2)) (then (return (i32.const 2)))) ;; word
    (if (i32.eq (local.get $spra) (i32.const 3)) (then (return (i32.const 4)))) ;; long
    (if (i32.eq (local.get $spra) (i32.const 7)) (then (return (i32.const 3)))) ;; triple
    ;; 4,5,6 = variable length, first byte is size
    (i32.const -1)  ;; signal: read next byte for size
  )

  ;; Parse a grpprl (array of sprms) and extract character formatting
  ;; Returns flags in lower 16 bits, font_size in bits 16-31
  (func $parse_chp_sprms (param $ptr i32) (param $len i32) (param $flags i32) (param $font_size i32) (param $color i32) (param $font_index i32)
        (result i32 i32 i32 i32)
    (local $pos i32)
    (local $end i32)
    (local $opcode i32)
    (local $operand_size i32)
    (local $val i32)

    (local.set $pos (local.get $ptr))
    (local.set $end (i32.add (local.get $ptr) (local.get $len)))

    (block $done
      (loop $loop
        ;; Need at least 2 bytes for opcode
        (br_if $done (i32.ge_u (i32.add (local.get $pos) (i32.const 2)) (local.get $end)))

        (local.set $opcode (call $read_u16_le (local.get $pos)))
        (local.set $pos (i32.add (local.get $pos) (i32.const 2)))

        (local.set $operand_size (call $sprm_size (local.get $opcode)))

        ;; Handle variable-length sprms
        (if (i32.eq (local.get $operand_size) (i32.const -1))
          (then
            (local.set $operand_size (call $read_u8 (local.get $pos)))
            (local.set $pos (i32.add (local.get $pos) (i32.const 1)))
          )
        )

        ;; sprmCFBold = 0x0835
        (if (i32.eq (local.get $opcode) (i32.const 0x0835))
          (then
            (if (call $read_u8 (local.get $pos))
              (then (local.set $flags (i32.or (local.get $flags) (i32.const 1))))
              (else (local.set $flags (i32.and (local.get $flags) (i32.const 0xFFFFFFFE))))
            )
          )
        )

        ;; sprmCFItalic = 0x0836
        (if (i32.eq (local.get $opcode) (i32.const 0x0836))
          (then
            (if (call $read_u8 (local.get $pos))
              (then (local.set $flags (i32.or (local.get $flags) (i32.const 2))))
              (else (local.set $flags (i32.and (local.get $flags) (i32.const 0xFFFFFFFD))))
            )
          )
        )

        ;; sprmCFUl (underline) = 0x0838  (actually 0x2A33 for kul, but 0x0838 is simple)
        (if (i32.eq (local.get $opcode) (i32.const 0x0838))
          (then
            (if (call $read_u8 (local.get $pos))
              (then (local.set $flags (i32.or (local.get $flags) (i32.const 4))))
              (else (local.set $flags (i32.and (local.get $flags) (i32.const 0xFFFFFFFB))))
            )
          )
        )

        ;; sprmCFStrike = 0x0837
        (if (i32.eq (local.get $opcode) (i32.const 0x0837))
          (then
            (if (call $read_u8 (local.get $pos))
              (then (local.set $flags (i32.or (local.get $flags) (i32.const 8))))
              (else (local.set $flags (i32.and (local.get $flags) (i32.const 0xFFFFFFF7))))
            )
          )
        )

        ;; sprmCHps (font size) = 0x4A43
        (if (i32.eq (local.get $opcode) (i32.const 0x4A43))
          (then
            (local.set $font_size (call $read_u16_le (local.get $pos)))
          )
        )

        ;; Font index sprms (2-byte operand = font table index)
        ;; sprmCRgFtc0 = 0x4A4F (Word 2000+, ASCII/Latin)
        ;; sprmCFtcDefault = 0x4A3D (Word 97)
        (if (i32.or
              (i32.eq (local.get $opcode) (i32.const 0x4A4F))
              (i32.eq (local.get $opcode) (i32.const 0x4A3D))
            )
          (then
            (local.set $font_index (call $read_u16_le (local.get $pos)))
          )
        )

        ;; sprmCCv (24-bit RGB color) = 0x6870 — 4-byte operand (COLORREF: 0x00BBGGRR)
        (if (i32.eq (local.get $opcode) (i32.const 0x6870))
          (then
            (local.set $val (call $read_u32_le (local.get $pos)))
            ;; COLORREF is 0x00BBGGRR, we need 0x00RRGGBB
            (local.set $color
              (i32.or
                (i32.or
                  (i32.shl (i32.and (local.get $val) (i32.const 0xFF)) (i32.const 16))    ;; R
                  (i32.and (local.get $val) (i32.const 0xFF00))                             ;; G stays
                )
                (i32.shr_u (i32.and (local.get $val) (i32.const 0xFF0000)) (i32.const 16)) ;; B
              )
            )
          )
        )

        ;; sprmCIco (color index) = 0x2A42
        (if (i32.eq (local.get $opcode) (i32.const 0x2A42))
          (then
            (local.set $val (call $read_u8 (local.get $pos)))
            ;; Map ico index to RGB (simplified — standard 16 colors)
            ;; 0=auto(black), 1=black, 2=blue, 3=cyan, 4=green,
            ;; 5=magenta, 6=red, 7=yellow, 8=white, 9=dk blue,
            ;; 10=dk cyan, 11=dk green, 12=dk magenta, 13=dk red,
            ;; 14=dk yellow, 15=dk gray, 16=lt gray
            ;; For simplicity, treat 0,1 as black, others we'll map later
            (if (i32.le_u (local.get $val) (i32.const 1))
              (then (local.set $color (i32.const 0x000000)))
            )
            (if (i32.eq (local.get $val) (i32.const 2))
              (then (local.set $color (i32.const 0x0000FF)))
            )
            (if (i32.eq (local.get $val) (i32.const 3))
              (then (local.set $color (i32.const 0x00FFFF)))
            )
            (if (i32.eq (local.get $val) (i32.const 4))
              (then (local.set $color (i32.const 0x00FF00)))
            )
            (if (i32.eq (local.get $val) (i32.const 5))
              (then (local.set $color (i32.const 0xFF00FF)))
            )
            (if (i32.eq (local.get $val) (i32.const 6))
              (then (local.set $color (i32.const 0xFF0000)))
            )
            (if (i32.eq (local.get $val) (i32.const 7))
              (then (local.set $color (i32.const 0xFFFF00)))
            )
            (if (i32.eq (local.get $val) (i32.const 8))
              (then (local.set $color (i32.const 0xFFFFFF)))
            )
          )
        )

        ;; Advance past operand
        (local.set $pos (i32.add (local.get $pos) (local.get $operand_size)))
        (br $loop)
      )
    )

    (local.get $flags)
    (local.get $font_size)
    (local.get $color)
    (local.get $font_index)
  )

  ;; Write a CHP run record at CHP_BASE + index*28
  (func $write_chp_run (param $idx i32) (param $cp_start i32) (param $cp_end i32)
                        (param $flags i32) (param $font_size i32) (param $color i32) (param $font_index i32)
    (local $ptr i32)
    (local.set $ptr (i32.add (global.get $CHP_BASE) (i32.mul (local.get $idx) (i32.const 28))))
    (i32.store (local.get $ptr) (local.get $cp_start))
    (i32.store (i32.add (local.get $ptr) (i32.const 4)) (local.get $cp_end))
    (i32.store (i32.add (local.get $ptr) (i32.const 8)) (local.get $flags))
    (i32.store (i32.add (local.get $ptr) (i32.const 12)) (local.get $font_size))
    (i32.store (i32.add (local.get $ptr) (i32.const 16)) (local.get $color))
    (i32.store (i32.add (local.get $ptr) (i32.const 20)) (local.get $font_index))
  )

  (func $parse_chp
    (local $fc_plcfbtechpx i32)
    (local $lcb i32)
    (local $n i32)
    (local $plc_ptr i32)
    (local $i i32)
    (local $fc_start i32)
    (local $fc_end i32)
    (local $bte_pn i32)
    (local $fkp_ptr i32)
    (local $cfkp_crun i32)
    (local $j i32)
    (local $rgfc_j i32)
    (local $rgfc_j1 i32)
    (local $chpx_offset i32)
    (local $chpx_ptr i32)
    (local $chpx_size i32)
    (local $flags i32)
    (local $font_size i32)
    (local $color i32)
    (local $font_index i32)
    (local $cp_start i32)
    (local $cp_end i32)
    (local $istd i32)

    (local.set $fc_plcfbtechpx (i32.load (i32.add (global.get $FIB_BASE) (i32.const 16))))
    (local.set $lcb (i32.load (i32.add (global.get $FIB_BASE) (i32.const 20))))

    ;; If no CHP data, create one default run for all text
    (if (i32.eqz (local.get $lcb))
      (then
        (call $write_chp_run (i32.const 0) (i32.const 0)
          (i32.div_u (global.get $text_len) (i32.const 2))
          (global.get $style_default_flags) (global.get $style_default_font_size) (i32.const 0x000000) (global.get $style_default_font_index))
        (global.set $chp_run_count (i32.const 1))
        (return)
      )
    )

    ;; PlcBteChpx is in the table stream at fc_plcfbtechpx
    ;; Format: (n+1) FCs (u32) then n BTEs (4 bytes each: pn u32)
    ;; n = (lcb - 4) / 8
    (local.set $plc_ptr (i32.add (global.get $table_ptr) (local.get $fc_plcfbtechpx)))
    (local.set $n (i32.div_u (i32.sub (local.get $lcb) (i32.const 4)) (i32.const 8)))

    (global.set $chp_run_count (i32.const 0))

    ;; For each BTE, read the FKP page
    (local.set $i (i32.const 0))
    (block $bte_done
      (loop $bte_loop
        (br_if $bte_done (i32.ge_u (local.get $i) (local.get $n)))

        ;; BTE pn (page number in WordDocument stream, each page = 512 bytes)
        (local.set $bte_pn
          (call $read_u32_le
            (i32.add (local.get $plc_ptr)
              (i32.add
                (i32.mul (i32.add (local.get $n) (i32.const 1)) (i32.const 4))
                (i32.mul (local.get $i) (i32.const 4))
              )
            )
          )
        )

        ;; FKP is at page pn * 512 in the WordDocument stream
        (local.set $fkp_ptr
          (i32.add (global.get $worddoc_ptr) (i32.mul (local.get $bte_pn) (i32.const 512)))
        )

        ;; Last byte of FKP = crun (number of runs)
        (local.set $cfkp_crun
          (call $read_u8 (i32.add (local.get $fkp_ptr) (i32.const 511)))
        )

        ;; FKP layout:
        ;; [0 .. (crun)*4]        rgfc: (crun+1) FCs (u32)
        ;; [(crun+1)*4 .. ]       unused space and CHPX data
        ;; [511-crun .. 510]      rgb: crun byte offsets (offset*2 into FKP for CHPX)
        ;; [511]                  crun

        (local.set $j (i32.const 0))
        (block $run_done
          (loop $run_loop
            (br_if $run_done (i32.ge_u (local.get $j) (local.get $cfkp_crun)))

            ;; FC boundaries
            (local.set $rgfc_j
              (call $read_u32_le
                (i32.add (local.get $fkp_ptr) (i32.mul (local.get $j) (i32.const 4)))
              )
            )
            (local.set $rgfc_j1
              (call $read_u32_le
                (i32.add (local.get $fkp_ptr) (i32.mul (i32.add (local.get $j) (i32.const 1)) (i32.const 4)))
              )
            )

            ;; Convert FC to CP using piece table mapping
            (local.set $cp_start (call $fc_to_cp (local.get $rgfc_j)))
            (local.set $cp_end (call $fc_to_cp (local.get $rgfc_j1)))

            ;; rgb[j] is at FKP + (crun+1)*4 + j
            ;; Each byte * 2 = byte offset within FKP to find the CHPX

            (local.set $chpx_offset
              (i32.mul
                (call $read_u8
                  (i32.add (local.get $fkp_ptr)
                    (i32.add
                      (i32.mul (i32.add (local.get $cfkp_crun) (i32.const 1)) (i32.const 4))
                      (local.get $j)
                    )
                  )
                )
                (i32.const 2)
              )
            )

            ;; Default formatting from paragraph's style
            (local.set $istd (call $get_istd_at_cp (local.get $cp_start)))
            (if (i32.and
                  (i32.gt_u (global.get $style_count) (i32.const 0))
                  (i32.lt_u (local.get $istd) (global.get $style_count))
                )
              (then
                (local.set $flags (call $style_get_flags (local.get $istd)))
                (local.set $font_size (call $style_get_font_size (local.get $istd)))
                (local.set $font_index (call $style_get_font_index (local.get $istd)))
              )
              (else
                (local.set $flags (global.get $style_default_flags))
                (local.set $font_size (global.get $style_default_font_size))
                (local.set $font_index (global.get $style_default_font_index))
              )
            )
            (local.set $color (i32.const 0x000000))

            ;; If offset is 0, no CHPX (use defaults)
            (if (local.get $chpx_offset)
              (then
                ;; CHPX at fkp_ptr + chpx_offset
                ;; First byte = cb (size of grpprl)
                (local.set $chpx_ptr (i32.add (local.get $fkp_ptr) (local.get $chpx_offset)))
                (local.set $chpx_size (call $read_u8 (local.get $chpx_ptr)))

                ;; Parse sprms
                (call $parse_chp_sprms
                  (i32.add (local.get $chpx_ptr) (i32.const 1))
                  (local.get $chpx_size)
                  (local.get $flags)
                  (local.get $font_size)
                  (local.get $color)
                  (local.get $font_index)
                )
                (local.set $font_index)
                (local.set $color)
                (local.set $font_size)
                (local.set $flags)
              )
            )

            ;; Only write if valid CP range
            (if (i32.and
                  (i32.ge_s (local.get $cp_start) (i32.const 0))
                  (i32.gt_s (local.get $cp_end) (local.get $cp_start))
                )
              (then
                (call $write_chp_run
                  (global.get $chp_run_count)
                  (local.get $cp_start)
                  (local.get $cp_end)
                  (local.get $flags)
                  (local.get $font_size)
                  (local.get $color)
                  (local.get $font_index)
                )
                (global.set $chp_run_count (i32.add (global.get $chp_run_count) (i32.const 1)))
              )
            )

            (local.set $j (i32.add (local.get $j) (i32.const 1)))
            (br $run_loop)
          )
        )

        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $bte_loop)
      )
    )

    ;; If no runs found, create default
    (if (i32.eqz (global.get $chp_run_count))
      (then
        (call $write_chp_run (i32.const 0) (i32.const 0)
          (i32.div_u (global.get $text_len) (i32.const 2))
          (global.get $style_default_flags) (global.get $style_default_font_size) (i32.const 0x000000) (global.get $style_default_font_index))
        (global.set $chp_run_count (i32.const 1))
      )
    )
  )

  ;; ── FC to CP conversion ─────────────────────────────────────
  ;; Map a file character offset (FC) back to a character position (CP)
  ;; by searching the piece table

  (func $fc_to_cp (param $fc i32) (result i32)
    (local $i i32)
    (local $pcd_ptr i32)
    (local $pcd_fc_raw i32)
    (local $pcd_fc i32)
    (local $compressed i32)
    (local $cp_start i32)
    (local $cp_end i32)
    (local $char_count i32)
    (local $fc_offset i32)

    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (global.get $piece_count)))

        (local.set $cp_start
          (call $read_u32_le
            (i32.add (global.get $piece_cps_ptr) (i32.mul (local.get $i) (i32.const 4)))
          )
        )
        (local.set $cp_end
          (call $read_u32_le
            (i32.add (global.get $piece_cps_ptr) (i32.mul (i32.add (local.get $i) (i32.const 1)) (i32.const 4)))
          )
        )
        (local.set $char_count (i32.sub (local.get $cp_end) (local.get $cp_start)))

        (local.set $pcd_ptr
          (i32.add (global.get $piece_pcds_ptr) (i32.mul (local.get $i) (i32.const 8)))
        )
        (local.set $pcd_fc_raw (call $read_u32_le (i32.add (local.get $pcd_ptr) (i32.const 2))))
        (local.set $compressed (i32.and (i32.shr_u (local.get $pcd_fc_raw) (i32.const 30)) (i32.const 1)))
        (local.set $pcd_fc (i32.and (local.get $pcd_fc_raw) (i32.const 0x3FFFFFFF)))

        (if (local.get $compressed)
          (then
            ;; Compressed: FC range is [pcd_fc/2, pcd_fc/2 + char_count)
            ;; The incoming $fc should be in this range
            (local.set $fc_offset (i32.div_u (local.get $pcd_fc) (i32.const 2)))
            (if (i32.and
                  (i32.ge_u (local.get $fc) (local.get $fc_offset))
                  (i32.lt_u (local.get $fc) (i32.add (local.get $fc_offset) (local.get $char_count)))
                )
              (then
                (return (i32.add (local.get $cp_start) (i32.sub (local.get $fc) (local.get $fc_offset))))
              )
            )
          )
          (else
            ;; Uncompressed: FC range is [pcd_fc, pcd_fc + char_count*2)
            (if (i32.and
                  (i32.ge_u (local.get $fc) (local.get $pcd_fc))
                  (i32.lt_u (local.get $fc) (i32.add (local.get $pcd_fc) (i32.mul (local.get $char_count) (i32.const 2))))
                )
              (then
                (return (i32.add (local.get $cp_start)
                  (i32.div_u (i32.sub (local.get $fc) (local.get $pcd_fc)) (i32.const 2))
                ))
              )
            )
          )
        )

        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )

    ;; Not found — return -1
    (i32.const -1)
  )

  ;; ── PAP (Paragraph Properties) Parser ───────────────────────

  ;; PAP run format at PAP_BASE: array of 28-byte records
  ;; [0..3]   cp_start (i32)
  ;; [4..7]   cp_end (i32)
  ;; [8..11]  alignment (0=left, 1=center, 2=right, 3=justify)
  ;; [12..15] space_before (twips)
  ;; [16..19] space_after (twips)
  ;; [20..23] first_line_indent (twips, can be negative)
  ;; [24..27] reserved
  (global $pap_run_count (mut i32) (i32.const 0))

  (func $write_pap_run (param $idx i32) (param $cp_start i32) (param $cp_end i32)
                        (param $alignment i32) (param $space_before i32) (param $space_after i32)
                        (param $first_indent i32)
    (local $ptr i32)
    (local.set $ptr (i32.add (global.get $PAP_BASE) (i32.mul (local.get $idx) (i32.const 28))))
    (i32.store (local.get $ptr) (local.get $cp_start))
    (i32.store (i32.add (local.get $ptr) (i32.const 4)) (local.get $cp_end))
    (i32.store (i32.add (local.get $ptr) (i32.const 8)) (local.get $alignment))
    (i32.store (i32.add (local.get $ptr) (i32.const 12)) (local.get $space_before))
    (i32.store (i32.add (local.get $ptr) (i32.const 16)) (local.get $space_after))
    (i32.store (i32.add (local.get $ptr) (i32.const 20)) (local.get $first_indent))
  )

  ;; Parse PAP sprms, returns (alignment, space_before, space_after, first_indent)
  (func $parse_pap_sprms (param $ptr i32) (param $len i32)
        (param $align i32) (param $sb i32) (param $sa i32) (param $fi i32)
        (result i32 i32 i32 i32)
    (local $pos i32)
    (local $end i32)
    (local $opcode i32)
    (local $operand_size i32)

    (local.set $pos (local.get $ptr))
    (local.set $end (i32.add (local.get $ptr) (local.get $len)))

    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (i32.add (local.get $pos) (i32.const 2)) (local.get $end)))

        (local.set $opcode (call $read_u16_le (local.get $pos)))
        (local.set $pos (i32.add (local.get $pos) (i32.const 2)))

        (local.set $operand_size (call $sprm_size (local.get $opcode)))
        (if (i32.eq (local.get $operand_size) (i32.const -1))
          (then
            (local.set $operand_size (call $read_u8 (local.get $pos)))
            (local.set $pos (i32.add (local.get $pos) (i32.const 1)))
          )
        )

        ;; sprmPJc (justification) = 0x2403
        (if (i32.eq (local.get $opcode) (i32.const 0x2403))
          (then (local.set $align (call $read_u8 (local.get $pos))))
        )

        ;; sprmPDyaBefore (space before) = 0xA413
        (if (i32.eq (local.get $opcode) (i32.const 0xA413))
          (then (local.set $sb (call $read_u16_le (local.get $pos))))
        )

        ;; sprmPDyaAfter (space after) = 0xA414
        (if (i32.eq (local.get $opcode) (i32.const 0xA414))
          (then (local.set $sa (call $read_u16_le (local.get $pos))))
        )

        ;; sprmPDxaLeft1 (first line indent) = 0x8460
        (if (i32.eq (local.get $opcode) (i32.const 0x8460))
          (then (local.set $fi (i32.extend16_s (call $read_u16_le (local.get $pos)))))
        )

        (local.set $pos (i32.add (local.get $pos) (local.get $operand_size)))
        (br $loop)
      )
    )

    (local.get $align)
    (local.get $sb)
    (local.get $sa)
    (local.get $fi)
  )

  (func $parse_pap
    (local $fc_plcfbtepapx i32)
    (local $lcb i32)
    (local $n i32)
    (local $plc_ptr i32)
    (local $i i32)
    (local $bte_pn i32)
    (local $fkp_ptr i32)
    (local $cpara i32)
    (local $j i32)
    (local $rgfc_j i32)
    (local $rgfc_j1 i32)
    (local $papx_off_byte i32)
    (local $papx_ptr i32)
    (local $papx_word_count i32)
    (local $papx_istd i32)
    (local $grpprl_ptr i32)
    (local $grpprl_len i32)
    (local $alignment i32)
    (local $space_before i32)
    (local $space_after i32)
    (local $first_indent i32)
    (local $cp_start i32)
    (local $cp_end i32)

    (local.set $fc_plcfbtepapx (i32.load (i32.add (global.get $FIB_BASE) (i32.const 24))))
    (local.set $lcb (i32.load (i32.add (global.get $FIB_BASE) (i32.const 28))))

    ;; If no PAP data, one default paragraph for all text
    (if (i32.eqz (local.get $lcb))
      (then
        (call $write_pap_run (i32.const 0) (i32.const 0)
          (i32.div_u (global.get $text_len) (i32.const 2))
          (i32.const 0) (i32.const 0) (i32.const 0) (i32.const 0))
        (global.set $pap_run_count (i32.const 1))
        (return)
      )
    )

    ;; PlcBtePapx: (n+1) FCs + n BTEs (4 bytes each)
    (local.set $plc_ptr (i32.add (global.get $table_ptr) (local.get $fc_plcfbtepapx)))
    (local.set $n (i32.div_u (i32.sub (local.get $lcb) (i32.const 4)) (i32.const 8)))

    (global.set $pap_run_count (i32.const 0))

    (local.set $i (i32.const 0))
    (block $bte_done
      (loop $bte_loop
        (br_if $bte_done (i32.ge_u (local.get $i) (local.get $n)))

        (local.set $bte_pn
          (call $read_u32_le
            (i32.add (local.get $plc_ptr)
              (i32.add
                (i32.mul (i32.add (local.get $n) (i32.const 1)) (i32.const 4))
                (i32.mul (local.get $i) (i32.const 4))
              )
            )
          )
        )

        (local.set $fkp_ptr
          (i32.add (global.get $worddoc_ptr) (i32.mul (local.get $bte_pn) (i32.const 512)))
        )

        ;; cpara = last byte
        (local.set $cpara (call $read_u8 (i32.add (local.get $fkp_ptr) (i32.const 511))))

        ;; PapxFkp layout:
        ;; rgfc: (cpara+1) u32 at start
        ;; rgbx: cpara * 13-byte records starting at (cpara+1)*4
        ;; Each BX: 1 byte bOffset (word offset into FKP) + 12 bytes PHE (ignored)

        (local.set $j (i32.const 0))
        (block $run_done
          (loop $run_loop
            (br_if $run_done (i32.ge_u (local.get $j) (local.get $cpara)))

            (local.set $rgfc_j
              (call $read_u32_le
                (i32.add (local.get $fkp_ptr) (i32.mul (local.get $j) (i32.const 4)))
              )
            )
            (local.set $rgfc_j1
              (call $read_u32_le
                (i32.add (local.get $fkp_ptr) (i32.mul (i32.add (local.get $j) (i32.const 1)) (i32.const 4)))
              )
            )

            (local.set $cp_start (call $fc_to_cp (local.get $rgfc_j)))
            (local.set $cp_end (call $fc_to_cp (local.get $rgfc_j1)))

            ;; BX: bOffset at (cpara+1)*4 + j*13
            (local.set $papx_off_byte
              (call $read_u8
                (i32.add (local.get $fkp_ptr)
                  (i32.add
                    (i32.mul (i32.add (local.get $cpara) (i32.const 1)) (i32.const 4))
                    (i32.mul (local.get $j) (i32.const 13))
                  )
                )
              )
            )

            ;; Defaults
            (local.set $alignment (i32.const 0))
            (local.set $space_before (i32.const 0))
            (local.set $space_after (i32.const 0))
            (local.set $first_indent (i32.const 0))

            (if (local.get $papx_off_byte)
              (then
                ;; PAPX at fkp_ptr + papx_off_byte * 2
                (local.set $papx_ptr
                  (i32.add (local.get $fkp_ptr) (i32.mul (local.get $papx_off_byte) (i32.const 2)))
                )
                ;; First byte = cb (count of bytes). If 0, the second byte is cb2 and grpprl starts at +2
                (local.set $papx_word_count (call $read_u8 (local.get $papx_ptr)))
                (if (local.get $papx_word_count)
                  (then
                    ;; cb * 2 - 1 = total PAPX size after cb byte
                    ;; First 2 bytes after cb = istd
                    (local.set $papx_istd (call $read_u16_le (i32.add (local.get $papx_ptr) (i32.const 1))))
                    (local.set $grpprl_ptr (i32.add (local.get $papx_ptr) (i32.const 3)))
                    (local.set $grpprl_len
                      (i32.sub (i32.sub (i32.mul (local.get $papx_word_count) (i32.const 2)) (i32.const 1)) (i32.const 2))
                    )
                  )
                  (else
                    ;; cb=0: next byte is cb2
                    (local.set $papx_word_count (call $read_u8 (i32.add (local.get $papx_ptr) (i32.const 1))))
                    (local.set $papx_istd (call $read_u16_le (i32.add (local.get $papx_ptr) (i32.const 2))))
                    (local.set $grpprl_ptr (i32.add (local.get $papx_ptr) (i32.const 4)))
                    (local.set $grpprl_len
                      (i32.sub (i32.sub (i32.mul (local.get $papx_word_count) (i32.const 2)) (i32.const 1)) (i32.const 2))
                    )
                  )
                )

                (if (i32.gt_s (local.get $grpprl_len) (i32.const 0))
                  (then
                    (call $parse_pap_sprms
                      (local.get $grpprl_ptr) (local.get $grpprl_len)
                      (local.get $alignment) (local.get $space_before)
                      (local.get $space_after) (local.get $first_indent)
                    )
                    (local.set $first_indent)
                    (local.set $space_after)
                    (local.set $space_before)
                    (local.set $alignment)
                  )
                )
              )
            )

            (if (i32.and
                  (i32.ge_s (local.get $cp_start) (i32.const 0))
                  (i32.gt_s (local.get $cp_end) (local.get $cp_start))
                )
              (then
                (call $write_pap_run
                  (global.get $pap_run_count)
                  (local.get $cp_start) (local.get $cp_end)
                  (local.get $alignment) (local.get $space_before)
                  (local.get $space_after) (local.get $first_indent)
                )
                ;; Store istd at offset 24 of PAP run
                (i32.store
                  (i32.add
                    (i32.add (global.get $PAP_BASE) (i32.mul (global.get $pap_run_count) (i32.const 28)))
                    (i32.const 24)
                  )
                  (local.get $papx_istd)
                )
                (global.set $pap_run_count (i32.add (global.get $pap_run_count) (i32.const 1)))
              )
            )

            (local.set $j (i32.add (local.get $j) (i32.const 1)))
            (br $run_loop)
          )
        )

        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $bte_loop)
      )
    )

    ;; Default if none found
    (if (i32.eqz (global.get $pap_run_count))
      (then
        (call $write_pap_run (i32.const 0) (i32.const 0)
          (i32.div_u (global.get $text_len) (i32.const 2))
          (i32.const 0) (i32.const 0) (i32.const 0) (i32.const 0))
        (global.set $pap_run_count (i32.const 1))
      )
    )
  )

  ;; ── SEP (Section Properties) Parser ─────────────────────────
  ;; Parse PlcfSed to extract page dimensions and margins from the first section.
  ;; SEPX format: cb(u16) + grpprl(cb bytes)
  ;; Key sprms:
  ;;   sprmSXaPage (page width twips)   = 0xB01F
  ;;   sprmSYaPage (page height twips)  = 0xB020
  ;;   sprmSDxaLeft (left margin twips) = 0xB021
  ;;   sprmSDxaRight (right margin)     = 0xB022
  ;;   sprmSDyaTop (top margin twips)   = 0x9023
  ;;   sprmSDyaBottom (bottom margin)   = 0x9024

  (func $parse_sep
    (local $fc_plcfsed i32)
    (local $lcb i32)
    (local $plc_ptr i32)
    (local $n i32)
    (local $sed_ptr i32)
    (local $fc_sepx i32)
    (local $sepx_ptr i32)
    (local $cb i32)
    (local $pos i32)
    (local $end i32)
    (local $opcode i32)
    (local $operand_size i32)
    (local $page_w i32)
    (local $page_h i32)
    (local $margin_l i32)
    (local $margin_r i32)
    (local $margin_t i32)
    (local $margin_b i32)

    ;; Default page dimensions in twips (8.5x11 inches)
    (local.set $page_w (i32.const 12240))
    (local.set $page_h (i32.const 15840))
    ;; Default margins: 1 inch = 1440 twips (Word default is 1" left/right, 1" top/bottom)
    (local.set $margin_l (i32.const 1800))  ;; Word default is 1.25 inch = 1800 twips
    (local.set $margin_r (i32.const 1800))
    (local.set $margin_t (i32.const 1440))
    (local.set $margin_b (i32.const 1440))

    (local.set $fc_plcfsed (i32.load (i32.add (global.get $FIB_BASE) (i32.const 40))))
    (local.set $lcb (i32.load (i32.add (global.get $FIB_BASE) (i32.const 44))))

    ;; If no SED data, use defaults
    (if (i32.eqz (local.get $lcb)) (then
      (global.set $PAGE_WIDTH_PX (call $twips_to_px (local.get $page_w)))
      (global.set $PAGE_HEIGHT_PX (call $twips_to_px (local.get $page_h)))
      (global.set $MARGIN_LEFT_PX (call $twips_to_px (local.get $margin_l)))
      (global.set $MARGIN_RIGHT_PX (call $twips_to_px (local.get $margin_r)))
      (global.set $MARGIN_TOP_PX (call $twips_to_px (local.get $margin_t)))
      (global.set $MARGIN_BOTTOM_PX (call $twips_to_px (local.get $margin_b)))
      (global.set $MARGIN_PX (global.get $MARGIN_LEFT_PX))
      (return)
    ))

    ;; PlcfSed: (n+1) CPs (u32) + n SEDs (12 bytes each)
    ;; n = (lcb - 4) / 16
    (local.set $plc_ptr (i32.add (global.get $table_ptr) (local.get $fc_plcfsed)))
    (local.set $n (i32.div_u (i32.sub (local.get $lcb) (i32.const 4)) (i32.const 16)))

    ;; Only parse first section (index 0)
    (if (i32.gt_u (local.get $n) (i32.const 0))
      (then
        ;; SED[0] starts after (n+1) CPs
        (local.set $sed_ptr
          (i32.add (local.get $plc_ptr)
            (i32.mul (i32.add (local.get $n) (i32.const 1)) (i32.const 4))
          )
        )

        ;; SED format: fn(i16) + fcSepx(i32) + fnMpr(i16) + fcMpr(i32) = 12 bytes
        ;; fcSepx at offset 2
        (local.set $fc_sepx (call $read_u32_le (i32.add (local.get $sed_ptr) (i32.const 2))))

        ;; fcSepx of 0xFFFFFFFF means no SEPX
        (if (i32.lt_u (local.get $fc_sepx) (i32.const 0xFFFFFFFF))
          (then
            ;; SEPX is in the WordDocument stream
            (local.set $sepx_ptr (i32.add (global.get $worddoc_ptr) (local.get $fc_sepx)))
            ;; First 2 bytes = cb (grpprl byte count)
            (local.set $cb (call $read_u16_le (local.get $sepx_ptr)))

            (if (i32.gt_u (local.get $cb) (i32.const 0))
              (then
                (local.set $pos (i32.add (local.get $sepx_ptr) (i32.const 2)))
                (local.set $end (i32.add (local.get $pos) (local.get $cb)))

                (block $done
                  (loop $loop
                    (br_if $done (i32.ge_u (i32.add (local.get $pos) (i32.const 2)) (local.get $end)))

                    (local.set $opcode (call $read_u16_le (local.get $pos)))
                    (local.set $pos (i32.add (local.get $pos) (i32.const 2)))

                    (local.set $operand_size (call $sprm_size (local.get $opcode)))
                    (if (i32.eq (local.get $operand_size) (i32.const -1))
                      (then
                        (local.set $operand_size (call $read_u8 (local.get $pos)))
                        (local.set $pos (i32.add (local.get $pos) (i32.const 1)))
                      )
                    )

                    ;; sprmSXaPage = 0xB01F (page width)
                    (if (i32.eq (local.get $opcode) (i32.const 0xB01F))
                      (then (local.set $page_w (call $read_u16_le (local.get $pos))))
                    )
                    ;; sprmSYaPage = 0xB020 (page height)
                    (if (i32.eq (local.get $opcode) (i32.const 0xB020))
                      (then (local.set $page_h (call $read_u16_le (local.get $pos))))
                    )
                    ;; sprmSDxaLeft = 0xB021 (left margin)
                    (if (i32.eq (local.get $opcode) (i32.const 0xB021))
                      (then (local.set $margin_l (call $read_u16_le (local.get $pos))))
                    )
                    ;; sprmSDxaRight = 0xB022 (right margin)
                    (if (i32.eq (local.get $opcode) (i32.const 0xB022))
                      (then (local.set $margin_r (call $read_u16_le (local.get $pos))))
                    )
                    ;; sprmSDyaTop = 0x9023 (top margin, signed)
                    (if (i32.eq (local.get $opcode) (i32.const 0x9023))
                      (then (local.set $margin_t (i32.extend16_s (call $read_u16_le (local.get $pos)))))
                    )
                    ;; sprmSDyaBottom = 0x9024 (bottom margin, signed)
                    (if (i32.eq (local.get $opcode) (i32.const 0x9024))
                      (then (local.set $margin_b (i32.extend16_s (call $read_u16_le (local.get $pos)))))
                    )

                    (local.set $pos (i32.add (local.get $pos) (local.get $operand_size)))
                    (br $loop)
                  )
                )
              )
            )
          )
        )
      )
    )

    ;; Convert to pixels and set globals
    (global.set $PAGE_WIDTH_PX (call $twips_to_px (local.get $page_w)))
    (global.set $PAGE_HEIGHT_PX (call $twips_to_px (local.get $page_h)))
    (global.set $MARGIN_LEFT_PX (call $twips_to_px (local.get $margin_l)))
    (global.set $MARGIN_RIGHT_PX (call $twips_to_px (local.get $margin_r)))
    (global.set $MARGIN_TOP_PX (call $twips_to_px (local.get $margin_t)))
    (global.set $MARGIN_BOTTOM_PX (call $twips_to_px (local.get $margin_b)))
    ;; Keep MARGIN_PX as left margin for backward compat
    (global.set $MARGIN_PX (global.get $MARGIN_LEFT_PX))
  )

  ;; ── Image Table ──────────────────────────────────────────────
  ;; Scan Data stream for PICF entries containing BLIP image data.
  ;; Image table stored at SEP_BASE (reusing the 64KB SEP region):
  ;; Each entry: 16 bytes = data_ptr(u32), data_len(u32), width_px(f32), height_px(f32)
  ;; Max 256 images

  (global $IMG_TABLE i32 (i32.const 0x00294000))  ;; reuse SEP_BASE
  (global $img_count (mut i32) (i32.const 0))
  (global $img_cursor (mut i32) (i32.const 0))  ;; next image to render (for 0x01 chars)

  (func $scan_images
    (local $off i32)
    (local $lcb i32)
    (local $cb_header i32)
    (local $search_pos i32)
    (local $search_end i32)
    (local $rec_type i32)
    (local $rec_len i32)
    (local $blip_data_ptr i32)
    (local $blip_data_len i32)
    (local $img_w_px f32)
    (local $img_h_px f32)
    (local $entry_ptr i32)
    (local $png_w i32)
    (local $png_h i32)

    ;; No Data stream = no images
    (if (i32.eqz (global.get $data_len)) (then (return)))

    (global.set $img_count (i32.const 0))
    (local.set $off (i32.const 0))

    (block $done
      (loop $loop
        ;; Need at least 6 bytes for lcb + cbHeader
        (br_if $done (i32.ge_u (i32.add (local.get $off) (i32.const 6)) (global.get $data_len)))
        ;; Max 256 images
        (br_if $done (i32.ge_u (global.get $img_count) (i32.const 256)))

        (local.set $lcb (call $read_u32_le (i32.add (global.get $data_ptr) (local.get $off))))
        ;; lcb includes itself — valid range check
        (br_if $done (i32.lt_u (local.get $lcb) (i32.const 70)))
        (br_if $done (i32.gt_u (local.get $lcb) (global.get $data_len)))

        (local.set $cb_header (call $read_u16_le (i32.add (global.get $data_ptr) (i32.add (local.get $off) (i32.const 4)))))

        ;; Validate cbHeader = 0x44 (standard PICF)
        (if (i32.ne (local.get $cb_header) (i32.const 0x44))
          (then
            ;; Try skipping forward to find next PICF
            (local.set $off (i32.add (local.get $off) (local.get $lcb)))
            (br $loop)
          )
        )

        ;; Search for BLIP record (0xF01A-0xF01F) in SpContainer area
        (local.set $search_pos (i32.add (local.get $off) (local.get $cb_header)))
        (local.set $search_end (i32.add (local.get $off) (local.get $lcb)))
        (if (i32.gt_u (local.get $search_end) (global.get $data_len))
          (then (local.set $search_end (global.get $data_len)))
        )

        (local.set $blip_data_ptr (i32.const 0))
        (local.set $blip_data_len (i32.const 0))

        (block $blip_found
          (loop $search_loop
            (br_if $blip_found
              (i32.ge_u (i32.add (local.get $search_pos) (i32.const 8))
                (local.get $search_end)
              )
            )

            (local.set $rec_type
              (call $read_u16_le
                (i32.add (global.get $data_ptr)
                  (i32.add (local.get $search_pos) (i32.const 2))
                )
              )
            )

            ;; Check for BLIP types: 0xF01A-0xF01F
            (if (i32.and
                  (i32.ge_u (local.get $rec_type) (i32.const 0xF01A))
                  (i32.le_u (local.get $rec_type) (i32.const 0xF01F))
                )
              (then
                (local.set $rec_len
                  (call $read_u32_le
                    (i32.add (global.get $data_ptr)
                      (i32.add (local.get $search_pos) (i32.const 4))
                    )
                  )
                )

                ;; PNG (0xF01E) or JPEG (0xF01D): 8 header + 16 UID + 1 tag = 25 bytes before image
                (if (i32.or
                      (i32.eq (local.get $rec_type) (i32.const 0xF01E))
                      (i32.eq (local.get $rec_type) (i32.const 0xF01D))
                    )
                  (then
                    (local.set $blip_data_ptr
                      (i32.add (global.get $data_ptr)
                        (i32.add (local.get $search_pos) (i32.const 25))
                      )
                    )
                    (local.set $blip_data_len (i32.sub (local.get $rec_len) (i32.const 17)))
                  )
                )

                ;; For EMF/WMF (0xF01A/0xF01B): 8 + 16 UID + 34 metafile header = 58
                (if (i32.or
                      (i32.eq (local.get $rec_type) (i32.const 0xF01A))
                      (i32.eq (local.get $rec_type) (i32.const 0xF01B))
                    )
                  (then
                    (local.set $blip_data_ptr
                      (i32.add (global.get $data_ptr)
                        (i32.add (local.get $search_pos) (i32.const 58))
                      )
                    )
                    (local.set $blip_data_len (i32.sub (local.get $rec_len) (i32.const 50)))
                  )
                )

                (br $blip_found)
              )
            )

            (local.set $search_pos (i32.add (local.get $search_pos) (i32.const 1)))
            (br $search_loop)
          )
        )

        ;; If we found a BLIP, get dimensions and store in image table
        (if (i32.gt_u (local.get $blip_data_len) (i32.const 0))
          (then
            ;; Get image dimensions: for PNG, IHDR at offset 16 (width BE u32), 20 (height BE u32)
            ;; For JPEG, we'll pass 0x0 and let JS figure it out
            (local.set $img_w_px (f32.const 200.0))  ;; default
            (local.set $img_h_px (f32.const 150.0))

            ;; Check PNG signature (89 50 4E 47)
            (if (i32.and
                  (i32.eq (call $read_u8 (local.get $blip_data_ptr)) (i32.const 0x89))
                  (i32.eq (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 1))) (i32.const 0x50))
                )
              (then
                ;; PNG IHDR: width at +16 (big-endian u32), height at +20
                (local.set $png_w
                  (i32.or
                    (i32.or
                      (i32.shl (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 16))) (i32.const 24))
                      (i32.shl (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 17))) (i32.const 16))
                    )
                    (i32.or
                      (i32.shl (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 18))) (i32.const 8))
                      (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 19)))
                    )
                  )
                )
                (local.set $png_h
                  (i32.or
                    (i32.or
                      (i32.shl (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 20))) (i32.const 24))
                      (i32.shl (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 21))) (i32.const 16))
                    )
                    (i32.or
                      (i32.shl (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 22))) (i32.const 8))
                      (call $read_u8 (i32.add (local.get $blip_data_ptr) (i32.const 23)))
                    )
                  )
                )
                ;; Scale to fit content width (max ~624 px for standard margins)
                (local.set $img_w_px (f32.convert_i32_u (local.get $png_w)))
                (local.set $img_h_px (f32.convert_i32_u (local.get $png_h)))
                ;; Cap width to content area
                (if (f32.gt (local.get $img_w_px)
                      (f32.sub (global.get $PAGE_WIDTH_PX)
                        (f32.add (global.get $MARGIN_LEFT_PX) (global.get $MARGIN_RIGHT_PX))
                      )
                    )
                  (then
                    (local.set $img_h_px
                      (f32.mul (local.get $img_h_px)
                        (f32.div
                          (f32.sub (global.get $PAGE_WIDTH_PX)
                            (f32.add (global.get $MARGIN_LEFT_PX) (global.get $MARGIN_RIGHT_PX))
                          )
                          (local.get $img_w_px)
                        )
                      )
                    )
                    (local.set $img_w_px
                      (f32.sub (global.get $PAGE_WIDTH_PX)
                        (f32.add (global.get $MARGIN_LEFT_PX) (global.get $MARGIN_RIGHT_PX))
                      )
                    )
                  )
                )
              )
            )

            ;; Write image table entry
            (local.set $entry_ptr
              (i32.add (global.get $IMG_TABLE) (i32.mul (global.get $img_count) (i32.const 16)))
            )
            (i32.store (local.get $entry_ptr) (local.get $blip_data_ptr))
            (i32.store (i32.add (local.get $entry_ptr) (i32.const 4)) (local.get $blip_data_len))
            (f32.store (i32.add (local.get $entry_ptr) (i32.const 8)) (local.get $img_w_px))
            (f32.store (i32.add (local.get $entry_ptr) (i32.const 12)) (local.get $img_h_px))

            (global.set $img_count (i32.add (global.get $img_count) (i32.const 1)))
          )
        )

        ;; Next PICF: try off + lcb, but also scan forward for cbHeader=0x44
        (local.set $off (i32.add (local.get $off) (local.get $lcb)))

        ;; Scan for next valid PICF (cbHeader=0x44 with SpContainer)
        (block $found_next
          (loop $scan
            (br_if $done (i32.ge_u (i32.add (local.get $off) (i32.const 70)) (global.get $data_len)))
            (if (i32.eq
                  (call $read_u16_le (i32.add (global.get $data_ptr) (i32.add (local.get $off) (i32.const 4))))
                  (i32.const 0x44)
                )
              (then (br $found_next))
            )
            (local.set $off (i32.add (local.get $off) (i32.const 2)))
            (br $scan)
          )
        )

        (br $loop)
      )
    )
  )

  ;; ── Layout Engine ───────────────────────────────────────────
  ;; Builds layout data at LAYOUT_BASE.
  ;; Uses $measureText import for line breaking.
  ;;
  ;; Layout format:
  ;; HEADER (16 bytes): page_count(u32), total_lines(u32), reserved(u32), reserved(u32)
  ;; Then per page: PAGE_HDR (16 bytes): page_idx(u32), line_count(u32), y_start(f32), reserved(u32)
  ;; Then per line: LINE (16 bytes): y_pos(f32), seg_count(u32), x_start(f32), line_width(f32)
  ;; Then per segment: SEG (24 bytes): text_ptr(u32), text_len(u32), x_pos(f32), font_size(u32), flags(u32), color(u32)
  ;;
  ;; For simplicity in this pass, we use a flat array of segments grouped by page.
  ;; We store segments sequentially and track page boundaries.

  ;; Layout constants — defaults for 8.5 x 11 inches at 96 DPI
  ;; These are mutable: SEP parsing may update them
  (global $PAGE_WIDTH_PX  (mut f32) (f32.const 816.0))
  (global $PAGE_HEIGHT_PX (mut f32) (f32.const 1056.0))
  (global $MARGIN_PX      (mut f32) (f32.const 96.0))
  ;; Separate margins for more accurate layout
  (global $MARGIN_LEFT_PX   (mut f32) (f32.const 96.0))
  (global $MARGIN_RIGHT_PX  (mut f32) (f32.const 96.0))
  (global $MARGIN_TOP_PX    (mut f32) (f32.const 96.0))
  (global $MARGIN_BOTTOM_PX (mut f32) (f32.const 96.0))

  ;; Layout segment: 24 bytes
  ;; [0..3]   text_ptr (i32) — pointer into wasm memory
  ;; [4..7]   text_len (i32) — bytes (UTF-16LE)
  ;; [8..11]  x_pos (f32)
  ;; [12..15] y_pos (f32)
  ;; [16..19] flags (i32) — font_size in high 16, fmt flags in low 16
  ;; [20..23] color (i32)

  (global $layout_seg_count (mut i32) (i32.const 0))

  ;; Page info: store (start_seg_idx, seg_count) per page at LAYOUT_BASE
  ;; Page table: starts at LAYOUT_BASE, each entry 8 bytes
  ;; Max 1000 pages
  ;; Segment data: starts at LAYOUT_BASE + 8000

  (global $LAYOUT_PAGE_TABLE i32 (i32.const 0x002B4000))
  (global $LAYOUT_SEG_DATA   i32 (i32.const 0x002B5F40))  ;; +8000

  (func $write_layout_seg (param $idx i32) (param $text_ptr i32) (param $text_len i32)
                           (param $x f32) (param $y f32) (param $flags_and_size i32) (param $color i32)
    (local $ptr i32)
    (local.set $ptr (i32.add (global.get $LAYOUT_SEG_DATA) (i32.mul (local.get $idx) (i32.const 24))))
    (i32.store (local.get $ptr) (local.get $text_ptr))
    (i32.store (i32.add (local.get $ptr) (i32.const 4)) (local.get $text_len))
    (f32.store (i32.add (local.get $ptr) (i32.const 8)) (local.get $x))
    (f32.store (i32.add (local.get $ptr) (i32.const 12)) (local.get $y))
    (i32.store (i32.add (local.get $ptr) (i32.const 16)) (local.get $flags_and_size))
    (i32.store (i32.add (local.get $ptr) (i32.const 20)) (local.get $color))
  )

  ;; Find which CHP run covers a given CP
  (func $find_chp_at_cp (param $cp i32) (result i32)
    ;; Returns index into CHP_BASE, or 0 if not found (uses first run as default)
    (local $i i32)
    (local $ptr i32)
    (local $start i32)
    (local $end i32)
    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (global.get $chp_run_count)))
        (local.set $ptr (i32.add (global.get $CHP_BASE) (i32.mul (local.get $i) (i32.const 28))))
        (local.set $start (i32.load (local.get $ptr)))
        (local.set $end (i32.load (i32.add (local.get $ptr) (i32.const 4))))
        (if (i32.and
              (i32.ge_s (local.get $cp) (local.get $start))
              (i32.lt_s (local.get $cp) (local.get $end))
            )
          (then (return (local.get $i)))
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )
    (i32.const 0)
  )

  ;; Find PAP run index for a given CP
  (func $find_pap_at_cp (param $cp i32) (result i32)
    (local $i i32)
    (local $ptr i32)
    (local $start i32)
    (local $end i32)
    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (global.get $pap_run_count)))
        (local.set $ptr (i32.add (global.get $PAP_BASE) (i32.mul (local.get $i) (i32.const 28))))
        (local.set $start (i32.load (local.get $ptr)))
        (local.set $end (i32.load (i32.add (local.get $ptr) (i32.const 4))))
        (if (i32.and
              (i32.ge_s (local.get $cp) (local.get $start))
              (i32.lt_s (local.get $cp) (local.get $end))
            )
          (then (return (local.get $i)))
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )
    (i32.const 0)
  )

  ;; Get the style index (istd) for a given CP from PAP runs
  (func $get_istd_at_cp (param $cp i32) (result i32)
    (local $pap_idx i32)
    (local $ptr i32)
    (local.set $pap_idx (call $find_pap_at_cp (local.get $cp)))
    (local.set $ptr (i32.add (global.get $PAP_BASE) (i32.mul (local.get $pap_idx) (i32.const 28))))
    (i32.load (i32.add (local.get $ptr) (i32.const 24)))
  )

  ;; Find PAP run for a given CP, returns alignment (0=left,1=center,2=right,3=justify)
  (func $get_pap_alignment (param $cp i32) (result i32)
    (local $i i32)
    (local $ptr i32)
    (local $start i32)
    (local $end i32)
    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (global.get $pap_run_count)))
        (local.set $ptr (i32.add (global.get $PAP_BASE) (i32.mul (local.get $i) (i32.const 28))))
        (local.set $start (i32.load (local.get $ptr)))
        (local.set $end (i32.load (i32.add (local.get $ptr) (i32.const 4))))
        (if (i32.and
              (i32.ge_s (local.get $cp) (local.get $start))
              (i32.lt_s (local.get $cp) (local.get $end))
            )
          (then (return (i32.load (i32.add (local.get $ptr) (i32.const 8)))))
        )
        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )
    (i32.const 0)  ;; default left
  )

  ;; Adjust x positions of segments on a line for alignment
  ;; Shifts segments [start_seg, end_seg) based on alignment and line width
  (func $align_line (param $start_seg i32) (param $end_seg i32) (param $alignment i32) (param $line_end_x f32)
    (local $i i32)
    (local $seg_ptr i32)
    (local $line_width f32)
    (local $shift f32)
    (local $content_right f32)

    ;; Only adjust for center (1) or right (2)
    (if (i32.eqz (local.get $alignment)) (then (return)))
    (if (i32.ge_u (local.get $start_seg) (local.get $end_seg)) (then (return)))

    (local.set $content_right (f32.sub (global.get $PAGE_WIDTH_PX) (global.get $MARGIN_RIGHT_PX)))
    (local.set $line_width (f32.sub (local.get $line_end_x) (global.get $MARGIN_LEFT_PX)))

    ;; Center: shift right by (available - used) / 2
    (if (i32.eq (local.get $alignment) (i32.const 1))
      (then
        (local.set $shift
          (f32.div
            (f32.sub (f32.sub (local.get $content_right) (global.get $MARGIN_LEFT_PX)) (local.get $line_width))
            (f32.const 2.0)
          )
        )
      )
    )
    ;; Right: shift right by (available - used)
    (if (i32.eq (local.get $alignment) (i32.const 2))
      (then
        (local.set $shift
          (f32.sub (f32.sub (local.get $content_right) (global.get $MARGIN_LEFT_PX)) (local.get $line_width))
        )
      )
    )

    ;; Apply shift to all segments on this line
    (if (f32.gt (local.get $shift) (f32.const 0.0))
      (then
        (local.set $i (local.get $start_seg))
        (block $done
          (loop $loop
            (br_if $done (i32.ge_u (local.get $i) (local.get $end_seg)))
            (local.set $seg_ptr
              (i32.add (global.get $LAYOUT_SEG_DATA) (i32.mul (local.get $i) (i32.const 24)))
            )
            (f32.store (i32.add (local.get $seg_ptr) (i32.const 8))
              (f32.add (f32.load (i32.add (local.get $seg_ptr) (i32.const 8))) (local.get $shift))
            )
            (local.set $i (i32.add (local.get $i) (i32.const 1)))
            (br $loop)
          )
        )
      )
    )
  )

  ;; Twips to pixels: px = twips * 96 / 1440
  (func $twips_to_px (param $twips i32) (result f32)
    (f32.div
      (f32.mul (f32.convert_i32_s (local.get $twips)) (f32.const 96.0))
      (f32.const 1440.0)
    )
  )

  (func $do_layout
    (local $total_cps i32)         ;; total character positions
    (local $cp i32)                ;; current character position
    (local $cur_x f32)             ;; current x position
    (local $cur_y f32)             ;; current y position
    (local $content_width f32)     ;; page width - 2*margin
    (local $content_height f32)    ;; page height - 2*margin
    (local $line_height f32)       ;; current line height
    (local $page_num i32)
    (local $page_start_seg i32)
    (local $seg_count_on_page i32)
    (local $word_start i32)        ;; CP of word start
    (local $word_end i32)
    (local $in_field i32)          ;; 1 = inside field instruction (skip)
    (local $line_start_seg i32)    ;; first segment index of current line
    (local $para_align i32)        ;; current paragraph alignment
    (local $char_code i32)
    (local $chp_idx i32)
    (local $chp_ptr i32)
    (local $chp_flags i32)
    (local $chp_size i32)
    (local $chp_color i32)
    (local $chp_font_index i32)
    (local $font_name_ptr i32)
    (local $font_name_len i32)
    (local $word_width f32)
    (local $word_text_ptr i32)
    (local $word_text_len i32)
    (local $flags_and_size i32)
    (local $font_height f32)
    (local $space_width f32)
    (local $pap_idx i32)
    (local $pap_ptr i32)
    (local $pap_space_after i32)

    (local.set $total_cps (i32.div_u (global.get $text_len) (i32.const 2)))
    (local.set $content_width (f32.sub (global.get $PAGE_WIDTH_PX) (f32.add (global.get $MARGIN_LEFT_PX) (global.get $MARGIN_RIGHT_PX))))
    (local.set $content_height (f32.sub (global.get $PAGE_HEIGHT_PX) (f32.add (global.get $MARGIN_TOP_PX) (global.get $MARGIN_BOTTOM_PX))))

    (local.set $cur_x (global.get $MARGIN_LEFT_PX))
    (local.set $cur_y (global.get $MARGIN_TOP_PX))
    (local.set $line_height (f32.const 16.0))  ;; default ~12pt
    (local.set $page_num (i32.const 0))
    (local.set $page_start_seg (i32.const 0))
    (local.set $seg_count_on_page (i32.const 0))
    (global.set $layout_seg_count (i32.const 0))

    (local.set $cp (i32.const 0))

    ;; Initialize alignment for first paragraph
    (local.set $para_align (call $get_pap_alignment (i32.const 0)))

    (block $all_done
      (loop $cp_loop
        (br_if $all_done (i32.ge_u (local.get $cp) (local.get $total_cps)))

        ;; Read character at cp
        (local.set $char_code
          (call $read_u16_le
            (i32.add (global.get $text_ptr) (i32.mul (local.get $cp) (i32.const 2)))
          )
        )

        ;; Handle field codes: 0x13=begin, 0x14=separator, 0x15=end
        ;; Skip field instructions (between 0x13 and 0x14), show field results
        (if (i32.eq (local.get $char_code) (i32.const 0x13))
          (then
            (local.set $in_field (i32.const 1))
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )
        (if (i32.eq (local.get $char_code) (i32.const 0x14))
          (then
            (local.set $in_field (i32.const 0))
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )
        (if (i32.eq (local.get $char_code) (i32.const 0x15))
          (then
            (local.set $in_field (i32.const 0))
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )
        ;; Skip if inside field instruction
        (if (local.get $in_field)
          (then
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )

        ;; Handle paragraph break (0x0D = carriage return)
        (if (i32.eq (local.get $char_code) (i32.const 0x0D))
          (then
            ;; Align the current line before moving to next
            (call $align_line
              (local.get $line_start_seg)
              (global.get $layout_seg_count)
              (local.get $para_align)
              (local.get $cur_x)
            )

            ;; Get PAP space_after for current paragraph
            (local.set $pap_idx (call $find_pap_at_cp (local.get $cp)))
            (local.set $pap_ptr (i32.add (global.get $PAP_BASE) (i32.mul (local.get $pap_idx) (i32.const 28))))
            (local.set $pap_space_after (i32.load (i32.add (local.get $pap_ptr) (i32.const 16))))

            ;; New line after paragraph: line_height + space_after (twips→px)
            (local.set $cur_y (f32.add (local.get $cur_y)
              (f32.add (local.get $line_height)
                (if (result f32) (local.get $pap_space_after)
                  (then (call $twips_to_px (local.get $pap_space_after)))
                  (else (f32.const 4.0))
                )
              )
            ))
            (local.set $cur_x (global.get $MARGIN_LEFT_PX))
            (local.set $line_height (f32.const 16.0))
            (local.set $line_start_seg (global.get $layout_seg_count))

            ;; Update alignment for next paragraph
            (local.set $para_align (call $get_pap_alignment (i32.add (local.get $cp) (i32.const 1))))

            ;; Page break check
            (if (f32.ge (local.get $cur_y) (f32.sub (global.get $PAGE_HEIGHT_PX) (global.get $MARGIN_BOTTOM_PX)))
              (then
                ;; Save page info
                (i32.store
                  (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.mul (local.get $page_num) (i32.const 8)))
                  (local.get $page_start_seg)
                )
                (i32.store
                  (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.add (i32.mul (local.get $page_num) (i32.const 8)) (i32.const 4)))
                  (local.get $seg_count_on_page)
                )
                (local.set $page_num (i32.add (local.get $page_num) (i32.const 1)))
                (local.set $page_start_seg (global.get $layout_seg_count))
                (local.set $seg_count_on_page (i32.const 0))
                (local.set $cur_y (global.get $MARGIN_TOP_PX))
                (local.set $cur_x (global.get $MARGIN_LEFT_PX))
              )
            )

            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )

        ;; Handle page break (0x0C)
        (if (i32.eq (local.get $char_code) (i32.const 0x0C))
          (then
            (i32.store
              (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.mul (local.get $page_num) (i32.const 8)))
              (local.get $page_start_seg)
            )
            (i32.store
              (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.add (i32.mul (local.get $page_num) (i32.const 8)) (i32.const 4)))
              (local.get $seg_count_on_page)
            )
            (local.set $page_num (i32.add (local.get $page_num) (i32.const 1)))
            (local.set $page_start_seg (global.get $layout_seg_count))
            (local.set $seg_count_on_page (i32.const 0))
            (local.set $cur_y (global.get $MARGIN_TOP_PX))
            (local.set $cur_x (global.get $MARGIN_LEFT_PX))
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )

        ;; Handle embedded object (0x01) — inline image
        (if (i32.and
              (i32.eq (local.get $char_code) (i32.const 0x01))
              (i32.lt_u (global.get $img_cursor) (global.get $img_count))
            )
          (then
            ;; Write image segment: use negative text_len as signal to renderer
            ;; flags_and_size encodes image index in high 16 bits, 0xFFFF in low 16 as marker
            (call $write_layout_seg
              (global.get $layout_seg_count)
              ;; text_ptr = image data ptr (from image table)
              (i32.load (i32.add (global.get $IMG_TABLE) (i32.mul (global.get $img_cursor) (i32.const 16))))
              ;; text_len = image data len
              (i32.load (i32.add (i32.add (global.get $IMG_TABLE) (i32.mul (global.get $img_cursor) (i32.const 16))) (i32.const 4)))
              ;; x = cur_x
              (local.get $cur_x)
              ;; y = cur_y
              (local.get $cur_y)
              ;; flags: 0xFFFF marker in low 16, img_cursor in high 16
              (i32.or (i32.const 0xFFFF) (i32.shl (global.get $img_cursor) (i32.const 16)))
              ;; color = packed width(u16)|height(u16) in pixels (capped to u16)
              (i32.or
                (i32.and
                  (i32.trunc_f32_u (f32.load (i32.add (i32.add (global.get $IMG_TABLE) (i32.mul (global.get $img_cursor) (i32.const 16))) (i32.const 8))))
                  (i32.const 0xFFFF)
                )
                (i32.shl
                  (i32.and
                    (i32.trunc_f32_u (f32.load (i32.add (i32.add (global.get $IMG_TABLE) (i32.mul (global.get $img_cursor) (i32.const 16))) (i32.const 12))))
                    (i32.const 0xFFFF)
                  )
                  (i32.const 16)
                )
              )
            )
            (global.set $layout_seg_count (i32.add (global.get $layout_seg_count) (i32.const 1)))
            (local.set $seg_count_on_page (i32.add (local.get $seg_count_on_page) (i32.const 1)))

            ;; Advance y by image height + small gap
            (local.set $cur_y
              (f32.add (local.get $cur_y)
                (f32.add
                  (f32.load (i32.add (i32.add (global.get $IMG_TABLE) (i32.mul (global.get $img_cursor) (i32.const 16))) (i32.const 12)))
                  (f32.const 4.0)
                )
              )
            )
            (local.set $cur_x (global.get $MARGIN_LEFT_PX))
            (local.set $line_start_seg (global.get $layout_seg_count))

            ;; Page break check after image
            (if (f32.ge (local.get $cur_y) (f32.sub (global.get $PAGE_HEIGHT_PX) (global.get $MARGIN_BOTTOM_PX)))
              (then
                (i32.store
                  (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.mul (local.get $page_num) (i32.const 8)))
                  (local.get $page_start_seg)
                )
                (i32.store
                  (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.add (i32.mul (local.get $page_num) (i32.const 8)) (i32.const 4)))
                  (local.get $seg_count_on_page)
                )
                (local.set $page_num (i32.add (local.get $page_num) (i32.const 1)))
                (local.set $page_start_seg (global.get $layout_seg_count))
                (local.set $seg_count_on_page (i32.const 0))
                (local.set $cur_y (global.get $MARGIN_TOP_PX))
                (local.set $cur_x (global.get $MARGIN_LEFT_PX))
              )
            )

            (global.set $img_cursor (i32.add (global.get $img_cursor) (i32.const 1)))
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )

        ;; Skip control characters and special Word markers
        ;; 0x01=embedded obj (no image), 0x07=cell mark, 0x08=drawn obj, 0x0A=LF, etc.
        (if (i32.lt_u (local.get $char_code) (i32.const 0x20))
          (then
            ;; Tab (0x09): advance x by ~4 spaces worth
            (if (i32.eq (local.get $char_code) (i32.const 0x09))
              (then
                (local.set $cur_x (f32.add (local.get $cur_x) (f32.const 48.0)))
              )
            )
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $cp_loop)
          )
        )

        ;; Find word boundary (sequence of non-space, non-control chars)
        ;; Also break at CHP run boundaries so mid-word formatting changes render correctly
        (local.set $word_start (local.get $cp))
        (local.set $chp_idx (call $find_chp_at_cp (local.get $word_start)))
        (local.set $chp_ptr (i32.add (global.get $CHP_BASE) (i32.mul (local.get $chp_idx) (i32.const 28))))
        (block $word_done
          (loop $word_loop
            (br_if $word_done (i32.ge_u (local.get $cp) (local.get $total_cps)))
            (local.set $char_code
              (call $read_u16_le
                (i32.add (global.get $text_ptr) (i32.mul (local.get $cp) (i32.const 2)))
              )
            )
            ;; Break on space, CR, LF, tab, page break
            (br_if $word_done (i32.eq (local.get $char_code) (i32.const 0x20)))
            (br_if $word_done (i32.eq (local.get $char_code) (i32.const 0x0D)))
            (br_if $word_done (i32.eq (local.get $char_code) (i32.const 0x0A)))
            (br_if $word_done (i32.eq (local.get $char_code) (i32.const 0x09)))
            (br_if $word_done (i32.eq (local.get $char_code) (i32.const 0x0C)))
            (br_if $word_done (i32.lt_u (local.get $char_code) (i32.const 0x20)))
            ;; Break at CHP boundary: if cp >= chp_end, stop
            (br_if $word_done
              (i32.ge_s (local.get $cp)
                (i32.load (i32.add (local.get $chp_ptr) (i32.const 4)))
              )
            )
            (local.set $cp (i32.add (local.get $cp) (i32.const 1)))
            (br $word_loop)
          )
        )
        (local.set $word_end (local.get $cp))

        ;; Skip if empty word
        (if (i32.le_u (local.get $word_end) (local.get $word_start))
          (then
            ;; It's a space or tab — advance
            (if (i32.lt_u (local.get $cp) (local.get $total_cps))
              (then (local.set $cp (i32.add (local.get $cp) (i32.const 1))))
            )
            (br $cp_loop)
          )
        )

        ;; CHP already set during word boundary scan above
        (local.set $chp_flags (i32.load (i32.add (local.get $chp_ptr) (i32.const 8))))
        (local.set $chp_size (i32.load (i32.add (local.get $chp_ptr) (i32.const 12))))
        (local.set $chp_color (i32.load (i32.add (local.get $chp_ptr) (i32.const 16))))
        (local.set $chp_font_index (i32.load (i32.add (local.get $chp_ptr) (i32.const 20))))

        ;; Look up font name from FONT_TABLE
        (local.set $font_name_ptr (i32.const 0))
        (local.set $font_name_len (i32.const 0))
        (if (i32.and
              (i32.gt_u (global.get $font_count) (i32.const 0))
              (i32.lt_u (local.get $chp_font_index) (global.get $font_count))
            )
          (then
            (local.set $font_name_ptr
              (i32.load (i32.add (global.get $FONT_TABLE) (i32.mul (local.get $chp_font_index) (i32.const 8)))))
            (local.set $font_name_len
              (i32.load (i32.add (i32.add (global.get $FONT_TABLE) (i32.mul (local.get $chp_font_index) (i32.const 8))) (i32.const 4))))
          )
        )

        ;; Set font for measurement
        (call $setFont (local.get $chp_size)
          (i32.and (local.get $chp_flags) (i32.const 1))
          (i32.and (i32.shr_u (local.get $chp_flags) (i32.const 1)) (i32.const 1))
          (local.get $font_name_ptr)
          (local.get $font_name_len)
        )

        ;; Measure word
        (local.set $word_text_ptr
          (i32.add (global.get $text_ptr) (i32.mul (local.get $word_start) (i32.const 2)))
        )
        (local.set $word_text_len
          (i32.mul (i32.sub (local.get $word_end) (local.get $word_start)) (i32.const 2))
        )
        (local.set $word_width
          (call $measureText (local.get $word_text_ptr) (local.get $word_text_len))
        )

        ;; Font height in pixels (half-points / 2 * 96/72 = half-points * 2/3)
        (local.set $font_height
          (f32.mul (f32.convert_i32_u (local.get $chp_size)) (f32.const 0.6667))
        )

        ;; Measure space
        ;; Approximate space width as font_height * 0.3
        (local.set $space_width (f32.mul (local.get $font_height) (f32.const 0.3)))

        ;; Line break check: if word doesn't fit on current line
        (if (f32.gt
              (f32.add (local.get $cur_x) (local.get $word_width))
              (f32.sub (global.get $PAGE_WIDTH_PX) (global.get $MARGIN_RIGHT_PX))
            )
          (then
            ;; Align the finished line before wrapping
            (call $align_line
              (local.get $line_start_seg)
              (global.get $layout_seg_count)
              (local.get $para_align)
              (local.get $cur_x)
            )

            ;; Wrap to next line
            (local.set $cur_y (f32.add (local.get $cur_y) (local.get $line_height)))
            (local.set $cur_x (global.get $MARGIN_LEFT_PX))
            (local.set $line_height (f32.const 16.0))
            (local.set $line_start_seg (global.get $layout_seg_count))

            ;; Page break check
            (if (f32.ge (local.get $cur_y) (f32.sub (global.get $PAGE_HEIGHT_PX) (global.get $MARGIN_BOTTOM_PX)))
              (then
                (i32.store
                  (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.mul (local.get $page_num) (i32.const 8)))
                  (local.get $page_start_seg)
                )
                (i32.store
                  (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.add (i32.mul (local.get $page_num) (i32.const 8)) (i32.const 4)))
                  (local.get $seg_count_on_page)
                )
                (local.set $page_num (i32.add (local.get $page_num) (i32.const 1)))
                (local.set $page_start_seg (global.get $layout_seg_count))
                (local.set $seg_count_on_page (i32.const 0))
                (local.set $cur_y (global.get $MARGIN_TOP_PX))
              )
            )
          )
        )

        ;; Update line height to max of current and this font
        (if (f32.gt (local.get $font_height) (local.get $line_height))
          (then (local.set $line_height (local.get $font_height)))
        )

        ;; Write segment
        ;; Pack: low 4 bits = format flags, bits 4-11 = font_index, high 16 = font_size
        (local.set $flags_and_size
          (i32.or
            (i32.or
              (i32.and (local.get $chp_flags) (i32.const 0xF))
              (i32.shl (i32.and (local.get $chp_font_index) (i32.const 0xFF)) (i32.const 4))
            )
            (i32.shl (local.get $chp_size) (i32.const 16))
          )
        )

        (call $write_layout_seg
          (global.get $layout_seg_count)
          (local.get $word_text_ptr)
          (local.get $word_text_len)
          (local.get $cur_x)
          (local.get $cur_y)
          (local.get $flags_and_size)
          (local.get $chp_color)
        )
        (global.set $layout_seg_count (i32.add (global.get $layout_seg_count) (i32.const 1)))
        (local.set $seg_count_on_page (i32.add (local.get $seg_count_on_page) (i32.const 1)))

        ;; Advance x by word width + space
        (local.set $cur_x (f32.add (local.get $cur_x) (f32.add (local.get $word_width) (local.get $space_width))))

        ;; Skip trailing space
        (if (i32.lt_u (local.get $cp) (local.get $total_cps))
          (then
            (local.set $char_code
              (call $read_u16_le
                (i32.add (global.get $text_ptr) (i32.mul (local.get $cp) (i32.const 2)))
              )
            )
            (if (i32.eq (local.get $char_code) (i32.const 0x20))
              (then (local.set $cp (i32.add (local.get $cp) (i32.const 1))))
            )
          )
        )

        (br $cp_loop)
      )
    )

    ;; Save last page
    (i32.store
      (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.mul (local.get $page_num) (i32.const 8)))
      (local.get $page_start_seg)
    )
    (i32.store
      (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.add (i32.mul (local.get $page_num) (i32.const 8)) (i32.const 4)))
      (local.get $seg_count_on_page)
    )

    (global.set $page_count (i32.add (local.get $page_num) (i32.const 1)))
  )

  ;; ── Render ──────────────────────────────────────────────────
  ;; Walk layout segments for the requested page and call canvas imports

  (func $render (param $page i32)
    (local $page_info_ptr i32)
    (local $start_seg i32)
    (local $seg_count i32)
    (local $i i32)
    (local $seg_ptr i32)
    (local $text_ptr_val i32)
    (local $text_len_val i32)
    (local $x f32)
    (local $y f32)
    (local $flags_and_size i32)
    (local $flags i32)
    (local $font_size i32)
    (local $color i32)
    (local $font_index i32)
    (local $font_name_ptr i32)
    (local $font_name_len i32)

    ;; Bounds check
    (if (i32.ge_u (local.get $page) (global.get $page_count))
      (then (return))
    )

    ;; Set up page canvas
    (call $setPage (local.get $page) (global.get $PAGE_WIDTH_PX) (global.get $PAGE_HEIGHT_PX))

    ;; Read page table entry
    (local.set $page_info_ptr
      (i32.add (global.get $LAYOUT_PAGE_TABLE) (i32.mul (local.get $page) (i32.const 8)))
    )
    (local.set $start_seg (i32.load (local.get $page_info_ptr)))
    (local.set $seg_count (i32.load (i32.add (local.get $page_info_ptr) (i32.const 4))))

    ;; Render each segment
    (local.set $i (i32.const 0))
    (block $done
      (loop $loop
        (br_if $done (i32.ge_u (local.get $i) (local.get $seg_count)))

        (local.set $seg_ptr
          (i32.add (global.get $LAYOUT_SEG_DATA)
            (i32.mul (i32.add (local.get $start_seg) (local.get $i)) (i32.const 24))
          )
        )

        (local.set $text_ptr_val (i32.load (local.get $seg_ptr)))
        (local.set $text_len_val (i32.load (i32.add (local.get $seg_ptr) (i32.const 4))))
        (local.set $x (f32.load (i32.add (local.get $seg_ptr) (i32.const 8))))
        (local.set $y (f32.load (i32.add (local.get $seg_ptr) (i32.const 12))))
        (local.set $flags_and_size (i32.load (i32.add (local.get $seg_ptr) (i32.const 16))))
        (local.set $color (i32.load (i32.add (local.get $seg_ptr) (i32.const 20))))

        ;; Unpack: low 4 bits = format flags, bits 4-11 = font_index, high 16 = font_size
        (local.set $flags (i32.and (local.get $flags_and_size) (i32.const 0xF)))
        (local.set $font_index (i32.and (i32.shr_u (local.get $flags_and_size) (i32.const 4)) (i32.const 0xFF)))
        (local.set $font_size (i32.shr_u (local.get $flags_and_size) (i32.const 16)))

        ;; Check if this is an image segment (flags_and_size low 16 = 0xFFFF)
        (if (i32.eq (i32.and (local.get $flags_and_size) (i32.const 0xFFFF)) (i32.const 0xFFFF))
          (then
            ;; Image: text_ptr=image data, text_len=image data len
            ;; color = width(low16) | height(high16) packed
            (call $drawImage
              (local.get $text_ptr_val)
              (local.get $text_len_val)
              (local.get $x)
              (local.get $y)
              (f32.convert_i32_u (i32.and (local.get $color) (i32.const 0xFFFF)))
              (f32.convert_i32_u (i32.shr_u (local.get $color) (i32.const 16)))
            )
            (local.set $i (i32.add (local.get $i) (i32.const 1)))
            (br $loop)
          )
        )

        ;; Look up font name
        (local.set $font_name_ptr (i32.const 0))
        (local.set $font_name_len (i32.const 0))
        (if (i32.and
              (i32.gt_u (global.get $font_count) (i32.const 0))
              (i32.lt_u (local.get $font_index) (global.get $font_count))
            )
          (then
            (local.set $font_name_ptr
              (i32.load (i32.add (global.get $FONT_TABLE) (i32.mul (local.get $font_index) (i32.const 8)))))
            (local.set $font_name_len
              (i32.load (i32.add (i32.add (global.get $FONT_TABLE) (i32.mul (local.get $font_index) (i32.const 8))) (i32.const 4))))
          )
        )

        ;; Set font and color
        (call $setFont (local.get $font_size)
          (i32.and (local.get $flags) (i32.const 1))
          (i32.and (i32.shr_u (local.get $flags) (i32.const 1)) (i32.const 1))
          (local.get $font_name_ptr)
          (local.get $font_name_len)
        )
        (call $setColor (local.get $color))

        ;; Draw text (y needs offset by font ascent — approximate as font_height * 0.8)
        (call $fillText
          (local.get $text_ptr_val)
          (local.get $text_len_val)
          (local.get $x)
          (f32.add (local.get $y)
            (f32.mul (f32.convert_i32_u (local.get $font_size)) (f32.const 0.5333))
          )
        )

        ;; Underline (bit 2)
        (if (i32.and (local.get $flags) (i32.const 4))
          (then
            (call $setColor (local.get $color))
            (call $fillRect
              (local.get $x)
              (f32.add (local.get $y)
                (f32.mul (f32.convert_i32_u (local.get $font_size)) (f32.const 0.5667))
              )
              ;; Width: use measureText
              (call $measureText (local.get $text_ptr_val) (local.get $text_len_val))
              (f32.const 1.0)
            )
          )
        )

        ;; Strikethrough (bit 3)
        (if (i32.and (local.get $flags) (i32.const 8))
          (then
            (call $setColor (local.get $color))
            (call $fillRect
              (local.get $x)
              (f32.add (local.get $y)
                (f32.mul (f32.convert_i32_u (local.get $font_size)) (f32.const 0.3))
              )
              (call $measureText (local.get $text_ptr_val) (local.get $text_len_val))
              (f32.const 1.0)
            )
          )
        )

        (local.set $i (i32.add (local.get $i) (i32.const 1)))
        (br $loop)
      )
    )
  )

  ;; ── Exports ─────────────────────────────────────────────────

  (func $set_input (param $ptr i32) (param $len i32)
    (global.set $input_ptr (local.get $ptr))
    (global.set $input_len (local.get $len))
    ;; Set arena to start AFTER the input (page-aligned)
    (global.set $arena_base
      (i32.and
        (i32.add (i32.add (local.get $ptr) (local.get $len)) (i32.const 0xFFFF))
        (i32.const 0xFFFF0000)
      )
    )
    (global.set $arena_ptr (global.get $arena_base))
  )

  (func $get_text_ptr (result i32) (global.get $text_ptr))
  (func $get_text_len (result i32) (global.get $text_len))
  (func $get_page_count (result i32) (global.get $page_count))
  (func $get_error_code (result i32) (global.get $error_code))

  (export "set_input" (func $set_input))
  (export "parse" (func $parse))
  (export "render" (func $render))
  (export "get_text_ptr" (func $get_text_ptr))
  (export "get_text_len" (func $get_text_len))
  (export "get_page_count" (func $get_page_count))
  (export "get_error_code" (func $get_error_code))
)
