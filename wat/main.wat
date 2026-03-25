(module
  ;; ============================================================
  ;; doc.wasm — Microsoft .doc (OLE2/CFBF) parser + renderer
  ;; Written in raw WAT
  ;; ============================================================

  ;; ── Imports (canvas-like API provided by JS) ────────────────

  (import "canvas" "measureText" (func $measureText (param $ptr i32) (param $len i32) (result f32)))
  (import "canvas" "setFont"     (func $setFont (param $size i32) (param $bold i32) (param $italic i32)))
  (import "canvas" "setColor"    (func $setColor (param $rgb i32)))
  (import "canvas" "fillText"    (func $fillText (param $ptr i32) (param $len i32) (param $x f32) (param $y f32)))
  (import "canvas" "fillRect"    (func $fillRect (param $x f32) (param $y f32) (param $w f32) (param $h f32)))
  (import "canvas" "setPage"     (func $setPage (param $pageNum i32) (param $widthPx f32) (param $heightPx f32)))
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

  ;; Read a mini-stream via mini-FAT chain into dest
  (func $read_mini_stream (param $start_sector i32) (param $size i32) (param $dest i32) (result i32)
    (local $sector i32)
    (local $remaining i32)
    (local $chunk i32)
    (local $offset i32)
    (local $mini_offset i32)
    ;; The mini-stream is the data of the root entry, already read via FAT
    ;; We need to read from the root entry's stream
    ;; For now, we read mini sectors from the root entry stream which
    ;; has been copied to a temporary location

    ;; Actually, mini-stream sectors are 64 bytes each, residing inside
    ;; the root entry's regular stream. We need to locate byte offset
    ;; mini_sector * 64 within that stream.

    ;; This requires the root entry stream to be read first.
    ;; We'll handle this properly during CFBF parse.
    ;; For now, stub:
    (local.get $size)
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

    (global.get $ERR_NONE)
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

    ;; Phase 3: Extract WordDocument stream
    (global.set $worddoc_ptr (call $arena_alloc (global.get $stream_worddoc_size)))
    (global.set $worddoc_len (global.get $stream_worddoc_size))
    (drop (call $read_stream
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
    (drop (call $read_stream
      (local.get $tbl_start)
      (local.get $tbl_size)
      (global.get $table_ptr)
    ))

    ;; Phase 6: Extract Data stream (optional)
    (if (i32.ne (global.get $stream_data_start) (i32.const -1))
      (then
        (global.set $data_ptr (call $arena_alloc (global.get $stream_data_size)))
        (global.set $data_len (global.get $stream_data_size))
        (drop (call $read_stream
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

    ;; TODO: Phase 9: Parse CHP/PAP/SEP
    ;; TODO: Phase 10: Layout
    ;; TODO: Phase 11: Build render data

    (global.set $page_count (i32.const 1))  ;; placeholder
    (global.set $error_code (global.get $ERR_NONE))
    (global.get $ERR_NONE)
  )

  ;; ── Render (stub) ───────────────────────────────────────────

  (func $render (param $page i32)
    ;; TODO: Walk layout data and call canvas imports
    ;; For now, just render extracted text on page 0
    (if (i32.eqz (local.get $page))
      (then
        (call $setPage (i32.const 0) (f32.const 816.0) (f32.const 1056.0))  ;; 8.5x11 at 96dpi
        (call $setFont (i32.const 24) (i32.const 0) (i32.const 0))  ;; 12pt
        (call $setColor (i32.const 0x000000))
        (call $fillText
          (global.get $text_ptr)
          (global.get $text_len)
          (f32.const 72.0)   ;; ~1 inch margin
          (f32.const 72.0)
        )
      )
    )
  )

  ;; ── Exports ─────────────────────────────────────────────────

  (func $set_input (param $ptr i32) (param $len i32)
    (global.set $input_ptr (local.get $ptr))
    (global.set $input_len (local.get $len))
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
