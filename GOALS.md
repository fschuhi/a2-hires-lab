# a2-hires-lab -- Goals and Roadmap

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

**Charter:** This file answers: where is the project going, in what order, and what happens next. It holds the strategic vision, the phased roadmap, and goals that need a strategy discussion before they are actionable. The _Current Session Pointer_ below is the single canonical "where we are / what's next" -- keep it to a few lines, update it, don't grow it; `FIRST_PROMPT.md` sends the reader here first. Concrete, startable work lives in `TODO.md`; the resolved-work record lives in `HISTORY.md` or `CHANGELOG.md`(on the heap, out of the per-session dump); architecture, contract, and settled decisions live in `README.md`.

---

## 📍 Current Session Pointer

**Where we are:**
Deliverable 1 (Sprite Editor/Viewer) is implemented: sheet layout, `Util`, `NTSCColor`, and `SpriteEditor` (both buttons, idempotent) are all built and importable. Verified against the first worked sprite from Chapter 3 page 8 -- the color viewer reproduces the solid-blue "5" shape exactly. Along the way we found and fixed a real gap in the NTSC color decision table: a colored 0-pixel (sandwiched between two 1s) must take its hue from its flanking 1-bit's column, not its own -- otherwise a repeating `0x55`-style byte renders as alternating stripes instead of the solid fill it's actually used for on real hardware. `RowColors` also now takes an absolute `iBaseCol` (plus optional true screen neighbors) instead of assuming the sprite sits at column 0, since color is a function of absolute screen position, not sprite-local position.

**What's next:**
Verify the second (mixed-byte1) worked sprite from page 8, and the byte-boundary case from `TODO.md` (`HB0 ≠ HB1` on one row). Decide how the `Hex0`/`Hex1` masked-display columns (`PDF0`/`PDF1` in the current scratch copy) should live in the shipped layout, then fold that into the sheet-layout script. After that, either continue hardening Deliverable 1 or move to Phase 2 (Sprite Inventory).

---

## 🎯 Strategic vision

An Excel workbook that makes the Apple II hi-res graphics system tangible — not as a general-purpose bitmap editor, but as a lab tightly coupled to the Lode Runner disassembly. Each sheet illuminates one layer of the graphics machinery: how pixels become bytes, how bytes become colors, how the shift tables work, how sprites land on the memory-mapped screen.

The workbook also serves as a ground-truth reference for the `papple2` emulator: the NTSC color rules, once verified visually in Excel, become test cases that drive improvements to `Display.update_hires`.

---

## 🗺️ Phased roadmap

### Phase 1 — Sprite Editor/Viewer ← we are here
Build the core worksheet: pixel grid editor, NTSC color viewer, hex/bits byte display, two-button workflow (pixels → bytes, bytes → pixels). VBA-driven, no conditional formatting.

### Phase 2 — Sprite Inventory
Add a sheet with all 104 Lode Runner sprites from `sprite_data.asm`, reorganized from interleaved layout. Selectable into the editor via dropdown + Load button. Named sprite labels from the EQU defines.

### Phase 3 — Pixel Shifter
Visualize the table-lookup chain from Chapter 3 §3.3: enter a 7-bit pattern and shift amount, watch the indirection through `PIXEL_SHIFT_PAGES`, `PIXEL_SHIFT_TABLE`, and `PIXEL_PATTERN_TABLE`. Scratchpad cells mirror 6502 register state at each step.

### Phase 4 — Sprite Shifter
Apply the Pixel Shifter to a full sprite (all 22 bytes), reproducing `COMPUTE_SHIFTED_SPRITE`. Show the OR step for the overlapping middle byte. Output loadable back into the editor.

### Phase 5 — Memory Map Viewer
Two sheets (HGR1, HGR2) with small cells showing the 8 KB graphics pages, making the non-consecutive row layout visible.

### Ongoing — `papple2` integration
Derive test fixtures from the workbook's NTSC rendering. Write failing tests, then fix `Display.update_hires` to implement the full adjacency-based color model.
