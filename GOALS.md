# a2-hires-lab -- Goals and Roadmap

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

**Charter:** This file answers: where is the project going, in what order, and what happens next. It holds the strategic vision, the phased roadmap, and goals that need a strategy discussion before they are actionable. The _Current Session Pointer_ below is the single canonical "where we are / what's next" -- keep it to a few lines, update it, don't grow it; `FIRST_PROMPT.md` sends the reader here first. Concrete, startable work lives in `TODO.md`; the resolved-work record lives in `HISTORY.md` or `CHANGELOG.md`(on the heap, out of the per-session dump); architecture, contract, and settled decisions live in `README.md`.

---

## 📍 Current Session Pointer

## Where we are -- 2026-09-19

- **Phase 1 (Sprite Editor/Viewer) Complete:** Built the foundational 11x14 editing grid with pixel/byte dual-representation and strict idempotent macro contracts. The full NTSC nearest-neighbor color decision table is implemented in `NTSCColor.bas` and verified pixel-for-pixel against Chapter 3's worked examples.
- **Phase 2 (Sprite Inventory) Complete:** Reflowed `sprite_data.asm` from its byte-position-major 6502 storage format (stride 104) into human-readable 22-byte sprite rows. Added the address lookup table and loader macro on `Sprite Loader`.
- **Phase 3 (Pixel Shifter) Complete:** Built the unified 7-shift `Pixel Shift Table` with a single universal formula mapping all 1,792 entries directly to `pixel_shift_table.asm`. Constructed the live 2-stage dictionary engine on `Pixel Shifter` (`PIXEL_SHIFT_PAGES` -> `pixel_shift_table.asm` -> `pixel_pattern_table.asm`) with both `TOROW` and Ariexcel-navigable 2D matrix formulas, fully documenting the 512-pattern mechanics and architecture in `README.md`.

Phase 2 (Sprite Inventory) has working sprite-table infrastructure: `sprite_data.asm` is split into individual bytes; `SPRITE_DATA` reflows the raw 154-line/2288-byte stream (six 16-byte lines then one 8-byte line, repeating) into a clean 16-column grid; `Sprite Loader` computes a selected sprite's 22 byte addresses (the "why" of the byte-position-major storage layout is in `README.md`), and a `LoadSpriteFromTable` macro pulls those bytes into whichever editor sheet is active. All verified against sprite `$01`, both visually and by independently recomputing its bytes from the raw file. `papple2` is public on https://github.com/fschuhi/papple2.

## What's Next: Phase 4 (Sprite Shifter)

Begin Phase 4 by implementing the full horizontal sprite shifting engine based on `COMPUTE_SHIFTED_SPRITE` (Chapter 3, page 10):

1. **Inventory & Disassembly Analysis:**
   - Catalog the routine from `main.nw` that accepts a sprite index (`0..103`) and shift amount (`0..6`) and populates the 33-byte `BLOCK_DATA` area (11 rows × 3 bytes).
   - Document the middle-byte bitwise OR step where shifted Byte 0 overflow merges with the beginning of shifted Byte 1.

2. **Workbook Architecture (`Sprite Shifter` Sheet):**
   - Implement a dedicated `Sprite Shifter` sheet featuring:
     - Sprite selection (via index) and Shift input (`0..6`).
     - A structured `BLOCK_DATA` memory staging range ($11 \times 3$ raw bytes).
     - A 21-pixel wide Color Viewer (3 screen bytes × 7 dots, 11 rows high) utilizing `NTSCColor.bas` to visualize the sprite smoothly crossing byte boundaries.
     - A VBA macro or live formulas linking the Phase 3 shift machinery directly to the 11 rows of sprite data.

---

## 🎯 Strategic vision

An Excel workbook that makes the Apple II hi-res graphics system tangible — not as a general-purpose bitmap editor, but as a lab tightly coupled to the Lode Runner disassembly. Each sheet illuminates one layer of the graphics machinery: how pixels become bytes, how bytes become colors, how the shift tables work, how sprites land on the memory-mapped screen.

The workbook also serves as a ground-truth reference for the `papple2` emulator: the NTSC color rules, once verified visually in Excel, become test cases that drive improvements to `Display.update_hires`.

---

## 🗺️ Phased roadmap

### Phase 1 — Sprite Editor/Viewer -- done 2026-09-17
Build the core worksheet: pixel grid editor, NTSC color viewer, hex/bits byte display, two-button workflow (pixels → bytes, bytes → pixels). VBA-driven, no conditional formatting.

### Phase 2 — Sprite Inventory ← we are here
Add a sheet with all 104 Lode Runner sprites from `sprite_data.asm`, reorganized from interleaved layout. Selectable into the editor via dropdown + Load button. Named sprite labels from the EQU defines. Core machinery (reflow table, address table, load macro) done 2026-09-17; dropdown/picker UI postponed.

### Phase 3 — Pixel Shifter
Visualize the table-lookup chain from Chapter 3 §3.3: enter a 7-bit pattern and shift amount, watch the indirection through `PIXEL_SHIFT_PAGES`, `PIXEL_SHIFT_TABLE`, and `PIXEL_PATTERN_TABLE`. Scratchpad cells mirror 6502 register state at each step.

### Phase 4 — Sprite Shifter (next session, per 2026-09-17 discussion -- order vs. Phase 3 TBD)
Apply the Pixel Shifter to a full sprite (all 22 bytes), reproducing `COMPUTE_SHIFTED_SPRITE`. Show the OR step for the overlapping middle byte. Output loadable back into the editor.

### Phase 5 — Memory Map Viewer
Two sheets (HGR1, HGR2) with small cells showing the 8 KB graphics pages, making the non-consecutive row layout visible.

### Ongoing — `papple2` integration
Derive test fixtures from the workbook's NTSC rendering. Write failing tests, then fix `Display.update_hires` to implement the full adjacency-based color model.
