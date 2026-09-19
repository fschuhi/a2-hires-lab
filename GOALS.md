# a2-hires-lab -- Goals and Roadmap

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

**Charter:** This file answers: where is the project going, in what order, and what happens next. It holds the strategic vision, the phased roadmap, and goals that need a strategy discussion before they are actionable. The _Current Session Pointer_ below is the single canonical "where we are / what's next" -- keep it to a few lines, update it, don't grow it; `FIRST_PROMPT.md` sends the reader here first. Concrete, startable work lives in `TODO.md`; the resolved-work record lives in `HISTORY.md` or `CHANGELOG.md`(on the heap, out of the per-session dump); architecture, contract, and settled decisions live in `README.md`.

---

## 📍 Current Session Pointer

## Where we are -- 2026-09-19

- **Phase 1 (Sprite Editor/Viewer) Complete:** Built the foundational 11x14 editing grid with pixel/byte dual-representation and strict idempotent macro contracts. The full NTSC nearest-neighbor color decision table is implemented in `NTSCColor.bas` and verified pixel-for-pixel against Chapter 3's worked examples.
- **Phase 2 (Sprite Inventory) Complete:** Reflowed `sprite_data.asm` from its byte-position-major 6502 storage format (stride 104) into human-readable 22-byte sprite rows. Added the address lookup table and loader macro on `Sprite Loader`.
- **Phase 3 (Pixel Shifter) Complete:** Built the unified 7-shift `Pixel Shift Table` with a single universal formula mapping all 1,792 entries directly to `pixel_shift_table.asm`. Constructed the live 2-stage dictionary engine on `Pixel Shifter` (`PIXEL_SHIFT_PAGES` -> `pixel_shift_table.asm` -> `pixel_pattern_table.asm`) with both `TOROW` and Ariexcel-navigable 2D matrix formulas, fully documenting the 512-pattern mechanics and architecture in `README.md`.
- **Phase 4 (Sprite Shifter) Complete:** Built live formula engine on `Sprite Shifter` linking directly to `'Sprite (load)'` and 2D matrix lookups into `pixel_shift_table.asm` and `pixel_pattern_table.asm`. Implemented middle-byte `BITOR` merge and 21-column screen bitfield across all 11 rows. Added direct "rewritten in place" lookup model on `Pixel Shifter` bypassing indirection.

## What's Next: Polish, Shift Lookup Analysis & Literate-Source Sync (Prior to Phase 5)

1. **Shift Table Optimization Architecture (`README.md`):**
   - Document why the two-stage dictionary (`PIXEL_SHIFT_TABLE` -> `PIXEL_PATTERN_TABLE`) could be consolidated into a single direct 1,792-byte lookup table without indirection, saving 1,024 bytes of table storage and eliminating 6502 cycle overhead.

2. **Workbench Polish:**
   - Clean up formatting, labels, and auditing navigation on `Sprite Shifter` and `Pixel Shifter`.

3. **Literate-Source Sync (`load-runner`):**
   - Transfer key findings, lookup breakdown tables, and Mermaid visual mechanics flows into the working Lode Runner literate source.

4. **Phase 5 (Deferred):**
   - Memory Map Viewer (HGR1/HGR2 page visualization).

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
