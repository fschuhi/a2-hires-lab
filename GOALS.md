# a2-hires-lab -- Goals and Roadmap

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

**Charter:** This file answers: where is the project going, in what order, and what happens next. It holds the strategic vision, the phased roadmap, and goals that need a strategy discussion before they are actionable. The _Current Session Pointer_ below is the single canonical "where we are / what's next" -- keep it to a few lines, update it, don't grow it; `FIRST_PROMPT.md` sends the reader here first. Concrete, startable work lives in `TODO.md`; the resolved-work record lives in `HISTORY.md` or `CHANGELOG.md`(on the heap, out of the per-session dump); architecture, contract, and settled decisions live in `README.md`.

---

## 📍 Current Session Pointer

**Where we are:** -- 2026-09-17
Deliverable 1 (Sprite Editor/Viewer) is complete and hardened: both worked sprites from Chapter 3 page 8 are verified (the second, mixed-byte1 sprite this session, against the page-9 image pasted onto the sheet), the byte-boundary case (`HB0 != HB1`) has been tested on a dedicated sheet and confirmed to match real Apple II hardware behavior -- each byte's high bit governs only its own pixels' color, which the code already did correctly (NTSC color naming/RGB values and the hardware sourcing behind this are now in `README.md`). High-bit columns can be hidden via a button (`Button_HideHighBitColumns`).

Phase 2 (Sprite Inventory) has working sprite-table infrastructure: `sprite_data.asm` is split into individual bytes; `SPRITE_DATA` reflows the raw 154-line/2288-byte stream (six 16-byte lines then one 8-byte line, repeating) into a clean 16-column grid; `Sprite Loader` computes a selected sprite's 22 byte addresses (the "why" of the byte-position-major storage layout is in `README.md`), and a `LoadSpriteFromTable` macro pulls those bytes into whichever editor sheet is active. All verified against sprite `$01`, both visually and by independently recomputing its bytes from the raw file. `papple2` is public on https://github.com/fschuhi/papple2.

**What's next:**
Sprite picker/dropdown for Phase 2 postponed -- typed sprite number in `Sprite Loader!C4` is the interface for now. Next session: Pixel Shifter (Phase 3), as the precursor to the Sprite Shifter (Phase 4). For the beginning of the conversation, Let's work through the `PIXEL_SHIFTER_PREP.md` together and think about how to change existing sheets or add new ones that help me understand how things hang together. Suggestion: We add a new sheet to reorganize the `pixel_shift_table.asm`, with live links to the hex values of that sheet, so that the new sheet shows the map (key = index, value1 = low byte, value 2 = high byte). Feedback welcome!

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
