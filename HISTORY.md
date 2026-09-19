# a2-hires-lab -- History

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

- The resolved-work record: what was built and when (note date, or have the points in roughly reverse-chronological order).
- This is the trophy case -- kept in the repo, **out of the per-session filesdump** (so it no longer rides along every session).
- For *forward* work see `TODO.md`; for direction see `GOALS.md`; for the architecture as it stands see `README.md`.
- See "Workflow for the Whole Session (CRITICAL)" in `LLM_INSTRUCTIONS.md` for the interplay between `TODO.md` and this file. 

---

## 2026-09-19 -- Phase 4: Sprite Shifter & direct lookup model

- Built the complete horizontal sprite shifting engine on `Sprite Shifter`, linking directly to the raw 7-bit keys on `'Sprite (load)'` and reproducing the 11-row `COMPUTE_SHIFTED_SPRITE` pipeline.
- Implemented the two-stage dictionary lookups via 2D matrix arithmetic (`INT(...) + 1`, `MOD(...) + 1` across 16-byte boundaries) directly into `pixel_shift_table.asm` and `pixel_pattern_table.asm`, preserving full formula auditability.
- Modeled the 33-byte `BLOCK_DATA` staging area (11 rows × 3 bytes), implementing the exact middle-byte bitwise merge via `=BITOR(...)` to combine shifted Byte 0 overflow with shifted Byte 1 head while naturally preserving the high-bit color flag.
- Built the 21-column screen bitfield (`AA6:AY16`) with stripped high bits and LSB-first pixel reflection, visually confirming smooth horizontal sprite movement across screen byte boundaries for all shifts 0..6.
- Added the "rewritten in place" direct shift lookup model on `Pixel Shifter` via `Pixel Shift Pattern Table`, demonstrating that consolidating the two split tables into a single 1,792-byte direct lookup eliminates indirection and would save 1,024 bytes and 6502 cycles in the original game engine.®

## 2026-09-18 -- Phase 3: Pixel shifter

- Added the unified 7-shift mapping sheet `Pixel Shift Table`, laying out all 128 input patterns (keys 0..127) across 7 side-by-side shift blocks. Structured coordinates in rows 2 and 3 enabled a single universal formula across all 1,792 data cells, linking directly to the raw byte grid on `pixel_shift_table.asm`.
- Added physical address columns on `pixel_shift_table.asm` ($A200..$A8FF) and `pixel_pattern_table.asm` ($A900..$ACFF), visually anchoring the 2x 128-byte half-page split and hardware page dispatch (`PIXEL_SHIFT_PAGES`).
- Defined the named range `PixelShiftTable` on `pixel_shift_table.asm` to decouple VBA and sheet formula access from raw column shifts.
- Completed the `Pixel Shifter` sheet, implementing the complete 5-stage shift pipeline:
  1. Reversing screen pixel inputs (`0110100`) to 6502 storage bit order (`%0010110` / decimal 22) via `ReverseString`.
  2. Resolving the physical shift page via `PixelShiftPages` ($A2..$A8).
  3. Reading the split 16-bit pattern address (`Lo` from the first 128 bytes, `Hi` from the second 128 bytes).
  4. Resolving the two output bytes from the 512-entry gallery in `pixel_pattern_table.asm`.
  5. Stripping the high color bit (bit 7) and reflecting the resulting bits back into screen pixel order across columns C:P.
- Provided a dedicated second evaluation block on `Pixel Shifter` using 2D matrix arithmetic (`INT(offset / COLUMNS) + 1`, `MOD(offset, COLUMNS) + 1`) to ensure full transparency and navigability via Ariexcel and the Excel formula auditing detective.
- Integrated the full `PIXEL_SHIFTER_PREP.md` mechanics, mathematical breakdown (512 unique shapes across widths 1..7), 3-column mapping table, and Mermaid visual flow diagram directly into `README.md`.

## 2026-09-17 -- Phase 2: Sprite inventory

- Verified the second (mixed-byte1) worked sprite from Chapter 3 page 9 against the color viewer.
- Tested the byte-boundary case (`HB0 != HB1`) on a dedicated sheet; confirmed against real Apple II hardware sources that each pixel's color comes from its own byte's high bit.
- Updated `NTSCColor.bas`'s RGB values to historically accurate NTSC hues (kept the name "Violet").
- Added `Button_HideHighBitColumns`.
- Built Phase 2 core machinery: `sprite_data.asm` byte-split, `SPRITE_DATA` reflow table (flattens the raw 154-line/2288-byte stream into a clean 16-column grid), `Sprite Loader` address table (`AddrByte0`/`AddrByte1`), and `LoadSpriteFromTable` macro. Refactored `SpriteEditor.bas` to pass `ws` explicitly rather than relying on a hardcoded sheet name; added `Buttons.bas`. Verified against sprite `$01`.
- Colored the two high-bit-set bytes found in `sprite_data.asm` (sprites 102/103) on `SPRITE_DATA`.

## 2026-09-10 -- Multiple viewer/editor sheets

- We now use local ranges for the Sprite Editor/Viewer sheets so that `sub`s and buttons work on the active sheet. 
- Added Buttons on the sprite sheets.
- Experimented with high bit columns (`HB0`, `HB1`) in order to show the 11x14 sprite without interspersed `HB0` column. Decision: hide the column instead of moving them around, so that the 2x8 bits layout is consistent. 

## 2026-09-09 -- Phase 1: Sprite Editor/Viewer built and colour model fixed

Implemented the full Deliverable 1 stack from the design doc: sheet layout (`openpyxl`, with a multi-area named-range bug fixed along the way), `Util` (hex/bit helpers, plus a `HexToBits` addition so Bits0/Bits1 run as live formulas), `NTSCColor`, and `SpriteEditor` with both idempotent buttons. `NTSCColor` also picked up a generalization -- absolute screen column and true left/right neighbors instead of an assumed isolated sprite at column 0 -- and, more importantly, a real fix: a colored 0-pixel takes its hue from its flanking 1-bit's column, not its own, without which a repeating `0x55` byte rendered as alternating stripes instead of the solid fill it's meant to produce. Verified by reproducing the chapter's worked "5"-shaped sprite pixel-for-pixel. Masked-hex display columns for comparing directly against the chapter's byte values were discussed but not yet built.

## 2026-09-08 / 2026-09-09 — Project inception and design (Opus 4.6)

Set up the `a2-hires-lab` project as a standalone Excel workbook for
exploring Apple II hi-res graphics, built around the Lode Runner
disassembly data (Chapter 3 of `main.nw`).

Over two sessions we worked out the design document
(`a2-hires-lab-design.md`) through back-and-forth discussion:

- Settled on VBA macros with cell background coloring (no conditional formatting) as the rendering approach 
- Designed the Sprite Editor/Viewer sheet layout: 11×14 pixel grid, per-byte high-bit toggles, hex and bits display columns, color viewer area.
- Formalized the NTSC artifact color rules as a nearest-neighbor decision table with a color lookup, verified against the chapter's examples and Tilleul's prior art.
- Defined five deliverables (Editor/Viewer, Inventory, Pixel Shifter, Sprite Shifter, Memory Map Viewer) plus a "Further ideas" section.
- Explained the sprite data interleaving in `sprite_data.asm` as a 6502 lookup optimization (indirect-indexed addressing, no multiply).
- Identified that `Apple.py`'s `update_hires` uses a simplified per-pixel color model without adjacency — the workbook's VBA will become the ground truth for fixing it.
- Found and reviewed François Vander Linden's `bitmap_creator` spreadsheet as prior art; confirmed it validates the Excel approach but is conceptually distant (screen-fragment editor, formula-driven, no sprite or game-data awareness).
- Established naming conventions: Hungarian notation for VBA, no module prefix, `cl` prefix for classes.
- Prepared the seed package for the build session (design doc, ASM data files, `Apple.py`, Chapter 3 PDF).
