# a2-hires-lab -- History

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

- The resolved-work record: what was built and when (note date, or have the points in roughly reverse-chronological order).
- This is the trophy case -- kept in the repo, **out of the per-session filesdump** (so it no longer rides along every session).
- For *forward* work see `TODO.md`; for direction see `GOALS.md`; for the architecture as it stands see `README.md`.
- See "Workflow for the Whole Session (CRITICAL)" in `LLM_INSTRUCTIONS.md` for the interplay between `TODO.md` and this file. 

---

## 2026-09-22 -- Phase 5: screen memory, sprite placement, NTSC screen

Closed the chain from `BLOCK_DATA` to the screen. A shifted Lode Runner sprite can now be placed anywhere on HGR page 1 and shows up in NTSC colours. Phase 5 was built directly instead of going through the planned design session.

**The hires sheets.** `Hires Memory` holds HGR page 1 as 192 lines of 40 bytes in address order, stepping over the 8 unused bytes at the end of each 128-byte block. `Hires Pixels` and `Hires HB` show the same bytes in screen order, as 7-pixel strings and as high bits; their addresses come from `ROW_TO_OFFSET_LO` / `ROW_TO_OFFSET_HI` OR `$20`, and `MATCH` finds each line in `Hires Memory`. All 192 addresses were checked against the standard formula `$2000 + (y mod 8) * $400 + ((y div 8) mod 8) * $80 + (y div 64) * $28`; they all agree.

**`ScreenMemory.bas`.** `PlaceShiftedSprite(row, col)` splits the pixel column into byte (`x div 7`) and shift (`x mod 7`), writes the shift into `Shift` so `Sprite Shifter` recomputes `BlockData`, and writes the 11 x 3 bytes into `Hires Memory`, cut off at the right and at the bottom. `PaintScreen` is the video circuit: it reads lines from memory and colours `Hires Screen` with `NTSCColor.PixelColor`, knowing nothing about sprites. After placing a sprite, the area it wrote plus one pixel on each side is repainted, because neighbouring pixels can change colour. `ClearScreen` works like `HGR`: zero page 1, paint the screen. `LineAddress` gives the address of a screen row. Black is shown as light grey `RGB(191,191,191)`. Buttons: `Button_PlaceShiftedSprite` (selected cell = sprite's top-left pixel) and `Button_ClearScreen`, replacing `Button_ClearMemory`.

**Two screen sheets.** `Hires Screen (debug)` keeps the formulas and shows which bits are set; `Hires Screen` has no formulas and is painted from memory. Same rows and columns, side by side, so switching between them shows bits versus colours at the same cells -- including coloured pixels with no bit set.

**Simplifications, now in `TODO.md`.** Full-resolution columns instead of the game's half columns (`HALF_SCREEN_COL_BYTE_TABLE` / `HALF_SCREEN_COL_SHIFT_TABLE`, needed because 280 does not fit in a byte), no masked merge with `PIXEL_MASK0` / `PIXEL_MASK1`, no erasing.

**Documentation.** New README chapter "Screen Memory and the Hires Screen": the hires sheets with a Mermaid data-flow chart, the memory map, placing a shifted sprite, painting as the video circuit, and how to try it out. Sections open with pointers into Chapter 3 of `main.nw`, so they can become annotations in `a2-lode-runner`. Two new images: `img/player_orange_blue.jpg` (column parity swaps blue and orange) and `img/player_edge_color.jpg` (colours change where two sprites meet).

**Tooling.** Changes from LLM sessions now arrive as patches applied with `make patch` (documented in the README together with `make help`). The new `Modules` module exports the VBA before patching, so the patched files can be imported back; `ExportProjectModules` was missing `Buttons`, which had left `src/bas/Buttons.bas` stale. `SetSilentApplicationState` / `RevertApplicationState` from `B_.bas` speed up sprite loading and placement. New named ranges: `BlockData`, `Shift`, `HiresScreen`, `HiresMemory`.

## 2026-09-20 -- Public release: README overhaul, licensing, VBA export, first tests

Repo made public. Session spent reviewing the whole project for how it reads from the outside, rather than on Phase 5.

**README rewritten for visitors.** Reordered so the first screens answer "what is this and what can I try": "What this is" with a one-line hook at the direct-table finding, a "What's in the workbook" table of sheets replacing the phase-based status table, "Design choices", and "Running" moved up and opened with two new paragraphs (Excel version requirements, macro security and the Unblock step on Windows). Added a "Project Structure" tree and shortened the table of contents. Phase numbers left the README; they live in `GOALS.md` now.

**6502 documentation corrected.** The assembly snippets written by Gemini were wrong: `(TMP_PTR+128),Y` is not valid 6502, and "two single-cycle indexed loads" mis-stated the cost (`LDA abs,Y` is 4 cycles). Replaced with an excerpt of the real `COMPUTE_SHIFTED_SPRITE`, which patches the shift page into its own lookup instructions, plus a plain-language explanation of that self-modifying code. Noted that every shift-table offset is even, which is why the routine can use a bare `ADC #$01` for the second byte. Also fixed the Mermaid diagram's worked example (`$A95A` holds `$B0 $81`, not `$94 $82`), dropped the invented "at 60 Hz" claim, and removed the collision-detection hypothesis, which nothing supports.

**Direct-table saving quantified and verified.** Beyond the 1,024 bytes, the single 1,792-byte table would cut `COMPUTE_SHIFTED_SPRITE` from about 1,824 to about 1,208 cycles per call (56 per row, 616 per sprite: roughly a third), which matters because every sprite draw and erase calls it. Independently checked all 896 pattern/shift combinations through both tables against the arithmetic: no mismatch, all 512 pattern entries used, all distinct. The chapter's cross-reference shows `PIXEL_PATTERN_TABLE` has no other consumer, so collapsing the two stages is safe.

**Licensing split.** Code (VBA, formulas, Python tools) is MIT via `LICENSE`; documentation and the Xekri-derived data stay CC BY-SA 4.0 via `LICENSE-CC-BY-SA-4.0.md`. Reasoning: ShareAlike is inherited only where his material is adapted, MIT matches the intent that the code be reusable, and the technical ideas are free regardless since copyright covers wording, not facts. Added two-line SPDX headers to the VBA modules and a license line on `Intro`.

**VBA export tool.** `tools/export_vba.py` plus `make export-vba`, wired into both filesdump targets. The workbook is the source of truth; `src/bas/` is a one-way export for reading on GitHub and for LLM sessions. Written because the hand-maintained copies had already drifted.

**First tests.** `tests/test_shift_tables.py` (10 tests) re-runs the table verification from the `.asm` files, so the README's claims can be reproduced with `make test` -- which previously failed, because `pytest` found no tests at all.

**Workbench became a real workbench.** `Intro` sheet with feature list and section hide/show buttons, `Versions` sheet, jump-station navigation, `Worksheets matrix`, data-flow boxes and labels on the complex sheets, backed by a subset of my general Excel library (`B_`, `JumpStation_`, `WorksheetsMatrix_`).

**Cleanups.** Removed leftover template artifacts (two stray UserForms, the broken `AllTags` name, four `GET.CELL` XLM names, German LAMBDAs), the deprecated `PixelShifter.bas`, the dead `SHEET_NAME` constant, and `Apple.py`. Renamed `SPRITE_DATA` to `Sprite Data` and removed the VBA constant that would have broken on the rename. Deleted the obsolete `a2-hires-lab-design.md`. Untracked the stencil-managed collaboration files. Trimmed requirements (`openpyxl` and `pyxll` out, `oletools` in) and restructured `TODO.md` by theme.

**Decisions taken.** `papple2` work belongs in `a2-lode-runner`, where running subroutines against the disassembly makes sense, not in a spreadsheet. PyXLL dropped: native Python is the better direction, and Excel stays prototypal on purpose. LLM collaboration is stated plainly in the README rather than hidden. `probotron`'s Robotron 2084 sprite mechanics noted as the direction that would make this a compendium of Apple II graphics techniques.

## 2026-09-19 -- Phase 4: Sprite Shifter & direct lookup model

- Built the complete horizontal sprite shifting engine on `Sprite Shifter`, linking directly to the raw 7-bit keys on `'Sprite (load)'` and reproducing the 11-row `COMPUTE_SHIFTED_SPRITE` pipeline.
- Implemented the two-stage dictionary lookups via 2D matrix arithmetic (`INT(...) + 1`, `MOD(...) + 1` across 16-byte boundaries) directly into `pixel_shift_table.asm` and `pixel_pattern_table.asm`, preserving full formula auditability.
- Modeled the 33-byte `BLOCK_DATA` staging area (11 rows × 3 bytes), implementing the exact middle-byte bitwise merge via `=BITOR(...)` to combine shifted Byte 0 overflow with shifted Byte 1 head while naturally preserving the high-bit color flag.
- Built the 21-column screen bitfield (`AA6:AY16`) with stripped high bits and LSB-first pixel reflection, visually confirming smooth horizontal sprite movement across screen byte boundaries for all shifts 0..6.
- Added the "rewritten in place" direct shift lookup model on `Pixel Shifter` via `Pixel Shift Pattern Table`, demonstrating that consolidating the two split tables into a single 1,792-byte direct lookup eliminates indirection and would save 1,024 bytes and 6502 cycles in the original game engine.

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
- Provided a dedicated second evaluation block on `Pixel Shifter` using 2D matrix arithmetic (`INT(offset / COLUMNS) + 1`, `MOD(offset, COLUMNS) + 1`) to ensure full transparency and navigability via Arixcel and the Excel formula auditing detective.
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
