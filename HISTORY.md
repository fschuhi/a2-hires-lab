# a2-hires-lab -- History

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

- The resolved-work record: what was built and when (note date, or have the points in roughly reverse-chronological order).
- This is the trophy case -- kept in the repo, **out of the per-session filesdump** (so it no longer rides along every session).
- For *forward* work see `TODO.md`; for direction see `GOALS.md`; for the architecture as it stands see `README.md`.
- See "Workflow for the Whole Session (CRITICAL)" in `LLM_INSTRUCTIONS.md` for the interplay between `TODO.md` and this file. 

---

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
