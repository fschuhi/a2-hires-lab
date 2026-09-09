# a2-hires-lab -- History

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

- The resolved-work record: what was built and when (note date, or have the points in roughly reverse-chronological order).
- This is the trophy case -- kept in the repo, **out of the per-session filesdump** (so it no longer rides along every session).
- For *forward* work see `TODO.md`; for direction see `GOALS.md`; for the architecture as it stands see `README.md`.
- See "Workflow for the Whole Session (CRITICAL)" in `LLM_INSTRUCTIONS.md` for the interplay between `TODO.md` and this file. 

---

## 2026-09-08 / 2026-09-09 — Project inception and design (Opus 4.6)

Set up the `a2-hires-lab` project as a standalone Excel workbook for
exploring Apple II hi-res graphics, built around the Lode Runner
disassembly data (Chapter 3 of `main.nw`).

Over two sessions we worked out the design document
(`a2-hires-lab-design.md`) through back-and-forth discussion:

- Settled on VBA macros with cell background coloring (no conditional
  formatting) as the rendering approach.
- Designed the Sprite Editor/Viewer sheet layout: 11×14 pixel grid,
  per-byte high-bit toggles, hex and bits display columns, color
  viewer area.
- Formalized the NTSC artifact color rules as a nearest-neighbor
  decision table with a color lookup, verified against the chapter's
  examples and Tilleul's prior art.
- Defined five deliverables (Editor/Viewer, Inventory, Pixel Shifter,
  Sprite Shifter, Memory Map Viewer) plus a "Further ideas" section.
- Explained the sprite data interleaving in `sprite_data.asm` as a
  6502 lookup optimization (indirect-indexed addressing, no multiply).
- Identified that `Apple.py`'s `update_hires` uses a simplified
  per-pixel color model without adjacency — the workbook's VBA will
  become the ground truth for fixing it.
- Found and reviewed François Vander Linden's `bitmap_creator`
  spreadsheet as prior art; confirmed it validates the Excel approach
  but is conceptually distant (screen-fragment editor, formula-driven,
  no sprite or game-data awareness).
- Established naming conventions: Hungarian notation for VBA, no
  module prefix, `cl` prefix for classes.
- Prepared the seed package for the build session (design doc, ASM
  data files, `Apple.py`, Chapter 3 PDF).
