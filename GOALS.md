# a2-hires-lab -- Goals and Roadmap

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

**Charter:** This file answers: where is the project going, in what order, and what happens next. It holds the strategic vision, the phased roadmap, and goals that need a strategy discussion before they are actionable. The _Current Session Pointer_ below is the single canonical "where we are / what's next" -- keep it to a few lines, update it, don't grow it; `FIRST_PROMPT.md` sends the reader here first. Concrete, startable work lives in `TODO.md`; the resolved-work record lives in `HISTORY.md` (on the heap, out of the per-session dump); architecture, contract, and settled decisions live in `README.md`.

---

## 📍 Current Session Pointer

## Where we are -- 2026-09-20

Phases 1-4 are complete and the repo is public. The 2026-09-20 session reviewed the whole project from a visitor's point of view: `README.md` rewritten and its 6502 documentation corrected and quantified, licensing split (MIT for code, CC BY-SA 4.0 for docs and Xekri-derived data), `make export-vba`, and the first tests. Details in `HISTORY.md`.

## What's next

1. **Phase 5 design session.** Phase 5 is underdetermined; it needs a discussion that ends in a design document under `docs/`, not a coding session. See the roadmap below.
2. **Then the literate-source sync into `a2-lode-runner`.** The markdown -> HTML pipeline gets sorted out first, outside this project.

---

## 🎯 Strategic vision

An Excel workbook that makes the Apple II hi-res graphics system tangible — not as a general-purpose bitmap editor, but as a lab tightly coupled to the Lode Runner disassembly. Each sheet illuminates one layer of the graphics machinery: how pixels become bytes, how bytes become colors, how the shift tables work, how sprites land on the memory-mapped screen.

Excel stays prototypal on purpose. When something needs to run rather than be inspected, it belongs in native Python (`papple2`) or in the disassembly project (`a2-lode-runner`), not bolted onto the workbook.

The second audience is the disassembly itself. Findings that started here -- the corrected shift-lookup documentation, the cycle comparison, the verification tests -- flow back into `a2-lode-runner`'s literate source, which is also the path towards taking that project public.

**Open question:** is `a2-hires-lab` a Lode Runner lab or an Apple II graphics lab? Adding `probotron`'s Robotron 2084 sprite mechanics (see `TODO.md`) only makes sense under the second reading, and it would turn the workbook into a compendium of the different ways Apple II games do graphics. The answer shapes the README's first sentence, the sheet structure, and how the data folders are laid out. It does not need deciding yet, but it should not decide itself by accident.

---

## 🗺️ Phased roadmap

### Phase 1 — Sprite Editor/Viewer -- done 2026-09-17
Core worksheet: pixel grid editor, NTSC color viewer, hex/bits byte display, two-button workflow (pixels → bytes, bytes → pixels). VBA-driven, no conditional formatting.

### Phase 2 — Sprite Inventory -- done 2026-09-17
All 104 Lode Runner sprites from `sprite_data.asm`, reorganized from the interleaved layout, loadable into the editor. Dropdown/picker UI postponed; the typed sprite number remains the interface (see `TODO.md`).

### Phase 3 — Pixel Shifter -- done 2026-09-18
The table-lookup chain from Chapter 3 §3.3: enter a 7-bit pattern and a shift amount, follow the indirection through `PIXEL_SHIFT_PAGES`, `PIXEL_SHIFT_TABLE` and `PIXEL_PATTERN_TABLE`, in three equivalent formula styles.

### Phase 4 — Sprite Shifter -- done 2026-09-19
The Pixel Shifter applied to a whole sprite (all 22 bytes), reproducing `COMPUTE_SHIFTED_SPRITE` including the `BITOR` merge of the overlapping middle byte, plus the direct-lookup model that led to the architectural analysis in `README.md`.

### Phase 5 — Closing the chain: from shifted bytes to the screen (needs a design session)
The workbook currently stops at three shifted bytes in `BLOCK_DATA`; nothing shows them reaching the screen. Chapter 3 §3.4 and §3.5 hold the missing machinery. Two halves, 5a first because it makes 5b possible:

- **5a — Memory map.** All 192 screen rows resolved to addresses via `ROW_TO_OFFSET_LO` / `ROW_TO_OFFSET_HI`, for both HGR pages, making the non-consecutive row layout visible, including where the status line sits.
- **5b — Sprite placement.** Game row and column to screen row, byte offset and shift amount (`GET_BYTE_AND_SHIFT_FOR_COL`), then the masked merge of the three `BLOCK_DATA` bytes into the screen bytes with `PIXEL_MASK0` / `PIXEL_MASK1`, the way `DRAW_SPRITE` does it.

Scope, sheet layout and how much of `DRAW_SPRITE` to reproduce are open. The design session decides them and produces a design document under `docs/`.

### Ongoing — Literate-source sync into `a2-lode-runner`
Carry the findings into the Lode Runner literate source: mapping tables, the corrected `COMPUTE_SHIFTED_SPRITE` documentation, the cycle comparison, the Mermaid diagram, the verification tests. Sequenced after Phase 5, so that the whole graphics pipeline travels at once rather than a story that stops halfway. Counter-argument worth keeping in view: Xekri's Ultima work and his markdown/HTML pipeline are current, and goodwill has a half-life -- if that window matters more than completeness, this moves ahead of Phase 5.

### Ongoing — `papple2`
Low priority within this project. Test fixtures derived from the workbook and the adjacency work on `Display.update_hires` are better done in `a2-lode-runner`, where running subroutines against the disassembly makes sense.
