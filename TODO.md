# a2-hires-lab -- TODO

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

## Charter 

- Forward-looking only -- concrete, startable work: tasks specified well enough that next-session-me can begin within ten minutes, plus investigation items, test specs, and scratchpad ideas awaiting promotion or deletion.
- Items are unordered within their theme sections; open questions are marked _Needs investigation_ in the bullet.
- When an item is completed, record its durable outcome in `HISTORY.md` during the same session while the evidence and rationale are fresh, then strike it through in `TODO.md` with a concise handover note.
- Retain struck-through items through the next session because `TODO.md` is included in the standard filesdump while `HISTORY.md` normally is not; at the end of that next session, remove the already-archived items from `TODO.md`. Strategic direction, ordering, and milestones live in `GOALS.md` -- anything that needs a strategy discussion before it is actionable goes there.
- Architecture, contract, and settled decisions live in `README.md`.

---

## Deliverable 1: Sprite Editor/Viewer

- ~~Implement the `.xlsm` workbook per `a2-hires-lab-design.md`~~ -- done 2026-09-09, see `HISTORY.md`
- ~~Test with the two example sprites from Chapter 3 page 8 (one all-high-bit, one mixed)~~ -- both verified, second sprite confirmed 2026-09-17 against the page-9 image pasted onto the sheet, see `HISTORY.md`
- ~~Verify byte boundary color behavior: set HB0 ≠ HB1 and check that the color viewer shows the correct palette per byte~~ -- tested 2026-09-17 on a dedicated sheet; confirmed each pixel's color uses its own byte's high bit, matching real Apple II hardware's per-byte half-pixel-delay mechanism (sources in `README.md`), see `HISTORY.md`
- ~~_Needs investigation:_ `Hex0`/`Hex1` store the byte as the game actually holds it (pixels + high bit OR'd in), which doesn't match the chapter's own printed byte values~~ -- resolved: `PDF0`/`PDF1` masked-hex display columns added, see `README.md`
- ~~High-bit hide/show button~~ -- done 2026-09-17 (`Button_HideHighBitColumns`)
- `Worksheet_Change` hook, replacing the button-click workflow with live updates on edit -- still open, being handled directly in VBA

## Future deliverables

- Sprite Inventory sheet (Phase 2) — reorganize `sprite_data.asm` interleaved layout into per-sprite blocks. Core machinery done 2026-09-17: `SPRITE_DATA` reflow table, `Sprite Loader` address table, `LoadSpriteFromTable` macro -- see `HISTORY.md`. Sprite picker/dropdown UI postponed; typed sprite number (`Sprite Loader!C4`) is the interface for now.
- _Needs investigation, low priority:_ two bytes in `sprite_data.asm` (sprites 102/103, row 10 byte 2) have their high bit set in storage, unlike every other byte in the table -- colored on `SPRITE_DATA` 2026-09-17. Possibly a disassembly-reconstruction artifact rather than deliberate game data; revisit if it ever matters for round-trip export.
- Pixel Shifter sheet (Phase 3) — import `pixel_shift_table.asm` and `pixel_pattern_table.asm` onto data sheet
- Sprite Shifter sheet (Phase 4) — decide on 21-column viewer mode vs. crop-to-14. Discussed 2026-09-17 as next session's focus; check dependency on Phase 3's table-lookup scaffolding first.
- Save sprite: export editor content back to `sprite_data.asm` interleaved format, enabling round-trip editing

## `papple2` integration

- Derive test fixtures (sprite bytes + expected pixel RGB per cell) from the workbook once Deliverable 1 is verified
- Write failing tests for `Display.update_hires` adjacency logic
- ~~_Needs investigation:_ Cross-byte high-bit boundary — does the real hardware produce a visible fringe, and should `papple2` reproduce it?~~ -- yes, confirmed 2026-09-17: each byte's high bit delays that byte's own pixels by half a pixel clock, so a `HB0 != HB1` boundary produces a physical stagger, not just a clean color split (sources in `README.md`). Still open: whether/how `Display.update_hires` should model this -- not critical for `a2-hires-lab`, revisit when deriving test fixtures.

## Scratchpad / ideas

- PyXll bridge: call `papple2` color logic from Excel, compare against VBA rendering side-by-side
- Half-pixel / 560-column viewer mode for NTSC phase-shift visualization
- Emulator-driven verification: load Lode Runner in `papple2`, run `DRAW_SPRITE_PAGE1`, capture screen buffer, compare
