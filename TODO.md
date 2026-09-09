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

- Implement the `.xlsm` workbook per `a2-hires-lab-design.md`
- Test with the two example sprites from Chapter 3 page 8 (one all-high-bit, one mixed)
- Verify byte boundary color behavior: set HB0 ≠ HB1 and check that the color viewer shows the correct palette per byte

## Future deliverables (parked until Phase 1 is done)

- Sprite Inventory sheet (Phase 2) — reorganize `sprite_data.asm` interleaved layout into per-sprite blocks
- Pixel Shifter sheet (Phase 3) — import `pixel_shift_table.asm` and `pixel_pattern_table.asm` onto data sheet
- Sprite Shifter sheet (Phase 4) — decide on 21-column viewer mode vs. crop-to-14
- Save sprite: export editor content back to `sprite_data.asm` interleaved format, enabling round-trip editing

## `papple2` integration

- Derive test fixtures (sprite bytes + expected pixel RGB per cell) from the workbook once Deliverable 1 is verified
- Write failing tests for `Display.update_hires` adjacency logic
- _Needs investigation:_ Cross-byte high-bit boundary — does the real hardware produce a visible fringe, and should `papple2` reproduce it?

## Scratchpad / ideas

- PyXll bridge: call `papple2` color logic from Excel, compare against VBA rendering side-by-side
- Half-pixel / 560-column viewer mode for NTSC phase-shift visualization
- Emulator-driven verification: load Lode Runner in `papple2`, run `DRAW_SPRITE_PAGE1`, capture screen buffer, compare
