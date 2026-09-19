# a2-hires-lab -- TODO

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

## Charter 

- Forward-looking only -- concrete, startable work: tasks specified well enough that next-session-me can begin within ten minutes, plus investigation items, test specs, and scratchpad ideas awaiting promotion or deletion.
- Items are unordered within their theme sections; open questions are marked _Needs investigation_ in the bullet.
- When an item is completed, record its durable outcome in `HISTORY.md` during the same session while the evidence and rationale are fresh, then strike it through in `TODO.md` with a concise handover note.
- Retain struck-through items through the next session because `TODO.md` is included in the standard filesdump while `HISTORY.md` normally is not; at the end of that next session, remove the already-archived items from `TODO.md`. Strategic direction, ordering, and milestones live in `GOALS.md` -- anything that needs a strategy discussion before it is actionable goes there.
- Architecture, contract, and settled decisions live in `README.md`.

---

## Deliverable Enhancements

- Sprite Editor enhancements: `Worksheet_Change` hook to replace the button-click workflow with live updates on edit -- still open, being handled directly in VBA.
- Sprite Inventory enhancements (Phase 2): sprite picker/dropdown UI postponed; typed sprite number (`Sprite Loader!C4`) remains the interface for now.
- Sprite Shifter visual painter (Phase 4): add a lightweight VBA macro button to paint the 21-cell row background colors underneath the screen bitfield using `NTSCColor.RowColors`.
- _Needs investigation, low priority:_ two bytes in `sprite_data.asm` (sprites 102/103, row 10 byte 2) have their high bit set in storage, unlike every other byte in the table -- colored on `SPRITE_DATA` 2026-09-17. Possibly a disassembly-reconstruction artifact rather than deliberate game data; revisit if it ever matters for round-trip export.
- ~~Sprite Shifter sheet (Phase 4) -- decide on 21-column viewer mode vs. crop-to-14. Apply shift mechanics across all 22 sprite bytes to reproduce `COMPUTE_SHIFTED_SPRITE` and show the middle-byte OR step.~~ -- done 2026-09-19: built live formula engine on `Sprite Shifter` linking directly to `'Sprite (load)'` and 2D matrix lookups into `pixel_shift_table.asm` and `pixel_pattern_table.asm`; implemented middle-byte `BITOR` merge and 21-column screen bitfield across all 11 rows.
- Save sprite: export editor content back to `sprite_data.asm` interleaved format, enabling round-trip editing.

## Immediate Next Steps (Prior to Phase 5)

- Workbench polishing: clean up formatting, range labels, and sheet navigation across `Sprite Shifter` and `Pixel Shifter`.
- Shift lookup analysis & documentation: document why the two-stage dictionary (`PIXEL_SHIFT_TABLE` -> `PIXEL_PATTERN_TABLE`) could be consolidated into a single direct 1,792-byte lookup table without indirection, saving 1,024 bytes and 6502 cycles. Add architecture section to `README.md`.
- Literate-source sync: integrate findings, mapping tables, and Mermaid visual mechanics into the working Lode Runner literate source in `load-runner`.

## `papple2` integration

- Derive test fixtures (sprite bytes + expected pixel RGB per cell) from the workbook once Deliverable 1 is verified.
- Write failing tests for `Display.update_hires` adjacency logic.

## Scratchpad / ideas

- PyXll bridge: call `papple2` color logic from Excel, compare against VBA rendering side-by-side.
- Half-pixel / 560-column viewer mode for NTSC phase-shift visualization.
- Emulator-driven verification: load Lode Runner in `papple2`, run `DRAW_SPRITE_PAGE1`, capture screen buffer, compare.
