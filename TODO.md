# a2-hires-lab -- TODO

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

## Charter 

- Forward-looking only -- concrete, startable work: tasks specified well enough that next-session-me can begin within ten minutes, plus investigation items, test specs, and scratchpad ideas awaiting promotion or deletion.
- Items are unordered within their theme sections; open questions are marked _Needs investigation_ in the bullet.
- When an item is completed, record its durable outcome in `HISTORY.md` during the same session while the evidence and rationale are fresh, then strike it through in `TODO.md` with a concise handover note[cite: 2].
- Retain struck-through items through the next session because `TODO.md` is included in the standard filesdump while `HISTORY.md` normally is not; at the end of that next session, remove the already-archived items from `TODO.md`[cite: 2]. Strategic direction, ordering, and milestones live in `GOALS.md` -- anything that needs a strategy discussion before it is actionable goes there[cite: 2].
- Architecture, contract, and settled decisions live in `README.md`[cite: 2].

---

## Deliverable Enhancements

- Sprite Editor enhancements: `Worksheet_Change` hook to replace the button-click workflow with live updates on edit -- still open, being handled directly in VBA[cite: 2].
- Sprite Inventory enhancements (Phase 2): sprite picker/dropdown UI postponed; typed sprite number (`Sprite Loader!C4`) remains the interface for now[cite: 2].
- _Needs investigation, low priority:_ two bytes in `sprite_data.asm` (sprites 102/103, row 10 byte 2) have their high bit set in storage, unlike every other byte in the table -- colored on `SPRITE_DATA` 2026-09-17[cite: 2]. Possibly a disassembly-reconstruction artifact rather than deliberate game data; revisit if it ever matters for round-trip export[cite: 2].
- ~~Pixel Shifter sheet (Phase 3) -- import `pixel_shift_table.asm` and `pixel_pattern_table.asm` onto data sheet~~ -- done 2026-09-18: built unified 7-shift `Pixel Shift Table` with universal formula linking 1,792 entries; implemented full 2-stage dictionary engine (`Pixel Shift Pages` -> `pixel_shift_table.asm` -> `pixel_pattern_table.asm`) on `Pixel Shifter` with both `TOROW` and Ariexcel-friendly 2D matrix formulas; verified across shifts 0..6[cite: 5, 12]; architecture added to `README.md`[cite: 2].
- Sprite Shifter sheet (Phase 4) -- decide on 21-column viewer mode vs. crop-to-14[cite: 2]. Apply shift mechanics across all 22 sprite bytes to reproduce `COMPUTE_SHIFTED_SPRITE` and show the middle-byte OR step[cite: 2].
- Save sprite: export editor content back to `sprite_data.asm` interleaved format, enabling round-trip editing[cite: 2].

## `papple2` integration

- Derive test fixtures (sprite bytes + expected pixel RGB per cell) from the workbook once Deliverable 1 is verified[cite: 2].
- Write failing tests for `Display.update_hires` adjacency logic[cite: 2].

## Scratchpad / ideas

- PyXll bridge: call `papple2` color logic from Excel, compare against VBA rendering side-by-side[cite: 2].
- Half-pixel / 560-column viewer mode for NTSC phase-shift visualization[cite: 2].
- Emulator-driven verification: load Lode Runner in `papple2`, run `DRAW_SPRITE_PAGE1`, capture screen buffer, compare[cite: 2].
