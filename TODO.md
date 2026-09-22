# a2-hires-lab -- TODO

(Note: "I" in the following paragraphs refer to the user, "you" to you as the AI model.)

## Charter 

- Forward-looking only -- concrete, startable work: tasks specified well enough that next-session-me can begin within ten minutes, plus investigation items, test specs, and scratchpad ideas awaiting promotion or deletion.
- Items are unordered within their theme sections; open questions are marked _Needs investigation_ in the bullet.
- When an item is completed, record its durable outcome in `HISTORY.md` during the same session while the evidence and rationale are fresh, then strike it through in `TODO.md` with a concise handover note.
- Retain struck-through items through the next session because `TODO.md` is included in the standard filesdump while `HISTORY.md` normally is not; at the end of that next session, remove the already-archived items from `TODO.md`. Strategic direction, ordering, and milestones live in `GOALS.md` -- anything that needs a strategy discussion before it is actionable goes there.
- Architecture, contract, and settled decisions live in `README.md`.

---

## Workbook

- Sprite placement, half-screen columns: the game can't hold a pixel column 0-279 in one byte, so it works with half columns (0-139, one per double pixel) and turns them into byte and shift with `HALF_SCREEN_COL_BYTE_TABLE` / `HALF_SCREEN_COL_SHIFT_TABLE` (Chapter 3, near `GET_SCREEN_COORDS_FOR`). `PlaceShiftedSprite` uses full-resolution x with `\ 7` and `Mod 7` instead. Reproduce the game's way.
- Sprite placement, masked merge: `PlaceShiftedSprite` overwrites the three bytes. The game merges them into the screen bytes with `PIXEL_MASK0` / `PIXEL_MASK1` in `DRAW_SPRITE`.
- Sprite placement, erasing: nothing removes a placed sprite except `Button_ClearScreen`.
- Hires sheets, small cleanups from the 2026-09-22 review: the `B` columns on `Hires Pixels` / `Hires HB` are leftover copies of the skip flag from `Hires Memory` and unused; `'row_to_offset_hi_table.asm'!A2` says `row_to_offset_lo_table.asm`; `D` means memory ordinal on `Hires Memory` but screen row on the other two; `'Sprite Shifter'!I7` is `=1` instead of a reference to `'Sprite (load)'`.
- `Hires Memory`: format `HiresMemory` as Text. Typed bytes like `33` work only because `HEX2BIN` reads the number as hex text; something like `1E2` turns into scientific notation first.
- `NTSCColor` refactoring: get rid of `FIRST_COL` and `LAST_COL` in favor of local named range access. Do this before the Sprite Shifter row painter.
- `Sprite (load)`: spin control next to the sprite number cell `'Sprite (load)'!X1`, so one can flip through the sprite inventory easily.
- Sprite Editor: `Worksheet_Change` hook that runs `LoadSpriteFromTable` when `'Sprite (load)'!X1` changes -- higher priority than the other event-based updates.
- Sprite Editor: `Worksheet_Change` hook to replace the button-click workflow with live updates on edit -- still open, being handled directly in VBA.
- Sprite Inventory: sprite picker/dropdown UI postponed; the typed sprite number in `'Sprite (load)'!X1` remains the interface for now.
- Workbook documentation: cell comments across the sheets, so that the "> comments <" navigation has something to show, plus text boxes with short explanations and pointers into `README.md`.
- _Low priority:_ Sprite Shifter row painter -- a lightweight VBA macro button that paints the 21-cell row background colors underneath the screen bitfield using `NTSCColor.RowColors`. Only after the `NTSCColor` refactoring. Seeing that the shifting works is enough for now. Note: `ScreenMemory.PaintScreen` shows one way to do it, colouring cells from bytes via `PixelColor`.
- _Low priority:_ Save sprite -- export editor content back to `sprite_data.asm` interleaved format, enabling round-trip editing.

## Tests

In the style of `tests/test_shift_tables.py`: plain `pytest`, reading the `.asm` files, no new dependencies.

- Direct table equivalence: build the 1,792-byte direct table from `pixel_shift_table.asm` and `pixel_pattern_table.asm`, then check that it yields the same two bytes as the two-stage lookup for all 896 pattern/shift combinations. This is the central claim of the "Architectural Analysis" section in `README.md`.
- Shift table invariants: every offset in `pixel_shift_table.asm` is even (which is why the game can use a plain `ADC #$01` for the second byte), and every page byte lies between `$A9` and `$AC`.
- Sprite data structure: `sprite_data.asm` holds 2,288 bytes; the interleaving stride is 104; every byte has bit 7 set except the two known exceptions below.

## Data questions

- _Needs investigation, low priority:_ two bytes in `sprite_data.asm` (sprites 102/103, row 10 byte 2) have their high bit set in storage, unlike every other byte in the table -- colored on `Sprite Data` 2026-09-17. Possibly a disassembly-reconstruction artifact rather than deliberate game data; revisit if it ever matters for round-trip export.

## Documentation

- Literate-source sync: integrate the findings into the Lode Runner literate source in `a2-lode-runner` -- mapping tables, the corrected `COMPUTE_SHIFTED_SPRITE` excerpt, the cycle comparison, the Mermaid visual mechanics diagram, the verification test, and the README chapter "Screen Memory and the Hires Screen", whose sections already point into Chapter 3. Strategic framing in `GOALS.md`.

## `papple2` integration

Low priority here: the sensible home for `papple2` work is `a2-lode-runner`, where running subroutines against the disassembly makes more sense than in a spreadsheet.

- Derive test fixtures (sprite bytes + expected pixel RGB per cell) from the workbook.
- Write failing tests for `Display.update_hires` adjacency logic.

## Scratchpad / ideas

- `probotron` sprite mechanics: bring the Robotron 2084 sprite handling from the `papple2`-based `probotron` workbench into this workbook. Low urgency, strategically important: it would turn `a2-hires-lab` into a compendium of the different ways Apple II games do graphics.
- Half-pixel / 560-column viewer mode for NTSC phase-shift visualization.
- Emulator-driven verification: load Lode Runner in `papple2`, run `DRAW_SPRITE_PAGE1`, capture screen buffer, compare.
