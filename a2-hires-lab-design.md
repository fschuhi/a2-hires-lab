# a2-hires-lab -- Design Document

An Excel workbook for exploring Apple II hi-res graphics mechanics, built around the data and routines documented in the Lode Runner disassembly (Chapter 3 of `main.nw`).

The workbook uses VBA macros to render colored cells and compute byte ↔ pixel conversions. No conditional formatting; all visual updates run through explicit VBA code triggered by button clicks (later: worksheet change events).

The project is standalone from the `lode-runner` disassembly project but draws on its data and documentation.


## Deliverables roadmap

| #  | Sheet(s)              | Status    | Description                                                  |
|----|-----------------------|-----------|--------------------------------------------------------------|
| 1  | Sprite Editor/Viewer  | **first** | Enter pixels, see colors, derive bytes -- or load from inventory |
| 2  | Sprite Inventory      | planned   | All 104 sprites from `sprite_data.asm`, selectable into the editor |
| 3  | Pixel Shifter         | planned   | 7-bit pattern × shift amount → two result bytes via table lookups |
| 4  | Sprite Shifter        | planned   | Full `COMPUTE_SHIFTED_SPRITE` for 11-row sprite, output 3 bytes/row |
| 5  | Memory Map Viewer     | idea      | Two sheets with small cells showing HGR page pixel/color state |

Each deliverable builds on the previous. The Sprite Editor/Viewer is the foundation; the others add sheets and VBA modules incrementally.


## Relation to `papple2`

The VBA color logic implements the NTSC artifact color rules from the chapter (page 7). The same rules -- once verified visually in the workbook -- become test cases for `papple2`'s `Display.update_hires`, which currently uses a simplified per-pixel model without adjacency.

Workflow: build correct color rendering in VBA → derive test fixtures (sprite bytes + expected pixel colors) → write failing tests for `papple2` → fix `update_hires` to pass them.

The PyXll bridge (calling Python from Excel) is a possible future direction but out of scope for now.


## Prior art

François Vander Linden's **bitmap_creator** (v1.0.1, January 2021) is an Excel spreadsheet for editing Apple II hi-res bitmaps:
https://github.com/tilleul/apple2/tree/master/tools/bitmap%20editor

His accompanying documentation covers the hi-res color rules in detail with Applesoft examples:
https://www.callapple.org/programming/new-excel-based-apple-ii-graphics-tool/

Key differences from `a2-hires-lab`:

- **Formula-driven, no VBA.** Color rendering uses conditional formatting formulas. This makes the workbook self-contained (works in LibreOffice) but buries the color logic in cell formulas rather than readable, testable code.
- **Screen-fragment editor, not sprite-focused.** His working area is 70 × 192 pixels (10 bytes × full screen height). No notion of sprites, no inventory, no shift mechanics.
- **No connection to game data.** No sprite data tables, no lookup-chain visualization, no emulator test output.

The research in his documentation and workbook is valuable reference material. His color rules confirm our NTSC decision table. The workbook itself is not a basis for our work -- the conceptual distance is too large -- but it validates that the approach (Excel + colored cells for hi-res visualization) works.


---


## Deliverable 1: Sprite Editor/Viewer


### Overview

One worksheet with four functional areas:

1. **Editor** -- an 11 × 14 grid of cells where the user enters 0 or 1, plus two high-bit toggle cells per row. This is the primary input.
2. **Bits display** -- two 7-character binary strings per row, showing the bit pattern for each byte (MSB left, matching the chapter's "Bits" column on page 8).
3. **Byte display** -- two hex values per row, derived from the pixel grid, high bits, and bit patterns.
4. **Color viewer** -- an 11 × 14 block of cells whose background color shows the NTSC rendering of the current pixel data.

Two buttons control the data flow:

- **"Update Viewer"** -- reads the pixel grid and high bits, computes bits/bytes, and paints the color viewer. Use after editing pixels.
- **"Load from Bytes"** -- reads the hex byte values (and high bits), decomposes them into pixel bits, populates the editor grid, and paints the color viewer. Use after pasting or loading byte data from the Inventory.

A later iteration will hook `Worksheet_Change` so the viewer updates on every edit, replacing the button clicks.


### Sheet layout

The layout below uses fixed cell ranges so that VBA code can address
them by name or constant.

```
      A      B  C  D  E  F  G  H     I     J  K  L  M  N  O  P     Q       R     S       T     U
  ┌──────┬─────────────────────────┬─────┬────────────────────────┬─────┬───────────┬───────────────┐
1 │      │            a2-hires-lab -- Sprite Editor/Viewer        │     │           │               │
  ├──────┼─────────────────────────┼─────┼────────────────────────┼─────┼───────────┼───────────────┤
2 │      │  (instructions / notes area)                           │     │           │               │
  ├──────┼─────────────────────────┼─────┼────────────────────────┼─────┼───────────┼───────────────┤
3 │ Row  │  ←── Byte 0 (cols 0–6) ─┤ HB0 │ ←── Byte 1 (cols 7–13)─┤ HB1 │ Hex0  Hex1│  Bits0   Bits1│
  ├──────┼──┬──┬──┬──┬──┬──┬───────┼─────┼──┬──┬──┬──┬──┬──┬──────┼─────┼─────┬─────┼───────┬───────┤
4 │  0   │  │  │  │  │  │  │       │  ●  │  │  │  │  │  │  │      │  ●  │  xx │  xx │0000000│0000000│
5 │  1   │  │  │  │  │  │  │       │     │  │  │  │  │  │  │      │     │     │     │       │       │
  │ ...  │  │  │  │  │  │  │       │     │  │  │  │  │  │  │      │     │     │     │       │       │
14│ 10   │  │  │  │  │  │  │       │     │  │  │  │  │  │  │      │     │     │     │       │       │
  └──────┴──┴──┴──┴──┴──┴──┴───────┴─────┴──┴──┴──┴──┴──┴──┴──────┴─────┴─────┴─────┴───────┴───────┘

      B  C  D  E  F  G  H  J  K  L  M  N  O  P
  ┌──────────────────────────────────────────────┐
16│           Color Viewer (11 × 14)             │
  ├──┬──┬──┬──┬──┬──┬──┬──┬──┬──┬──┬──┬──┬───────┤
17│  │  │  │  │  │  │  │  │  │  │  │  │  │       │    ← sprite row 0
  │ ...                                          │
27│  │  │  │  │  │  │  │  │  │  │  │  │  │       │    ← sprite row 10
  └──┴──┴──┴──┴──┴──┴──┴──┴──┴──┴──┴──┴──┴───────┘
```

Column I sits between the two 7-pixel groups to visually separate the bytes. The Color Viewer reuses the same columns (B–H, J–P) so pixel columns line up vertically between editor and viewer.

Named ranges (defined in VBA constants or Name Manager):

| Name              | Range           | Description                                   |
|-------------------|-----------------|-----------------------------------------------|
| `EditorByte0`     | B4:H14          | Byte 0 pixel columns (0–6) for rows 0–10     |
| `EditorByte1`     | J4:P14          | Byte 1 pixel columns (7–13) for rows 0–10    |
| `HighBit0`        | I4:I14          | High-bit toggle for byte 0, per row           |
| `HighBit1`        | Q4:Q14          | High-bit toggle for byte 1, per row           |
| `HexByte0`        | R4:R14          | Computed hex string for byte 0, per row       |
| `HexByte1`        | S4:S14          | Computed hex string for byte 1, per row       |
| `BitsByte0`       | T4:T14          | 7-char binary string for byte 0, per row      |
| `BitsByte1`       | U4:U14          | 7-char binary string for byte 1, per row      |
| `ViewerGrid`      | B17:H27, J17:P27| Color viewer area (background-colored cells)  |
| `SpriteNumCell`   | (TBD)           | For loading from inventory -- sprite index     |


### Pixel ↔ byte conversion

Pixels are numbered 0–13 left to right (screen column order). Byte 0 holds pixels 0–6, byte 1 holds pixels 7–13.

Within each byte, **bit 0 = leftmost pixel of that byte** (LSB-first display, matching the Apple II convention: bit 0 → column 0, bit 6 → column 6). Bit 7 is the high bit (color control, not pixel data).

**Pixels → byte** (Update Viewer direction):

```
byte0 = (iHB0 * 128) + (pixel[6] * 64) + (pixel[5] * 32)
      + (pixel[4] * 16) + (pixel[3] * 8) + (pixel[2] * 4)
      + (pixel[1] * 2) + pixel[0]

byte1 = (iHB1 * 128) + (pixel[13] * 64) + (pixel[12] * 32)
      + (pixel[11] * 16) + (pixel[10] * 8) + (pixel[9] * 4)
      + (pixel[8] * 2) + pixel[7]
```

**Byte → pixels** (Load from Bytes direction):

```
pixel[N] = (aiBytes(N \ 7) >> (N Mod 7)) And 1    ' N = 0..13
iHB0     = (aiBytes(0) >> 7) And 1
iHB1     = (aiBytes(1) >> 7) And 1
```

**Bits display:** The 7-char binary string shows bits 6 down to 0 (MSB left), matching the "Bits" column in the chapter's tables on page 8. This is the *storage* bit order, not the pixel order -- bit 6 is at the left of the string but corresponds to the rightmost pixel of that byte. The bits display exists so the user can read it alongside the chapter.


### NTSC artifact color rules

These rules determine the color of each pixel in the viewer. They are the formalization of the prose on page 7 of Chapter 3.

The chapter introduces these as "some rules." The patterns it lists (`11`, `010`, `101`, `01010`, `11011`) are illustrations, not separate rules -- they all follow from the nearest-neighbor decision table below. Tilleul's additional observations (colored pixels always odd in number, minimum distances between same-color dots) are consequences of these rules, not additions.

**Inputs per pixel:**

| Symbol | Meaning                                                     |
|--------|-------------------------------------------------------------|
| V      | Value of this pixel (0 or 1)                                |
| L      | Value of the pixel to the left (0 if at left edge)          |
| R      | Value of the pixel to the right (0 if at right edge)        |
| C      | Screen column of this pixel (0–13 within the sprite)        |
| H      | High bit of the byte containing this pixel                  |

**Decision table:**

| V | L or R = 1? | L = 1 and R = 1? | Result    |
|---|-------------|-------------------|-----------|
| 0 | --           | yes               | COLORED   |
| 0 | --           | no                | BLACK     |
| 1 | yes         | --                 | WHITE     |
| 1 | no          | --                 | COLORED   |

In words:

- A **1-pixel** that has at least one 1-neighbor is **white**.
- A **1-pixel** with 0 on both sides is **colored**.
- A **0-pixel** between two 1-neighbors is **colored**.
- All other 0-pixels are **black**.

**Color lookup** (when result is COLORED):

| High bit | Column parity | Color  |
|----------|---------------|--------|
| 0        | even (0,2,4…) | Violet |
| 0        | odd  (1,3,5…) | Green  |
| 1        | even (0,2,4…) | Blue   |
| 1        | odd  (1,3,5…) | Orange |

**RGB values** (matching `Apple.py`):

| Color  | RGB              | Hex       |
|--------|------------------|-----------|
| Black  | (0, 0, 0)        | `#000000` |
| White  | (255, 255, 255)  | `#FFFFFF` |
| Green  | (0, 255, 0)      | `#00FF00` |
| Violet | (255, 0, 255)    | `#FF00FF` |
| Orange | (255, 192, 0)    | `#FFC000` |
| Blue   | (0, 192, 255)    | `#00C0FF` |

**Boundary note:** The left neighbor of pixel 7 is pixel 6. These belong to different bytes with potentially different high bits. The adjacency check crosses the byte boundary; the color lookup uses the high bit of the byte containing the pixel being colored.

**Edge assumption:** Pixels outside the sprite (to the left of column 0, to the right of column 13) are treated as 0. This matches the chapter's example renderings. In the real game, context pixels from adjacent screen bytes would matter -- the Sprite Shifter (deliverable 4) will address this.

**Completeness note:** The cross-byte high-bit boundary is the one edge case beyond the basic nearest-neighbor rules. The chapter says "you can only select one pair of colors per byte"; Tilleul calls mixed-high-bit boundaries "color problems." Our rules handle this correctly: the adjacency check is palette-blind (it only looks at 0/1), and the color lookup uses each pixel's own byte's high bit. A visual "fringe" artifact at the boundary is a real NTSC effect that our cell-per-pixel rendering cannot show (it would need half-pixel resolution). This is a known simplification.


### Pseudocode for the color rendering

```
Sub UpdateViewer()
    Dim aiPixels(0 To 10, 0 To 13) As Long   ' from editor grid
    Dim aiHB(0 To 10, 0 To 1) As Long        ' high bits per row per byte

    ' 1. Read editor grid into aiPixels() and aiHB()
    '    (read from cell ranges)

    ' 2. For each row and column, determine color
    Dim iRow As Long, iCol As Long
    For iRow = 0 To 10
        For iCol = 0 To 13
            Dim iV As Long: iV = aiPixels(iRow, iCol)
            Dim iL As Long: iL = IIf(iCol > 0, aiPixels(iRow, iCol - 1), 0)
            Dim iR As Long: iR = IIf(iCol < 13, aiPixels(iRow, iCol + 1), 0)
            Dim iH As Long: iH = aiHB(iRow, iCol \ 7)

            Dim lColor As Long  ' RGB color value

            If iV = 1 Then
                If iL = 1 Or iR = 1 Then
                    lColor = RGB(255, 255, 255)         ' WHITE
                Else
                    lColor = ColoredPixel(iCol, iH)     ' COLORED
                End If
            Else  ' iV = 0
                If iL = 1 And iR = 1 Then
                    lColor = ColoredPixel(iCol, iH)     ' COLORED
                Else
                    lColor = RGB(0, 0, 0)               ' BLACK
                End If
            End If

            ViewerCell(iRow, iCol).Interior.Color = lColor
        Next iCol
    Next iRow

    ' 3. Compute and display hex byte values and bits strings
    For iRow = 0 To 10
        Dim lByte0 As Long: lByte0 = PixelsToByteValue(aiPixels, iRow, 0, aiHB(iRow, 0))
        Dim lByte1 As Long: lByte1 = PixelsToByteValue(aiPixels, iRow, 1, aiHB(iRow, 1))
        HexByte0Cell(iRow).Value = ByteToHex(lByte0)
        HexByte1Cell(iRow).Value = ByteToHex(lByte1)
        BitsByte0Cell(iRow).Value = ByteToBits(lByte0 And &H7F)
        BitsByte1Cell(iRow).Value = ByteToBits(lByte1 And &H7F)
    Next iRow
End Sub

Function ColoredPixel(iCol As Long, iHighBit As Long) As Long
    Dim iParity As Long: iParity = iCol Mod 2
    If iHighBit = 0 Then
        If iParity = 0 Then
            ColoredPixel = RGB(255, 0, 255)         ' Violet
        Else
            ColoredPixel = RGB(0, 255, 0)           ' Green
        End If
    Else
        If iParity = 0 Then
            ColoredPixel = RGB(0, 192, 255)         ' Blue
        Else
            ColoredPixel = RGB(255, 192, 0)         ' Orange
        End If
    End If
End Function
```


### VBA module structure

```
Modules/
├── Util               Hex/decimal/bit conversion helpers
├── NTSCColor          ColoredPixel(), full-row color scan
├── SpriteEditor       UpdateViewer(), LoadFromBytes(), PixelsToByteValue()
└── SpriteInventory    (deliverable 2) ParseSpriteData(), SelectSprite()
```

Class modules use a `cl` prefix (e.g. `clSprite`) if needed in later deliverables.

**`Util` functions:**

| Function              | Signature                                   | Purpose                        |
|-----------------------|---------------------------------------------|--------------------------------|
| `ByteToHex`           | `(lVal As Long) As String`                  | 0–255 → "00"–"FF"             |
| `HexToByte`           | `(sHex As String) As Long`                  | "00"–"FF" → 0–255             |
| `ByteToBits`          | `(lVal As Long) As String`                  | 0–127 → "0000000"–"1111111"   |
| `BitsToByte`          | `(sBits As String) As Long`                 | inverse of ByteToBits          |
| `GetBit`              | `(lVal As Long, iBit As Long) As Long`      | extract one bit                |
| `SetBit`              | `(lVal As Long, iBit As Long, iB As Long) As Long` | set one bit             |

**Naming convention:** Hungarian notation throughout. `i` Integer, `l` Long, `s` String, `v` Variant, `ai` array of integers, `al` array of longs, `rng` Range. No `mod` prefix on module names; `cl` prefix on class names.


### Data notes

**High-bit default:** The sprite data in `sprite_data.asm` stores 7-bit values (0x00–0x7F). The pixel pattern table (`pixel_pattern_table.asm`) has bit 7 set in every output byte -- this is how "the game automatically sets the high bit." In the editor, HB0 and HB1 default to 1 to match game behavior, but the user can clear them to see the green/violet palette.

**Sprite data layout in `sprite_data.asm`:** The file contains 2288 bytes = 104 sprites × 22 bytes/sprite. The layout is interleaved: all first-bytes of all 104 sprites appear first (104 bytes), then all second-bytes (104 bytes), and so on for all 22 byte-positions.

This interleaving is a 6502 performance optimization. The code uses `LDA (TMP_PTR),Y` where Y holds the sprite number and TMP_PTR points to the start of a byte-position block. Reading any sprite's byte at position B is a single indirect-indexed load -- no multiplication needed. If the data were serial (22 bytes per sprite in sequence), retrieving sprite S at byte-position B would require computing offset = S × 22 + B, and the 6502 has no multiply instruction.

To extract sprite S, byte-position B: offset = B × 104 + S. This gives the byte at row (B \ 2), byte-within-row (B Mod 2) of sprite S.


---


## Deliverable 2: Sprite Inventory (sketch)

A second worksheet containing a table of all 104 sprites, reorganized from the interleaved layout into one block of 11 rows × 2 hex bytes per sprite.

A dropdown or index cell on the Editor/Viewer sheet selects a sprite number (0–103). A "Load" button copies the byte data into the hex cells and triggers `LoadFromBytes`.

Sprite names from the `defines` (SPRITE_EMPTY = 0, SPRITE_BRICK = 1, etc.) appear as labels.


## Deliverable 3: Pixel Shifter (sketch)

A worksheet implementing the table-lookup chain from page 24 of Chapter 3.

**Inputs:** a 7-bit pixel pattern (0–127) and a shift amount (0–6).

**Outputs:** two bytes representing the shifted pattern.

**Visible steps** (each a labeled cell or small group of cells):

1. Shift amount → `PIXEL_SHIFT_PAGES` → page number
2. Page + pattern → offset table entry → offset
3. Page + pattern → page table entry → result page
4. Result page + offset → `PIXEL_PATTERN_TABLE` → byte 0
5. Result page + (offset+1) → `PIXEL_PATTERN_TABLE` → byte 1

Each intermediate value is shown in a cell, with the relevant table row/column highlighted (via VBA background color) to show "this is the entry being read."

The data from `pixel_shift_table.asm` (1792 bytes = 7 × 256) and `pixel_pattern_table.asm` (512 × 2 bytes) lives on a hidden or separate data sheet.

A scratchpad area mirrors the 6502 registers (A, X, Y) at each step of the `COMPUTE_SHIFTED_SPRITE` inner loop to make the code's execution tangible.


## Deliverable 4: Sprite Shifter (sketch)

Applies the Pixel Shifter logic to all 22 bytes of the sprite currently in the Editor/Viewer.

**Input:** shift amount (0–6), sprite data from the editor.

**Output:** 11 rows × 3 bytes (the shifted sprite in `BLOCK_DATA` format). The OR step for the overlapping middle byte (where byte 1 of the first source byte and byte 0 of the second source byte are combined) is shown explicitly.

The result can be loaded back into the Editor/Viewer. Since the shifted sprite spans 3 bytes (up to 21 pixel columns), the viewer would need to expand. This implies either:

- A wider editor/viewer mode (21 columns, 3 high-bit toggles), or
- Showing only the leftmost 14 columns (which is what the game does when drawing: only 10 pixels out of 14 are used per sprite).

Decision deferred until deliverable 1 is built and we can see what feels right.


## Deliverable 5: Memory Map Viewer (idea)

Two worksheets (HGR1 and HGR2) with small cells representing the 8 KB memory-mapped graphics pages. Each cell is one byte (or one
pixel -- TBD). Coloring shows which pixels are set and what colors they produce. Could be used to visualize the non-consecutive row layout (the "rows are not consecutive in memory" aspect from section 3.4).

Not designed yet. Depends on having the NTSC color logic solid from deliverable 1.


## Further ideas

- **Save sprite:** Export the current editor content back to `sprite_data.asm` format (interleaved layout). This would let the workbook round-trip: load a sprite from inventory, edit it, save it back, and eventually regenerate the full `.asm` file.

- **PyXll bridge:** Once `papple2`'s `update_hires` is correct, call it from Excel via PyXll to render sprites with the full emulator color model instead of the VBA approximation.

- **Emulator-driven verification:** Load the full Lode Runner game image in `papple2`, run `DRAW_SPRITE_PAGE1` with known inputs, capture the screen buffer, and compare against the workbook's rendering.

- **Half-pixel / 560-column mode:** A viewer mode that uses two cells per pixel column to show the NTSC half-pixel shift between palettes. Low priority -- the 280-column model is correct for sprite work.


---


## Seed package for the build session

The following files should be included in the conversation that builds the first deliverable:

| File                          | Purpose                                              |
|-------------------------------|------------------------------------------------------|
| `a2-hires-lab-design.md`      | This document (the build contract)                   |
| `sprite_data.asm`             | Raw sprite byte data (needed for deliverable 2, but useful for testing deliverable 1) |
| `pixel_pattern_table.asm`     | Shifted pixel patterns (deliverable 3 data, but confirms high-bit behavior) |
| `pixel_shift_table.asm`       | Shift offset/page lookup (deliverable 3 data)        |
| `Apple.py`                    | Reference for current (simplified) color logic and RGB values |
| `chapter-3.pdf`               | Full chapter for context -- color rules, sprite examples, routine documentation |

Optionally include `main-chapter-3.md` (the LaTeX/Noweb source) if the build session needs to extract data programmatically from the chapter's tables.

**Build order within the session:**

1. Create the workbook with the Editor/Viewer sheet layout
2. Implement `Util` (hex/bit conversions)
3. Implement `NTSCColor` (color decision logic)
4. Implement `SpriteEditor` (`UpdateViewer`, `LoadFromBytes`, button wiring)
5. Test with the two example sprites from page 8 of the chapter
6. If time allows, add the Sprite Inventory sheet and `LoadSprite`
