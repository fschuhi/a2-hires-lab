# Lode Runner Shift Engine: Architecture & Mechanics

To simulate the *Lode Runner* shift engine in the `a2-hires-lab` VBA environment, it is critical to understand that the game performs absolutely no runtime bit-shifting. Instead, it uses a highly optimized, two-stage dictionary lookup to translate any 7-bit sprite pixel sequence and a required horizontal shift (0 to 6) into a ready-to-blit 14-bit (2-byte) screen output.

## 1. The Core Engineering Constraint

The Apple II high-resolution screen draws pixels least-significant-bit (LSB) first. A single sprite byte in *Lode Runner* contains 7 active pixels (with the high bit always set to $1$ for color artifacting). 

When moving a sprite horizontally, shifting a 7-pixel window to the right by any amount from 0 to 6 pixels causes the visual output to overflow into a second byte. Therefore, any shifted pattern spans a 14-pixel wide, 2-byte window.

## 2. The Pattern Table (`pixel_pattern_table.asm`)

The `pixel_pattern_table.asm` is a pre-rendered gallery of output shapes. It contains no instructions, just pairs of raw screen bytes (`BYTE byte0, byte1`). 

### Why Exactly 512 Entries (With Zero Padding)?
A naive lookup table mapping every 7-bit input ($2^7 = 128$) across 7 shift positions ($0 \dots 6$) would require 896 two-byte pairs. However, many combinations produce the exact same 14-pixel visual state. For example, a single pixel at position 0 shifted by 4 looks identical to a single pixel at position 4 shifted by 0.

The engine only stores **mathematically unique** 14-pixel visual outcomes. When shifting a 7-pixel pattern by 0 to 6 places, the active pixels only ever land in columns $0 \dots 12$ of the 14-pixel field (column 13 is never touched). 

If we count every valid 13-bit output state that could originate from a contiguous 7-pixel window (meaning the distance between the first $1$ and the last $1$ is at most 6), we can break it down by the pixel width $W$:

*   **Width 1:** 1 shape $\times\ 13$ valid placement positions $= 13$
*   **Width 2:** 1 shape ($2^0$) $\times\ 12$ positions $= 12$
*   **Width 3:** 2 shapes ($2^1$) $\times\ 11$ positions $= 22$
*   **Width 4:** 4 shapes ($2^2$) $\times\ 10$ positions $= 40$
*   **Width 5:** 8 shapes ($2^3$) $\times\ 9$ positions $= 72$
*   **Width 6:** 16 shapes ($2^4$) $\times\ 8$ positions $= 128$
*   **Width 7:** 32 shapes ($2^5$) $\times\ 7$ positions $= 224$

Summing these gives exactly **511 unique non-zero visual states**. Adding the $1$ all-zero pattern brings the grand total to **exactly 512 unique visual outcomes**.

Because there are exactly 512 combinations, there are no dummy bytes or padding. At 2 bytes per entry, the table requires exactly 1,024 bytes, perfectly filling four consecutive 256-byte 6502 memory pages (`$A900` to `$ACFF`).

## 3. The Shift Table (`pixel_shift_table.asm`)

To find a specific pattern in that 512-entry gallery, the engine consults the index: `pixel_shift_table.asm`. 

Because the 6502 processor uses 8-bit index registers ($0 \dots 255$), it cannot easily retrieve a 16-bit target address from a single array. To solve this, the lookup table for each shift amount (0 to 6) is allocated its own 256-byte page, mathematically split into two 128-byte halves:

1.  **Low Bytes (Offsets):** The first 128 bytes of the page (indices `$00 \dots$7F`).
2.  **High Bytes (Pages):** The second 128 bytes of the page (indices `$80 \dots$FF`).

### The 3-Column Mapping Model
When building the "Pixel Shifter" in Excel, it helps to visualize these split arrays as a unified 3-column map. The array index acts as the input key. Note that the high bit of the input pattern is always $0$ (values $0 \dots 127$), representing just the 7 raw pixels.

Here is how the raw assembly for Shift 0 translates into a readable mapping structure:

| Key: Input Byte (Binary) | Value 1: Low Byte (Offset) | Value 2: High Byte (Page) | Resulting Target Address |
| :--- | :--- | :--- | :--- |
| `%00000000` (0) | `$00` (Index 0) | `$A9` (Index 128) | `$A900` |
| `%00000001` (1) | `$02` (Index 1) | `$A9` (Index 129) | `$A902` |
| `%00000010` (2) | `$04` (Index 2) | `$A9` (Index 130) | `$A904` |
| `%00000011` (3) | `$12` (Index 3) | `$A9` (Index 131) | `$A912` |
| ... | ... | ... | ... |
| `%00100000` (32) | `$0C` (Index 32) | `$A9` (Index 160) | `$A90C` |
| `%00100001` (33) | `$02` (Index 33) | `$AA` (Index 161) | `$AA02` |
| `%00100010` (34) | `$84` (Index 34) | `$A9` (Index 162) | `$A984` |
| `%00100011` (35) | `$12` (Index 35) | `$AA` (Index 163) | `$AA12` |

### Why the High Bytes Alternate (`$A9`, `$AA`, etc.)
Looking at the High Byte column above, starting at input `33`, the values begin bouncing between `$A9` and `$AA`. This occurs because the target addresses in the Pattern Table are crossing the physical hardware boundary separating memory page `$A900` and `$AA00`. 

By storing the high and low bytes in these independent 128-byte arrays, the engine can pull the exact 16-bit target address using two extremely fast, single-cycle indexed load instructions (`LDA .offset_table, Y` and `LDA .page_table, Y`).