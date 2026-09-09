# Chapter 3: Apple II Graphics

Source: `main.nw`, lines 142-1557.

Navigation is generated from the full source. Defines and Uses follow the bundled weaver's identifier index; they are not a complete assembly symbol analysis. References lists included chunks; Referenced by lists chunks that include this named chunk. Locators refer to the original source, not this export.

Hi-res graphics on the Apple II is odd. Graphics are memory-mapped, not exactly
consecutively, and bits don't always correspond to pixels. Color especially is
odd, compared to today's luxurious 32-bit per pixel RGBA.

The Apple II has two hi-res graphics pages, and maps the area from `$2000-$3FFF` to
high-res graphics page 1 (HGR1), and `$4000-$5FFF` to page 2 (HGR2).

We have routines to clear these screens.

**Chunk:** `routines` (lines 152-182)

```asm
    ORG     $7A51
CLEAR_HGR1:
    SUBROUTINE

    LDA     #$20                ; Start at $2000
    LDX     #$40                ; End at $4000 (but not including)
    BNE     CLEAR_PAGE          ; Unconditional jump

CLEAR_HGR2:
    SUBROUTINE

    LDA     #$40                ; Start at $4000
    LDX     #$60                ; End at $6000 (but not including)
    ; fallthrough

CLEAR_PAGE:
    STA     TMP_PTR+1           ; Start with the page in A.
    LDA     #$00
    STA     TMP_PTR
    TAY
    LDA     #$80                ; fill byte = 0x80

.loop:
    STA     (TMP_PTR),Y
    INY
    BNE     .loop
    INC     TMP_PTR+1
    CPX     TMP_PTR+1
    BNE     .loop               ; while TMP_PTR != X * 0x100
    RTS
```

**Previous:** None  
**Next:** `routines` (lines 451-547)  
**Defines:** `CLEAR_HGR1` -- used in `put status` (lines 1463-1555), `splash screen` (lines 4042-4067), `level editor` (lines 9257-9343); `CLEAR_HGR2` -- used in `put status` (lines 1463-1555), `construct and display high score screen` (lines 3870-3885), `return handler` (lines 4925-5050), `bad data disk` (lines 8490-8510), `dont manipulate master disk` (lines 8513-8532), `editor edit level` (lines 9492-9564), `level editor key functions` (lines 9605-9760)  
**Uses:** `TMP_PTR` -- `defines` (lines 132-139)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

## Pixels and their color

First we'll talk about pixels. Nominally, the resolution of the hi-res graphics screen
is 280 pixels wide by 192 pixels tall. In the memory map, each row is represented
by 40 bytes. The high bit of each byte is not used for pixel data, but is used to
control color.

Here are some rules for how these bytes are turned into pixels:

```latex
\begin{itemize}
  \item Pixels are drawn to the screen from byte data least significant bit first.
        This means that for the first byte bit 0 is column 0, bit 1 is column 1,
        and so on.
  \item A pattern of [[11]] results in two white pixels at the [[1]] positions.
  \item A pattern of [[010]] results at least in a colored pixel at the [[1]] position.
  \item A pattern of [[101]] results at least in a colored pixel at the [[0]] position.
  \item So, a pattern of [[01010]] results in at least three consecutive colored
        pixels starting from the first [[1]] to the last [[1]]. The last [[0]] bit
        would also be colored if followed by a [[1]].
  \item Likewise, a pattern of [[11011]] results in two white pixels, a colored pixel,
        and then two more white pixels.
  \item The color of a [[010]] pixel depends on the column that the [[1]] falls on, and
        also whether the high bit of its byte was set or not. 
  \item The color of a [[11011]] pixel depends on the column that the [[0]] falls on, and
        also whether the high bit of its byte was set or not.

        \begin{center}
        \begin{tabular}{@{}rcc@{}} \toprule
        & Odd & Even \\ \cmidrule(r){2-3}
        High bit clear & Green & Violet \\
        High bit set & Orange & Blue \\ \bottomrule
        \end{tabular}
        \end{center}

        The implication is that you can only select one pair of colors per byte.
\end{itemize}
```


An example would probably be good here. We will take one of the sprites from the game.


```latex
\begin{center}
\begin{tabular}{@{}rcc@{}} \toprule
Bytes & Bits & Pixel Data \\ \cmidrule{1-3}
[[00 00]] & [[0000000 0000000]] & [[00000000000000]] \\
[[00 00]] & [[0000000 0000000]] & [[00000000000000]] \\
[[00 00]] & [[0000000 0000000]] & [[00000000000000]] \\
[[55 00]] & [[1010101 0000000]] & [[10101010000000]] \\
[[41 00]] & [[1000001 0000000]] & [[10000010000000]] \\
[[01 00]] & [[0000001 0000000]] & [[10000000000000]] \\
[[55 00]] & [[1010101 0000000]] & [[10101010000000]] \\
[[50 00]] & [[1010000 0000000]] & [[00001010000000]] \\
[[50 00]] & [[1010000 0000000]] & [[00001010000000]] \\
[[51 00]] & [[1010001 0000000]] & [[10001010000000]] \\
[[55 00]] & [[1010101 0000000]] & [[10101010000000]] \\ \bottomrule
\end{tabular}
\end{center}
```


The game automatically sets the high bit of each byte, so we know we're going to see
orange and blue. Assuming that the following bits are all zero, and we place the
sprite starting at column 0, we should see this:


```latex
\begin{center}
\begin{tabular}{@{}rcccccccccccccc@{}}
 0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 1 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 2 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 3 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 4 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 5 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 6 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 7 & \bk0 & \bk0 & \bk0 & \bk0 & \bl0 & \bl0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 8 & \bk0 & \bk0 & \bk0 & \bk0 & \bl0 & \bl0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
 9 & \bl0 & \bk0 & \bk0 & \bk0 & \bl0 & \bl0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
10 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
\end{tabular}
\end{center}
```


Here is a more complex sprite:


```latex
\begin{center}
\begin{tabular}{@{}rcc@{}} \toprule
Bytes & Bits & Pixel Data \\ \cmidrule{1-3}
[[40 00]] & [[1000000 0000000]] & [[00000010000000]] \\
[[60 01]] & [[1100000 0000001]] & [[00000111000000]] \\
[[60 01]] & [[1100000 0000001]] & [[00000111000000]] \\
[[70 00]] & [[1110000 0000000]] & [[00001110000000]] \\
[[6C 01]] & [[1101100 0000001]] & [[00110111000000]] \\
[[36 06]] & [[0110110 0000110]] & [[01101100110000]] \\
[[30 00]] & [[0110000 0000000]] & [[00001100000000]] \\
[[70 00]] & [[1110000 0000000]] & [[00001110000000]] \\
[[5E 01]] & [[1011110 0000001]] & [[01111011000000]] \\
[[40 01]] & [[1000000 0000001]] & [[00000011000000]] \\
[[40 01]] & [[1000000 0000001]] & [[00000011000000]] \\ \bottomrule
\end{tabular}
\end{center}
```



```latex
\begin{center}
\begin{tabular}{@{}rcccccccccccccc@{}}
0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bl0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
1 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bw0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
2 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bw0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
3 & \bk0 & \bk0 & \bk0 & \bk0 & \bw0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
4 & \bk0 & \bk0 & \bw0 & \bw0 & \bo0 & \bw0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
5 & \bk0 & \bw0 & \bw0 & \bl0 & \bw0 & \bw0 & \bk0 & \bk0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 \\
6 & \bk0 & \bk0 & \bk0 & \bk0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
7 & \bk0 & \bk0 & \bk0 & \bk0 & \bw0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
8 & \bk0 & \bw0 & \bw0 & \bw0 & \bw0 & \bl0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
9 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
10 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bw0 & \bw0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 & \bk0 \\
\end{tabular}
\end{center}
```


Take note of the orange and blue pixels. All the patterns noted in the rules above are used.


## The sprites

Lode Runner defines 104 sprites, each being 11 rows, with two bytes per row. The first bytes of
all 104 sprites are in the table first, then the second bytes, then the third bytes, and so on.
Later we will see that only the leftmost 10 pixels out of the 14-pixel description is used.

**Chunk:** `tables` (lines 301-304)

```asm
    ORG     $AD00
SPRITE_DATA:
    INCLUDE "sprite_data.asm"
```

**Previous:** None  
**Next:** `tables` (lines 383-386)  
**Defines:** `SPRITE_DATA` -- used in `routines` (lines 451-547)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

```latex
\input{sprite_tables.tex}
```

**Chunk:** `defines` (lines 309-324)

```asm
SPRITE_EMPTY        EQU     #$00
SPRITE_BRICK        EQU     #$01
SPRITE_STONE        EQU     #$02
SPRITE_LADDER       EQU     #$03
SPRITE_ROPE         EQU     #$04
SPRITE_TRAP         EQU     #$05
SPRITE_INVISIBLE_LADDER       EQU     #$06
SPRITE_GOLD         EQU     #$07
SPRITE_GUARD        EQU     #$08
SPRITE_PLAYER       EQU     #$09
SPRITE_ALLWHITE     EQU     #$0A
SPRITE_BRICK_FILL0  EQU     #$37
SPRITE_BRICK_FILL1  EQU     #$38
SPRITE_GUARD_EGG0   EQU     #$39
SPRITE_GUARD_EGG1   EQU     #$3A
```

**Previous:** `defines` (lines 132-139)  
**Next:** `defines` (lines 335-336)  
**Defines:** None  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

## Shifting sprites

This is all very good if we're going to draw sprites exactly on 7-pixel
boundaries, but what if we want to draw them starting at other columns?
In general, such a shifted sprite would straddle three bytes, and Lode
Runner sets aside an area of memory at the end of zero page for 11 rows
of three bytes that we'll write to when we want to compute the data for
a shifted sprite.

**Chunk:** `defines` (lines 335-336)

```asm
BLOCK_DATA      EQU     $DF     ; 33 bytes
```

**Previous:** `defines` (lines 309-324)  
**Next:** `defines` (lines 446-448)  
**Defines:** `BLOCK_DATA` -- used in `routines` (lines 451-547), `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

Lode Runner also contains tables which show how to shift any arbitrary
7-pixel pattern right by any amount from zero to six pixels.

For example, suppose we start with a pixel pattern of `0110001`, and we want to
shift that right by three bits. The 14-bit result would be `0000110 0010000`.
However, we have to break that up into bytes, reverse the bits (remember that
each byte's bits are output as pixels least significant bit first), and set
their high bits, so we end up with `10110000 10000100`.

Now, given a shift amount and a pixel pattern, we should be able to find the
two-byte shifted pattern. Lode Runner accomplishes this with table lookups as follows:


```latex
\vspace{1em}
```


```latex
\begin{tikzpicture}
  [basicbox/.style={draw,rectangle,inner sep=0pt,minimum width=1.5cm,minimum height=1.5cm,fill=blue!10},
   pageoffsets/.style={basicbox,minimum height=1cm,text height=1.5ex,text depth=.25ex},
   multilinebox/.style={basicbox,text width=1cm,align=center}]
  \node (pixelshiftpages) at (0,0) [multilinebox] {pixel shift pages};
  \node (start) at (-3,0) {};
  \draw [->] (start) -- (pixelshiftpages) node [above,text width=1cm,align=center,midway] {shift amount};
  \node (offsets0) [pageoffsets,anchor=north,below right=0 and 2 of pixelshiftpages.north east] {offsets};
  \node (pages0) [pageoffsets,below=0 of offsets0.south] {pages};
  \node (offsets1) [pageoffsets,below=0 of pages0.south,fill=blue!30] {offsets};
  \node (pages1) [pageoffsets,below=0 of offsets1.south,fill=blue!30] {pages};
  \node (offsets2) [pageoffsets,below=0 of pages1.south] {offsets};
  \node (pages2) [pageoffsets,below=0 of offsets2.south] {pages};
  \draw [->] (pixelshiftpages) -- (offsets1.north west) {};
  \node (pixelpattern) [above left=1 and 0 of offsets0.north west] {pixel pattern};
  \draw [->] (pixelpattern.south) |- (offsets1.west) {};
  \draw [->] (pixelpattern.south) |- (pages1.west) {};
  \node (patterntable) [multilinebox,minimum height=4cm,text width=1.2cm,anchor=north west,below right=0 and 2 of offsets0.north east] {pixel pattern table};
  \node (join) [inner sep=0pt,below right=0 and 1 of offsets1.south east] {};
  \draw (offsets1.east) -- (join);
  \draw (pages1.east) -- (join);
  \draw [->] (join) -- ([yshift=5mm]patterntable.west);
\end{tikzpicture}
```


```latex
\vspace{1em}
```


The pixel pattern table is a table of every possible pattern of 7 consecutive pixels
spread out over two bytes. This table is 512 entries, each entry being two bytes.
A naive table would have redundancy. For example the pattern `0000100` starting
at column 0 is exactly the same as the pattern `0001000` starting at column 1.
This table eliminates that redundancy.

**Chunk:** `tables` (lines 383-386)

```asm
    ORG     $A900
PIXEL_PATTERN_TABLE:
    INCLUDE "pixel_pattern_table.asm"
```

**Previous:** `tables` (lines 301-304)  
**Next:** `tables` (lines 394-397)  
**Defines:** `PIXEL_PATTERN_TABLE` -- no indexed uses  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

Now we just need tables which index into `PIXEL_PATTERN_TABLE` for every
7-pixel pattern and shift value. This table works by having the page number
for the shifted pixel pattern at index `shift * 0x100 + 0x80 + pattern`
and the offset at index `shift * 0x100 + pattern`.

**Chunk:** `tables` (lines 394-397)

```asm
    ORG     $A200
PIXEL_SHIFT_TABLE:
    INCLUDE "pixel_shift_table.asm"
```

**Previous:** `tables` (lines 383-386)  
**Next:** `tables` (lines 404-407)  
**Defines:** `PIXEL_SHIFT_TABLE` -- no indexed uses  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

Rather than multiplying the shift value by `0x100`, we instead define
another table which holds the page numbers for the shift tables for each
shift value.

**Chunk:** `tables` (lines 404-407)

```asm
    ORG     $84C1
PIXEL_SHIFT_PAGES:
    HEX     A2 A3 A4 A5 A6 A7 A8
```

**Previous:** `tables` (lines 394-397)  
**Next:** `tables` (lines 558-563)  
**Defines:** `PIXEL_SHIFT_PAGES` -- used in `routines` (lines 451-547)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

So we can get shifted pixels by indexing into all these tables.

Now we can define a routine that will take a sprite number and a pixel shift
amount, and write the shifted pixel data into the `BLOCK_DATA` area. The
routine first shifts the first byte of the sprite into a two-byte area. Then
it shifts the second byte of the sprite, and combines that two-byte result
with the first. Thus, we shift two bytes of sprite data into a three-byte
result.


```latex
\begin{center}
\begin{tikzpicture}
  [basicbox/.style={draw,rectangle,inner sep=0pt,minimum width=1.5cm,minimum height=0.5cm,fill=blue!10}]
  \node (spriterowbyte0) at (0,0) [basicbox] {};
  \node (spriterowbyte1) [basicbox,right=0 of spriterowbyte0.east] {};
  \node (spriterowlabel) [left=0.1 of spriterowbyte0.west] {sprite row};
  \node (shifted0byte0) [basicbox,below left=1 and 0 of spriterowbyte0.south west] {};
  \node (shifted0byte1) [basicbox,right=0 of shifted0byte0.east] {};
  \node (shifted1byte0) [basicbox,below right=2 and 0 of spriterowbyte0.south west] {};
  \node (shifted1byte1) [basicbox,right=0 of shifted1byte0.east] {};
  \node (orlabel) [below=0 of shifted0byte1] {OR};
  \draw [->] (spriterowbyte0.south) -- (shifted0byte0.north east)
    node [left,text width=1cm,align=center,midway] {shift};
  \draw [->] (spriterowbyte1.south) to [auto, bend left=45] node {shift} (shifted1byte0.north east);
  \node (result0) [basicbox,below left=0.5 and 0 of shifted1byte0.south west] {};
  \node (result1) [basicbox,right=0 of result0.east] {};
  \node (result2) [basicbox,right=0 of result1.east] {};
  \draw [->] (shifted0byte0) -- (result0) {};
  \draw [->] (shifted1byte0) -- (result1) {};
  \draw [->] (shifted1byte1) -- (result2) {};
  \node (blocklabel) [right=0.1 of result2.east] {block data};
\end{tikzpicture}
\end{center}
```


Rather than load addresses from the tables and store them, the routine
modifies its own instructions with those addresses.

**Chunk:** `defines` (lines 446-448)

```asm
ROW_COUNT       EQU     $1D
SPRITE_NUM      EQU     $1E
```

**Previous:** `defines` (lines 335-336)  
**Next:** `defines` (lines 566-569)  
**Defines:** `ROW_COUNT` -- used in `routines` (lines 451-547), `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121), `record hi score data` (lines 8305-8487); `SPRITE_NUM` -- used in `routines` (lines 451-547), `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121), `dead code` (lines 3783-3831), `get player sprite and coord data` (lines 4557-4574), `check for input` (lines 4702-4757), `get guard sprite and coords` (lines 6558-6575), `editor edit level` (lines 9492-9564)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `routines` (lines 451-547)

```asm
    ORG     $8438
COMPUTE_SHIFTED_SPRITE:
    SUBROUTINE
    ; Enter routine with X set to pixel shift amount and
    ; SPRITE_NUM containing the sprite number to read.

.offset_table       EQU $A000               ; Target addresses in read
.page_table         EQU $A080               ; instructions. The only truly
.shift_ptr_byte0    EQU $A000               ; necessary value here is the
.shift_ptr_byte1    EQU $A000               ; 0x80 in .shift_ptr_byte0.

    LDA     #$0B                            ; 11 rows
    STA     ROW_COUNT
    LDA     #<SPRITE_DATA
    STA     TMP_PTR
    LDA     #>SPRITE_DATA
    STA     TMP_PTR+1                       ; TMP_PTR = SPRITE_DATA
    LDA     PIXEL_SHIFT_PAGES,X 
    STA     .rd_offset_table + 2
    STA     .rd_page_table + 2
    STA     .rd_offset_table2 + 2
    STA     .rd_page_table2 + 2             ; Fix up pages in lookup instructions
                                            ; based on shift amount (X).

    LDX     #$00                            ; X is the offset into BLOCK_DATA.

.loop:                                      ; === LOOP === (over all 11 rows)
    LDY     SPRITE_NUM
    LDA     (TMP_PTR),Y 
    TAY                                     ; Get sprite pixel data.

.rd_offset_table:
    LDA     .offset_table,Y                 ; Load offset for shift amount.
    STA     .rd_shift_ptr_byte0 + 1
    CLC
    ADC     #$01
    STA     .rd_shift_ptr_byte1 + 1         ; Fix up instruction offsets with it.
.rd_page_table:
    LDA     .page_table,Y                   ; Load page for shift amount.
    STA     .rd_shift_ptr_byte0 + 2
    STA     .rd_shift_ptr_byte1 + 2         ; Fix up instruction page with it.

.rd_shift_ptr_byte0:
    LDA     .shift_ptr_byte0                ; Read shifted pixel data byte 0
    STA     BLOCK_DATA,X                    ; and store in block data byte 0.
.rd_shift_ptr_byte1:
    LDA     .shift_ptr_byte1                ; Read shifted pixel data byte 1
    STA     BLOCK_DATA+1,X                  ; and store in block data byte 1.

    LDA     TMP_PTR
    CLC
    ADC     #$68
    STA     TMP_PTR
    LDA     TMP_PTR+1
    ADC     #$00
    STA     TMP_PTR+1                       ; TMP_PTR++

    ; Now basically do the same thing with the second sprite byte

    LDY     SPRITE_NUM
    LDA     (TMP_PTR),Y 
    TAY                                     ; Get sprite pixel data.

.rd_offset_table2:
    LDA     .offset_table,Y                 ; Load offset for shift amount.
    STA     .rd_shift_ptr2_byte0 + 1
    CLC
    ADC     #$01
    STA     .rd_shift_ptr2_byte1 + 1        ; Fix up instruction offsets with it.
.rd_page_table2:
    LDA     .page_table,Y                   ; Load page for shift amount.
    STA     .rd_shift_ptr2_byte0 + 2
    STA     .rd_shift_ptr2_byte1 + 2        ; Fix up instruction page with it.

.rd_shift_ptr2_byte0:
    LDA     .shift_ptr_byte0                ; Read shifted pixel data byte 0
    ORA     BLOCK_DATA+1,X                  ; OR with previous block data byte 1
    STA     BLOCK_DATA+1,X                  ; and store in block data byte 1.
.rd_shift_ptr2_byte1:
    LDA     .shift_ptr_byte1                ; Read shifted pixel data byte 1
    STA     BLOCK_DATA+2,X                  ; and store in block data byte 2.

    LDA     TMP_PTR
    CLC
    ADC     #$68
    STA     TMP_PTR
    LDA     TMP_PTR+1
    ADC     #$00
    STA     TMP_PTR+1                       ; TMP_PTR++

    INX
    INX
    INX                                     ; X += 3
    DEC     ROW_COUNT                       ; ROW_COUNT--
    BNE     .loop                           ; loop while ROW_COUNT > 0
    RTS
```

**Previous:** `routines` (lines 152-182)  
**Next:** `routines` (lines 572-584)  
**Defines:** `COMPUTE_SHIFTED_SPRITE` -- used in `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121)  
**Uses:** `BLOCK_DATA` -- `defines` (lines 335-336); `PIXEL_SHIFT_PAGES` -- `tables` (lines 404-407); `ROW_COUNT` -- `defines` (lines 446-448); `SPRITE_DATA` -- `tables` (lines 301-304); `SPRITE_NUM` -- `defines` (lines 446-448); `TMP_PTR` -- `defines` (lines 132-139)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

## Memory mapped graphics


Within a screen row, consecutive bytes map to consecutive pixels. However, rows
themselves are not consecutive in memory.

To make it easy to convert a row number from 0 to 191 to a base address, Lode Runner has
a table and a routine to use that table.

**Chunk:** `tables` (lines 558-563)

```asm
    ORG     $1A85
ROW_TO_OFFSET_LO:
    INCLUDE "row_to_offset_lo_table.asm"
ROW_TO_OFFSET_HI:
    INCLUDE "row_to_offset_hi_table.asm"
```

**Previous:** `tables` (lines 404-407)  
**Next:** `tables` (lines 619-656)  
**Defines:** `ROW_TO_OFFSET_HI` -- used in `routines` (lines 572-584), `routines` (lines 589-605); `ROW_TO_OFFSET_LO` -- used in `routines` (lines 572-584), `routines` (lines 589-605)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `defines` (lines 566-569)

```asm
ROW_ADDR        EQU     $0C     ; 2 bytes
ROW_ADDR2       EQU     $0E     ; 2 bytes
HGR_PAGE        EQU     $1F     ; 0x20 for HGR1, 0x40 for HGR2
```

**Previous:** `defines` (lines 446-448)  
**Next:** `defines` (lines 614-616)  
**Defines:** `HGR_PAGE` -- used in `routines` (lines 572-584), `routines` (lines 808-908), `splash screen` (lines 4042-4067), `anims` (lines 9066-9143); `ROW_ADDR` -- used in `routines` (lines 572-584), `routines` (lines 589-605), `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121), `level draw routine` (lines 2704-2728), `draw wipe block` (lines 3095-3117), `[[ROW_ADDR = $9E00 + LEVELNUM * $0100]]` (lines 3535-3541), `Copy data from [[ROW_ADDR]] into [[DISK_BUFFER]]` (lines 3544-3550), `splash screen loop` (lines 4084-4111), `show anim line` (lines 9146-9233); `ROW_ADDR2` -- used in `routines` (lines 589-605), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121), `level draw routine` (lines 2704-2728), `draw wipe block` (lines 3095-3117)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `routines` (lines 572-584)

```asm
    ORG     $7A31
ROW_TO_ADDR:
    SUBROUTINE
    ; Enter routine with Y set to row. Base address
    ; (for column 0) will be placed in ROW_ADDR.

    LDA     ROW_TO_OFFSET_LO,Y 
    STA     ROW_ADDR
    LDA     ROW_TO_OFFSET_HI,Y 
    ORA     HGR_PAGE
    STA     ROW_ADDR+1
    RTS
```

**Previous:** `routines` (lines 451-547)  
**Next:** `routines` (lines 589-605)  
**Defines:** `ROW_TO_ADDR` -- used in `routines` (lines 808-908), `splash screen loop` (lines 4084-4111), `show anim line` (lines 9146-9233)  
**Uses:** `HGR_PAGE` -- `defines` (lines 566-569); `ROW_ADDR` -- `defines` (lines 566-569); `ROW_TO_OFFSET_HI` -- `tables` (lines 558-563); `ROW_TO_OFFSET_LO` -- `tables` (lines 558-563)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

There's also a routine to load the address for both page 1 and page 2.

**Chunk:** `routines` (lines 589-605)

```asm
    ORG     $7A3E
ROW_TO_ADDR_FOR_BOTH_PAGES:
    SUBROUTINE
    ; Enter routine with Y set to row. Base address
    ; (for column 0) will be placed in ROW_ADDR (for page 1)
    ; and ROW_ADDR2 (for page 2).

    LDA     ROW_TO_OFFSET_LO,Y 
    STA     ROW_ADDR
    STA     ROW_ADDR2
    LDA     ROW_TO_OFFSET_HI,Y 
    ORA     #$20
    STA     ROW_ADDR+1
    EOR     #$60
    STA     ROW_ADDR2+1
    RTS
```

**Previous:** `routines` (lines 572-584)  
**Next:** `routines` (lines 663-677)  
**Defines:** `ROW_TO_ADDR_FOR_BOTH_PAGES` -- used in `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121), `Draw wipe for south part` (lines 2985-3007), `Draw wipe for north part` (lines 3010-3033), `Draw wipe for north2 part` (lines 3036-3059), `Draw wipe for south2 part` (lines 3062-3088)  
**Uses:** `ROW_ADDR` -- `defines` (lines 566-569); `ROW_ADDR2` -- `defines` (lines 566-569); `ROW_TO_OFFSET_HI` -- `tables` (lines 558-563); `ROW_TO_OFFSET_LO` -- `tables` (lines 558-563)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

Lode Runner's screens are organized into 28 sprites across by 17 sprites
down. To convert between sprite coordinates and screen coordinates and vice-versa, we
use tables and lookup routines. Each sprite is 10 pixels across by 11 pixels down.

Note that the last row is used for the status, so actually the game screen is 16 sprites vertically.

**Chunk:** `defines` (lines 614-616)

```asm
MAX_GAME_COL        EQU     #27     ; 0x1B
MAX_GAME_ROW        EQU     #15     ; 0x0F
```

**Previous:** `defines` (lines 566-569)  
**Next:** `defines` (lines 778-785)  
**Defines:** None  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `tables` (lines 619-656)

```asm
    ORG     $1C35
HALF_SCREEN_COL_TABLE:
    ; 28 cols of 5 double-pixels each
    HEX     00 05 0a 0f 14 19 1e 23 28 2d 32 37 3c 41 46 4b
    HEX     50 55 5a 5f 64 69 6e 73 78 7d 82 87
SCREEN_ROW_TABLE:
    ; 17 rows of 11 pixels each
    HEX     00 0B 16 21 2C 37 42 4D 58 63 6E 79 84 8F 9A A5
    HEX     B5
COL_BYTE_TABLE:
    ; Byte number
    HEX     00 01 02 04 05 07 08 0A 0B 0C 0E 0F 11 12 14 15
    HEX     16 18 19 1B 1C 1E 1F 20 22 23 25 26
COL_SHIFT_TABLE:
    ; Right shift amount
    HEX     00 03 06 02 05 01 04 00 03 06 02 05 01 04 00 03
    HEX     06 02 05 01 04 00 03 06 02 05 01 04
HALF_SCREEN_COL_BYTE_TABLE:
    HEX     00 00 00 00 01 01 01 02 02 02 02 03 03 03 04 04
    HEX     04 04 05 05 05 06 06 06 06 07 07 07 08 08 08 08
    HEX     09 09 09 0A 0A 0A 0A 0B 0B 0B 0C 0C 0C 0C 0D 0D
    HEX     0D 0E 0E 0E 0E 0F 0F 0F 10 10 10 10 11 11 11 12
    HEX     12 12 12 13 13 13 14 14 14 14 15 15 15 16 16 16
    HEX     16 17 17 17 18 18 18 18 19 19 19 1A 1A 1A 1A 1B
    HEX     1B 1B 1C 1C 1C 1C 1D 1D 1D 1E 1E 1E 1E 1F 1F 1F
    HEX     20 20 20 20 21 21 21 22 22 22 22 23 23 23 24 24
    HEX     24 24 25 25 25 26 26 26 26 27 27 27
HALF_SCREEN_COL_SHIFT_TABLE:
    HEX     00 02 04 06 01 03 05 00 02 04 06 01 03 05 00 02
    HEX     04 06 01 03 05 00 02 04 06 01 03 05 00 02 04 06
    HEX     01 03 05 00 02 04 06 01 03 05 00 02 04 06 01 03
    HEX     05 00 02 04 06 01 03 05 00 02 04 06 01 03 05 00
    HEX     02 04 06 01 03 05 00 02 04 06 01 03 05 00 02 04
    HEX     06 01 03 05 00 02 04 06 01 03 05 00 02 04 06 01
    HEX     03 05 00 02 04 06 01 03 05 00 02 04 06 01 03 05
    HEX     00 02 04 06 01 03 05 00 02 04 06 01 03 05 00 02
    HEX     04 06 01 03 05 00 02 04 06 01 03 05
```

**Previous:** `tables` (lines 558-563)  
**Next:** `tables` (lines 722-725)  
**Defines:** `COL_BYTE_TABLE` -- used in `routines` (lines 683-696), `routines` (lines 808-908); `COL_SHIFT_TABLE` -- used in `routines` (lines 683-696), `routines` (lines 808-908); `HALF_SCREEN_COL_BYTE_TABLE` -- used in `routines` (lines 702-715); `HALF_SCREEN_COL_SHIFT_TABLE` -- used in `routines` (lines 702-715); `HALF_SCREEN_COL_TABLE` -- used in `routines` (lines 663-677); `SCREEN_ROW_TABLE` -- used in `routines` (lines 663-677), `routines` (lines 808-908)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

Here is the routine to return the screen coordinates for the given sprite coordinates.
The reason that `GET_SCREEN_COORDS_FOR` returns half the screen column coordinate
is that otherwise the screen column coordinate wouldn't fit in a register.

**Chunk:** `routines` (lines 663-677)

```asm
    ORG     $885D
GET_SCREEN_COORDS_FOR:
    SUBROUTINE
    ; Enter routine with Y set to sprite row (0-16) and
    ; X set to sprite column (0-27). On return, Y will be set to
    ; screen row, and X is set to half screen column.

    LDA     SCREEN_ROW_TABLE,Y 
    PHA
    LDA     HALF_SCREEN_COL_TABLE,X 
    TAX                         ; X = HALF_SCREEN_COL_TABLE[X]
    PLA
    TAY                         ; Y = SCREEN_ROW_TABLE[Y]
    RTS
```

**Previous:** `routines` (lines 589-605)  
**Next:** `routines` (lines 683-696)  
**Defines:** `GET_SCREEN_COORDS_FOR` -- used in `routines` (lines 728-745), `routines` (lines 754-771), `routines` (lines 808-908), `handle timers` (lines 4114-4290), `check for gold picked up by player` (lines 4615-4659), `try digging left` (lines 5785-5913), `try digging right` (lines 5916-6046), `do ladders` (lines 6277-6343), `check for gold picked up by guard` (lines 6434-6478), `move guard` (lines 6751-6972), `guard drop gold` (lines 7554-7589)  
**Uses:** `HALF_SCREEN_COL_TABLE` -- `tables` (lines 619-656); `SCREEN_ROW_TABLE` -- `tables` (lines 619-656)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

This routine takes a sprite column and converts it to
the memory-mapped byte offset and right-shift amount.

**Chunk:** `routines` (lines 683-696)

```asm
    ORG     $8868
GET_BYTE_AND_SHIFT_FOR_COL:
    SUBROUTINE
    ; Enter routine with X set to sprite column. On
    ; return, A will be set to screen column byte number
    ; and X will be set to an additional right shift amount.

    LDA     COL_BYTE_TABLE,X 
    PHA                         ; A = COL_BYTE_TABLE[X]
    LDA     COL_SHIFT_TABLE,X 
    TAX                         ; X = COL_SHIFT_TABLE[X]
    PLA
    RTS
```

**Previous:** `routines` (lines 663-677)  
**Next:** `routines` (lines 702-715)  
**Defines:** `GET_BYTE_AND_SHIFT_FOR_COL` -- used in `routines` (lines 808-908)  
**Uses:** `COL_BYTE_TABLE` -- `tables` (lines 619-656); `COL_SHIFT_TABLE` -- `tables` (lines 619-656)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

This routine takes half the screen column coordinate and converts it to
the memory-mapped byte offset and right-shift amount.

**Chunk:** `routines` (lines 702-715)

```asm
    ORG     $8872
GET_BYTE_AND_SHIFT_FOR_HALF_SCREEN_COL:
    SUBROUTINE
    ; Enter routine with X set to half screen column. On
    ; return, A will be set to screen column byte number
    ; and X will be set to an additional right shift amount.

    LDA     HALF_SCREEN_COL_BYTE_TABLE,X 
    PHA                         ; A = HALF_SCREEN_COL_BYTE_TABLE[X]
    LDA     HALF_SCREEN_COL_SHIFT_TABLE,X 
    TAX                         ; X = HALF_SCREEN_COL_SHIFT_TABLE[X]
    PLA
    RTS
```

**Previous:** `routines` (lines 683-696)  
**Next:** `routines` (lines 728-745)  
**Defines:** `GET_BYTE_AND_SHIFT_FOR_HALF_SCREEN_COL` -- used in `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121)  
**Uses:** `HALF_SCREEN_COL_BYTE_TABLE` -- `tables` (lines 619-656); `HALF_SCREEN_COL_SHIFT_TABLE` -- `tables` (lines 619-656)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

We also have some utility routines that let us take a sprite row or column and
get its screen row or half column, but offset in either row or column by anywhere from
`-2` to `+2`.

**Chunk:** `tables` (lines 722-725)

```asm
    ORG     $888A
ROW_OFFSET_TABLE:
    HEX     FB FD 00 02 04
```

**Previous:** `tables` (lines 619-656)  
**Next:** `tables` (lines 748-751)  
**Defines:** `ROW_OFFSET_TABLE` -- used in `routines` (lines 728-745)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `routines` (lines 728-745)

```asm
    ORG     $887C
GET_SCREEN_ROW_OFFSET_IN_X_FOR:
    SUBROUTINE
    ; Enter routine with X set to offset+2 (in double-pixels) and
    ; Y set to sprite row. On return, X will retain its value and
    ; Y will be set to the screen row.

    TXA
    PHA
    JSR     GET_SCREEN_COORDS_FOR
    PLA
    TAX                                 ; Restore X
    TYA
    CLC
    ADC     ROW_OFFSET_TABLE,X
    TAY
    RTS
```

**Previous:** `routines` (lines 702-715)  
**Next:** `routines` (lines 754-771)  
**Defines:** `GET_SCREEN_ROW_OFFSET_IN_X_FOR` -- used in `get player sprite and coord data` (lines 4557-4574), `get guard sprite and coords` (lines 6558-6575)  
**Uses:** `GET_SCREEN_COORDS_FOR` -- `routines` (lines 663-677); `ROW_OFFSET_TABLE` -- `tables` (lines 722-725)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `tables` (lines 748-751)

```asm
    ORG     $889D
COL_OFFSET_TABLE:
    HEX     FE FF 00 01 02
```

**Previous:** `tables` (lines 722-725)  
**Next:** `tables` (lines 788-805)  
**Defines:** `COL_OFFSET_TABLE` -- used in `routines` (lines 754-771)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `routines` (lines 754-771)

```asm
    ORG     $888F
GET_HALF_SCREEN_COL_OFFSET_IN_Y_FOR:
    SUBROUTINE
    ; Enter routine with Y set to offset+2 (in double-pixels) and
    ; X set to sprite column. On return, Y will retain its value and
    ; X will be set to the half screen column.

    TYA
    PHA
    JSR     GET_SCREEN_COORDS_FOR
    PLA
    TAY                                 ; Restore Y
    TXA
    CLC
    ADC     COL_OFFSET_TABLE,Y
    TAX
    RTS
```

**Previous:** `routines` (lines 728-745)  
**Next:** `routines` (lines 808-908)  
**Defines:** `GET_HALF_SCREEN_COL_OFFSET_IN_Y_FOR` -- used in `get player sprite and coord data` (lines 4557-4574), `get guard sprite and coords` (lines 6558-6575)  
**Uses:** `COL_OFFSET_TABLE` -- `tables` (lines 748-751); `GET_SCREEN_COORDS_FOR` -- `routines` (lines 663-677)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

Now we can finally write the routines that draw a sprite on the screen. We have one
routine that draws a sprite at a given game row and game column.
There are two entry points, one to draw on HGR1, and one for HGR2.

**Chunk:** `defines` (lines 778-785)

```asm
ROWNUM          EQU     $1B
COLNUM          EQU     $1C
MASK0           EQU     $50
MASK1           EQU     $51
COL_SHIFT_AMT   EQU     $71
GAME_COLNUM     EQU     $85
GAME_ROWNUM     EQU     $86
```

**Previous:** `defines` (lines 614-616)  
**Next:** `defines` (lines 1013-1014)  
**Defines:** `COLNUM` -- used in `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121); `COL_SHIFT_AMT` -- used in `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121); `GAME_COLNUM` -- used in `routines` (lines 808-908), `put char` (lines 1220-1248), `put digit` (lines 1294-1312), `add and update score` (lines 1396-1448), `put status` (lines 1463-1555), `get level from keyboard` (lines 2187-2316), `level draw routine` (lines 2490-2494), `level draw routine` (lines 2674-2684), `level draw routine` (lines 2733-2766), `uncompress row data` (lines 3721-3756), `dead code` (lines 3783-3831), `construct and display high score screen` (lines 3870-3885), `handle timers` (lines 4114-4290), `check for gold picked up by player` (lines 4615-4659), `return handler` (lines 4925-5050), `try digging left` (lines 5785-5913), `try digging right` (lines 5916-6046), `drop player in hole` (lines 6049-6092), `do ladders` (lines 6277-6343), `check for gold picked up by guard` (lines 6434-6478), `guard resurrections` (lines 6590-6679), `move guard` (lines 6751-6972), `guard drop gold` (lines 7554-7589), `record hi score data` (lines 8305-8487), `bad data disk` (lines 8490-8510), `dont manipulate master disk` (lines 8513-8532), `game loop` (lines 8889-9017), `level editor` (lines 9257-9343), `editor edit level` (lines 9492-9564), `get key for edit level` (lines 9567-9578), `level editor key functions` (lines 9605-9760); `GAME_ROWNUM` -- used in `routines` (lines 808-908), `put char` (lines 1220-1248), `add and update score` (lines 1396-1448), `put status` (lines 1463-1555), `level draw routine` (lines 2332-2342), `level draw routine` (lines 2540-2554), `level draw routine` (lines 2590-2616), `level draw routine` (lines 2637-2657), `level draw routine` (lines 2674-2684), `level draw routine` (lines 2733-2766), `Initialize level counts` (lines 3655-3679), `uncompress level data` (lines 3692-3708), `next compressed row for [[row_loop]]` (lines 3759-3763), `dead code` (lines 3783-3831), `construct and display high score screen` (lines 3870-3885), `splash screen` (lines 4042-4067), `splash screen loop` (lines 4084-4111), `handle timers` (lines 4114-4290), `ready yourself` (lines 4354-4363), `no button pressed` (lines 4392-4405), `long delay attract mode` (lines 4522-4532), `check for gold picked up by player` (lines 4615-4659), `return handler` (lines 4925-5050), `try digging left` (lines 5785-5913), `try digging right` (lines 5916-6046), `drop player in hole` (lines 6049-6092), `do ladders` (lines 6277-6343), `check for gold picked up by guard` (lines 6434-6478), `guard resurrections` (lines 6590-6679), `move guard` (lines 6751-6972), `guard drop gold` (lines 7554-7589), `record hi score data` (lines 8305-8487), `bad data disk` (lines 8490-8510), `dont manipulate master disk` (lines 8513-8532), `game loop` (lines 8889-9017), `show anim line` (lines 9146-9233), `level editor` (lines 9257-9343), `editor edit level` (lines 9492-9564), `get key for edit level` (lines 9567-9578), `level editor key functions` (lines 9605-9760); `MASK0` -- used in `routines` (lines 808-908), `access hi score data` (lines 8246-8294); `MASK1` -- used in `routines` (lines 808-908); `ROWNUM` -- used in `routines` (lines 808-908), `erase sprite at screen coordinate` (lines 919-1007), `draw sprite at screen coordinate` (lines 1017-1121), `determine guard left right limits` (lines 7100-7197)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `tables` (lines 788-805)

```asm
    ORG     $8328
PIXEL_MASK0:
    BYTE    %00000000
    BYTE    %00000001
    BYTE    %00000011
    BYTE    %00000111
    BYTE    %00001111
    BYTE    %00011111
    BYTE    %00111111
PIXEL_MASK1:
    BYTE    %11111000
    BYTE    %11110000
    BYTE    %11100000
    BYTE    %11000000
    BYTE    %10000000
    BYTE    %11111110
    BYTE    %11111100
```

**Previous:** `tables` (lines 748-751)  
**Next:** `tables` (lines 1718-1721)  
**Defines:** `PIXEL_MASK0` -- used in `routines` (lines 808-908); `PIXEL_MASK1` -- used in `routines` (lines 808-908)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `routines` (lines 808-908)

```asm
    ORG     $82AA
DRAW_SPRITE_PAGE1:
    SUBROUTINE
    ; Enter routine with A set to sprite number to draw,
    ; GAME_ROWNUM set to the row to draw it at, and GAME_COLNUM
    ; set to the column to draw it at.

    STA     SPRITE_NUM
    LDA     #$20                ; Page number for HGR1
    BNE     DRAW_SPRITE         ; Actually unconditional jump

DRAW_SPRITE_PAGE2:
    SUBROUTINE
    ; Enter routine with A set to sprite number to draw,
    ; GAME_ROWNUM set to the row to draw it at, and GAME_COLNUM
    ; set to the column to draw it at.

    STA     SPRITE_NUM
    LDA     #$40                ; Page number for HGR2
    ; fallthrough

DRAW_SPRITE:
    STA     HGR_PAGE
    LDY     GAME_ROWNUM
    JSR     GET_SCREEN_COORDS_FOR
    STY     ROWNUM              ; ROWNUM = SCREEN_ROW_TABLE[GAME_ROWNUM]

    LDX     GAME_COLNUM
    JSR     GET_BYTE_AND_SHIFT_FOR_COL
    STA     COLNUM              ; COLNUM = COL_BYTE_TABLE[GAME_COLNUM]
    STX     COL_SHIFT_AMT       ; COL_SHIFT_AMT = COL_SHIFT_TABLE[GAME_COLNUM]

    LDA     PIXEL_MASK0,X 
    STA     MASK0               ; MASK0 = PIXEL_MASK0[COL_SHIFT_AMT]
    LDA     PIXEL_MASK1,X 
    STA     MASK1               ; MASK1 = PIXEL_MASK1[COL_SHIFT_AMT]

    JSR     COMPUTE_SHIFTED_SPRITE

    LDA     #$0B
    STA     ROW_COUNT
    LDX     #$00
    LDA     COL_SHIFT_AMT
    CMP     #$05
    BCS     .need_3_bytes       ; If COL_SHIFT_AMT >= 5, we need to alter three
                                ; screen bytes, otherwise just two bytes.

.loop1:
    LDY     ROWNUM
    JSR     ROW_TO_ADDR
    LDY     COLNUM
    LDA     (ROW_ADDR),Y 
    AND     MASK0
    ORA     BLOCK_DATA,X 
    STA     (ROW_ADDR),Y        ; screen[COLNUM] =
                                ;   screen[COLNUM] & MASK0 | BLOCK_DATA[i]

    INX                         ; X++
    INY                         ; Y++
    LDA     (ROW_ADDR),Y 
    AND     MASK1
    ORA     BLOCK_DATA,X 
    STA     (ROW_ADDR),Y        ; screen[COLNUM+1] =
                                ;   screen[COLNUM+1] & MASK1 | BLOCK_DATA[i+1]

    INX
    INX                         ; X += 2
    INC     ROWNUM              ; ROWNUM++
    DEC     ROW_COUNT           ; ROW_COUNT--
    BNE     .loop1              ; loop while ROW_COUNT > 0
    RTS

.need_3_bytes
    LDY     ROWNUM
    JSR     ROW_TO_ADDR
    LDY     COLNUM
    LDA     (ROW_ADDR),Y 
    AND     MASK0
    ORA     BLOCK_DATA,X 
    STA     (ROW_ADDR),Y        ; screen[COLNUM] =
                                ;   screen[COLNUM] & MASK0 | BLOCK_DATA[i]

    INX                         ; X++
    INY                         ; Y++
    LDA     BLOCK_DATA,X 
    STA     (ROW_ADDR),Y        ; screen[COLNUM+1] = BLOCK_DATA[i+1]

    INX                         ; X++
    INY                         ; Y++
    LDA     (ROW_ADDR),Y 
    AND     MASK1
    ORA     BLOCK_DATA,X 
    STA     (ROW_ADDR),Y        ; screen[COLNUM+2] =
                                ;   screen[COLNUM+2] & MASK1 | BLOCK_DATA[i+2]

    INX                         ; X++
    INC     ROWNUM              ; ROWNUM++
    DEC     ROW_COUNT           ; ROW_COUNT--
    BNE     .need_3_bytes       ; loop while ROW_COUNT > 0
    RTS
```

**Previous:** `routines` (lines 754-771)  
**Next:** `routines` (lines 2830-2867)  
**Defines:** `DRAW_SPRITE_PAGE1` -- used in `put char` (lines 1220-1248), `put digit` (lines 1294-1312), `wait for key page1` (lines 2074-2124), `handle timers` (lines 4114-4290), `try digging left` (lines 5785-5913), `try digging right` (lines 5916-6046), `drop player in hole` (lines 6049-6092), `guard resurrections` (lines 6590-6679), `editor edit level` (lines 9492-9564); `DRAW_SPRITE_PAGE2` -- used in `put char` (lines 1220-1248), `put digit` (lines 1294-1312), `wait for key` (lines 2026-2071), `level draw routine` (lines 2674-2684), `level draw routine` (lines 2733-2766), `handle timers` (lines 4114-4290), `check for gold picked up by player` (lines 4615-4659), `return handler` (lines 4925-5050), `drop player in hole` (lines 6049-6092), `do ladders` (lines 6277-6343), `check for gold picked up by guard` (lines 6434-6478), `guard resurrections` (lines 6590-6679), `move guard` (lines 6751-6972), `guard drop gold` (lines 7554-7589)  
**Uses:** `BLOCK_DATA` -- `defines` (lines 335-336); `COLNUM` -- `defines` (lines 778-785); `COL_BYTE_TABLE` -- `tables` (lines 619-656); `COL_SHIFT_AMT` -- `defines` (lines 778-785); `COL_SHIFT_TABLE` -- `tables` (lines 619-656); `COMPUTE_SHIFTED_SPRITE` -- `routines` (lines 451-547); `GAME_COLNUM` -- `defines` (lines 778-785); `GAME_ROWNUM` -- `defines` (lines 778-785); `GET_BYTE_AND_SHIFT_FOR_COL` -- `routines` (lines 683-696); `GET_SCREEN_COORDS_FOR` -- `routines` (lines 663-677); `HGR_PAGE` -- `defines` (lines 566-569); `MASK0` -- `defines` (lines 778-785); `MASK1` -- `defines` (lines 778-785); `PIXEL_MASK0` -- `tables` (lines 788-805); `PIXEL_MASK1` -- `tables` (lines 788-805); `ROWNUM` -- `defines` (lines 778-785); `ROW_ADDR` -- `defines` (lines 566-569); `ROW_COUNT` -- `defines` (lines 446-448); `ROW_TO_ADDR` -- `routines` (lines 572-584); `SCREEN_ROW_TABLE` -- `tables` (lines 619-656); `SPRITE_NUM` -- `defines` (lines 446-448)  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

There is a different routine which erases a sprite at a given screen coordinate.
It does this by drawing the inverse of the sprite on page 1, then drawing the sprite data
from page 2 (the background page) onto page 1.

Upon entry, the Y register needs to be set to the screen row coordinate (0-191). However, the
X register needs to be set to half the screen column coordinate (0-139) because otherwise
the maximum coordinate (279) wouldn't fit in a register.

**Chunk:** `erase sprite at screen coordinate` (lines 919-1007)

```asm
    ORG     $8336
ERASE_SPRITE_AT_PIXEL_COORDS:
    SUBROUTINE
    ; Enter routine with A set to sprite number to draw,
    ; Y set to the screen row to erase it at, and X
    ; set to *half* the screen column to erase it at.

    STY     ROWNUM
    STA     SPRITE_NUM
    JSR     GET_BYTE_AND_SHIFT_FOR_HALF_SCREEN_COL
    STA     COLNUM
    STX     COL_SHIFT_AMT
    JSR     COMPUTE_SHIFTED_SPRITE

    LDX     #$0B
    STX     ROW_COUNT
    LDX     #$00
    LDA     COL_SHIFT_AMT
    CMP     #$05
    BCS     .need_3_bytes       ; If COL_SHIFT_AMT >= 5, we need to alter three
                                ; screen bytes, otherwise just two bytes.

.loop1:
    LDY     ROWNUM
    JSR     ROW_TO_ADDR_FOR_BOTH_PAGES
    LDY     COLNUM
    LDA     BLOCK_DATA,X
    EOR     #$7F
    AND     (ROW_ADDR),Y
    ORA     (ROW_ADDR2),Y
    STA     (ROW_ADDR),Y            ; screen[COLNUM] =
                                    ;   (screen[COLNUM] & (BLOCK_DATA[i] ^ 0x7F)) |
                                    ;   screen2[COLNUM]

    INX                             ; X++
    INY                             ; Y++
    LDA     BLOCK_DATA,X
    EOR     #$7F
    AND     (ROW_ADDR),Y
    ORA     (ROW_ADDR2),Y
    STA     (ROW_ADDR),Y            ; screen[COLNUM+1] =
                                    ;   (screen[COLNUM+1] & (BLOCK_DATA[i+1] ^ 0x7F)) |
                                    ;   screen2[COLNUM+1]

    INX                             ; X++
    INX                             ; X++
    INC     ROWNUM
    DEC     ROW_COUNT
    BNE     .loop1
    RTS

.need_3_bytes:
    LDY     ROWNUM
    JSR     ROW_TO_ADDR_FOR_BOTH_PAGES
    LDY     COLNUM
    LDA     BLOCK_DATA,X
    EOR     #$7F
    AND     (ROW_ADDR),Y
    ORA     (ROW_ADDR2),Y
    STA     (ROW_ADDR),Y            ; screen[COLNUM] =
                                    ;   (screen[COLNUM] & (BLOCK_DATA[i] ^ 0x7F)) |
                                    ;   screen2[COLNUM]

    INX                             ; X++
    INY                             ; Y++
    LDA     BLOCK_DATA,X
    EOR     #$7F
    AND     (ROW_ADDR),Y
    ORA     (ROW_ADDR2),Y
    STA     (ROW_ADDR),Y            ; screen[COLNUM+1] =
                                    ;   (screen[COLNUM+1] & (BLOCK_DATA[i+1] ^ 0x7F)) |
                                    ;   screen2[COLNUM+1]

    INX                             ; X++
    INY                             ; Y++
    LDA     BLOCK_DATA,X
    EOR     #$7F
    AND     (ROW_ADDR),Y
    ORA     (ROW_ADDR2),Y
    STA     (ROW_ADDR),Y            ; screen[COLNUM+2] =
                                    ;   (screen[COLNUM+2] & (BLOCK_DATA[i+2] ^ 0x7F)) |
                                    ;   screen2[COLNUM+2]

    INX                             ; X++
    INC     ROWNUM
    DEC     ROW_COUNT
    BNE     .need_3_bytes
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `ERASE_SPRITE_AT_PIXEL_COORDS` -- used in `handle timers` (lines 4114-4290), `check for gold picked up by player` (lines 4615-4659), `try moving up` (lines 5383-5471), `try moving down` (lines 5502-5558), `try moving left` (lines 5592-5663), `try moving right` (lines 5668-5742), `try digging left` (lines 5785-5913), `try digging right` (lines 5916-6046), `move player` (lines 6097-6255), `check for gold picked up by guard` (lines 6434-6478), `move guard` (lines 6751-6972), `try guard move left` (lines 7200-7289), `try guard move right` (lines 7292-7384), `try guard move up` (lines 7387-7473), `try guard move down` (lines 7476-7546)  
**Uses:** `BLOCK_DATA` -- `defines` (lines 335-336); `COLNUM` -- `defines` (lines 778-785); `COL_SHIFT_AMT` -- `defines` (lines 778-785); `COMPUTE_SHIFTED_SPRITE` -- `routines` (lines 451-547); `GET_BYTE_AND_SHIFT_FOR_HALF_SCREEN_COL` -- `routines` (lines 702-715); `ROWNUM` -- `defines` (lines 778-785); `ROW_ADDR` -- `defines` (lines 566-569); `ROW_ADDR2` -- `defines` (lines 566-569); `ROW_COUNT` -- `defines` (lines 446-448); `ROW_TO_ADDR_FOR_BOTH_PAGES` -- `routines` (lines 589-605); `SPRITE_NUM` -- `defines` (lines 446-448)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

And then there's the corresponding routine to draw a sprite at the given coordinates. The
routine also sets whether the active and the background screens differ in `SCREENS_DIFFER`.

**Chunk:** `defines` (lines 1013-1014)

```asm
SCREENS_DIFFER      EQU     $52
```

**Previous:** `defines` (lines 778-785)  
**Next:** `defines` (lines 1216-1217)  
**Defines:** `SCREENS_DIFFER` -- used in `draw sprite at screen coordinate` (lines 1017-1121), `draw player` (lines 1128-1141)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `draw sprite at screen coordinate` (lines 1017-1121)

```asm
    ORG     $83A7
DRAW_SPRITE_AT_PIXEL_COORDS:
    SUBROUTINE
    ; Enter routine with A set to sprite number to draw,
    ; Y set to the screen row to draw it at, and X
    ; set to *half* the screen column to draw it at.

    STY     ROWNUM
    STA     SPRITE_NUM
    JSR     GET_BYTE_AND_SHIFT_FOR_HALF_SCREEN_COL
    STA     COLNUM
    STX     COL_SHIFT_AMT
    JSR     COMPUTE_SHIFTED_SPRITE

    LDA     #$0B
    STA     ROW_COUNT
    LDX     #$00
    STX     SCREENS_DIFFER      ; SCREENS_DIFFER = 0
    LDA     COL_SHIFT_AMT
    CMP     #$05
    BCS     .need_3_bytes       ; If COL_SHIFT_AMT >= 5, we need to alter three
                                ; screen bytes, otherwise just two bytes.

.loop1:
    LDY     ROWNUM
    JSR     ROW_TO_ADDR_FOR_BOTH_PAGES
    LDY     COLNUM
    LDA     (ROW_ADDR),Y
    EOR     (ROW_ADDR2),Y
    AND     BLOCK_DATA,X
    ORA     SCREENS_DIFFER
    STA     SCREENS_DIFFER          ; SCREENS_DIFFER |= 
                                    ;   ( (screen[COLNUM] ^ screen2[COLNUM]) &
                                    ;     BLOCK_DATA[i])
    LDA     BLOCK_DATA,X            
    ORA     (ROW_ADDR),Y            
    STA     (ROW_ADDR),Y            ; screen[COLNUM] |= BLOCK_DATA[i]

    INX                             ; X++
    INY                             ; Y++
    LDA     (ROW_ADDR),Y
    EOR     (ROW_ADDR2),Y
    AND     BLOCK_DATA,X
    ORA     SCREENS_DIFFER
    STA     SCREENS_DIFFER          ; SCREENS_DIFFER |= 
                                    ;   ( (screen[COLNUM+1] ^ screen2[COLNUM+1]) &
                                    ;     BLOCK_DATA[i+1])
    LDA     BLOCK_DATA,X            
    ORA     (ROW_ADDR),Y            
    STA     (ROW_ADDR),Y            ; screen[COLNUM+1] |= BLOCK_DATA[i+1]

    INX                             ; X++
    INX                             ; X++
    INC     ROWNUM
    DEC     ROW_COUNT
    BNE     .loop1
    RTS

.need_3_bytes:
    LDY     ROWNUM
    JSR     ROW_TO_ADDR_FOR_BOTH_PAGES
    LDY     COLNUM
    LDA     (ROW_ADDR),Y
    EOR     (ROW_ADDR2),Y
    AND     BLOCK_DATA,X
    ORA     SCREENS_DIFFER
    STA     SCREENS_DIFFER          ; SCREENS_DIFFER |= 
                                    ;   ( (screen[COLNUM] ^ screen2[COLNUM]) &
                                    ;     BLOCK_DATA[i])
    LDA     BLOCK_DATA,X            
    ORA     (ROW_ADDR),Y            
    STA     (ROW_ADDR),Y            ; screen[COLNUM] |= BLOCK_DATA[i]

    INX                             ; X++
    INY                             ; Y++
    LDA     (ROW_ADDR),Y
    EOR     (ROW_ADDR2),Y
    AND     BLOCK_DATA,X
    ORA     SCREENS_DIFFER
    STA     SCREENS_DIFFER          ; SCREENS_DIFFER |= 
                                    ;   ( (screen[COLNUM+1] ^ screen2[COLNUM+1]) &
                                    ;     BLOCK_DATA[i+1])
    LDA     BLOCK_DATA,X            
    ORA     (ROW_ADDR),Y            
    STA     (ROW_ADDR),Y            ; screen[COLNUM+1] |= BLOCK_DATA[i+1]

    INX                             ; X++
    INY                             ; Y++
    LDA     (ROW_ADDR),Y
    EOR     (ROW_ADDR2),Y
    AND     BLOCK_DATA,X
    ORA     SCREENS_DIFFER
    STA     SCREENS_DIFFER          ; SCREENS_DIFFER |= 
                                    ;   ( (screen[COLNUM+2] ^ screen2[COLNUM+2]) &
                                    ;     BLOCK_DATA[i+2])
    LDA     BLOCK_DATA,X            
    ORA     (ROW_ADDR),Y            
    STA     (ROW_ADDR),Y            ; screen[COLNUM+2] |= BLOCK_DATA[i+2]

    INX                             ; X++
    INC     ROWNUM
    DEC     ROW_COUNT
    BNE     .need_3_bytes
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `DRAW_SPRITE_AT_PIXEL_COORDS` -- used in `draw player` (lines 1128-1141), `try digging left` (lines 5785-5913), `try digging right` (lines 5916-6046), `do ladders` (lines 6277-6343), `guard resurrections` (lines 6590-6679), `move guard` (lines 6751-6972), `try guard move left` (lines 7200-7289), `try guard move right` (lines 7292-7384), `try guard move up` (lines 7387-7473), `guard drop gold` (lines 7554-7589)  
**Uses:** `BLOCK_DATA` -- `defines` (lines 335-336); `COLNUM` -- `defines` (lines 778-785); `COL_SHIFT_AMT` -- `defines` (lines 778-785); `COMPUTE_SHIFTED_SPRITE` -- `routines` (lines 451-547); `GET_BYTE_AND_SHIFT_FOR_HALF_SCREEN_COL` -- `routines` (lines 702-715); `ROWNUM` -- `defines` (lines 778-785); `ROW_ADDR` -- `defines` (lines 566-569); `ROW_ADDR2` -- `defines` (lines 566-569); `ROW_COUNT` -- `defines` (lines 446-448); `ROW_TO_ADDR_FOR_BOTH_PAGES` -- `routines` (lines 589-605); `SCREENS_DIFFER` -- `defines` (lines 1013-1014); `SPRITE_NUM` -- `defines` (lines 446-448)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

There is a special routine to draw the player sprite at the player's location. If
the two pages at the player's location are different and the player didn't pick up
gold (which would explain the difference), then the player is killed.

**Chunk:** `draw player` (lines 1128-1141)

```asm
    ORG     $6C02
DRAW_PLAYER:
    SUBROUTINE

    JSR     GET_SPRITE_AND_SCREEN_COORD_AT_PLAYER
    JSR     DRAW_SPRITE_AT_PIXEL_COORDS
    LDA     SCREENS_DIFFER
    BEQ     .end
    LDA     DIDNT_PICK_UP_GOLD
    BEQ     .end
    LSR     ALIVE       ; Set player as dead
.end
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `DRAW_PLAYER` -- used in `try moving up` (lines 5383-5471), `try moving left` (lines 5592-5663), `try moving right` (lines 5668-5742), `try digging left` (lines 5785-5913), `try digging right` (lines 5916-6046), `move player` (lines 6097-6255)  
**Uses:** `ALIVE` -- `defines` (lines 3623-3624); `DIDNT_PICK_UP_GOLD` -- `defines` (lines 4611-4612); `DRAW_SPRITE_AT_PIXEL_COORDS` -- `draw sprite at screen coordinate` (lines 1017-1121); `GET_SPRITE_AND_SCREEN_COORD_AT_PLAYER` -- `get player sprite and coord data` (lines 4557-4574); `SCREENS_DIFFER` -- `defines` (lines 1013-1014)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

## Printing strings


Now that we can put sprites onto the screen at any game coordinate, we can
also have some routines that print strings. We saw above that we have
letter and number sprites, plus some punctuation. Letters and punctuation
are always blue, while numbers are always orange.

There is a basic routine to put a character at the current `GAME_COLNUM`
and `GAME_ROWNUM`, incrementing this "cursor", and putting it at the beginning
of the next line if we "print" a newline character.

We first define a routine to convert the ASCII code of a character to
its sprite number. Lode Runner sets the high bit of the code to
make it be treated as ASCII.

**Chunk:** `char to sprite num` (lines 1159-1210)

```asm
    ORG     $7B2A
CHAR_TO_SPRITE_NUM:
    SUBROUTINE
    ; Enter routine with A set to the ASCII code of the
    ; character to convert to sprite number, with the high bit set.
    ; The sprite number is returned in A.

    CMP     #$C1                    ; 'A' -> sprite 69
    BCC     .not_letter
    CMP     #$DB                    ; 'Z' -> sprite 94
    BCC     .letter

.not_letter:
    ; On return, we will subtract 0x7C from X to
    ; get the actual sprite. This is to make A-Z
    ; easier to handle.
    LDX     #$7C
    CMP     #$A0                    ; ' ' -> sprite 0
    BEQ     .end
    LDX     #$DB
    CMP     #$BE                    ; '>' -> sprite 95
    BEQ     .end
    INX
    CMP     #$AE                    ; '.' -> sprite 96
    BEQ     .end
    INX
    CMP     #$A8                    ; '(' -> sprite 97
    BEQ     .end
    INX
    CMP     #$A9                    ; ')' -> sprite 98
    BEQ     .end
    INX
    CMP     #$AF                    ; '/' -> sprite 99
    BEQ     .end
    INX
    CMP     #$AD                    ; '-' -> sprite 100
    BEQ     .end
    INX
    CMP     #$BC                    ; '<' -> sprite 101
    BEQ     .end
    LDA     #$10                    ; sprite 16: just one of the man sprites
    RTS

.end:
    TXA

.letter:
    SEC
    SBC     #$7C                    
    RTS

```

**Previous:** None  
**Next:** None  
**Defines:** `CHAR_TO_SPRITE_NUM` -- used in `put char` (lines 1220-1248), `record hi score data` (lines 8305-8487)  
**Uses:** None  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

Now we can define the routine to put a character on the screen at the
current position.

**Chunk:** `defines` (lines 1216-1217)

```asm
DRAW_PAGE   EQU     $87     ; 0x20 for page 1, 0x40 for page 2
```

**Previous:** `defines` (lines 1013-1014)  
**Next:** `defines` (lines 1256-1257)  
**Defines:** `DRAW_PAGE` -- used in `put char` (lines 1220-1248), `put digit` (lines 1294-1312), `put status` (lines 1463-1555), `construct and display high score screen` (lines 3870-3885), `show high score page` (lines 4030-4035), `splash screen` (lines 4042-4067), `record hi score data` (lines 8305-8487), `bad data disk` (lines 8490-8510), `dont manipulate master disk` (lines 8513-8532), `level editor` (lines 9257-9343), `editor edit level` (lines 9492-9564), `level editor key functions` (lines 9605-9760)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `put char` (lines 1220-1248)

```asm
    ORG     $7B64
PUT_CHAR:
    SUBROUTINE
    ; Enter routine with A set to the ASCII code of the
    ; character to put on the screen, with the high bit set.

    CMP     #$8D
    BEQ     NEWLINE                 ; If newline, do NEWLINE instead.
    JSR     CHAR_TO_SPRITE_NUM
    LDX     DRAW_PAGE
    CPX     #$40
    BEQ     .draw_to_page2

    JSR     DRAW_SPRITE_PAGE1
    INC     GAME_COLNUM
    RTS

.draw_to_page2
    JSR     DRAW_SPRITE_PAGE2
    INC     GAME_COLNUM
    RTS

NEWLINE:
    SUBROUTINE
    INC     GAME_ROWNUM
    LDA     #$00
    STA     GAME_COLNUM
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `NEWLINE` -- used in `next high score row` (lines 4016-4023); `PUT_CHAR` -- used in `put string` (lines 1260-1288), `draw high score row number` (lines 3914-3932), `draw high score initials` (lines 3943-3963), `record hi score data` (lines 8305-8487)  
**Uses:** `CHAR_TO_SPRITE_NUM` -- `char to sprite num` (lines 1159-1210); `DRAW_PAGE` -- `defines` (lines 1216-1217); `DRAW_SPRITE_PAGE1` -- `routines` (lines 808-908); `DRAW_SPRITE_PAGE2` -- `routines` (lines 808-908); `GAME_COLNUM` -- `defines` (lines 778-785); `GAME_ROWNUM` -- `defines` (lines 778-785)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

The `PUT_STRING` routine uses `PUT_CHAR` to put a string on the screen. Rather than take
an address pointing to a string, instead it uses the return address as the source for data.
It then has to fix up the actual return address at the end to be just after the zero-terminating
byte of the string.

**Chunk:** `defines` (lines 1256-1257)

```asm
SAVED_RET_ADDR      EQU     $10     ; 2 bytes
```

**Previous:** `defines` (lines 1216-1217)  
**Next:** `defines` (lines 1322-1325)  
**Defines:** `SAVED_RET_ADDR` -- used in `put string` (lines 1260-1288), `load sound data` (lines 1614-1651)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `put string` (lines 1260-1288)

```asm
    ORG     $86E0
PUT_STRING:
    SUBROUTINE

    PLA
    STA     SAVED_RET_ADDR
    PLA
    STA     SAVED_RET_ADDR+1
    BNE     .next

.loop:
    LDY     #$00
    LDA     (SAVED_RET_ADDR),Y
    BEQ     .end
    JSR     PUT_CHAR

.next:
    INC     SAVED_RET_ADDR
    BNE     .loop
    INC     SAVED_RET_ADDR+1
    BNE     .loop

.end:
    LDA     SAVED_RET_ADDR+1
    PHA
    LDA     SAVED_RET_ADDR
    PHA
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `PUT_STRING` -- used in `put status` (lines 1463-1555), `hit key to continue` (lines 2148-2167), `draw high score table header` (lines 3888-3900), `draw high score row number` (lines 3914-3932), `draw high score initials` (lines 3943-3963), `draw high score level` (lines 3966-3979), `bad data disk` (lines 8490-8510), `dont manipulate master disk` (lines 8513-8532), `editor initialize disk` (lines 8613-8731), `editor clear high scores` (lines 8741-8786), `level editor` (lines 9257-9343), `editor clear level` (lines 9350-9376), `editor move level` (lines 9390-9440), `editor play level` (lines 9443-9471), `editor edit level` (lines 9474-9489), `level editor key functions` (lines 9605-9760)  
**Uses:** `PUT_CHAR` -- `put char` (lines 1220-1248); `SAVED_RET_ADDR` -- `defines` (lines 1256-1257)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

Like `PUT_CHAR`, we also have `PUT_DIGIT` which draws the sprite corresponding
to digits 0 to 9 at the current position, incrementing the cursor.

**Chunk:** `put digit` (lines 1294-1312)

```asm
    ORG     $7B15
PUT_DIGIT:
    SUBROUTINE
    ; Enter routine with A set to the digit to put on the screen.

    CLC
    ADC     #$3B                    ; '0' -> sprite 59, '9' -> sprite 68.
    LDX     DRAW_PAGE
    CPX     #$40
    BEQ     .draw_to_page2
    JSR     DRAW_SPRITE_PAGE1
    INC     GAME_COLNUM
    RTS

.draw_to_page2:
    JSR     DRAW_SPRITE_PAGE2
    INC     GAME_COLNUM
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `PUT_DIGIT` -- used in `add and update score` (lines 1396-1448), `put status` (lines 1463-1555), `get level from keyboard` (lines 2187-2316), `draw high score row number` (lines 3914-3932), `draw high score level` (lines 3966-3979), `draw high score` (lines 3982-4013)  
**Uses:** `DRAW_PAGE` -- `defines` (lines 1216-1217); `DRAW_SPRITE_PAGE1` -- `routines` (lines 808-908); `DRAW_SPRITE_PAGE2` -- `routines` (lines 808-908); `GAME_COLNUM` -- `defines` (lines 778-785)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

## Numbers


We also need a way to put numbers on the screen.

First, a routine to convert a one-byte decimal number into hundreds,
tens, and units.

**Chunk:** `defines` (lines 1322-1325)

```asm
HUNDREDS        EQU     $89
TENS            EQU     $8A
UNITS           EQU     $8B
```

**Previous:** `defines` (lines 1256-1257)  
**Next:** `defines` (lines 1381-1382)  
**Defines:** `HUNDREDS` -- used in `to decimal3` (lines 1328-1354), `put status` (lines 1463-1555), `get level from keyboard` (lines 2187-2316), `draw high score level` (lines 3966-3979); `TENS` -- used in `to decimal3` (lines 1328-1354), `bcd to decimal2` (lines 1359-1374), `add and update score` (lines 1396-1448), `put status` (lines 1463-1555), `get level from keyboard` (lines 2187-2316), `draw high score level` (lines 3966-3979), `draw high score` (lines 3982-4013); `UNITS` -- used in `to decimal3` (lines 1328-1354), `bcd to decimal2` (lines 1359-1374), `add and update score` (lines 1396-1448), `put status` (lines 1463-1555), `get level from keyboard` (lines 2187-2316), `draw high score level` (lines 3966-3979), `draw high score` (lines 3982-4013)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

**Chunk:** `to decimal3` (lines 1328-1354)

```asm
    ORG     $7AF8
TO_DECIMAL3:
    SUBROUTINE
    ; Enter routine with A set to the number to convert.

    LDX     #$00
    STX     TENS
    STX     HUNDREDS

.loop1:
    CMP     #100
    BCC     .loop2
    INC     HUNDREDS
    SBC     #100
    BNE     .loop1

.loop2:
    CMP     #10
    BCC     .end
    INC     TENS
    SBC     #10
    BNE     .loop2

.end:
    STA     UNITS
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `TO_DECIMAL3` -- used in `put status` (lines 1463-1555), `get level from keyboard` (lines 2187-2316), `draw high score level` (lines 3966-3979)  
**Uses:** `HUNDREDS` -- `defines` (lines 1322-1325); `TENS` -- `defines` (lines 1322-1325); `UNITS` -- `defines` (lines 1322-1325)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

There's also a routine to convert a BCD byte to tens and units.

**Chunk:** `bcd to decimal2` (lines 1359-1374)

```asm
    ORG     $7AE9
BCD_TO_DECIMAL2:
    SUBROUTINE
    ; Enter routine with A set to the BCD number to convert.

    STA     TENS
    AND     #$0F
    STA     UNITS
    LDA     TENS
    LSR
    LSR
    LSR
    LSR
    STA     TENS
    RTS
```

**Previous:** None  
**Next:** None  
**Defines:** `BCD_TO_DECIMAL2` -- used in `add and update score` (lines 1396-1448), `draw high score` (lines 3982-4013)  
**Uses:** `TENS` -- `defines` (lines 1322-1325); `UNITS` -- `defines` (lines 1322-1325)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

## Score and status


Lode Runner stores your score as an 8-digit BCD number.

**Chunk:** `defines` (lines 1381-1382)

```asm
SCORE       EQU     $8E     ; 4 bytes, BCD format, tens/units in first byte.
```

**Previous:** `defines` (lines 1322-1325)  
**Next:** `defines` (lines 1454-1456)  
**Defines:** `SCORE` -- used in `add and update score` (lines 1396-1448), `put status` (lines 1463-1555), `draw high score table header` (lines 3888-3900), `handle timers` (lines 4114-4290), `check for gold picked up by player` (lines 4615-4659), `move guard` (lines 6751-6972), `record hi score data` (lines 8305-8487), `editor clear high scores` (lines 8741-8786), `Initialize game data` (lines 8813-8862), `game loop` (lines 8889-9017), `level editor` (lines 9257-9343)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

The score is always put on the screen at row 16 column 5, but
only the last 7 digits. Row 16 is the status line, as can be
seen at the bottom of this screenshot.


```latex
\begin{center}
\includegraphics[width=\columnwidth]{screen}
\end{center}
```


There's a routine to add a 4-digit BCD
number to the score and then update it on the screen.

**Chunk:** `add and update score` (lines 1396-1448)

```asm
    ORG     $7A92
ADD_AND_UPDATE_SCORE:
    SUBROUTINE
    ; Enter routine with A set to BCD tens/units and
    ; Y set to BCD thousands/hundreds.

    CLC
    SED                         ; Turn on BCD addition mode.
    ADC     SCORE
    STA     SCORE
    TYA
    ADC     SCORE+1
    STA     SCORE+1
    LDA     #$00
    ADC     SCORE+2
    STA     SCORE+2
    LDA     #$00
    ADC     SCORE+3
    STA     SCORE+3             ; SCORE += param
    CLD                         ; Turn off BCD addition mode.

    LDA     #5
    STA     GAME_COLNUM
    LDA     #16
    STA     GAME_ROWNUM

    LDA     SCORE+3
    JSR     BCD_TO_DECIMAL2
    LDA     UNITS               ; Note we skipped TENS.
    JSR     PUT_DIGIT

    LDA     SCORE+2
    JSR     BCD_TO_DECIMAL2
    LDA     TENS
    JSR     PUT_DIGIT
    LDA     UNITS
    JSR     PUT_DIGIT

    LDA     SCORE+1
    JSR     BCD_TO_DECIMAL2
    LDA     TENS
    JSR     PUT_DIGIT
    LDA     UNITS
    JSR     PUT_DIGIT

    LDA     SCORE
    JSR     BCD_TO_DECIMAL2
    LDA     TENS
    JSR     PUT_DIGIT
    LDA     UNITS
    JMP     PUT_DIGIT           ; tail call

```

**Previous:** None  
**Next:** None  
**Defines:** `ADD_AND_UPDATE_SCORE` -- used in `put status` (lines 1463-1555), `handle timers` (lines 4114-4290), `check for gold picked up by player` (lines 4615-4659), `move guard` (lines 6751-6972), `game loop` (lines 8889-9017)  
**Uses:** `BCD_TO_DECIMAL2` -- `bcd to decimal2` (lines 1359-1374); `GAME_COLNUM` -- `defines` (lines 778-785); `GAME_ROWNUM` -- `defines` (lines 778-785); `PUT_DIGIT` -- `put digit` (lines 1294-1312); `SCORE` -- `defines` (lines 1381-1382); `TENS` -- `defines` (lines 1322-1325); `UNITS` -- `defines` (lines 1322-1325)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)

The other elements in the status line are the number of men
(i.e. lives) and the current level.

**Chunk:** `defines` (lines 1454-1456)

```asm
LEVELNUM    EQU     $A6
LIVES       EQU     $98
```

**Previous:** `defines` (lines 1381-1382)  
**Next:** `defines` (lines 1608-1611)  
**Defines:** `LEVELNUM` -- used in `put status` (lines 1463-1555), `get level from keyboard` (lines 2187-2316), `[[ROW_ADDR = $9E00 + LEVELNUM * $0100]]` (lines 3535-3541), `button pressed at startup` (lines 4410-4421), `timed out waiting for button or keypress` (lines 4462-4486), `ctrl handlers` (lines 4762-4772), `record hi score data` (lines 8305-8487), `game loop` (lines 8889-9017), `level editor key functions` (lines 9605-9760); `LIVES` -- used in `put status` (lines 1463-1555), `ctrl handlers` (lines 4762-4772), `ctrl handlers` (lines 4777-4788), `dead code` (lines 4794-4802), `ctrl handlers` (lines 4821-4831), `check for mode 1 input` (lines 5071-5132), `Initialize game data` (lines 8813-8862), `game loop` (lines 8889-9017), `level editor` (lines 9257-9343)  
**Uses:** None  
**References:** None  
**Referenced by:** `*` (lines 10172-10181)

Here are the routines to put the lives and level number on
the status line. Lives starts at column 16, and level number
starts at column 25.

**Chunk:** `put status` (lines 1463-1555)

```asm
    ORG     $7A70
PUT_STATUS_LIVES:
    SUBROUTINE

    LDA     LIVES
    LDX     #16
    ; fallthrough

PUT_STATUS_BYTE:
    SUBROUTINE
    ; Puts the number in A as a three-digit decimal on the screen
    ; at row 16, column X.

    STX     GAME_COLNUM
    JSR     TO_DECIMAL3
    LDA     #16
    STA     GAME_ROWNUM
    LDA     HUNDREDS
    JSR     PUT_DIGIT
    LDA     TENS
    JSR     PUT_DIGIT
    LDA     UNITS
    JMP     PUT_DIGIT           ; tail call

PUT_STATUS_LEVEL:
    SUBROUTINE

    LDA     LEVELNUM
    LDX     #25
    BNE     PUT_STATUS_BYTE     ; Unconditional jump

    ORG     $79AD
PUT_STATUS:
    SUBROUTINE

    JSR     CLEAR_HGR1
    JSR     CLEAR_HGR2

PUT_STATUS_DRAW:
    LDY     #$27
    LDA     DRAW_PAGE
    CMP     #$40
    BEQ     .draw_line_on_page_2

.draw_line_on_page_1:
    LDA     #$AA
    STA     $2350,Y
    STA     $2750,Y
    STA     $2B50,Y
    STA     $2F50,Y
    DEY
    LDA     #$D5
    STA     $2350,Y
    STA     $2750,Y
    STA     $2B50,Y
    STA     $2F50,Y
    DEY
    BPL     .draw_line_on_page_1
    BMI     .end        ; Unconditional

.draw_line_on_page_2:
    LDA     #$AA
    STA     $4350,Y
    STA     $4750,Y
    STA     $4B50,Y
    STA     $4F50,Y
    DEY
    LDA     #$D5
    STA     $4350,Y
    STA     $4750,Y
    STA     $4B50,Y
    STA     $4F50,Y
    DEY
    BPL     .draw_line_on_page_2

.end:
    LDA     #$10
    STA     GAME_ROWNUM
    LDA     #$00
    STA     GAME_COLNUM

    ; "SCORE        MEN    LEVEL   "
    JSR     PUT_STRING
    HEX     D3 C3 CF D2 C5 A0 A0 A0 A0 A0 A0 A0 A0 CD C5 CE
    HEX     A0 A0 A0 A0 CC C5 D6 C5 CC A0 A0 A0 00

    JSR     PUT_STATUS_LIVES
    JSR     PUT_STATUS_LEVEL
    LDA     #$00
    TAY
    JMP     ADD_AND_UPDATE_SCORE        ; tailcall

```

**Previous:** None  
**Next:** None  
**Defines:** `PUT_STATUS` -- used in `Initialize game data` (lines 8813-8862); `PUT_STATUS_LEVEL` -- used in `iris wipe` (lines 2784-2821); `PUT_STATUS_LIVES` -- used in `iris wipe` (lines 2784-2821), `ctrl handlers` (lines 4777-4788), `game loop` (lines 8889-9017)  
**Uses:** `ADD_AND_UPDATE_SCORE` -- `add and update score` (lines 1396-1448); `CLEAR_HGR1` -- `routines` (lines 152-182); `CLEAR_HGR2` -- `routines` (lines 152-182); `DRAW_PAGE` -- `defines` (lines 1216-1217); `GAME_COLNUM` -- `defines` (lines 778-785); `GAME_ROWNUM` -- `defines` (lines 778-785); `HUNDREDS` -- `defines` (lines 1322-1325); `LEVELNUM` -- `defines` (lines 1454-1456); `LIVES` -- `defines` (lines 1454-1456); `PUT_DIGIT` -- `put digit` (lines 1294-1312); `PUT_STRING` -- `put string` (lines 1260-1288); `SCORE` -- `defines` (lines 1381-1382); `TENS` -- `defines` (lines 1322-1325); `TO_DECIMAL3` -- `to decimal3` (lines 1328-1354); `UNITS` -- `defines` (lines 1322-1325)  
**References:** None  
**Referenced by:** `routines` (lines 10020-10169)
