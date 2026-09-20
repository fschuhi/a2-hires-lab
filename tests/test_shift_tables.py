"""Verify the game's shift tables against the arithmetic they encode.

Lode Runner shifts a 7-pixel pattern by 0 to 6 pixels with two table lookups:
`pixel_shift_table.asm` gives the address of an entry in
`pixel_pattern_table.asm`, and that entry holds the two output bytes.

These tests re-derive every result from first principles -- shift the pattern,
split it into two 7-bit halves, set bit 7 -- and compare it with what the
tables actually say. They back the claim in `README.md` that a single direct
1,792-byte table would produce exactly the same output as the two-stage
lookup, since that claim only holds if the two stages resolve as described.

The data files are Xekri's disassembly output and are read as they are; no
part of this suite touches the workbook.
"""

import re
from pathlib import Path

import pytest

DATA_DIR = Path(__file__).resolve().parents[1] / "data" / "lode_runner_reveng"
SHIFT_TABLE_PATH = DATA_DIR / "pixel_shift_table.asm"
PATTERN_TABLE_PATH = DATA_DIR / "pixel_pattern_table.asm"

PATTERN_TABLE_ORIGIN = 0xA900  # where PIXEL_PATTERN_TABLE is assembled
SHIFT_AMOUNTS = range(7)  # the game shifts right by 0 to 6 pixels
PATTERN_COUNT = 128  # every 7-bit pixel pattern
ENTRY_COUNT = 512  # entries in the pattern table, two bytes each

# One 256-byte page per shift amount: the low bytes (offsets) live in the
# first half, the high bytes (pages) in the second.
PAGE_SIZE = 256
PAGE_HALF = 128


def parse_hex_table(path: Path) -> list[int]:
    """Read a table written as `HEX 00 02 04 ...` lines."""
    values: list[int] = []
    for line in path.read_text().splitlines():
        stripped = line.strip()
        if not stripped.startswith("HEX"):
            continue
        values.extend(int(token, 16) for token in stripped[len("HEX") :].split())
    return values


def parse_byte_table(path: Path) -> list[int]:
    """Read a table written as `BYTE %10000000, %10000000` lines."""
    values: list[int] = []
    for line in path.read_text().splitlines():
        stripped = line.strip()
        if not stripped.startswith("BYTE"):
            continue
        values.extend(int(token, 2) for token in re.findall(r"%([01]+)", stripped))
    return values


@pytest.fixture(scope="module")
def shift_table() -> list[int]:
    return parse_hex_table(SHIFT_TABLE_PATH)


@pytest.fixture(scope="module")
def pattern_table() -> list[int]:
    return parse_byte_table(PATTERN_TABLE_PATH)


def expected_bytes(pattern: int, shift: int) -> tuple[int, int]:
    """Return the two output bytes for a pattern shifted right by `shift`.

    Pixels reach the screen least significant bit first, so shifting a pattern
    right on screen means shifting its value left. Bit 7 carries the color
    pair and is always set in the table's output.
    """
    shifted = pattern << shift
    return (shifted & 0x7F) | 0x80, ((shifted >> 7) & 0x7F) | 0x80


def entry_index(shift_table: list[int], pattern: int, shift: int) -> int:
    """Resolve one pattern and shift to an entry number in the pattern table."""
    page_base = shift * PAGE_SIZE
    low_byte = shift_table[page_base + pattern]
    high_byte = shift_table[page_base + PAGE_HALF + pattern]
    address = high_byte * 256 + low_byte
    offset = address - PATTERN_TABLE_ORIGIN
    assert offset % 2 == 0, f"odd offset {offset} for pattern {pattern}, shift {shift}"
    return offset // 2


def test_table_sizes(shift_table: list[int], pattern_table: list[int]) -> None:
    assert len(shift_table) == len(SHIFT_AMOUNTS) * PAGE_SIZE
    assert len(pattern_table) == ENTRY_COUNT * 2


@pytest.mark.parametrize("shift", SHIFT_AMOUNTS)
def test_lookup_matches_arithmetic(
    shift_table: list[int], pattern_table: list[int], shift: int
) -> None:
    """Every pattern, followed through both tables, gives the shifted bytes."""
    for pattern in range(PATTERN_COUNT):
        index = entry_index(shift_table, pattern, shift)
        found = (pattern_table[2 * index], pattern_table[2 * index + 1])
        assert found == expected_bytes(
            pattern, shift
        ), f"pattern {pattern:#04x}, shift {shift}"


def test_pattern_entries_are_unique(pattern_table: list[int]) -> None:
    """No two entries hold the same pair of bytes.

    This is why the table has 512 entries: that is the number of distinct
    outputs of the 896 pattern/shift combinations.
    """
    entries = list(zip(pattern_table[0::2], pattern_table[1::2]))
    assert len(set(entries)) == ENTRY_COUNT


def test_every_entry_is_reachable(shift_table: list[int]) -> None:
    """The shift table reaches all 512 entries, so none is dead weight."""
    reached = {
        entry_index(shift_table, pattern, shift)
        for shift in SHIFT_AMOUNTS
        for pattern in range(PATTERN_COUNT)
    }
    assert reached == set(range(ENTRY_COUNT))
