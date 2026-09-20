"""Export the VBA modules of the workbook into plain-text files.

The workbook is the single source of truth for the VBA code. The exported
files exist so that the code can be read on GitHub and included in the
filesdump for LLMs that cannot open the workbook. They are never imported
back into Excel, so this is a one-way export.

What gets exported:

- Standard modules -> <name>.bas
- Class modules and sheet/workbook modules -> <name>.cls, but only if they
  contain code beyond the attribute lines and `Option` statements. Most sheet
  modules are empty, and exporting them would only add clutter.

What is left alone:

- UserForms. A form's layout lives in a binary .frx file that only the VBA
  editor can write, so a script-generated .frm would be incomplete. Existing
  .frm/.frx files (exported from the VBA editor) are never touched.
- Files in the output folder that no longer match a module in the workbook.
  They are listed as "stale" but not deleted, because deleting files is a
  decision for a human.

Files are written as CRLF and Windows-1252, byte for byte the way the VBA
editor exports them (see `.gitattributes`), and only when their content
actually changed, so an unchanged module never shows up in `git status`.

Usage:
    python tools/export_vba.py [--xlsm a2-hires-lab.xlsm] [--out src/bas]
"""

import argparse
import sys
from pathlib import Path

from oletools.olevba import VBA_Parser

DEFAULT_XLSM = Path("a2-hires-lab.xlsm")
DEFAULT_OUT = Path("src/bas")
ENCODING = "cp1252"  # what the VBA editor writes on a Western-European Windows
EXPORTED_SUFFIXES = {".bas", ".cls"}


def has_real_code(code: str) -> bool:
    """Tell whether a module contains more than boilerplate.

    Every module starts with `Attribute` lines, and most sheet modules
    contain nothing else except `Option Explicit`. Only modules with actual
    code are worth exporting.
    """
    for line in code.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("Attribute ") or stripped.startswith("Option "):
            continue
        return True
    return False


def to_vbe_bytes(code: str) -> bytes:
    """Convert extracted code into the bytes the VBA editor would write.

    oletools already returns CRLF line endings, but normalizing first makes
    the result independent of that detail.
    """
    normalized = code.replace("\r\n", "\n").replace("\r", "\n")
    return normalized.replace("\n", "\r\n").encode(ENCODING)


def extract_modules(xlsm: Path) -> dict[str, str]:
    """Return {file name: code} for every VBA module in the workbook.

    The file name (e.g. `Util.bas`, `Tabelle1.cls`, `UserFormComment.frm`)
    comes from oletools, which derives the extension from the module type.
    """
    parser = VBA_Parser(str(xlsm))
    try:
        if not parser.detect_vba_macros():
            return {}
        return {
            vba_filename: code
            for _, _, vba_filename, code in parser.extract_all_macros()
        }
    finally:
        parser.close()


def export(xlsm: Path, out_dir: Path) -> int:
    """Export the modules and print a summary. Returns the process exit code."""
    if not xlsm.is_file():
        print(f"Error: workbook not found: {xlsm}", file=sys.stderr)
        return 1

    modules = extract_modules(xlsm)
    if not modules:
        print(f"Error: no VBA project found in {xlsm}", file=sys.stderr)
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    written: list[str] = []
    unchanged: list[str] = []
    forms: list[str] = []
    expected: set[str] = set()

    for filename, code in sorted(modules.items(), key=lambda kv: kv[0].lower()):
        suffix = Path(filename).suffix.lower()
        if suffix == ".frm":
            forms.append(filename)
            continue
        if suffix not in EXPORTED_SUFFIXES:
            continue
        if suffix == ".cls" and not has_real_code(code):
            continue

        expected.add(filename)
        target = out_dir / filename
        data = to_vbe_bytes(code)
        if target.is_file() and target.read_bytes() == data:
            unchanged.append(filename)
            continue
        target.write_bytes(data)
        written.append(filename)

    stale = sorted(
        path.name
        for path in out_dir.iterdir()
        if path.suffix.lower() in EXPORTED_SUFFIXES and path.name not in expected
    )
    missing_forms = [name for name in forms if not (out_dir / name).is_file()]

    print(f"Exported VBA from {xlsm} into {out_dir}/")
    print(f"  written:   {', '.join(written) if written else '-'}")
    print(f"  unchanged: {len(unchanged)} module(s)")
    if forms:
        print(f"  forms (not touched, export from the VBA editor): {', '.join(forms)}")
    if missing_forms:
        print(
            f"  WARNING: forms with no .frm in {out_dir}/: {', '.join(missing_forms)}"
        )
    if stale:
        print(f"  stale (not in workbook, not deleted): {', '.join(stale)}")
    return 0


def main() -> int:
    arg_parser = argparse.ArgumentParser(
        description="Export the workbook's VBA modules into plain-text files."
    )
    arg_parser.add_argument(
        "--xlsm", type=Path, default=DEFAULT_XLSM, help="workbook to read"
    )
    arg_parser.add_argument(
        "--out", type=Path, default=DEFAULT_OUT, help="folder to write into"
    )
    args = arg_parser.parse_args()
    return export(args.xlsm, args.out)


if __name__ == "__main__":
    sys.exit(main())
