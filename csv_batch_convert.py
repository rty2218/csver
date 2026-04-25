#!/usr/bin/env python3
"""
Batch convert local CSV files to XLSX, fixed-width TXT tables, and Markdown tables.

Examples:
  python3 csv_batch_convert.py
  python3 csv_batch_convert.py ./csv_files -o ./converted
  python3 csv_batch_convert.py ./csv_files --recursive --infer-types
  python3 csv_batch_convert.py a.csv b.csv --in-place
"""

from __future__ import annotations

import argparse
import csv
import glob
import io
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional
from unicodedata import combining, east_asian_width
from xml.sax.saxutils import escape


DEFAULT_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "cp936", "big5", "latin-1")
INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
MAX_XLSX_ROWS = 1_048_576
MAX_XLSX_COLS = 16_384
FORMAT_MAP = {
    "xlsx": ("xlsx",),
    "txt": ("txt",),
    "md": ("md",),
    "all": ("xlsx", "txt", "md"),
}


def main() -> int:
    args = parse_args()
    csv_files = collect_csv_files(args.inputs, args.recursive)
    if not csv_files:
        print("No CSV files found.")
        return 1

    output_dir = Path(args.output_dir).expanduser().resolve()
    if not args.in_place:
        output_dir.mkdir(parents=True, exist_ok=True)

    used_outputs: set[Path] = set()
    formats = resolve_formats(args.format)
    converted = 0

    for csv_path in csv_files:
        try:
            base_dir = csv_path.parent if args.in_place else output_dir
            base_dir.mkdir(parents=True, exist_ok=True)
            outputs, encoding, delimiter = convert_one_csv(
                csv_path,
                base_dir=base_dir,
                used_outputs=used_outputs,
                formats=formats,
                encoding=args.encoding,
                delimiter=args.delimiter,
                infer_types=args.infer_types,
                no_header=args.no_header,
                max_col_width=args.max_col_width,
            )

            converted += 1
            print(f"Converted: {csv_path}")
            print(f"  encoding={encoding}, delimiter={delimiter_label(delimiter)}")
            for output_path in outputs:
                print(f"  -> {output_path}")
        except Exception as exc:  # Keep batch conversion moving after one bad file.
            print(f"Failed: {csv_path}: {exc}")

    print(f"Done. Converted {converted}/{len(csv_files)} CSV file(s).")
    return 0 if converted else 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Batch convert CSV files to .xlsx, fixed-width .txt, and Markdown .md tables."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="CSV files or folders. Defaults to the current folder.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="converted",
        help="Output folder when not using --in-place. Default: ./converted",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Find CSV files recursively inside input folders.",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Write output files next to each source CSV.",
    )
    parser.add_argument(
        "--encoding",
        default="auto",
        help="CSV encoding. Use 'auto' to try utf-8/gb18030/cp936/big5/latin-1. Default: auto",
    )
    parser.add_argument(
        "--delimiter",
        default="auto",
        help="CSV delimiter. Use 'auto' to sniff. Examples: ',', ';', tab, '|'. Default: auto",
    )
    parser.add_argument(
        "--format",
        choices=sorted(FORMAT_MAP),
        default="all",
        help="Output type: xlsx, txt, md, or all. Default: all",
    )
    parser.add_argument(
        "--infer-types",
        action="store_true",
        help="Write obvious numbers and booleans as typed XLSX values and right-align Markdown numeric columns.",
    )
    parser.add_argument(
        "--no-header",
        action="store_true",
        help="For Markdown output, generate Column 1... headers instead of treating the first CSV row as headers.",
    )
    parser.add_argument(
        "--max-col-width",
        type=int,
        default=60,
        help="Maximum display width for TXT columns. Use 0 for no wrapping. Default: 60",
    )
    return parser.parse_args()


def resolve_formats(format_name: str) -> tuple[str, ...]:
    try:
        return FORMAT_MAP[format_name]
    except KeyError:
        raise ValueError(f"unknown format: {format_name}") from None


def convert_one_csv(
    csv_path: Path,
    base_dir: Path,
    used_outputs: set[Path],
    formats: tuple[str, ...],
    encoding: str,
    delimiter: str,
    infer_types: bool,
    no_header: bool,
    max_col_width: int,
) -> tuple[list[Path], str, str]:
    base_dir.mkdir(parents=True, exist_ok=True)
    rows, detected_encoding, detected_delimiter = read_csv_rows(csv_path, encoding, delimiter)
    base_name = csv_path.stem
    outputs: list[Path] = []

    if "xlsx" in formats:
        xlsx_path = unique_path(base_dir / f"{base_name}.xlsx", used_outputs)
        write_xlsx(rows, xlsx_path, sheet_name=base_name, infer_types=infer_types)
        outputs.append(xlsx_path)

    if "txt" in formats:
        txt_path = unique_path(base_dir / f"{base_name}.txt", used_outputs)
        write_txt_table(rows, txt_path, max_col_width=max_col_width)
        outputs.append(txt_path)

    if "md" in formats:
        md_path = unique_path(base_dir / f"{base_name}.md", used_outputs)
        write_markdown_table(rows, md_path, no_header=no_header, infer_align=infer_types)
        outputs.append(md_path)

    return outputs, detected_encoding, detected_delimiter


def collect_csv_files(inputs: list[str], recursive: bool) -> list[Path]:
    raw_inputs = inputs or ["."]
    files: list[Path] = []

    for raw in raw_inputs:
        path = Path(raw).expanduser()
        if path.is_dir():
            iterator = path.rglob("*") if recursive else path.iterdir()
            files.extend(sorted(p for p in iterator if p.is_file() and p.suffix.lower() == ".csv"))
        elif path.is_file() and path.suffix.lower() == ".csv":
            files.append(path)
        elif any(ch in raw for ch in "*?[]"):
            files.extend(
                sorted(Path(p) for p in glob.glob(raw, recursive=recursive) if Path(p).is_file() and Path(p).suffix.lower() == ".csv")
            )

    seen: set[Path] = set()
    unique: list[Path] = []
    for file_path in files:
        resolved = file_path.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(resolved)
    return unique


def read_csv_rows(path: Path, encoding: str, delimiter: str) -> tuple[list[list[str]], str, str]:
    encodings = (encoding,) if encoding.lower() != "auto" else DEFAULT_ENCODINGS
    last_error: Optional[Exception] = None

    for candidate in encodings:
        try:
            with path.open("r", encoding=candidate, newline="") as handle:
                text = handle.read()
            dialect = sniff_dialect(text, delimiter)
            rows = [[cell for cell in row] for row in csv.reader(io.StringIO(text, newline=""), dialect)]
            return rows, candidate, dialect.delimiter
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

    raise ValueError(f"could not read with encodings {encodings}: {last_error}")


def sniff_dialect(text: str, delimiter: str) -> csv.Dialect:
    if delimiter.lower() != "auto":
        if delimiter.lower() == "tab":
            delimiter = "\t"
        return type("ManualDialect", (csv.excel,), {"delimiter": delimiter})

    sample = text[:8192]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        return csv.get_dialect("excel")


def delimiter_label(delimiter: str) -> str:
    return "TAB" if delimiter == "\t" else repr(delimiter)


def normalize_rows(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []
    width = max(len(row) for row in rows)
    return [row + [""] * (width - len(row)) for row in rows]


def write_xlsx(rows: list[list[str]], output_path: Path, sheet_name: str, infer_types: bool) -> None:
    rows = normalize_rows(rows)
    if len(rows) > MAX_XLSX_ROWS:
        raise ValueError(f"too many rows for XLSX: {len(rows)} > {MAX_XLSX_ROWS}")
    if rows and len(rows[0]) > MAX_XLSX_COLS:
        raise ValueError(f"too many columns for XLSX: {len(rows[0])} > {MAX_XLSX_COLS}")

    sheet_xml = build_sheet_xml(rows, infer_types=infer_types)
    workbook_xml = build_workbook_xml(clean_sheet_name(sheet_name))
    core_xml = build_core_xml()

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("docProps/app.xml", APP_XML)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS_XML)
        zf.writestr("xl/styles.xml", STYLES_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def build_sheet_xml(rows: list[list[str]], infer_types: bool) -> str:
    row_count = max(len(rows), 1)
    col_count = max((len(row) for row in rows), default=1)
    dimension = f"A1:{column_name(col_count)}{row_count}"
    col_widths = estimate_col_widths(rows)

    cols_xml = "".join(
        f'<col min="{i}" max="{i}" width="{width:.1f}" customWidth="1"/>'
        for i, width in enumerate(col_widths, start=1)
    )

    sheet_rows: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells: list[str] = []
        for col_index, raw_value in enumerate(row, start=1):
            value = "" if raw_value is None else str(raw_value)
            ref = f"{column_name(col_index)}{row_index}"
            style = ' s="1"' if row_index == 1 else ""
            cells.append(build_cell_xml(ref, value, style=style, infer_types=infer_types and row_index != 1))
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    pane_xml = (
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft"/>'
        '</sheetView></sheetViews>'
        if rows
        else '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="{dimension}"/>'
        f"{pane_xml}"
        f"<cols>{cols_xml}</cols>"
        f'<sheetData>{"".join(sheet_rows)}</sheetData>'
        "</worksheet>"
    )


def build_cell_xml(ref: str, value: str, style: str, infer_types: bool) -> str:
    if value == "":
        return f'<c r="{ref}"{style}/>'

    if infer_types:
        typed = infer_xlsx_value(value)
        if typed is not None:
            kind, typed_value = typed
            if kind == "number":
                return f'<c r="{ref}"{style}><v>{typed_value}</v></c>'
            if kind == "bool":
                return f'<c r="{ref}" t="b"{style}><v>{typed_value}</v></c>'

    escaped = escape(value, {'"': "&quot;"})
    preserve = ' xml:space="preserve"' if value[:1].isspace() or value[-1:].isspace() else ""
    return f'<c r="{ref}" t="inlineStr"{style}><is><t{preserve}>{escaped}</t></is></c>'


def infer_xlsx_value(value: str) -> Optional[tuple[str, str]]:
    stripped = value.strip()
    if stripped == "":
        return None
    lowered = stripped.lower()
    if lowered in {"true", "yes"}:
        return ("bool", "1")
    if lowered in {"false", "no"}:
        return ("bool", "0")

    normalized = stripped.replace(",", "")
    if re.fullmatch(r"[+-]?\d+", normalized):
        if len(normalized.lstrip("+-")) > 1 and normalized.lstrip("+-").startswith("0"):
            return None
        return ("number", normalized)
    if re.fullmatch(r"[+-]?(\d+\.\d*|\d*\.\d+)([eE][+-]?\d+)?", normalized) or re.fullmatch(
        r"[+-]?\d+[eE][+-]?\d+", normalized
    ):
        return ("number", normalized)
    return None


def estimate_col_widths(rows: list[list[str]]) -> list[float]:
    if not rows:
        return [12.0]
    col_count = max(len(row) for row in rows)
    widths: list[float] = []
    for col_index in range(col_count):
        max_width = 8
        for row in rows[:500]:
            text = row[col_index] if col_index < len(row) else ""
            max_width = max(max_width, min(display_width(str(text)) + 2, 60))
        widths.append(float(max_width))
    return widths


def build_workbook_xml(sheet_name: str) -> str:
    escaped_name = escape(sheet_name, {'"': "&quot;"})
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        f'<sheet name="{escaped_name}" sheetId="1" r:id="rId1"/>'
        "</sheets>"
        "</workbook>"
    )


def build_core_xml() -> str:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>csv_batch_convert.py</dc:creator>"
        "<cp:lastModifiedBy>csv_batch_convert.py</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def clean_sheet_name(name: str) -> str:
    cleaned = INVALID_SHEET_CHARS.sub("_", name).strip("'").strip()
    if not cleaned:
        cleaned = "Sheet1"
    return cleaned[:31]


def column_name(index: int) -> str:
    result = []
    while index:
        index, remainder = divmod(index - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def write_txt_table(rows: list[list[str]], output_path: Path, max_col_width: int) -> None:
    rows = normalize_rows(rows)
    if not rows:
        output_path.write_text("", encoding="utf-8")
        return

    wrapped_rows: list[list[list[str]]] = []
    col_count = len(rows[0])
    widths = [0] * col_count

    for row in rows:
        wrapped_row: list[list[str]] = []
        for col_index, cell in enumerate(row):
            lines = wrap_cell(clean_for_text(cell), max_col_width)
            wrapped_row.append(lines)
            widths[col_index] = max(widths[col_index], *(display_width(line) for line in lines))
        wrapped_rows.append(wrapped_row)

    separator = "+" + "+".join("-" * (width + 2) for width in widths) + "+"
    out_lines = [separator]
    numeric_cols = numeric_columns(rows[1:] if len(rows) > 1 else rows)

    for row_index, wrapped_row in enumerate(wrapped_rows):
        height = max(len(cell_lines) for cell_lines in wrapped_row)
        for line_index in range(height):
            cells: list[str] = []
            for col_index, cell_lines in enumerate(wrapped_row):
                text = cell_lines[line_index] if line_index < len(cell_lines) else ""
                align_right = row_index > 0 and col_index in numeric_cols
                cells.append(pad_cell(text, widths[col_index], align_right=align_right))
            out_lines.append("| " + " | ".join(cells) + " |")
        out_lines.append(separator)

    output_path.write_text("\n".join(out_lines) + "\n", encoding="utf-8")


def wrap_cell(value: str, max_width: int) -> list[str]:
    raw_lines = value.splitlines() or [""]
    if max_width <= 0:
        return raw_lines

    result: list[str] = []
    for raw_line in raw_lines:
        if raw_line == "":
            result.append("")
            continue
        current = ""
        current_width = 0
        for char in raw_line:
            char_width = display_width(char)
            if current and current_width + char_width > max_width:
                result.append(current)
                current = char
                current_width = char_width
            else:
                current += char
                current_width += char_width
        result.append(current)
    return result


def pad_cell(value: str, width: int, align_right: bool) -> str:
    pad = max(width - display_width(value), 0)
    if align_right:
        return " " * pad + value
    return value + " " * pad


def display_width(value: str) -> int:
    total = 0
    for char in value:
        if combining(char):
            continue
        total += 2 if east_asian_width(char) in {"F", "W"} else 1
    return total


def clean_for_text(value: str) -> str:
    return str(value).replace("\t", "    ").replace("\r\n", "\n").replace("\r", "\n")


def write_markdown_table(rows: list[list[str]], output_path: Path, no_header: bool, infer_align: bool) -> None:
    rows = normalize_rows(rows)
    if not rows:
        output_path.write_text("", encoding="utf-8")
        return

    if no_header:
        headers = [f"Column {i}" for i in range(1, len(rows[0]) + 1)]
        body = rows
    else:
        headers = [cell if str(cell).strip() else f"Column {i}" for i, cell in enumerate(rows[0], start=1)]
        body = rows[1:]

    numeric_cols = numeric_columns(body) if infer_align else set()
    separator = [("---:" if i in numeric_cols else "---") for i in range(len(headers))]

    md_rows = [headers, separator, *body]
    lines = ["| " + " | ".join(markdown_cell(cell) for cell in row) + " |" for row in md_rows]
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def markdown_cell(value: str) -> str:
    text = str(value).replace("\\", "\\\\").replace("|", "\\|")
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    return text.strip()


def numeric_columns(rows: list[list[str]]) -> set[int]:
    if not rows:
        return set()
    col_count = max(len(row) for row in rows)
    numeric: set[int] = set()
    for col_index in range(col_count):
        values = [str(row[col_index]).strip() for row in rows if col_index < len(row) and str(row[col_index]).strip()]
        if values and all(infer_xlsx_value(value) and infer_xlsx_value(value)[0] == "number" for value in values):
            numeric.add(col_index)
    return numeric


def unique_path(path: Path, used_outputs: set[Path]) -> Path:
    path = path.resolve()
    candidate = path
    counter = 2
    while candidate in used_outputs or candidate.exists():
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        counter += 1
    used_outputs.add(candidate)
    return candidate


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

WORKBOOK_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>csv_batch_convert.py</Application>
</Properties>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""


if __name__ == "__main__":
    raise SystemExit(main())
