from __future__ import annotations

import argparse
from pathlib import Path

from bom_alternates import load_bom_alternates
from excess_builder import build_excess_rows
from formatting import write_excess_workbook
from workbook_io import find_bom_file, find_shortage_source, load_shortage_records


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_DIR = PROJECT_ROOT / "input"
DEFAULT_OUTPUT_FILE = PROJECT_ROOT / "output" / "excess_report.xlsx"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a usable Excess report from BOM.xlsx and a shortage worksheet."
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT_DIR,
        help=f"Input folder. Default: {DEFAULT_INPUT_DIR}",
    )
    parser.add_argument(
        "--output-file",
        type=Path,
        default=DEFAULT_OUTPUT_FILE,
        help=f"Output xlsx file. Default: {DEFAULT_OUTPUT_FILE}",
    )
    args = parser.parse_args()

    input_dir = args.input_dir.resolve()
    output_file = args.output_file.resolve()

    bom_file = find_bom_file(input_dir)
    shortage_source = find_shortage_source(input_dir, bom_file)

    print(f"BOM: {bom_file}")
    print(f"Shortage: {shortage_source.path} [{shortage_source.sheet_name}]")

    alternates = load_bom_alternates(bom_file)
    shortage_records = load_shortage_records(shortage_source)
    result = build_excess_rows(shortage_records, alternates)
    write_excess_workbook(
        result.rows,
        output_file,
        result.row1_headers,
        result.row2_headers,
        result.part_column_count,
    )

    print(f"Read shortage parts: {result.input_part_count}")
    print(f"Generated Excess rows: {result.output_row_count}")
    print(f"Part columns: {result.part_column_count}")
    print(f"Output: {output_file}")


if __name__ == "__main__":
    main()
