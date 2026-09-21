# Usage:
#   cd ~/buyer-excess-reports
#   python3 src/main.py
#
# Optional:
#   python3 src/main.py --customer RAKEN
#   python3 src/main.py --customer AVTC
#
# Input folders:
#   input/RAKEN/BOM.xlsx + one Excel file with a shortage sheet
#   input/AVTC/BOM.xlsx  + one Excel file with a shortage sheet
#
# Output files:
#   output/RAKEN/excess_report.xlsx
#   output/AVTC/excess_report.xlsx

from __future__ import annotations

import sys

from runner import main as runner_main


def main() -> int:
    return runner_main()


if __name__ == "__main__":
    sys.exit(main())
