#!/usr/bin/env python3
"""Generate Buyer Excess reports from BOM and shortage inputs."""

from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

def main() -> int:
    if not sys.argv[1:] or sys.argv[1:] == ["--gui"]:
        from gui import main as gui_main
        return gui_main()
    from runner import main as runner_main
    return runner_main()


if __name__ == "__main__":
    sys.exit(main())
