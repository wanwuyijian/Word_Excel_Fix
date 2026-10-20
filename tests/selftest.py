"""源码模式下的自检入口（转发到 wxfixer.selftest）。

用法:
    python tests/selftest.py [--samples samples] [--out tests/out]
"""

from __future__ import annotations

import os
import sys

SRC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
sys.path.insert(0, SRC)

from wxfixer.selftest import main  # noqa: E402
from wxfixer.utils import logging_setup  # noqa: E402

if __name__ == "__main__":
    logging_setup.setup_logging(console=False)
    sys.exit(main())
