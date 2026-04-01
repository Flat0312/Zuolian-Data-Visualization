from __future__ import annotations

import runpy
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent / "app" / "frontend"
APP_ENTRY = APP_DIR / "app.py"


def main() -> None:
    if not APP_ENTRY.exists():
        raise FileNotFoundError(f"找不到应用入口：{APP_ENTRY}")

    app_dir_text = str(APP_DIR)
    if app_dir_text not in sys.path:
        sys.path.insert(0, app_dir_text)

    runpy.run_path(str(APP_ENTRY), run_name="__main__")


if __name__ == "__main__":
    main()
