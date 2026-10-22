"""PostToolUse hook: if the just-edited/written file is SCRAPING-SOURCES.md, regenerate the
matching .docx from it (see build_scraping_sources_docx.py). Wired up in .claude/settings.json.

Reads the tool-call JSON Claude Code pipes in on stdin. Deliberately never raises or exits
non-zero - a hook failing here must never block the actual Edit/Write it's reacting to.
"""

import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
TARGET_BASENAME = "SCRAPING-SOURCES.md"


def main():
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return

    tool_input = payload.get("tool_input") or {}
    file_path = tool_input.get("file_path") or payload.get("tool_response", {}).get("filePath") or ""
    if Path(file_path).name != TARGET_BASENAME:
        return

    generator = HERE / "build_scraping_sources_docx.py"
    if not generator.exists():
        return

    try:
        subprocess.run([sys.executable, str(generator)], cwd=str(HERE), timeout=60)
    except Exception:
        pass  # best-effort - never let a regeneration failure affect the triggering edit


if __name__ == "__main__":
    main()
