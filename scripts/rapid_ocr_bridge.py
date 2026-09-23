from __future__ import annotations

import json
import os
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")


def main() -> int:
    try:
        tool_root = Path(sys.argv[1]).resolve()
        image_paths = [Path(value).resolve() for value in sys.argv[2:]]
        internal = tool_root / "_internal"
        if hasattr(os, "add_dll_directory"):
            for directory, _children, files in os.walk(internal):
                if any(name.lower().endswith(".dll") for name in files):
                    try:
                        os.add_dll_directory(directory)
                    except OSError:
                        pass
        sys.path.insert(0, str(internal))
        sys._MEIPASS = str(internal)  # type: ignore[attr-defined]
        from rapidocr import RapidOCR

        engine = RapidOCR(params={"Global.text_score": 0.2})
        rows = []
        for image_path in image_paths:
            result = engine(str(image_path))
            texts = list(getattr(result, "txts", None) or [])
            scores = [float(score) for score in (getattr(result, "scores", None) or [])]
            rows.append({
                "text": "\n".join(map(str, texts)),
                "confidence": sum(scores) / len(scores) if scores else 0.0,
                "backend": "rapidocr_ppocrv6",
            })
        print(json.dumps(rows[0] if len(rows) == 1 else {"results": rows}, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({"text": "", "confidence": 0.0, "backend": "rapidocr_ppocrv6", "error": str(exc)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
