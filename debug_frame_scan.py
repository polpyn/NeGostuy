#!/usr/bin/env python3
"""
Сканирование и симуляция патча рамки без запуска сервера.

  py debug_frame_scan.py путь\\к\\файлу.docx
  py debug_frame_scan.py путь\\к\\файлу.docx папка_вывода

В папке:
00_scan.txt           — где найдены {{nomer}}, {{list}}, PAGE, размеры шрифта
  word__headerN_*       — исходный XML колонтитулов
  word__headerN_SIM_*   — patch_header_xml: с номером / без номера (плейсхолдер убирается)
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def main() -> None:
    if len(sys.argv) < 2:
        print(__doc__.strip())
        sys.exit(1)

    docx = Path(sys.argv[1])
    if not docx.is_file():
        print("Файл не найден:", docx)
        sys.exit(1)

    out = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "debug_frame" / "manual_scan"
    out.mkdir(parents=True, exist_ok=True)

    from core.frame_debug import format_scan_report, scan_docx_xml
    from core.frame_placeholders import patch_header_xml

    scan = scan_docx_xml(str(docx))
    (out / "00_scan.txt").write_text(format_scan_report(scan), encoding="utf-8")

    with zipfile.ZipFile(docx, "r") as z:
        for name in sorted(z.namelist()):
            if not re.match(r"word/(header|footer)\d+\.xml$", name):
                continue
            raw = z.read(name).decode("utf-8")
            safe = name.replace("/", "__")
            (out / f"{safe}_extracted.xml").write_text(raw, encoding="utf-8")

            dry_no = patch_header_xml(raw, None)
            dry_z = patch_header_xml(raw, "DEBUG_ZACH_123")
            (out / f"{safe}_SIM_no_zachet.xml").write_text(dry_no, encoding="utf-8")
            (out / f"{safe}_SIM_with_zachet.xml").write_text(dry_z, encoding="utf-8")

    doc_xml = "word/document.xml"
    with zipfile.ZipFile(docx, "r") as z:
        if doc_xml in z.namelist():
            d = z.read(doc_xml).decode("utf-8")
            (out / "word__document_extracted.xml").write_text(d, encoding="utf-8")
            if "sectPr" in d:
                m = re.search(r"<w:sectPr.*?</w:sectPr>", d, re.DOTALL)
                if m:
                    (out / "word__document_sectPr_only.xml").write_text(
                        m.group(0), encoding="utf-8"
                    )

    print("OK ->", out.resolve())


if __name__ == "__main__":
    main()
