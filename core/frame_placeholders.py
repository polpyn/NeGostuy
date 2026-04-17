"""
Подстановки в штампе рамки (.docx): {{nomer}}, {{list}}.

Word часто режет {{nomer}} на несколько w:t — чиним через regex по сырому XML.
{{list}} заменяем на поле PAGE (в шаблоне рамки уже w:pgNumType w:start="2").
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


_RPR_INNER = r"<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>"

# Word вставляет proofErr между run'ами — допускаем любое число подряд.
_PROOF_GAP = r"(?:\s*<w:proofErr\b[^>]*/>)*"

# Word разбивает «{{nomer}}» на три run: {{ | nomer | }} (регистр «nomer» любой)
_SPLIT_NOMER = re.compile(
    r"<w:r[^>]*>(?:" + _RPR_INNER + r")?<w:t[^>]*>\{\{</w:t></w:r>"
    + _PROOF_GAP
    + r"<w:r[^>]*>(?P<mid>" + _RPR_INNER + r"<w:t[^>]*>nomer</w:t></w:r>)"
    + _PROOF_GAP
    + r"<w:r[^>]*>(?:" + _RPR_INNER + r")?<w:t[^>]*>\}\}</w:t></w:r>",
    re.DOTALL | re.IGNORECASE,
)

# Один run: {{nomer}} или с пробелами
_CONTIG_NOMER = re.compile(
    r"<w:r[^>]*>(?:" + _RPR_INNER + r")?<w:t[^>]*>\{\{\s*nomer\s*\}\}</w:t></w:r>",
    re.DOTALL | re.IGNORECASE,
)

# {{list}} — один run: один w:rPr (без «сквозного» .*? через вложенные </w:rPr>)
_LIST_RUN = re.compile(
    r"<w:r[^>]*>(?P<rpr><w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)"
    r"<w:t[^>]*>\{\{list\}\}</w:t></w:r>",
    re.DOTALL,
)


def _replace_nomer(xml: str, zachet: str) -> str:
    z = (zachet or "").strip()
    tokens = (
        "{{nomer}}",
        "{{ nomer }}",
        "{{nomer }}",
        "{{ nomer}}",
        "{{Nomer}}",
        "{{ Nomer }}",
    )

    if z:
        val = _xml_escape(z)

        def repl_split(m: re.Match) -> str:
            mid = m.group("mid")
            rpr_m = re.search(r"<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>", mid, re.DOTALL)
            rpr = rpr_m.group(0) if rpr_m else "<w:rPr/>"
            return f"<w:r>{rpr}<w:t xml:space=\"preserve\">{val}</w:t></w:r>"

        xml = _SPLIT_NOMER.sub(repl_split, xml)

        def repl_contig2(m: re.Match) -> str:
            rpr_m = re.search(
                r"<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>", m.group(0), re.DOTALL
            )
            rpr = rpr_m.group(0) if rpr_m else "<w:rPr/>"
            return f"<w:r>{rpr}<w:t xml:space=\"preserve\">{val}</w:t></w:r>"

        xml = _CONTIG_NOMER.sub(repl_contig2, xml)

        for token in tokens:
            if token in xml:
                xml = xml.replace(token, val)
        return xml

    # Пустой номер: убираем плейсхолдер, иначе в штампе остаётся «{{nomer}}».
    xml = _SPLIT_NOMER.sub("", xml)
    xml = _CONTIG_NOMER.sub("", xml)
    for token in tokens:
        if token in xml:
            xml = xml.replace(token, "")
    return xml


def _rpr_page_field_14pt(template_rpr_xml: str) -> str:
    """Жёстко 14 pt (28 half-points) для поля PAGE."""
    has_bold = bool(
        re.search(r"<w:b\s*/>|<w:b>\s*</w:b>|<w:bCs\s*/>", template_rpr_xml)
    )
    bold_el = "<w:b/>" if has_bold else ""
    return (
        "<w:rPr>"
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
        'w:cs="Times New Roman" w:eastAsia="Times New Roman"/>'
        f"{bold_el}"
        '<w:sz w:val="28"/><w:szCs w:val="28"/>'
        '<w:lang w:val="ru-RU" w:eastAsia="ru-RU"/>'
        "</w:rPr>"
    )


def _replace_list_with_page_field(xml: str) -> str:
    """Номер листа — поле PAGE (в ramka sectPr уже w:pgNumType w:start=\"2\")."""

    def repl(m: re.Match) -> str:
        rpr = _rpr_page_field_14pt(m.group("rpr") or "")
        return (
            '<w:fldSimple w:instr=" PAGE \\* MERGEFORMAT ">'
            f"<w:r>{rpr}<w:t>1</w:t></w:r>"
            "</w:fldSimple>"
        )
    # В некоторых шаблонах {{list}} дублируется в AlternateContent (VML + fallback),
    # из-за чего в штампе может появляться «повтор номера страницы». Оставляем одно поле.
    xml = _LIST_RUN.sub(repl, xml, count=1)
    xml = _LIST_RUN.sub("", xml)
    xml = xml.replace("{{list}}", "")
    return xml


def patch_header_xml(xml: str, zachet_number: str | None) -> str:
    z = (zachet_number or "").strip()
    s = _replace_list_with_page_field(xml)
    s = _replace_nomer(s, z)
    return s


def _normalize_document_section_xml(xml: str) -> str:
    """
    Убираем «особую первую страницу»: иначе первая страница идёт с header type=first,
    остальные с default — отступ тела от верха визуально разный.
    """
    xml = xml.replace("<w:titlePg/>", "")
    xml = xml.replace("<w:titlePg></w:titlePg>", "")
    xml = re.sub(
        r'<w:headerReference\s+w:type="first"\s+r:id="[^"]*"\s*/>',
        "",
        xml,
    )
    return xml


def patch_docx_headers(docx_path: str, zachet_number: str | None = None) -> None:
    """Правит word/header*.xml и sectPr в document.xml (штамп, поля, единый колонтитул)."""
    from .frame_debug import (
        resolve_frame_debug_dir,
        write_after_and_summary,
        write_patch_debug_session,
    )

    debug_parent = resolve_frame_debug_dir()
    session: str | None = None
    before: dict[str, str] = {}
    if debug_parent:
        with zipfile.ZipFile(docx_path, "r") as zin:
            for info in zin.infolist():
                fn = info.filename
                if re.match(r"word/(header|footer)\d+\.xml$", fn) or fn == "word/document.xml":
                    before[fn] = zin.read(fn).decode("utf-8")
        session = write_patch_debug_session(
            debug_parent, docx_path, zachet_number, before
        )

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    try:
        with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(
            tmp.name, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if re.match(r"word/(header|footer)\d+\.xml$", info.filename):
                    text = data.decode("utf-8")
                    new_text = patch_header_xml(text, zachet_number)
                    data = new_text.encode("utf-8")
                elif info.filename == "word/document.xml":
                    text = data.decode("utf-8")
                    text = _normalize_document_section_xml(text)
                    zn = (zachet_number or "").strip()
                    text = _replace_nomer(text, zn)
                    data = text.encode("utf-8")
                zout.writestr(info, data, compress_type=info.compress_type)
        shutil.move(tmp.name, docx_path)
        if session and before:
            write_after_and_summary(session, docx_path, before)
    except Exception:
        try:
            import os

            os.unlink(tmp.name)
        except OSError:
            pass
        raise
