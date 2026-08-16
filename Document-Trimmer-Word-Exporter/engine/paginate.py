"""Page-boundary detection for Word documents.

Engine selection (in priority order):
  1. MS Word COM  - Word's own layout engine (Windows only). Exact fidelity.
  2. Word markers - <w:lastRenderedPageBreak/> markers Word embeds in docx
                    files, reproducing Word's pagination on Linux/Render.
  3. Explicit breaks - fallback counting explicit page/section breaks.
"""

import os

from . import convert, docx_trim, lo_paginate, word_com, word_markers
from .docx_trim import build_units, _q


def _explicit_break_pagination(docx_path):
    root = docx_trim.load_document_xml(docx_path)
    body = root.find(_q("body"))
    units = build_units(body)
    if not units:
        return 1, [0]

    unit_page = []
    current_page = 1
    next_new_page = False
    for u in units:
        new_page = False
        ends_section = False
        if u["kind"] == "p":
            node = u["node"]
            pPr = node.find(_q("pPr"))
            if pPr is not None:
                if pPr.find(_q("pageBreakBefore")) is not None:
                    new_page = True
                if pPr.find(_q("sectPr")) is not None:
                    ends_section = True
                    new_page = True
            if not new_page:
                for br in node.iter(_q("br")):
                    if br.get(_q("type")) == "page":
                        new_page = True
                        break
        if next_new_page:
            current_page += 1
            next_new_page = False
        elif new_page:
            current_page += 1
        if u["kind"] == "p" and ends_section:
            next_new_page = True
        unit_page.append(current_page)

    page_count = max(unit_page)
    by_page = {}
    for i, p in enumerate(unit_page):
        by_page.setdefault(p, []).append(i)
    boundaries = []
    last = -1
    for p in range(1, page_count + 1):
        lst = by_page.get(p)
        if lst:
            last = max(last, max(lst))
        boundaries.append(last)
    return page_count, boundaries


def paginate_docx(docx_path):
    """Return (page_count, boundaries); boundaries[page0] = last unit index."""
    # 1) PDF-backed layout pagination (Fast, exact visual rendering matching PDF preview/export).
    try:
        res = lo_paginate.paginate_pdf_backed(docx_path)
        if res is not None and res[1] and res[1][0] != -1:
            return res
    except Exception:
        pass

    # 2) MS Word COM (Windows, exact Word layout).
    if word_com.word_available():
        try:
            res = word_com.paginate(docx_path)
            if res is not None and res[1] and res[1][0] != -1:
                return res
        except Exception:
            pass

    # 3) Word's own recorded pagination markers.
    try:
        res = word_markers.paginate(docx_path)
        if res is not None and res[1] and res[1][0] != -1:
            return res
    except Exception:
        pass

    # 4) Explicit page/section break counting.
    return _explicit_break_pagination(docx_path)
