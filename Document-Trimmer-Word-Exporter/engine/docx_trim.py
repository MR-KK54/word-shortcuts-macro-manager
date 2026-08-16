"""High-fidelity DOCX page-range splitting.

The splitter clones the source package (styles, headers, footers, numbering,
theme, media) and rewrites only word/document.xml so that every layout element
is preserved exactly. Explicit page breaks are inserted at source page
boundaries so each output page mirrors the source page, and all trailing
"breaks" (empty paragraphs, page/line break runs, section breaks) are stripped
from the last page of every split file.
"""

import os
import re
import zipfile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = {"w": W_NS}

DOCUMENT_PART = "word/document.xml"


def load_document_xml(path):
    """Parse word/document.xml from a docx (zip) package."""
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read(DOCUMENT_PART))


def _q(tag):
    return "{%s}%s" % (W_NS, tag)


def _strip_punct(text):
    return re.findall(r"[a-z0-9]+", text.lower())


def _iter_text(elem):
    if elem.tag == _q("t"):
        yield elem.text or ""
    for child in elem:
        yield from _iter_text(child)


def _element_text(elem):
    return "".join(_iter_text(elem))


def _element_words(elem):
    return _strip_punct(_element_text(elem))


def _iter_body_content_nodes(parent, body_index):
    """Yield dicts for content units inside parent element."""
    for child in parent:
        tag = child.tag
        if tag == _q("p"):
            yield {"kind": "p", "body_index": body_index, "node": child}
        elif tag == _q("tbl"):
            rows = child.findall(_q("tr"))
            if not rows:
                yield {"kind": "tbl", "body_index": body_index, "table": child, "rows": []}
            else:
                for row in rows:
                    yield {"kind": "row", "body_index": body_index, "table": child, "row": row}
        elif tag in (_q("sdt"), _q("customXml"), _q("ins"), _q("del"), _q("smartTag")):
            sdtContent = child.find(_q("sdtContent"))
            container = sdtContent if sdtContent is not None else child
            yield from _iter_body_content_nodes(container, body_index)


def build_units(body):
    """Build a flat list of content units (paragraphs and table rows).

    Each unit is a dict: {'kind': 'p'|'row', 'body_index', 'node'|('table','row')}
    """
    if body is None:
        return []
    units = []
    for idx, child in enumerate(body):
        units.extend(list(_iter_body_content_nodes([child], idx)))
    return units


def unit_words(unit):
    if unit["kind"] == "p":
        return _element_words(unit["node"])
    return _element_words(unit["row"])


def _make_page_break_paragraph():
    p = etree.fromstring(
        '<w:p xmlns:w="%s"><w:r><w:br w:type="page"/></w:r></w:p>' % W_NS
    )
    return p


def _has_explicit_break(unit):
    """True if the paragraph/row already forces a page break before it or at its end."""
    if unit["kind"] == "p":
        node = unit["node"]
        pPr = node.find(_q("pPr"))
        if pPr is not None:
            if pPr.find(_q("pageBreakBefore")) is not None:
                return True
            if pPr.find(_q("sectPr")) is not None:
                return True  # section break starts a new page
        # an explicit page break run at the end of the paragraph forces the next
        # unit onto a fresh page (nothing after it). Scan runs backwards and
        # inspect run children so a trailing page break inside the last run is
        # found even when the same run also carries text before it.
        for r in reversed(node.findall(_q("r"))):
            for child in reversed(list(r)):
                if child.tag == _q("rPr"):
                    continue
                if child.tag == _q("t"):
                    if (child.text or "").strip():
                        return False  # trailing text -> no trailing break
                    continue
                if child.tag == _q("br"):
                    if child.get(_q("type")) == "page":
                        return True
                    continue  # line break, keep scanning backwards
                # drawing / pict / object / other content -> no trailing break
                return False
    return False


def _needs_forced_break(units, fu):
    """Decide whether a page break must be inserted before unit index fu.

    If the source already begins a new page there naturally (explicit page
    break paragraph, pageBreakBefore, section break), inserting another break
    would produce a blank page, so we return False.
    """
    if fu <= 0 or fu >= len(units):
        return False
    # Break already expressed on the unit itself.
    if _has_explicit_break(units[fu]):
        return False
    # Previous unit is an explicit page-break paragraph -> next unit is a new page.
    prev = units[fu - 1]
    if prev["kind"] == "p":
        if _has_explicit_break(prev):
            return False
    return True


def _add_page_break_before(paragraph):
    pPr = paragraph.find(_q("pPr"))
    if pPr is None:
        pPr = etree.SubElement(paragraph, _q("pPr"))
        paragraph.insert(0, pPr)
    if pPr.find(_q("pageBreakBefore")) is None:
        pb = etree.Element(_q("pageBreakBefore"))
        pPr.insert(0, pb)


def _has_text(elem):
    for t in elem.iter(_q("t")):
        if (t.text or "").strip():
            return True
    return False


def _paragraph_has_only_breaks(paragraph):
    """True if the paragraph has no text and only page/line break runs."""
    if _has_text(paragraph):
        return False
    pPr = paragraph.find(_q("pPr"))
    if pPr is not None and pPr.find(_q("sectPr")) is not None:
        return False  # section break paragraph is handled separately
    for r in paragraph.findall(_q("r")):
        for br in r.findall(_q("br")):
            if br.get(_q("type")) in (None, "textWrapping", "page"):
                continue
        # run with drawing/pict etc. is content
        if r.find(_q("drawing")) is not None or r.find(_q("pict")) is not None or r.find(_q("object")) is not None:
            return False
    return True


def _strip_trailing_break_runs(paragraph):
    """Remove trailing page-break/line-break runs from the end of a paragraph.

    Also removes a page/line break that sits at the very end of the last run,
    even when that run also contains text before the break.
    """
    runs = paragraph.findall(_q("r"))
    while runs:
        r = runs[-1]
        # strip a trailing break that is the last child of this run
        children = list(r)
        if children and children[-1].tag == _q("br"):
            last = children[-1]
            if last.get(_q("type")) in (None, "textWrapping", "page"):
                r.remove(last)
                if not list(r):
                    paragraph.remove(r)
                    runs = paragraph.findall(_q("r"))
                continue
        has_br = any(
            br.get(_q("type")) in (None, "textWrapping", "page")
            for br in r.findall(_q("br"))
        )
        has_text = _has_text(r)
        has_drawing = r.find(_q("drawing")) is not None or r.find(_q("pict")) is not None
        if has_text or has_drawing or not has_br:
            break
        paragraph.remove(r)
        runs = paragraph.findall(_q("r"))


def _strip_trailing_empty_rows(table):
    rows = table.findall(_q("tr"))
    while rows:
        r = rows[-1]
        if _has_text(r):
            break
        if r.find(_q("tc")) is None:
            break
        table.remove(r)
        rows = table.findall(_q("tr"))


def _is_section_break_paragraph(paragraph):
    pPr = paragraph.find(_q("pPr"))
    if pPr is None:
        return False
    return pPr.find(_q("sectPr")) is not None


def _find_final_sectpr(body):
    children = list(body)
    if children and children[-1].tag == _q("sectPr"):
        return children[-1]
    return None


def _governing_sectpr(body, units, ue):
    """Return (sectPr_node_to_use_for_body_level, keep_sectpr_para: bool).

    Finds the section whose properties govern the last kept content unit and
    returns a deepcopy of its sectPr so headers/footers of the final output
    page match the source. Returns (None, ...) if the body-level sectPr governs.
    """
    final_sectpr = _find_final_sectpr(body)
    if ue >= len(units):
        return final_sectpr, False
    body_index = units[ue]["body_index"]
    children = list(body)
    # Walk forward from the last content unit to find the section that ends at/after it.
    for i in range(body_index, len(children)):
        child = children[i]
        if child.tag == _q("p"):
            pPr = child.find(_q("pPr"))
            if pPr is not None:
                sectPr = pPr.find(_q("sectPr"))
                if sectPr is not None:
                    if sectPr is final_sectpr:
                        return final_sectpr, False
                    return deepcopy_or_none(sectPr), i
    return final_sectpr, False


def deepcopy_or_none(node):
    return etree.fromstring(etree.tostring(node))


def _build_new_body(original_body, us, ue, break_unit_indices):
    """Build a new <w:body> containing only units in [us, ue] with page breaks."""
    units = build_units(original_body)
    children = list(original_body)

    body_units = {}
    for ui, u in enumerate(units):
        body_units.setdefault(u["body_index"], []).append(ui)

    new_body = etree.Element(_q("body"))
    new_children = []

    for i, child in enumerate(children):
        if child.tag == _q("sectPr"):
            continue  # final sectPr re-added at the end
        uis = body_units.get(i, [])
        if not uis:
            continue
        if child.tag == _q("p"):
            ui = uis[0]
            if not (us <= ui <= ue):
                continue
            new_p = etree.fromstring(etree.tostring(child))
            if ui in break_unit_indices:
                _add_page_break_before(new_p)
            new_children.append(new_p)
        elif child.tag == _q("tbl"):
            kept = [ui for ui in uis if us <= ui <= ue]
            if not kept:
                continue
            # split into segments at break boundaries
            segments = []
            cur = []
            for ui in kept:
                if ui in break_unit_indices and cur:
                    segments.append(cur)
                    cur = [ui]
                else:
                    cur.append(ui)
            if cur:
                segments.append(cur)

            for s_idx, seg in enumerate(segments):
                if s_idx > 0:
                    new_children.append(_make_page_break_paragraph())
                seg_tbl = etree.fromstring(etree.tostring(child))
                keep_set = set(seg)
                row_nodes = seg_tbl.findall(_q("tr"))
                for rn in row_nodes:
                    seg_tbl.remove(rn)
                for j, ui in enumerate(uis):
                    if ui in keep_set:
                        seg_tbl.append(row_nodes[j])
                new_children.append(seg_tbl)
        else:
            kept = [ui for ui in uis if us <= ui <= ue]
            if kept:
                new_children.append(etree.fromstring(etree.tostring(child)))

    # Trailing break removal on the last page ---------------------------------
    replacement_final_sectpr = None
    while new_children:
        last = new_children[-1]
        if last.tag == _q("p"):
            if _is_section_break_paragraph(last):
                pPr = last.find(_q("pPr"))
                sectPr = pPr.find(_q("sectPr"))
                replacement_final_sectpr = deepcopy_or_none(sectPr)
                new_children.pop()
                continue
            if _has_text(last):
                _strip_trailing_break_runs(last)
                break
            if _paragraph_has_only_breaks(last):
                new_children.pop()
                continue
            _strip_trailing_break_runs(last)
            break
        elif last.tag == _q("tbl"):
            _strip_trailing_empty_rows(last)
            if _has_text(last):
                break
            new_children.pop()
        else:
            new_children.pop()

    # Final section properties -------------------------------------------------
    final_sectpr = None
    if replacement_final_sectpr is not None:
        final_sectpr = replacement_final_sectpr
    else:
        governing, _ = _governing_sectpr(original_body, units, ue)
        if governing is not None and governing is not _find_final_sectpr(original_body):
            final_sectpr = deepcopy_or_none(governing)
        else:
            final = _find_final_sectpr(original_body)
            if final is not None:
                final_sectpr = deepcopy_or_none(final)

    for nc in new_children:
        new_body.append(nc)
    if final_sectpr is not None:
        type_elem = final_sectpr.find(_q("type"))
        if type_elem is not None and type_elem.get(_q("val")) in ("nextPage", "oddPage", "evenPage"):
            type_elem.set(_q("val"), "continuous")
        new_body.append(final_sectpr)

    return new_body


def split_docx_range(src_docx, start_page, end_page, boundaries, out_path):
    """Split src_docx into a new file containing source pages start_page..end_page.

    boundaries: list where boundaries[page0] = last unit index of that page.
    """
    page_count = len(boundaries)
    if start_page < 1:
        start_page = 1
    if end_page > page_count:
        end_page = page_count
    if start_page > page_count:
        start_page = page_count
    if end_page < start_page:
        raise ValueError(f"No content on pages {start_page}-{end_page}.")

    us = (boundaries[start_page - 2] + 1) if start_page > 1 else 0
    ue = boundaries[end_page - 1]
    if ue < us:
        raise ValueError(f"No content on pages {start_page}-{end_page}.")

    units = build_units(load_document_xml(src_docx).find(_q("body")))
    break_unit_indices = set()
    for p in range(start_page + 1, end_page + 1):
        fu = (boundaries[p - 2] + 1) if p > 1 else 0
        if fu > us and fu <= ue and _needs_forced_break(units, fu):
            break_unit_indices.add(fu)

    _rewrite_document_xml(src_docx, out_path, us, ue, break_unit_indices, start_page, end_page)


def _parse_xml(zf, path):
    return etree.fromstring(zf.read(path))


def _rewrite_document_xml(src_docx, out_path, us, ue, break_unit_indices, start_page, end_page):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(src_docx) as zin:
        names = zin.namelist()
        root = _parse_xml(zin, "word/document.xml")
        body = root.find(_q("body"))
        new_body = _build_new_body(body, us, ue, break_unit_indices)
        root.replace(body, new_body)

        # Set the document title in core properties when available.
        xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name == "word/document.xml":
                    zout.writestr(name, xml)
                else:
                    data = zin.read(name)
                    if name == "docProps/core.xml" and _try_set_title(data, out_path):
                        zout.writestr(name, _try_set_title(data, out_path))
                    else:
                        zout.writestr(name, data)
    if os.path.exists(out_path):
        os.remove(out_path)
    os.replace(tmp, out_path)


def _try_set_title(core_xml, out_path):
    """Best-effort set of the dc:title in docProps/core.xml. Returns bytes or None."""
    try:
        root = etree.fromstring(core_xml)
    except Exception:
        return None
    title = os.path.splitext(os.path.basename(out_path))[0]
    ns = {"dc": "http://purl.org/dc/elements/1.1/"}
    el = root.find("dc:title", ns)
    if el is None:
        el = etree.SubElement(root, "{http://purl.org/dc/elements/1.1/}title")
    el.text = title
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
