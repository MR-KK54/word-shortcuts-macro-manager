import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import paginate, docx_trim, verify

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
src = os.path.join(OUT, "verify_loop_src.docx")

# Simulate a buggy output: valid 2-page split but with an extra trailing page-break
# paragraph appended -> would render as 3 pages.
out = os.path.join(OUT, "auto_correct_test.docx")
pc, boundaries = paginate.paginate_docx(src)
docx_trim.split_docx_range(src, 3, 4, boundaries, out)

# Append a trailing page-break paragraph + empty paragraph to force an extra page.
import zipfile
from lxml import etree
root = docx_trim.load_document_xml(out)
body = root.find(docx_trim._q("body"))
pb = etree.SubElement(body, docx_trim._q("p"))
r = etree.SubElement(pb, docx_trim._q("r"))
etree.SubElement(r, docx_trim._q("br"), attrib={docx_trim._q("type"): "page"})
etree.SubElement(body, docx_trim._q("p"))
tmp = out + ".tmp"
with zipfile.ZipFile(out) as zin:
    names = zin.namelist()
    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            if name == "word/document.xml":
                zout.writestr(name, xml)
            else:
                zout.writestr(name, zin.read(name))
if os.path.exists(out):
    os.remove(out)
os.replace(tmp, out)

# Now verify with auto-correct.
report = verify.verify_export(src, out, 3, 4)
print("report pass:", report["pass"])
print("page_count:", report["page_count"])
print("corrected:", report["corrected"])
print("note:", report["note"])
for e in report["pages"]:
    print("  p%d match=%s text=%s %s" % (e["page"], e["match"], e["text_match"], e["detail"]))
if report["pass"] and report["corrected"] >= 1:
    print("AUTO-CORRECT OK")
else:
    print("AUTO-CORRECT FAILED")
    sys.exit(1)
