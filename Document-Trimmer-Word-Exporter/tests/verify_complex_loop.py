import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import paginate, docx_trim, verify

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
src = os.path.join(OUT, "complex.docx")
pc, boundaries = paginate.paginate_docx(src)
print("complex.docx pages:", pc, boundaries)

ok = 0
total = 0
for start, end in [(1, 1), (1, 4), (5, 6), (7, 10), (2, 3)]:
    total += 1
    out = os.path.join(OUT, "verify_complex_%d_%d.docx" % (start, end))
    docx_trim.split_docx_range(src, start, end, boundaries, out)
    r = verify.verify_export(src, out, start, end)
    passed = "PASS" if r["pass"] else "FAIL"
    if r["pass"]:
        ok += 1
    print("  %d-%d: %s expected=%s actual=%s corrected=%s" % (
        start, end, passed, r["page_count"]["expected"], r["page_count"]["actual"], r["corrected"]))
    for e in r["pages"]:
        if not e["match"] or not e["text_match"]:
            print("      p%d match=%s text=%s %s" % (e["page"], e["match"], e["text_match"], e["detail"]))

print("%d/%d passed" % (ok, total))
sys.exit(0 if ok == total else 1)
