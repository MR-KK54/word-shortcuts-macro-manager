import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import docx_trim

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
LOCAL = os.path.join(OUT, "complex_1_4.docx")
DEPLOYED = os.path.join(os.environ.get("TEMP", "."), "deployed_complex_out.docx")
SRC = os.path.join(OUT, "complex.docx")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def summarize(path, label):
    if not os.path.exists(path):
        print(label, "MISSING")
        return
    root = docx_trim.load_document_xml(path)
    body = root.find(W + "body")
    kids = list(body)
    n_p = sum(1 for k in kids if k.tag == W + "p")
    n_tbl = sum(1 for k in kids if k.tag == W + "tbl")
    n_sect = sum(1 for k in kids if k.tag == W + "sectPr")
    text = []
    for p in kids[:6]:
        if p.tag == W + "p":
            t = "".join(x.text or "" for x in p.iter(W + "t"))
            text.append(t[:40])
    print(f"\n=== {label} ===")
    print(f"  body children: {len(kids)} (p={n_p}, tbl={n_tbl}, sectPr={n_sect})")
    print(f"  first paras: {text}")
    # last 2 elements
    print(f"  last elements: {[k.tag.split('}')[-1] for k in kids[-2:]]}")


summarize(SRC, "SOURCE complex.docx")
summarize(LOCAL, "LOCAL output (Word-exact) complex_1_4.docx")
summarize(DEPLOYED, "DEPLOYED output (Render) complex_pages_1-4.docx")
