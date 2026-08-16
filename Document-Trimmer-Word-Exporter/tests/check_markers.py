import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import docx_trim

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

for f in ["complex.docx", "sample.docx", "complex_word_saved.docx", "sample_word_saved.docx"]:
    path = os.path.join(OUT, f)
    if not os.path.exists(path):
        print(f, "MISSING")
        continue
    root = docx_trim.load_document_xml(path)
    n_markers = len(root.findall(".//" + W + "lastRenderedPageBreak"))
    n_explicit = len([br for br in root.iter(W + "br") if br.get(W + "type") == "page"])
    n_sec = len(root.findall(".//" + W + "sectPr"))
    print(f"{f}: markers={n_markers} explicit_pgbreaks={n_explicit} sectPr={n_sec}")
