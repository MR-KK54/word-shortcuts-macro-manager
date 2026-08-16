import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pymupdf
from engine import word_com

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
CMP = os.path.join(OUT, "render_cmp")
os.makedirs(CMP, exist_ok=True)


def render_to_pdf(word, path, pdf_out):
    doc = word.Documents.Open(FileName=os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    doc.SaveAs2(FileName=os.path.abspath(pdf_out), FileFormat=17)
    doc.Close(SaveChanges=0)


def page_image(pdf_path, page, png_path, width=1100):
    with pymupdf.open(pdf_path) as d:
        pg = d[page - 1]
        zoom = width / 1100
        pix = pg.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
        pix.save(png_path)


def compare_png(a, b):
    pa = pymupdf.open(a).load_page(0).get_pixmap()
    pb = pymupdf.open(b).load_page(0).get_pixmap()
    if pa.width != pb.width or pa.height != pb.height:
        return False, f"size {pa.width}x{pa.height} vs {pb.width}x{pb.height}"
    da, db = pa.samples, pb.samples
    n = len(da)
    diffs = 0
    step = 3
    for i in range(0, n, step * 4):
        if da[i] != db[i] or da[i + 1] != db[i + 1] or da[i + 2] != db[i + 2]:
            diffs += 1
    total = n // (step * 4)
    pct = 100.0 * diffs / total if total else 0.0
    return pct < 1.0, f"{pct:.2f}% pixels differ"


word = word_com._open_word()
try:
    src_pdf = os.path.join(CMP, "cmp_src.pdf")
    out_pdf = os.path.join(CMP, "cmp_out.pdf")
    render_to_pdf(word, os.path.join(OUT, "complex.docx"), src_pdf)
    render_to_pdf(word, os.path.join(OUT, "deployed_complex_out.docx"), out_pdf)
    for sp in [1, 2, 3, 4]:
        a = os.path.join(CMP, f"d_src_{sp}.png")
        b = os.path.join(CMP, f"d_out_{sp}.png")
        page_image(src_pdf, sp, a)
        page_image(out_pdf, sp, b)
        ok, detail = compare_png(a, b)
        status = "MATCH" if ok else "DIFF"
        print(f"src p{sp} vs deployed out p{sp}: {status} ({detail})")
finally:
    word.Quit()
