import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import word_com

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
os.makedirs(os.path.join(OUT, "render_cmp"), exist_ok=True)

word = word_com._open_word()


def render_to_pdf(path, pdf_out):
    doc = word.Documents.Open(FileName=os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    doc.SaveAs2(FileName=os.path.abspath(pdf_out), FileFormat=17)
    doc.Close(SaveChanges=0)


def page_image(pdf_path, page, png_path, width=1100):
    import pymupdf
    with pymupdf.open(pdf_path) as d:
        pg = d[page - 1]
        zoom = width / 1100
        pix = pg.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
        pix.save(png_path)
    return pix.width, pix.height


def compare_png(a, b):
    import pymupdf
    pa = pymupdf.open(a).load_page(0).get_pixmap()
    pb = pymupdf.open(b).load_page(0).get_pixmap()
    if pa.width != pb.width or pa.height != pb.height:
        return False, f"size {pa.width}x{pa.height} vs {pb.width}x{pb.height}"
    da = pa.samples
    db = pb.samples
    n = len(da)
    diffs = 0
    step = 3  # sample stride
    for i in range(0, n, step * 4):
        if da[i] != db[i] or da[i + 1] != db[i + 1] or da[i + 2] != db[i + 2]:
            diffs += 1
    total = n // (step * 4)
    pct = 100.0 * diffs / total if total else 0.0
    return pct < 1.0, f"{pct:.2f}% pixels differ"


cases = [
    # (source, output, list of (source_page -> output_page))
    ("complex.docx", "complex_1_4.docx", [(1, 1), (2, 2), (3, 3), (4, 4)]),
    ("complex.docx", "complex_5_6.docx", [(5, 1), (6, 2)]),
    ("complex.docx", "complex_7_10.docx", [(7, 1), (8, 2), (9, 3), (10, 4)]),
]

try:
    for src, out, mapping in cases:
        src_pdf = os.path.join(OUT, "render_cmp", src.replace(".docx", "_src.pdf"))
        out_pdf = os.path.join(OUT, "render_cmp", out.replace(".docx", "_out.pdf"))
        render_to_pdf(os.path.join(OUT, src), src_pdf)
        render_to_pdf(os.path.join(OUT, out), out_pdf)
        print(f"\n=== {src} -> {out} ===")
        for sp, op in mapping:
            a = os.path.join(OUT, "render_cmp", f"{src}_{sp}.png")
            b = os.path.join(OUT, "render_cmp", f"{out}_{op}.png")
            page_image(src_pdf, sp, a)
            page_image(out_pdf, op, b)
            ok, detail = compare_png(a, b)
            print(f"  src p{sp} vs out p{op}: {'MATCH' if ok else 'DIFF'} ({detail})")
finally:
    word.Quit()
