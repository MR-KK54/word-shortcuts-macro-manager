import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import word_com

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")

word = word_com._open_word()
try:
    for name in ["split_1_2.docx", "split_3_3.docx", "split_2_3.docx"]:
        path = os.path.join(OUT, name)
        doc = word.Documents.Open(FileName=os.path.abspath(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        doc.Repaginate()
        pages = int(doc.ComputeStatistics(2))
        # header text of first section
        hdr = ""
        try:
            hdr = doc.Sections(1).Headers(1).Range.Text.strip()
        except Exception:
            hdr = ""
        # last paragraphs text (strip trailing empty)
        last_texts = []
        for i in range(max(1, doc.Paragraphs.Count - 3), doc.Paragraphs.Count + 1):
            try:
                t = (doc.Paragraphs(i).Range.Text or "").rstrip("\r")
                last_texts.append(repr(t[:50]))
            except Exception:
                pass
        # is there a page break in the very last paragraph?
        last = doc.Paragraphs(doc.Paragraphs.Count)
        has_pgbrk = False
        try:
            rng = last.Range
            rng.End = rng.Start + 1
            has_pgbrk = (rng.Text or "") == "\x0c"
        except Exception:
            pass
        print(f"{name}: pages={pages} header='{hdr}' last_paras={last_texts} last_is_pgbrk={has_pgbrk}")
        doc.Close(SaveChanges=0)
finally:
    word.Quit()
