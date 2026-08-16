import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import word_com

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
word = word_com._open_word()
try:
    for n in ["complex_1_4.docx", "complex_5_6.docx", "complex_7_10.docx"]:
        p = os.path.join(OUT, n)
        doc = word.Documents.Open(FileName=os.path.abspath(p), ReadOnly=True, AddToRecentFiles=False, Visible=False)
        doc.Repaginate()
        pages = int(doc.ComputeStatistics(2))
        nsec = doc.Sections.Count
        # header of the FIRST and LAST section
        try:
            h1 = doc.Sections(1).Headers(1).Range.Text.strip()
        except Exception:
            h1 = ""
        try:
            hl = doc.Sections(nsec).Headers(1).Range.Text.strip()
        except Exception:
            hl = ""
        # last paragraph before section mark
        lastpara = (doc.Paragraphs(doc.Paragraphs.Count).Range.Text or "").rstrip("\r")
        print(f"{n}: pages={pages} sections={nsec} header_first='{h1}' header_last='{hl}' last_para={lastpara!r}")
        doc.Close(SaveChanges=0)
finally:
    word_com._shutdown_word(word)
