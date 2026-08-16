import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import word_com

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
src = os.path.join(OUT, "sample.docx")
pdf = os.path.join(OUT, "sample.pdf")
word_com.convert(src, "pdf", pdf)
print("pdf created:", os.path.getsize(pdf), "bytes")
