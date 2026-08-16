import os, sys, re, tempfile, shutil, pymupdf
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import convert, word_com, renderer
from engine.docx_trim import load_document_xml, build_units, _element_words, _q

docx_path = r"C:\Users\Hxtreme\Downloads\202100_1.docx"

root = load_document_xml(docx_path)
body = root.find(_q("body"))
units = build_units(body)

rend = renderer.Renderer(os.path.join(tempfile.gettempdir(), "doc_trim_lo_pag"))
pdf_file, _ = rend._pdf_for(docx_path)

with pymupdf.open(pdf_file) as doc:
    page_texts = [" ".join(re.findall(r"[a-z0-9]+", page.get_text().lower())) for page in doc]

print(f"Total PDF pages: {len(page_texts)}, Total units: {len(units)}")

# Find repetitive header/footer N-grams that appear across >3 pages
from collections import Counter
ngram_counts = Counter()
for pt in page_texts:
    words = pt.split()
    seen_in_page = set()
    for i in range(len(words) - 2):
        gram = " ".join(words[i:i+3])
        seen_in_page.add(gram)
    for g in seen_in_page:
        ngram_counts[g] += 1

repeated_grams = {g for g, count in ngram_counts.items() if count >= 3}
print(f"Identified {len(repeated_grams)} repeated header/footer 3-grams.")

unit_page = []
p_idx = 0

for idx, u in enumerate(units):
    node = u["node"] if u["kind"] == "p" else u["row"]
    text_words = _element_words(node)
    if not text_words:
        unit_page.append(p_idx + 1)
        continue

    # Filter out repeated header 3-grams
    content_words = []
    w_str = " ".join(text_words)
    for w in text_words:
        if len(w) >= 3 and w not in ("plataforma", "brasil", "saude", "gov", "br", "http", "https", "visao", "pesquisador"):
            content_words.append(w)

    if not content_words:
        unit_page.append(p_idx + 1)
        continue

    # Check if content_words match p_idx or next pages
    # Search for best matching page starting from p_idx
    best_p = p_idx
    best_score = -1

    # Only look ahead up to 5 pages from current p_idx
    max_search = min(len(page_texts), p_idx + 6)
    for pi in range(p_idx, max_search):
        pt = page_texts[pi]
        score = sum(1 for w in content_words if w in pt)
        if score > best_score:
            best_score = score
            best_p = pi

    # Advance p_idx if best_p > p_idx and match score is strong
    if best_p > p_idx and best_score >= min(2, len(content_words)):
        p_idx = best_p

    unit_page.append(p_idx + 1)

page_count = max(len(page_texts), max(unit_page) if unit_page else 1)
by_page = {}
for i, p in enumerate(unit_page):
    by_page.setdefault(p, []).append(i)

boundaries = []
last = -1
for p in range(1, page_count + 1):
    if p in by_page:
        last = max(by_page[p])
    boundaries.append(last)

print("Calculated boundaries (first 10):", boundaries[:10])
