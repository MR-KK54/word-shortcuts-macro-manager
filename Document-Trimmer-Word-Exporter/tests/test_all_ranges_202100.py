import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import paginate, docx_trim, ranges, verify, renderer
import pymupdf

src = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_sample", "202100_1.docx")
out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_sample", "range_tests")
os.makedirs(out_dir, exist_ok=True)

rend = renderer.Renderer(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "renders"))

pc, bd = paginate.paginate_docx(src)
print("Source document page count:", pc)
print("Source document boundaries (first 10):", bd[:10])

test_specs = ["1-1", "1-2", "1-3", "2-3", "2-4", "3-5", "1-end"]
results = []

for spec in test_specs:
    ranges_list = ranges.parse_range_spec(spec, pc)
    for start, end in ranges_list:
        out_name = f"split_{start}_{end}.docx"
        out_path = os.path.join(out_dir, out_name)
        docx_trim.split_docx_range(src, start, end, bd, out_path)
        verify.verify_export(src, out_path, start, end)
        
        pdf_path, _ = rend._pdf_for(out_path)
        with pymupdf.open(pdf_path) as doc:
            actual_pages = doc.page_count
        expected_pages = end - start + 1
        status = "PASS" if actual_pages == expected_pages else "FAIL"
        results.append((spec, f"{start}-{end}", expected_pages, actual_pages, status))

print("\n=== COMPREHENSIVE RANGE SPLIT VERIFICATION RESULTS ===")
print(f"{'Spec':<10} | {'Range':<10} | {'Expected':<10} | {'Actual':<10} | {'Status'}")
print("-" * 55)
all_pass = True
for spec, r_str, exp, act, st in results:
    print(f"{spec:<10} | {r_str:<10} | {exp:<10} | {act:<10} | {st}")
    if st != "PASS":
        all_pass = False

if all_pass:
    print("\nALL RANGE SPECIFICATIONS CUT WITH 100% ACCURACY!")
else:
    print("\nSOME RANGES FAILED VERIFICATION.")
    sys.exit(1)
