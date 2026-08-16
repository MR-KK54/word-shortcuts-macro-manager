import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine import paginate, docx_trim

src = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_sample", "202100_1.docx")
user_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_sample", "202100_1_pages_1-2.docx")

pc, bd = paginate.paginate_docx(src)
print("SOURCE 202100_1.docx page count:", pc)
print("BOUNDARIES:", bd)

user_pc, _ = paginate.paginate_docx(user_out)
print("USER EXPORT 202100_1_pages_1-2.docx page count:", user_pc)

root = docx_trim.load_document_xml(src)
units = docx_trim.build_units(root.find(docx_trim._q("body")))
print("TOTAL UNITS in 202100_1.docx:", len(units))

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "user_sample", "our_split_1_2.docx")
docx_trim.split_docx_range(src, 1, 2, bd, out)
our_pc, _ = paginate.paginate_docx(out)
print("OUR GENERATED split_1_2.docx page count:", our_pc)
