import os, sys, shutil, pymupdf
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import server, engine.word_com as wc, engine.renderer as r

src = r"C:\Users\Hxtreme\Downloads\202100_1.docx"
job = server._new_job({
    "files": ["202100_1.docx"],
    "range": "1-2",
    "format": "docx",
    "output_dir": "./exported_pages",
    "naming_pattern": "{original_name}_pages_{start_page}-{end_page}",
    "overwrite": True,
    "clear_storage_after_export": False,
    "engine_mode": "trimming",
    "visible": False
})

# Copy src into UPLOADS_DIR
os.makedirs(server.UPLOADS_DIR, exist_ok=True)
shutil.copy(src, os.path.join(server.UPLOADS_DIR, "202100_1.docx"))

outputs, expected = server._process_one(job, os.path.join(server.UPLOADS_DIR, "202100_1.docx"))
print("EXPORT OUTPUTS:", outputs)

out_file = os.path.join(server.JOBS_DIR, job["id"], "outputs", outputs[0])
print("EXPORTED FILE PATH:", out_file)

# Check MS Word COM page count
word = wc._open_word()
doc = wc._open_doc(word, os.path.abspath(out_file))
word_pages = int(doc.ComputeStatistics(2))
doc.Close(0)
wc._shutdown_word(word)

print(f"\n=======================================================")
print(f"FINAL EXPORTED FILE MS WORD PAGE COUNT: {word_pages}")
print(f"=======================================================\n")

if word_pages == 2:
    print("SUCCESS! EXPORT FILE IS EXACTLY 2 PAGES IN MS WORD!")
else:
    print(f"FAILED! EXPORT FILE IS {word_pages} PAGES IN MS WORD.")
    sys.exit(1)
