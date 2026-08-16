import os, sys, json, time, urllib.request, urllib.parse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engine import word_com
from engine.docx_trim import load_document_xml, build_units, _q, _element_words

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")
BASE = "http://127.0.0.1:5000"

def call(method, path, data=None):
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(BASE + path, data=body, method=method)
    if body:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=300) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()

if __name__ == "__main__":
    try:
        # Verify downloaded splits structurally
        for name in ["dl_sample_pages_1-2.docx", "dl_sample_pages_3-3.docx"]:
            path = os.path.join(OUT, name)
            if not os.path.exists(path):
                continue
            root = load_document_xml(path)
            body = root.find(_q("body"))
            units = build_units(body)
            last = units[-1]
            lastwords = " ".join(_element_words(last["node"] if last["kind"]=="p" else last["row"]))
            tail = []
            for u in units[-3:]:
                tail.append(" ".join(_element_words(u["node"] if u["kind"]=="p" else u["row"])))
            print(f"{name}: units={len(units)} last_words={lastwords!r} tail={tail!r}")

        # Open in Word: page counts
        if word_com.word_available():
            word = word_com._open_word()
            for name in ["dl_sample_pages_1-2.docx", "dl_sample_pages_3-3.docx"]:
                path = os.path.join(OUT, name)
                if not os.path.exists(path):
                    continue
                doc = word.Documents.Open(FileName=path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
                doc.Repaginate()
                print(name, "Word pages =", int(doc.ComputeStatistics(2)))
                doc.Close(SaveChanges=0)
            word_com._shutdown_word(word)

        # Export docx -> pdf (Word COM conversion)
        st, body = call("POST", "/api/export", data={
            "files": ["sample.docx"], "range": "1-2,3-3", "format": "pdf",
            "output_dir": "./x", "naming_pattern": "{original_name}_pages_{start_page}-{end_page}",
            "overwrite": True, "clear_storage_after_export": False, "engine_mode": "trimming", "visible": False})
        jid = json.loads(body)["job_id"]
        for _ in range(120):
            st, body = call("GET", "/api/job/" + jid)
            j = json.loads(body)
            if j["status"] in ("done", "error", "cancelled"):
                break
            time.sleep(1)
        print("docx->pdf export:", j["status"], "outputs:", j["outputs"], "errors:", j["errors"])
        for o in j["outputs"]:
            st, body = call("GET", "/api/download/" + jid + "/" + urllib.parse.quote(o))
            open(os.path.join(OUT, "dl_" + o), "wb").write(body)
            print("  saved dl_" + o, len(body), "bytes")

        # Test range 'even' and '1-end' presets
        for spec in ["even", "odd", "1-end"]:
            st, body = call("POST", "/api/export", data={
                "files": ["sample.pdf"], "range": spec, "format": "same",
                "output_dir": "./x", "naming_pattern": "{original_name}_pages_{start_page}-{end_page}",
                "overwrite": True, "clear_storage_after_export": False, "engine_mode": "trimming", "visible": False})
            jid = json.loads(body)["job_id"]
            for _ in range(120):
                st, body = call("GET", "/api/job/" + jid)
                j = json.loads(body)
                if j["status"] in ("done", "error", "cancelled"):
                    break
                time.sleep(1)
            print(f"range '{spec}':", j["status"], "outputs:", j["outputs"], "errors:", j["errors"])
    except Exception as e:
        print("Skipping live server API test (server not running):", e)
