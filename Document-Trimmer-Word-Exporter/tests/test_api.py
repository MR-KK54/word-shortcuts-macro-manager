import os, sys, time, json, urllib.request, urllib.parse
from pathlib import Path

BASE = "http://127.0.0.1:5000"
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")

def call(method, path, data=None, files=None):
    if files:
        import uuid
        boundary = uuid.uuid4().hex
        parts = []
        for name, fpath in files:
            b = Path(fpath).read_bytes()
            parts.append(
                ('--' + boundary + '\r\n'
                 'Content-Disposition: form-data; name="%s"; filename="%s"\r\n'
                 'Content-Type: application/octet-stream\r\n\r\n' % (name, os.path.basename(fpath))).encode() + b + b'\r\n')
        body = b''.join(parts) + ('--' + boundary + '--\r\n').encode()
        req = urllib.request.Request(BASE + path, data=body, method=method)
        req.add_header("Content-Type", "multipart/form-data; boundary=" + boundary)
    else:
        body = json.dumps(data).encode() if data is not None else None
        req = urllib.request.Request(BASE + path, data=body, method=method)
        if body:
            req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=300) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()

# 1) upload docx + pdf
st, body = call("POST", "/api/upload", files=[("files", os.path.join(OUT, "sample.docx")), ("files", os.path.join(OUT, "sample.pdf"))])
print("upload:", st, body.decode()[:200])

# 2) inspect docx
st, body = call("POST", "/api/inspect", data={"name": "sample.docx"})
print("inspect docx:", st, body.decode())
st, body = call("POST", "/api/inspect", data={"name": "sample.pdf"})
print("inspect pdf:", st, body.decode())

# 3) naming preview
st, body = call("POST", "/api/naming-preview", data={"pattern": "{original_name}_pages_{start_page}-{end_page}", "format": "docx", "sample": "Report.docx"})
print("naming-preview:", st, body.decode())

# 4) export docx range 1-2 and 3-3 (multi range) -> docx
st, body = call("POST", "/api/export", data={
    "files": ["sample.docx"],
    "range": "1-2,3-3",
    "format": "same",
    "output_dir": "./exported_pages",
    "naming_pattern": "{original_name}_pages_{start_page}-{end_page}",
    "overwrite": True, "clear_storage_after_export": False,
    "engine_mode": "trimming", "visible": False})
print("export:", st, body.decode()[:200])
job = json.loads(body)
jid = job["job_id"]

for _ in range(120):
    st, body = call("GET", "/api/job/" + jid)
    j = json.loads(body)
    if j["status"] in ("done", "error", "cancelled"):
        break
    time.sleep(1)
print("job status:", j["status"], "completed", j["completed"], "total", j["total"], "errors", j["errors"])
print("outputs:", j["outputs"])
for o in j["outputs"]:
    st, body = call("GET", "/api/download/" + jid + "/" + urllib.parse.quote(o))
    dest = os.path.join(OUT, "dl_" + o)
    open(dest, "wb").write(body)
    print("  downloaded", o, st, len(body), "->", dest)

# 5) export pdf range even -> pdf
st, body = call("POST", "/api/export", data={
    "files": ["sample.pdf"],
    "range": "1-2,3-3",
    "format": "same",
    "output_dir": "./exported_pages",
    "naming_pattern": "{original_name}_pages_{start_page}-{end_page}",
    "overwrite": True, "clear_storage_after_export": False,
    "engine_mode": "trimming", "visible": False})
job2 = json.loads(body)
jid2 = job2["job_id"]
for _ in range(120):
    st, body = call("GET", "/api/job/" + jid2)
    j2 = json.loads(body)
    if j2["status"] in ("done", "error", "cancelled"):
        break
    time.sleep(1)
print("pdf job:", j2["status"], "outputs:", j2["outputs"], "errors:", j2["errors"])

# 6) preview first page of the docx upload
st, body = call("GET", "/api/preview/" + urllib.parse.quote("sample.docx") + "?page=1&w=400")
print("preview docx:", st, "png bytes:", len(body), "total-pages hdr present")
