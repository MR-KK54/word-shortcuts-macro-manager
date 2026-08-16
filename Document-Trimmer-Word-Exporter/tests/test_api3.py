import os, sys, json, time, uuid, urllib.request, urllib.parse
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

BASE = "http://127.0.0.1:5000"
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "runtime", "tests")


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


def upload(name, path):
    boundary = uuid.uuid4().hex
    data = open(path, "rb").read()
    body = (
        ('--' + boundary + '\r\nContent-Disposition: form-data; name="files"; filename="'
         + name + '"\r\nContent-Type: application/octet-stream\r\n\r\n').encode()
        + data + b'\r\n--' + boundary.encode() + b'--\r\n'
    )
    req = urllib.request.Request(BASE + "/api/upload", data=body, method="POST")
    req.add_header("Content-Type", "multipart/form-data; boundary=" + boundary)
    with urllib.request.urlopen(req) as r:
        return r.status, r.read()


if __name__ == "__main__":
    try:
        call("POST", "/api/clear-storage")
        print("upload:", upload("sample.docx", os.path.join(OUT, "sample.docx"))[0])

        st, b = call("POST", "/api/export", {
            "files": ["sample.docx"], "range": "1-2,3-3", "format": "same",
            "output_dir": "./x", "naming_pattern": "{original_name}_pages_{start_page}-{end_page}",
            "overwrite": True, "clear_storage_after_export": False, "engine_mode": "trimming", "visible": False})
        jid = json.loads(b)["job_id"]
        for _ in range(120):
            st, b = call("GET", "/api/job/" + jid)
            j = json.loads(b)
            if j["status"] in ("done", "error", "cancelled"):
                break
            time.sleep(1)
        print("job:", j["status"], "outputs:", j["outputs"], "errors:", j["errors"])

        from engine.docx_trim import load_document_xml, _q
        for o in j["outputs"]:
            st, b = call("GET", "/api/download/" + jid + "/" + urllib.parse.quote(o))
            dest = os.path.join(OUT, "new_" + o)
            open(dest, "wb").write(b)
            root = load_document_xml(dest)
            kids = list(root.find(_q("body")))
            tail = [(k.tag.split("}")[-1], "".join(t.text or "" for t in k.iter(_q("t")))[:25]) for k in kids[-3:]]
            print(f"  {o}: children={len(kids)} tail={tail}")
    except Exception as e:
        print("Skipping live server API test (server not running):", e)
