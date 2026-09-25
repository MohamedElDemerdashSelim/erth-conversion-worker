import io, os, re, time, uuid, zipfile, subprocess, secrets, shutil, json, base64, hmac, hashlib
from pathlib import Path, PurePosixPath
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from lxml import etree

app = FastAPI(title="ERTH Conversion Worker", version="2.0.0")
SECRET = os.getenv("ERTH_WORKER_SECRET", "")
ORIGIN = os.getenv("ERTH_ALLOWED_ORIGIN", "https://erthpub.com")
TTL = int(os.getenv("ERTH_JOB_TTL", "2700"))
JAR = os.getenv("EPUBCHECK_JAR", "/opt/epubcheck/epubcheck.jar")
PANDOC = os.getenv("PANDOC_BIN", "pandoc")
BASE = Path("/tmp/erth-jobs"); BASE.mkdir(parents=True, exist_ok=True)
JOBS = {}
app.add_middleware(CORSMiddleware, allow_origins=[ORIGIN], allow_credentials=False, allow_methods=["GET", "POST"], allow_headers=["*"])

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
XHTML_NS = "http://www.w3.org/1999/xhtml"
OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"
CONTAINER_NS = "urn:oasis:names:tc:opendocument:xmlns:container"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def b64decode(s):
    return base64.urlsafe_b64decode(s + "=" * ((4 - len(s) % 4) % 4))


def verify_ticket(ticket, filename, platform):
    if not SECRET or "." not in ticket:
        raise HTTPException(401, "invalid_ticket")
    encoded, sig = ticket.rsplit(".", 1)
    expected = hmac.new(SECRET.encode(), encoded.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(sig, expected):
        raise HTTPException(401, "invalid_ticket")
    try:
        claims = json.loads(b64decode(encoded))
    except Exception:
        raise HTTPException(401, "invalid_ticket")
    now = int(time.time())
    if claims.get("exp", 0) < now or claims.get("iat", now) > now + 30:
        raise HTTPException(401, "expired_ticket")
    if claims.get("scope") != "convert_docx":
        raise HTTPException(403, "invalid_scope")
    if claims.get("filename") != filename:
        raise HTTPException(403, "filename_mismatch")
    if claims.get("platform", "general") != platform:
        raise HTTPException(403, "platform_mismatch")
    return claims


def cleanup():
    now = time.time()
    for k, j in list(JOBS.items()):
        if now - j["created"] > TTL:
            shutil.rmtree(j["dir"], ignore_errors=True)
            JOBS.pop(k, None)


def _zip_xml(z, name):
    try:
        return etree.fromstring(z.read(name))
    except Exception:
        return None


def inspect_docx(data):
    if len(data) > 20 * 1024 * 1024:
        raise HTTPException(413, "file_too_large")
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except Exception:
        raise HTTPException(415, "invalid_docx")
    names = z.namelist()
    if "word/document.xml" not in names:
        raise HTTPException(415, "invalid_docx")
    root = _zip_xml(z, "word/document.xml")
    if root is None:
        raise HTTPException(415, "invalid_docx")
    words = len(re.findall(r"[\w\u0600-\u06ff]+", " ".join(root.xpath("//w:t/text()", namespaces=NS))))
    refs = root.xpath("//w:footnoteReference", namespaces=NS)
    footnote_ids = [r.get("{%s}id" % W_NS) for r in refs if r.get("{%s}id" % W_NS) not in (None, "-1", "0")]
    tables = len(root.xpath("//w:tbl", namespaces=NS))
    unsupported_footnote_tables = 0
    if "word/footnotes.xml" in names:
        fr = _zip_xml(z, "word/footnotes.xml")
        if fr is not None:
            unsupported_footnote_tables = len(fr.xpath("//w:footnote[w:tbl and number(@w:id) > 0]", namespaces=NS))
    stats = {
        "words": words,
        "pages": max(1, (words + 349) // 350),
        "tables": tables + unsupported_footnote_tables,
        "images": sum(1 for n in names if n.startswith("word/media/") and not n.endswith("/")),
        "footnotes": len(footnote_ids),
        "footnote_tables": unsupported_footnote_tables,
    }
    z.close()
    return stats


def _docx_title(data):
    try:
        doc = Document(io.BytesIO(data))
        title = (doc.core_properties.title or "").strip()
        if title:
            return title
        for p in doc.paragraphs[:20]:
            if p.text and p.style and (p.style.name or "").lower() in ("title", "subtitle"):
                return p.text.strip()
    except Exception:
        pass
    return "كتاب رقمي"


def _pandoc_version():
    try:
        p = subprocess.run([PANDOC, "--version"], capture_output=True, text=True, timeout=10)
        first = (p.stdout or p.stderr).splitlines()[0] if (p.stdout or p.stderr) else ""
        return first.strip()
    except Exception:
        return "unavailable"


def _epub_paths(raw_epub):
    with zipfile.ZipFile(raw_epub) as z:
        container = etree.fromstring(z.read("META-INF/container.xml"))
        ns = {"c": CONTAINER_NS}
        rootfile = container.xpath("string(//c:rootfile/@full-path)", namespaces=ns)
        if not rootfile:
            raise RuntimeError("missing_opf")
        return PurePosixPath(rootfile)


def _postprocess_epub(raw_epub, out_epub, title):
    work = Path(raw_epub).with_suffix(".unpacked")
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True)
    with zipfile.ZipFile(raw_epub) as z:
        z.extractall(work)

    opf_rel = _epub_paths(raw_epub)
    opf_path = work / Path(str(opf_rel))
    opf = etree.parse(str(opf_path))
    opf_root = opf.getroot()
    ns = {"opf": OPF_NS, "dc": DC_NS}

    languages = opf.xpath("//dc:language", namespaces=ns)
    if languages:
        languages[0].text = "ar"
    metadata = opf.xpath("//opf:metadata", namespaces=ns)
    if metadata and not languages:
        el = etree.Element("{%s}language" % DC_NS); el.text = "ar"; metadata[0].append(el)
    titles = opf.xpath("//dc:title", namespaces=ns)
    if titles:
        titles[0].text = title
    elif metadata:
        el = etree.Element("{%s}title" % DC_NS); el.text = title; metadata[0].append(el)
    spines = opf.xpath("//opf:spine", namespaces=ns)
    if spines:
        spines[0].set("page-progression-direction", "rtl")

    manifest = {}
    for item in opf.xpath("//opf:manifest/opf:item", namespaces=ns):
        manifest[item.get("id")] = (item.get("href"), item.get("media-type"), item.get("properties", ""))

    opf.write(str(opf_path), encoding="UTF-8", xml_declaration=True, pretty_print=True)

    opf_dir = opf_path.parent
    xhtml_files = []
    css_files = []
    for _id, (href, media_type, _props) in manifest.items():
        if not href:
            continue
        p = opf_dir / Path(href)
        if media_type == "application/xhtml+xml" and p.exists():
            xhtml_files.append(p)
        elif media_type == "text/css" and p.exists():
            css_files.append(p)

    for p in xhtml_files:
        try:
            tree = etree.parse(str(p))
            root = tree.getroot()
            root.set("lang", "ar")
            root.set("{%s}lang" % XML_NS, "ar")
            root.set("dir", "rtl")
            bodies = root.xpath("//*[local-name()='body']")
            if bodies:
                bodies[0].set("dir", "rtl")
            tree.write(str(p), encoding="UTF-8", xml_declaration=True, pretty_print=True, doctype="<!DOCTYPE html>")
        except Exception:
            continue

    erth_css = """
html,body{direction:rtl;text-align:right}
body{font-family:serif;line-height:1.9;color:#202733}
p{text-align:justify;text-justify:inter-word}
h1,h2,h3,h4,h5,h6{direction:rtl;text-align:right;line-height:1.5}
a[epub\\:type='noteref'],a.footnote-ref{text-decoration:none}
.footnotes{margin-top:2.5em;border-top:1px solid #ccc;padding-top:1em}
.footnote{line-height:1.75}
""".strip() + "\n"
    if css_files:
        for p in css_files:
            current = p.read_text(encoding="utf-8", errors="ignore")
            p.write_text(erth_css + current, encoding="utf-8")
    else:
        styles = opf_dir / "styles"; styles.mkdir(exist_ok=True)
        css = styles / "erth.css"; css.write_text(erth_css, encoding="utf-8")
        manifest_el = opf.xpath("//opf:manifest", namespaces=ns)[0]
        item = etree.Element("{%s}item" % OPF_NS, id="erth-css", href="styles/erth.css", **{"media-type": "text/css"})
        manifest_el.append(item)
        opf.write(str(opf_path), encoding="UTF-8", xml_declaration=True, pretty_print=True)

    with zipfile.ZipFile(out_epub, "w") as zout:
        mimetype = work / "mimetype"
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        zout.writestr(zi, mimetype.read_bytes() if mimetype.exists() else b"application/epub+zip")
        for p in sorted(work.rglob("*")):
            if not p.is_file() or p == mimetype:
                continue
            arc = p.relative_to(work).as_posix()
            zout.write(p, arc, compress_type=zipfile.ZIP_DEFLATED)
    shutil.rmtree(work, ignore_errors=True)


def _extract_preview(epub_path):
    chapters = []
    footnotes = 0
    with zipfile.ZipFile(epub_path) as z:
        opf_rel = _epub_paths(epub_path)
        opf_root = etree.fromstring(z.read(str(opf_rel)))
        ns = {"opf": OPF_NS}
        manifest = {}
        for item in opf_root.xpath("//opf:manifest/opf:item", namespaces=ns):
            manifest[item.get("id")] = (item.get("href"), item.get("media-type"), item.get("properties", ""))
        opf_dir = opf_rel.parent
        for itemref in opf_root.xpath("//opf:spine/opf:itemref", namespaces=ns):
            idref = itemref.get("idref")
            meta = manifest.get(idref)
            if not meta:
                continue
            href, media_type, props = meta
            if media_type != "application/xhtml+xml" or "nav" in props or "title_page" in (href or ""):
                continue
            rel = (opf_dir / PurePosixPath(href)).as_posix()
            if rel not in z.namelist():
                continue
            try:
                root = etree.fromstring(z.read(rel))
            except Exception:
                continue
            bodies = root.xpath("//*[local-name()='body']")
            if not bodies:
                continue
            body = bodies[0]
            footnotes += len(body.xpath(".//*[@epub:type='footnote']", namespaces={"epub":"http://www.idpf.org/2007/ops"}))
            heading = body.xpath("string((.//*[local-name()='h1' or local-name()='h2'])[1])").strip()
            title = heading or f"قسم {len(chapters)+1}"
            html = "".join(etree.tostring(c, encoding="unicode", method="html") for c in body)
            chapters.append({"title": title, "html": html})
    return chapters, footnotes


def build_epub(data, out):
    title = _docx_title(data)
    out = Path(out)
    jobdir = out.parent
    src = jobdir / "source.docx"
    raw = jobdir / "pandoc.epub"
    css = jobdir / "erth.css"
    src.write_bytes(data)
    css.write_text("html,body{direction:rtl;text-align:right} p{text-align:justify;text-justify:inter-word}", encoding="utf-8")

    cmd = [
        PANDOC, str(src), "-o", str(raw), "--to=epub3", "--toc",
        "--metadata", f"title={title}", "--metadata", "lang=ar", "--css", str(css)
    ]
    p = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    if p.returncode != 0 or not raw.exists():
        raise RuntimeError("pandoc_failed: " + ((p.stderr or p.stdout or "unknown")[-1500:]))
    _postprocess_epub(raw, out, title)
    chapters, footnotes = _extract_preview(out)
    if not chapters:
        raise RuntimeError("preview_extraction_failed")
    return title, chapters, {"footnotes": footnotes, "engine": "pandoc", "engine_version": _pandoc_version()}


def epubcheck(path):
    try:
        p = subprocess.run(["java", "-jar", JAR, str(path)], capture_output=True, text=True, timeout=120)
    except Exception as e:
        return {"status": "failed", "passed": False, "tail": str(e)[-1000:]}
    return {"status": "passed" if p.returncode == 0 else "failed", "passed": p.returncode == 0, "tail": (p.stdout + p.stderr)[-6000:]}


def get_job(jobid, token):
    cleanup(); j = JOBS.get(jobid)
    if not j or not hmac.compare_digest(j["token"], token):
        raise HTTPException(404, "missing_or_expired")
    return j


@app.get("/health")
def health():
    return {
        "ok": True,
        "service": "ERTH Conversion Worker",
        "version": "2.0.0",
        "engine": {"name": "pandoc", "available": shutil.which(PANDOC) is not None, "version": _pandoc_version()},
        "epubcheck": Path(JAR).exists(),
        "features": {"footnotes": True, "rtl_postprocess": True, "docx": True}
    }


@app.post("/v1/jobs")
async def create_job(file: UploadFile = File(...), platform: str = Form("general"), ticket: str = Form(...)):
    cleanup(); filename = file.filename or ""
    verify_ticket(ticket, filename, platform)
    if not filename.lower().endswith(".docx"):
        raise HTTPException(415, "docx_only")
    data = await file.read(); stats = inspect_docx(data)
    # MVP v2: text-first DOCX + footnotes. Tables/images stay review-only for now.
    if stats["pages"] > 300 or stats["tables"] or stats["images"]:
        raise HTTPException(422, detail={"status": "team_review", "analysis": stats, "reason": "outside_mvp_auto_scope"})
    jobid = str(uuid.uuid4()); token = secrets.token_urlsafe(32); d = BASE / jobid; d.mkdir(); out = d / "book.epub"
    try:
        title, chapters, conversion = build_epub(data, out)
        qa = epubcheck(out)
    except HTTPException:
        raise
    except Exception as e:
        JOBS[jobid] = {"id": jobid, "created": time.time(), "dir": str(d), "token": token, "status": "failed", "message": "Conversion failed", "error": str(e)[-2000:]}
        return {"ok": True, "job": {"id": jobid}, "access_token": token}
    if not qa["passed"]:
        JOBS[jobid] = {"id": jobid, "created": time.time(), "dir": str(d), "token": token, "status": "failed", "message": "EPUBCheck failed", "qa": qa, "conversion": conversion, "analysis": stats}
    else:
        JOBS[jobid] = {"id": jobid, "created": time.time(), "dir": str(d), "token": token, "status": "completed", "message": "تم التحويل", "qa": {"passed": True}, "epubcheck": qa, "title": title, "chapters": chapters, "conversion": conversion, "analysis": stats, "path": str(out), "name": Path(filename).stem + ".epub"}
    return {"ok": True, "job": {"id": jobid}, "access_token": token}


@app.get("/v1/jobs/{jobid}")
def job_status(jobid: str, token: str = Query(...)):
    j = get_job(jobid, token)
    payload = {"id": j["id"], "status": j["status"], "message": j.get("message", "")}
    if j["status"] == "failed":
        payload["error"] = j.get("error", "")
        payload["qa"] = j.get("qa", {})
    return {"ok": True, "job": payload}


@app.get("/v1/jobs/{jobid}/preview")
def preview(jobid: str, token: str = Query(...)):
    j = get_job(jobid, token)
    if j["status"] != "completed":
        raise HTTPException(409, "not_completed")
    return {"ok": True, "title": j["title"], "chapters": j["chapters"], "qa": j["qa"], "epubcheck": j["epubcheck"], "conversion": j.get("conversion", {}), "analysis": j.get("analysis", {})}


@app.get("/v1/jobs/{jobid}/download")
def download(jobid: str, token: str = Query(...)):
    j = get_job(jobid, token)
    if j["status"] != "completed" or not Path(j["path"]).exists():
        raise HTTPException(404, "missing")
    return FileResponse(j["path"], media_type="application/epub+zip", filename=j["name"])
