import io, os, re, time, uuid, zipfile, subprocess, secrets, shutil, json, base64, hmac, hashlib
from pathlib import Path
from html import escape
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from lxml import etree

app=FastAPI(title="ERTH Conversion Worker",version="1.0.0")
SECRET=os.getenv("ERTH_WORKER_SECRET","")
ORIGIN=os.getenv("ERTH_ALLOWED_ORIGIN","https://erthpub.com")
TTL=int(os.getenv("ERTH_JOB_TTL","2700"))
JAR=os.getenv("EPUBCHECK_JAR","/opt/epubcheck/epubcheck.jar")
BASE=Path("/tmp/erth-jobs"); BASE.mkdir(parents=True,exist_ok=True)
JOBS={}
app.add_middleware(CORSMiddleware,allow_origins=[ORIGIN],allow_credentials=False,allow_methods=["GET","POST"],allow_headers=["*"])

def b64decode(s):
    return base64.urlsafe_b64decode(s+"="*((4-len(s)%4)%4))

def verify_ticket(ticket, filename, platform):
    if not SECRET or "." not in ticket: raise HTTPException(401,"invalid_ticket")
    encoded,sig=ticket.rsplit(".",1)
    expected=hmac.new(SECRET.encode(),encoded.encode(),hashlib.sha256).hexdigest()
    if not hmac.compare_digest(sig,expected): raise HTTPException(401,"invalid_ticket")
    try: claims=json.loads(b64decode(encoded))
    except Exception: raise HTTPException(401,"invalid_ticket")
    now=int(time.time())
    if claims.get("exp",0)<now or claims.get("iat",now)>now+30: raise HTTPException(401,"expired_ticket")
    if claims.get("scope")!="convert_docx": raise HTTPException(403,"invalid_scope")
    if claims.get("filename")!=filename: raise HTTPException(403,"filename_mismatch")
    if claims.get("platform","general")!=platform: raise HTTPException(403,"platform_mismatch")
    return claims

def cleanup():
    now=time.time()
    for k,j in list(JOBS.items()):
        if now-j["created"]>TTL:
            shutil.rmtree(j["dir"],ignore_errors=True); JOBS.pop(k,None)

def inspect_docx(data):
    if len(data)>20*1024*1024: raise HTTPException(413,"file_too_large")
    try: z=zipfile.ZipFile(io.BytesIO(data))
    except Exception: raise HTTPException(415,"invalid_docx")
    names=z.namelist()
    if "word/document.xml" not in names: raise HTTPException(415,"invalid_docx")
    root=etree.fromstring(z.read("word/document.xml"))
    ns={"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    words=len(re.findall(r"[\w\u0600-\u06ff]+"," ".join(root.xpath("//w:t/text()",namespaces=ns))))
    return {"words":words,"pages":max(1,(words+349)//350),
      "tables":len(root.xpath("//w:tbl",namespaces=ns)),
      "images":sum(1 for n in names if n.startswith("word/media/") and not n.endswith("/")),
      "footnotes":1 if "word/footnotes.xml" in names else 0}

def heading_level(p):
    style=(p.style.name or "").lower() if p.style else ""
    m=re.search(r"(?:heading|head|title)\s*([1-6])?",style)
    return int(m.group(1) or 1) if m else 0

def para_html(p):
    out=[]
    for r in p.runs:
        t=escape(r.text,quote=True)
        if not t: continue
        if r.bold:t="<strong>"+t+"</strong>"
        if r.italic:t="<em>"+t+"</em>"
        out.append(t)
    return "".join(out).strip()

def build_epub(data,out):
    doc=Document(io.BytesIO(data)); title=(doc.core_properties.title or "").strip() or "كتاب رقمي"
    chapters=[]; cur={"title":title,"body":[]}
    for p in doc.paragraphs:
        html=para_html(p)
        if not html: continue
        level=heading_level(p)
        if level and level<=2:
            if cur["body"]:chapters.append(cur)
            cur={"title":p.text.strip() or "فصل","body":[]}
        elif level:cur["body"].append("<h%d>%s</h%d>"%(level,html,level))
        else:cur["body"].append("<p>"+html+"</p>")
    if cur["body"]:chapters.append(cur)
    if not chapters:raise HTTPException(422,"no_convertible_text")
    if len(chapters)>120:raise HTTPException(422,detail={"status":"team_review","reason":"too_many_chapters"})
    uid="urn:uuid:"+str(uuid.uuid4()); modified=time.strftime("%Y-%m-%dT%H:%M:%SZ",time.gmtime())
    manifest=[];spine=[];nav=[];preview=[]
    with zipfile.ZipFile(out,"w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"),"application/epub+zip",compress_type=zipfile.ZIP_STORED)
        z.writestr("META-INF/container.xml",'<?xml version="1.0" encoding="UTF-8"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="EPUB/package.opf" media-type="application/oebps-package+xml"/></rootfiles></container>')
        z.writestr("EPUB/styles.css","html{direction:rtl}body{direction:rtl;text-align:right;font-family:serif;line-height:1.9;margin:5%;color:#202733}p{text-align:justify;text-justify:inter-word;margin:0 0 1em}h1,h2,h3,h4,h5,h6{direction:rtl;text-align:right;line-height:1.5}")
        for i,c in enumerate(chapters,1):
            fn="chapter-%03d.xhtml"%i; ident="ch%d"%i
            body="<h1>"+escape(c["title"],quote=True)+"</h1>"+"".join(c["body"])
            x='<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="ar" lang="ar" dir="rtl"><head><meta charset="utf-8"/><title>'+escape(c["title"],quote=True)+'</title><link rel="stylesheet" type="text/css" href="styles.css"/></head><body>'+body+'</body></html>'
            z.writestr("EPUB/"+fn,x)
            manifest.append('<item id="%s" href="%s" media-type="application/xhtml+xml"/>'%(ident,fn));spine.append('<itemref idref="%s"/>'%ident)
            nav.append('<li><a href="%s">%s</a></li>'%(fn,escape(c["title"],quote=True)));preview.append({"title":c["title"],"html":body})
        z.writestr("EPUB/nav.xhtml",'<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE html><html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="ar" lang="ar" dir="rtl"><head><meta charset="utf-8"/><title>الفهرس</title></head><body><nav epub:type="toc" id="toc"><h1>الفهرس</h1><ol>'+"".join(nav)+'</ol></nav></body></html>')
        opf='<?xml version="1.0" encoding="UTF-8"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="book-id" xml:lang="ar" dir="rtl"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="book-id">'+uid+'</dc:identifier><dc:title>'+escape(title,quote=True)+'</dc:title><dc:language>ar</dc:language><meta property="dcterms:modified">'+modified+'</meta></metadata><manifest><item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/><item id="css" href="styles.css" media-type="text/css"/>'+"".join(manifest)+'</manifest><spine page-progression-direction="rtl">'+"".join(spine)+'</spine></package>'
        z.writestr("EPUB/package.opf",opf)
    return title,preview

def epubcheck(path):
    p=subprocess.run(["java","-jar",JAR,str(path)],capture_output=True,text=True,timeout=90)
    return {"status":"passed" if p.returncode==0 else "failed","passed":p.returncode==0,"tail":(p.stdout+p.stderr)[-4000:]}

def get_job(jobid,token):
    cleanup();j=JOBS.get(jobid)
    if not j or not hmac.compare_digest(j["token"],token):raise HTTPException(404,"missing_or_expired")
    return j

@app.get("/health")
def health():return {"ok":True,"service":"ERTH Conversion Worker","epubcheck":Path(JAR).exists()}

@app.post("/v1/jobs")
async def create_job(file:UploadFile=File(...),platform:str=Form("general"),ticket:str=Form(...)):
    cleanup();filename=file.filename or ""
    verify_ticket(ticket,filename,platform)
    if not filename.lower().endswith(".docx"):raise HTTPException(415,"docx_only")
    data=await file.read();stats=inspect_docx(data)
    if stats["pages"]>300 or stats["tables"] or stats["images"] or stats["footnotes"]:
        raise HTTPException(422,detail={"status":"team_review","analysis":stats,"reason":"outside_mvp_auto_scope"})
    jobid=str(uuid.uuid4());token=secrets.token_urlsafe(32);d=BASE/jobid;d.mkdir();out=d/"book.epub"
    title,chapters=build_epub(data,out);qa=epubcheck(out)
    if not qa["passed"]:
        JOBS[jobid]={"id":jobid,"created":time.time(),"dir":str(d),"token":token,"status":"failed","message":"EPUBCheck failed","qa":qa}
    else:
        JOBS[jobid]={"id":jobid,"created":time.time(),"dir":str(d),"token":token,"status":"completed","message":"تم التحويل","qa":{"passed":True},"epubcheck":qa,"title":title,"chapters":chapters,"path":str(out),"name":Path(filename).stem+".epub"}
    return {"ok":True,"job":{"id":jobid},"access_token":token}

@app.get("/v1/jobs/{jobid}")
def job_status(jobid:str,token:str=Query(...)):
    j=get_job(jobid,token);return {"ok":True,"job":{"id":j["id"],"status":j["status"],"message":j.get("message","")}}

@app.get("/v1/jobs/{jobid}/preview")
def preview(jobid:str,token:str=Query(...)):
    j=get_job(jobid,token)
    if j["status"]!="completed":raise HTTPException(409,"not_completed")
    return {"ok":True,"title":j["title"],"chapters":j["chapters"],"qa":j["qa"],"epubcheck":j["epubcheck"]}

@app.get("/v1/jobs/{jobid}/download")
def download(jobid:str,token:str=Query(...)):
    j=get_job(jobid,token)
    if j["status"]!="completed" or not Path(j["path"]).exists():raise HTTPException(404,"missing")
    return FileResponse(j["path"],media_type="application/epub+zip",filename=j["name"])
