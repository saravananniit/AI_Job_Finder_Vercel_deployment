# """AI Job Finder - City Based web search (FastAPI, Vercel-ready).

# Flow (mirrors the n8n workflow):
#   form -> build queries -> Serper web search -> normalize + dedupe
#        -> Firecrawl scrape (optional, FIRECRAWL_API_KEY env) -> format -> XLSX
# """
# import asyncio
# import io
# import os
# import re
# from pathlib import Path
# from urllib.parse import urlparse, urlunparse

# import httpx
# from fastapi import FastAPI, HTTPException
# from fastapi.responses import FileResponse, StreamingResponse
# from openpyxl import Workbook
# from openpyxl.styles import Alignment, Font, PatternFill
# from pydantic import BaseModel, Field

# app = FastAPI(title="AI Job Finder")

# SERPER_URL = "https://google.serper.dev/search"
# FIRECRAWL_URL = "https://api.firecrawl.dev/v1/scrape"
# FIRECRAWL_KEY = os.getenv("FIRECRAWL_API_KEY", "")
# SERPER_GL = os.getenv("SERPER_GL", "in")  # Google country code for results

# COLUMNS = [
#     "Job Title", "Company", "Location", "Experience", "Qualification",
#     "Posted Date", "Description", "Source", "Apply Link",
# ]

# # Google "tbs" time filters for the Posted within dropdown
# POSTED_TBS = {
#     "anytime": None,
#     "30 days": "qdr:m",
#     "14 days": "qdr:d14",
#     "7 days": "qdr:w",
#     "3 days": "qdr:d3",
#     "24 hours": "qdr:d",
# }

# QUERY_TEMPLATES = [
#     '{role} jobs in {city} {exp}',
#     '{role} job opening {city} apply',
#     '{role} hiring {city} careers',
# ]

# SKIP_DOMAINS = ("youtube.com", "facebook.com", "instagram.com", "twitter.com", "x.com", "pinterest.com")


# class FindJobsRequest(BaseModel):
#     roles: str = Field(..., min_length=1)
#     city: str = Field(..., min_length=1)
#     posted_within: str = "Anytime"
#     experience: str = "0-3 years"
#     max_results: int = Field(25, ge=1, le=100)
#     serper_api_key: str = Field(..., min_length=1)


# class ExportRequest(BaseModel):
#     jobs: list[dict]


# # ---------- 1. build search queries ----------
# def build_queries(roles: str, city: str, experience: str) -> list[tuple[str, str]]:
#     role_list = [r.strip() for r in re.split(r"[,\n;]+", roles) if r.strip()]
#     queries = []
#     for role in role_list:
#         for tpl in QUERY_TEMPLATES:
#             queries.append((role, tpl.format(role=role, city=city, exp=experience).strip()))
#     return queries


# # ---------- 2. web search (Serper) ----------
# async def serper_search(client: httpx.AsyncClient, key: str, query: str, tbs: str | None) -> list[dict]:
#     payload = {"q": query, "gl": SERPER_GL, "num": 10}
#     if tbs:
#         payload["tbs"] = tbs
#     r = await client.post(SERPER_URL, json=payload, headers={"X-API-KEY": key, "Content-Type": "application/json"})
#     if r.status_code in (401, 403):
#         raise HTTPException(401, "Serper rejected the API key. Check it and try again.")
#     if r.status_code == 429:
#         raise HTTPException(429, "Serper rate limit or credits exhausted.")
#     r.raise_for_status()
#     return r.json().get("organic", [])


# # ---------- 3. normalize + dedupe ----------
# def clean_url(url: str) -> str:
#     p = urlparse(url)
#     return urlunparse((p.scheme, p.netloc.lower().replace("www.", ""), p.path.rstrip("/"), "", "", ""))


# def split_title(title: str) -> tuple[str, str]:
#     """'Python Developer - Acme - Chennai | Naukri' -> ('Python Developer', 'Acme')."""
#     parts = [p.strip() for p in re.split(r"\s[-|–—]\s|\sat\s|\s@\s", title) if p.strip()]
#     job = parts[0] if parts else title
#     company = parts[1] if len(parts) > 1 else ""
#     return job, company


# def normalize(results: list[tuple[str, dict]]) -> list[dict]:
#     seen_links, seen_keys, jobs = set(), set(), []
#     for role, item in results:
#         link = item.get("link", "")
#         if not link or any(d in link for d in SKIP_DOMAINS):
#             continue
#         norm = clean_url(link)
#         title, company = split_title(item.get("title", ""))
#         key = re.sub(r"\W+", "", f"{title}{company}").lower()
#         if norm in seen_links or (company and key in seen_keys):
#             continue
#         seen_links.add(norm)
#         seen_keys.add(key)
#         jobs.append({
#             "role": role,
#             "title": title,
#             "company": company,
#             "snippet": item.get("snippet", ""),
#             "date": item.get("date", ""),
#             "link": link,
#             "source": urlparse(link).netloc.replace("www.", ""),
#         })
#     return jobs


# def interleave(jobs: list[dict], limit: int) -> list[dict]:
#     """Spread the max-results budget evenly across the searched roles."""
#     buckets: dict[str, list[dict]] = {}
#     for j in jobs:
#         buckets.setdefault(j["role"], []).append(j)
#     out = []
#     while len(out) < limit and any(buckets.values()):
#         for b in buckets.values():
#             if b and len(out) < limit:
#                 out.append(b.pop(0))
#     return out


# # ---------- 4. scrape pages (Firecrawl) ----------
# async def scrape(client: httpx.AsyncClient, sem: asyncio.Semaphore, url: str) -> dict:
#     if not FIRECRAWL_KEY:
#         return {}
#     async with sem:
#         try:
#             r = await client.post(
#                 FIRECRAWL_URL,
#                 json={"url": url, "formats": ["markdown"], "onlyMainContent": True, "timeout": 20000},
#                 headers={"Authorization": f"Bearer {FIRECRAWL_KEY}"},
#             )
#             if r.status_code != 200:
#                 return {}
#             return r.json().get("data", {}) or {}
#         except Exception:
#             return {}


# # ---------- 5. format ----------
# EXP_RE = re.compile(r"(\d{1,2}\s*(?:-|–|to)\s*\d{1,2}\s*\+?\s*(?:years?|yrs?)|\d{1,2}\s*\+\s*(?:years?|yrs?)|\d{1,2}\s*(?:years?|yrs?)\s+of\s+experience|fresher[s]?)", re.I)
# QUAL_RE = re.compile(r"(B\.?\s?E\.?|B\.?\s?Tech|M\.?\s?Tech|B\.?\s?Sc|M\.?\s?Sc|BCA|MCA|MBA|B\.?\s?Com|M\.?\s?Com|Bachelor(?:'s)?(?: of [A-Za-z ]{3,25})?|Master(?:'s)?(?: of [A-Za-z ]{3,25})?|Diploma|Graduate|Post[- ]?graduate|PhD)", re.I)
# POSTED_RE = re.compile(r"((?:posted|published)?\s*(?:\d+\+?\s*(?:minutes?|hours?|days?|weeks?|months?)\s+ago|today|yesterday|just now))", re.I)


# def clean_markdown(md: str) -> str:
#     md = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", md)
#     md = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", md)
#     md = re.sub(r"[#*_`>|]+", " ", md)
#     return re.sub(r"\s+", " ", md).strip()


# def first(rx: re.Pattern, text: str) -> str:
#     m = rx.search(text)
#     return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""


# def format_job(job: dict, page: dict, city: str, experience: str) -> dict:
#     md = page.get("markdown", "") or ""
#     meta = page.get("metadata", {}) or {}
#     text = clean_markdown(md)
#     blob = f"{job['title']} {job['snippet']} {text}"

#     company = job["company"] or meta.get("ogSiteName", "") or ""
#     description = (text[:700] if text else job["snippet"]).strip()
#     qual = ", ".join(dict.fromkeys(m.strip() for m in QUAL_RE.findall(text or job["snippet"])))[:120]
#     location = city if re.search(re.escape(city), blob, re.I) else (city + " (verify)")

#     return {
#         "Job Title": job["title"],
#         "Company": company,
#         "Location": location,
#         "Experience": first(EXP_RE, blob) or experience,
#         "Qualification": qual,
#         "Posted Date": job["date"] or first(POSTED_RE, blob),
#         "Description": description,
#         "Source": job["source"],
#         "Apply Link": job["link"],
#     }


# # ---------- routes ----------
# INDEX_HTML = Path(__file__).resolve().parent.parent / "public" / "index.html"


# @app.get("/", include_in_schema=False)
# def home():
#     # Local runs only: on Vercel, public/index.html is served statically at "/".
#     return FileResponse(INDEX_HTML)


# @app.post("/api/find-jobs")
# async def find_jobs(req: FindJobsRequest):
#     tbs = POSTED_TBS.get(req.posted_within.strip().lower())
#     queries = build_queries(req.roles, req.city, req.experience)
#     if not queries:
#         raise HTTPException(400, "Enter at least one job role.")

#     async with httpx.AsyncClient(timeout=30) as client:
#         batches = await asyncio.gather(*[serper_search(client, req.serper_api_key, q, tbs) for _, q in queries])
#         flat = [(role, item) for (role, _), items in zip(queries, batches) for item in items]
#         jobs = interleave(normalize(flat), req.max_results)
#         if not jobs:
#             return {"count": 0, "jobs": [], "scraped": bool(FIRECRAWL_KEY)}

#         sem = asyncio.Semaphore(8)
#         pages = await asyncio.gather(*[scrape(client, sem, j["link"]) for j in jobs])

#     rows = [format_job(j, p, req.city, req.experience) for j, p in zip(jobs, pages)]
#     return {"count": len(rows), "jobs": rows, "scraped": bool(FIRECRAWL_KEY)}


# @app.post("/api/export")
# def export_excel(req: ExportRequest):
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Jobs"
#     ws.append(COLUMNS)
#     for c in ws[1]:
#         c.font = Font(bold=True, color="FFFFFF")
#         c.fill = PatternFill("solid", fgColor="1F3A5F")
#     for job in req.jobs:
#         ws.append([str(job.get(col, ""))[:32000] for col in COLUMNS])
#     link_col = COLUMNS.index("Apply Link") + 1
#     for row in ws.iter_rows(min_row=2, min_col=link_col, max_col=link_col):
#         for c in row:
#             if c.value:
#                 c.hyperlink = c.value
#                 c.font = Font(color="0563C1", underline="single")
#     widths = [34, 24, 18, 16, 26, 16, 70, 22, 50]
#     for i, w in enumerate(widths, start=1):
#         ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
#     for row in ws.iter_rows(min_row=2):
#         for c in row:
#             c.alignment = Alignment(wrap_text=True, vertical="top")
#     ws.freeze_panes = "A2"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return StreamingResponse(
#         buf,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": 'attachment; filename="jobs.xlsx"'},
#     )


# # ---------- local dev: serve the page from the same origin ----------
# # On Vercel, public/index.html is served by the CDN, so this route is only used locally.
# PUBLIC_DIR = Path(__file__).resolve().parent.parent / "public"


# @app.get("/", include_in_schema=False)
# def home():
#     return FileResponse(PUBLIC_DIR / "index.html")

"""AI Job Finder - City Based web search (FastAPI, Vercel-ready).

Flow (mirrors the n8n workflow):
  form -> build queries -> Serper web search -> normalize + dedupe
       -> Firecrawl scrape (optional, FIRECRAWL_API_KEY env) -> format -> XLSX
"""
import asyncio
import io
import os
import re
from pathlib import Path
from urllib.parse import urlparse, urlunparse

import httpx
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pydantic import BaseModel, Field

app = FastAPI(title="AI Job Finder")

SERPER_URL = "https://google.serper.dev/search"
FIRECRAWL_URL = "https://api.firecrawl.dev/v1/scrape"
FIRECRAWL_KEY = os.getenv("FIRECRAWL_API_KEY", "")
SERPER_GL = os.getenv("SERPER_GL", "in")  # Google country code for results

COLUMNS = [
    "Job Title", "Company", "Location", "Experience", "Qualification",
    "Posted Date", "Description", "Source", "Apply Link",
]

# Google "tbs" time filters for the Posted within dropdown
POSTED_TBS = {
    "anytime": None,
    "30 days": "qdr:m",
    "14 days": "qdr:d14",
    "7 days": "qdr:w",
    "3 days": "qdr:d3",
    "24 hours": "qdr:d",
}

QUERY_TEMPLATES = [
    '{role} jobs in {city} {exp}',
    '{role} job opening {city} apply',
    '{role} hiring {city} careers',
]

SKIP_DOMAINS = ("youtube.com", "facebook.com", "instagram.com", "twitter.com", "x.com", "pinterest.com")


class FindJobsRequest(BaseModel):
    roles: str = Field(..., min_length=1)
    city: str = Field(..., min_length=1)
    posted_within: str = "Anytime"
    experience: str = "0-3 years"
    max_results: int = Field(25, ge=1, le=100)
    serper_api_key: str = Field(..., min_length=1)


class ExportRequest(BaseModel):
    jobs: list[dict]


# ---------- 1. build search queries ----------
def build_queries(roles: str, city: str, experience: str) -> list[tuple[str, str]]:
    role_list = [r.strip() for r in re.split(r"[,\n;]+", roles) if r.strip()]
    queries = []
    for role in role_list:
        for tpl in QUERY_TEMPLATES:
            queries.append((role, tpl.format(role=role, city=city, exp=experience).strip()))
    return queries


# ---------- 2. web search (Serper) ----------
async def serper_search(client: httpx.AsyncClient, key: str, query: str, tbs: str | None) -> list[dict]:
    payload = {"q": query, "gl": SERPER_GL, "num": 10}
    if tbs:
        payload["tbs"] = tbs
    r = await client.post(SERPER_URL, json=payload, headers={"X-API-KEY": key, "Content-Type": "application/json"})
    if r.status_code in (401, 403):
        raise HTTPException(401, "Serper rejected the API key. Check it and try again.")
    if r.status_code == 429:
        raise HTTPException(429, "Serper rate limit or credits exhausted.")
    r.raise_for_status()
    return r.json().get("organic", [])


# ---------- 3. normalize + dedupe ----------
def clean_url(url: str) -> str:
    p = urlparse(url)
    return urlunparse((p.scheme, p.netloc.lower().replace("www.", ""), p.path.rstrip("/"), "", "", ""))


def split_title(title: str) -> tuple[str, str]:
    """'Python Developer - Acme - Chennai | Naukri' -> ('Python Developer', 'Acme')."""
    parts = [p.strip() for p in re.split(r"\s[-|–—]\s|\sat\s|\s@\s", title) if p.strip()]
    job = parts[0] if parts else title
    company = parts[1] if len(parts) > 1 else ""
    return job, company


def normalize(results: list[tuple[str, dict]]) -> list[dict]:
    seen_links, seen_keys, jobs = set(), set(), []
    for role, item in results:
        link = item.get("link", "")
        if not link or any(d in link for d in SKIP_DOMAINS):
            continue
        norm = clean_url(link)
        title, company = split_title(item.get("title", ""))
        key = re.sub(r"\W+", "", f"{title}{company}").lower()
        if norm in seen_links or (company and key in seen_keys):
            continue
        seen_links.add(norm)
        seen_keys.add(key)
        jobs.append({
            "role": role,
            "title": title,
            "company": company,
            "snippet": item.get("snippet", ""),
            "date": item.get("date", ""),
            "link": link,
            "source": urlparse(link).netloc.replace("www.", ""),
        })
    return jobs


def interleave(jobs: list[dict], limit: int) -> list[dict]:
    """Spread the max-results budget evenly across the searched roles."""
    buckets: dict[str, list[dict]] = {}
    for j in jobs:
        buckets.setdefault(j["role"], []).append(j)
    out = []
    while len(out) < limit and any(buckets.values()):
        for b in buckets.values():
            if b and len(out) < limit:
                out.append(b.pop(0))
    return out


# ---------- 4. scrape pages (Firecrawl) ----------
async def scrape(client: httpx.AsyncClient, sem: asyncio.Semaphore, url: str) -> dict:
    if not FIRECRAWL_KEY:
        return {}
    async with sem:
        try:
            r = await client.post(
                FIRECRAWL_URL,
                json={"url": url, "formats": ["markdown"], "onlyMainContent": True, "timeout": 20000},
                headers={"Authorization": f"Bearer {FIRECRAWL_KEY}"},
            )
            if r.status_code != 200:
                return {}
            return r.json().get("data", {}) or {}
        except Exception:
            return {}


# ---------- 5. format ----------
EXP_RE = re.compile(r"(\d{1,2}\s*(?:-|–|to)\s*\d{1,2}\s*\+?\s*(?:years?|yrs?)|\d{1,2}\s*\+\s*(?:years?|yrs?)|\d{1,2}\s*(?:years?|yrs?)\s+of\s+experience|fresher[s]?)", re.I)
QUAL_RE = re.compile(r"(B\.?\s?E\.?|B\.?\s?Tech|M\.?\s?Tech|B\.?\s?Sc|M\.?\s?Sc|BCA|MCA|MBA|B\.?\s?Com|M\.?\s?Com|Bachelor(?:'s)?(?: of [A-Za-z ]{3,25})?|Master(?:'s)?(?: of [A-Za-z ]{3,25})?|Diploma|Graduate|Post[- ]?graduate|PhD)", re.I)
POSTED_RE = re.compile(r"((?:posted|published)?\s*(?:\d+\+?\s*(?:minutes?|hours?|days?|weeks?|months?)\s+ago|today|yesterday|just now))", re.I)


def clean_markdown(md: str) -> str:
    md = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", md)
    md = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", md)
    md = re.sub(r"[#*_`>|]+", " ", md)
    return re.sub(r"\s+", " ", md).strip()


def first(rx: re.Pattern, text: str) -> str:
    m = rx.search(text)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""


def format_job(job: dict, page: dict, city: str, experience: str) -> dict:
    md = page.get("markdown", "") or ""
    meta = page.get("metadata", {}) or {}
    text = clean_markdown(md)
    blob = f"{job['title']} {job['snippet']} {text}"

    company = job["company"] or meta.get("ogSiteName", "") or ""
    description = (text[:700] if text else job["snippet"]).strip()
    qual = ", ".join(dict.fromkeys(m.strip() for m in QUAL_RE.findall(text or job["snippet"])))[:120]
    location = city if re.search(re.escape(city), blob, re.I) else (city + " (verify)")

    return {
        "Job Title": job["title"],
        "Company": company,
        "Location": location,
        "Experience": first(EXP_RE, blob) or experience,
        "Qualification": qual,
        "Posted Date": job["date"] or first(POSTED_RE, blob),
        "Description": description,
        "Source": job["source"],
        "Apply Link": job["link"],
    }


# ---------- routes ----------
def _find_index_html() -> Path | None:
    here = Path(__file__).resolve().parent
    candidates = [
        here.parent / "public" / "index.html",  # api/index.py + public/index.html
        here / "public" / "index.html",          # index.py at project root
        here / "index.html",                     # index.html beside index.py
        Path.cwd() / "public" / "index.html",
        Path.cwd() / "index.html",
    ]
    return next((c for c in candidates if c.exists()), None)


@app.get("/", include_in_schema=False)
def home():
    # Local runs only: on Vercel, public/index.html is served statically at "/".
    page = _find_index_html()
    if not page:
        raise HTTPException(404, "index.html not found. Put it in public/ next to the api/ folder.")
    return FileResponse(page)


@app.post("/api/find-jobs")
async def find_jobs(req: FindJobsRequest):
    tbs = POSTED_TBS.get(req.posted_within.strip().lower())
    queries = build_queries(req.roles, req.city, req.experience)
    if not queries:
        raise HTTPException(400, "Enter at least one job role.")

    async with httpx.AsyncClient(timeout=30) as client:
        batches = await asyncio.gather(*[serper_search(client, req.serper_api_key, q, tbs) for _, q in queries])
        flat = [(role, item) for (role, _), items in zip(queries, batches) for item in items]
        jobs = interleave(normalize(flat), req.max_results)
        if not jobs:
            return {"count": 0, "jobs": [], "scraped": bool(FIRECRAWL_KEY)}

        sem = asyncio.Semaphore(8)
        pages = await asyncio.gather(*[scrape(client, sem, j["link"]) for j in jobs])

    rows = [format_job(j, p, req.city, req.experience) for j, p in zip(jobs, pages)]
    return {"count": len(rows), "jobs": rows, "scraped": bool(FIRECRAWL_KEY)}


@app.post("/api/export")
def export_excel(req: ExportRequest):
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3A5F")
    for job in req.jobs:
        ws.append([str(job.get(col, ""))[:32000] for col in COLUMNS])
    link_col = COLUMNS.index("Apply Link") + 1
    for row in ws.iter_rows(min_row=2, min_col=link_col, max_col=link_col):
        for c in row:
            if c.value:
                c.hyperlink = c.value
                c.font = Font(color="0563C1", underline="single")
    widths = [34, 24, 18, 16, 26, 16, 70, 22, 50]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="jobs.xlsx"'},
    )


# ---------- local dev: serve the page from the same origin ----------
# On Vercel, public/index.html is served by the CDN, so this route is only used locally.
PUBLIC_DIR = Path(__file__).resolve().parent.parent / "public"


@app.get("/", include_in_schema=False)
def home():
    return FileResponse(PUBLIC_DIR / "index.html")