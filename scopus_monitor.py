# -*- coding: utf-8 -*-
"""
KTF Elmi Monitorinq — Scopus statistikası

Exceldəki müəllimlərin Scopus Author ID-lərinə görə Elsevier API-dən
məqalələri, istinadları, h-index-i və jurnal kvartillərini çəkir və
dashboard.html yaradır. Hər işə salınanda məlumat yenilənir və əvvəlki
yoxlama ilə müqayisə olunur ("+13 son yoxlamadan", "Yeni").

İstifadə:
    python scopus_monitor.py              # bir dəfə yenilə
    python scopus_monitor.py --loop 24    # hər 24 saatdan bir yenilə (pəncərə açıq qalır)
    python scopus_monitor.py --demo       # internetsiz, nümunə məlumatla dizayna bax

Qeyd: Elsevier API yalnız universitetin şəbəkəsindən (IP) tam işləyir.
"""
import argparse
import csv
import datetime as dt
import io
import json
import os
import random
import re
import sqlite3
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser

try:
    import openpyxl
except ImportError:  # ilk işə salınma: lazım olan kitabxananı özü quraşdırır
    import subprocess
    print("openpyxl quraşdırılır...", flush=True)
    subprocess.call([sys.executable, "-m", "pip", "install", "--user", "--quiet", "openpyxl"])
    import site
    try:
        sys.path.append(site.getusersitepackages())
    except Exception:
        pass
    import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
# KTF_HOME: yazıla bilən qovluq (config, Excel, data, dashboard.html)
# KTF_RES:  proqramla gələn fayllar (şablon, ilkin config və Excel) — masaüstü tətbiqdə quraşdırma qovluğu
BASE = os.environ.get("KTF_HOME") or _HERE
RES = os.environ.get("KTF_RES") or _HERE
os.makedirs(BASE, exist_ok=True)
DATA = os.path.join(BASE, "data")
os.makedirs(DATA, exist_ok=True)

API = "https://api.elsevier.com/content"
UA = "Mozilla/5.0 (KTF Elmi Monitorinq)"


# ----------------------------------------------------------------- config
def load_config():
    path = os.path.join(BASE, "config.json")
    if not os.path.exists(path) and os.path.exists(os.path.join(RES, "config.json")):
        import shutil
        shutil.copyfile(os.path.join(RES, "config.json"), path)
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg.setdefault("faculty_short", "KTF")
    cfg.setdefault("faculty_name", "KTF")
    cfg.setdefault("excel_file", "muellimler.xlsx")
    cfg.setdefault("skip_sheets", ["Dropdown options"])
    cfg.setdefault("output_html", "dashboard.html")
    cfg.setdefault("years_in_chart", 10)
    cfg.setdefault("request_pause_sec", 0.2)
    return cfg


STATUS = {"running": False, "last": "", "error": None, "finished": None}


def log(msg):
    line = f"[{dt.datetime.now():%H:%M:%S}] {msg}"
    STATUS["last"] = msg
    print(line, flush=True)


# ----------------------------------------------------------------- excel
ID_RE = re.compile(r"(?<!\d)(\d{9,11})(?!\d)")


def parse_ids(raw):
    """'60094747600 \n57219319680', '57216204317/ 3', 57201364732.0 → ['...', '...']"""
    if raw is None:
        return []
    if isinstance(raw, float):
        raw = str(int(raw))
    out = []
    for m in ID_RE.findall(str(raw)):
        if m not in out:
            out.append(m)
    return out


def _clean(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v not in (None, "") else ""


# Excel vərəq adı 31 simvoldan uzun ola bilmir — dashboard-da göstəriləcək tam adlar
DEPT_DISPLAY_NAMES = {
    "Proseslərin Təhlükəsizliyi": "Proseslərin Təhlükəsizliyi Mühəndisliyi",
    "Biotexnologiya və qida müh.": "Biotexnologiya və qida mühəndisliyi",
    "Proses təhlükəsizliyi müh.": "Proseslərin Təhlükəsizliyi Mühəndisliyi",
}


def read_teachers(cfg):
    """Hər vərəq = kafedra. Sütunlar: Ad, Vəzifə, Scopus ID (bir neçə ola bilər), WoS, ORCID, Qeyd."""
    path = cfg["excel_file"]
    if not os.path.isabs(path):
        path = os.path.join(BASE, path)
    import warnings
    warnings.filterwarnings("ignore", module="openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True)
    depts = []
    for ws in wb.worksheets:
        if ws.title.strip() in cfg["skip_sheets"]:
            continue
        teachers, seen = [], {}
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue  # başlıq
            row = list(row) + [None] * 6
            name = _clean(row[0])
            if not name or name.lower().startswith("ad, soyad"):
                continue
            ids = parse_ids(row[2])
            key = name.lower()
            if key in seen:  # eyni müəllim təkrar yazılıb — ID-ləri birləşdir
                t = seen[key]
                for x in ids:
                    if x not in t["scopus_ids"]:
                        t["scopus_ids"].append(x)
                t["scopus_id"] = t["scopus_ids"][0] if t["scopus_ids"] else None
                continue
            t = {
                "name": name,
                "position": _clean(row[1]),
                "scopus_ids": ids,
                "scopus_id": ids[0] if ids else None,
                "raw_id": _clean(row[2]),
                "orcid": _clean(row[4]),
                "note": _clean(row[5]),
            }
            seen[key] = t
            teachers.append(t)
        title = ws.title.strip()
        names = dict(DEPT_DISPLAY_NAMES, **(cfg.get("department_names") or {}))
        depts.append({"name": names.get(title, title), "teachers": teachers})
    return depts


# ----------------------------------------------------------------- http
class AuthError(RuntimeError):
    pass


class Api:
    def __init__(self, key, pause, insttoken=""):
        self.key = key
        self.pause = pause
        self.insttoken = insttoken

    def get(self, path, params=None, tries=4):
        params = dict(params or {})
        params["apiKey"] = self.key
        params["httpAccept"] = "application/json"
        url = f"{API}{path}?{urllib.parse.urlencode(params, safe='()')}"
        for attempt in range(tries):
            try:
                headers = {"User-Agent": UA, "Accept": "application/json"}
                if self.insttoken:
                    headers["X-ELS-Insttoken"] = self.insttoken
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=40) as r:
                    time.sleep(self.pause)
                    return json.loads(r.read().decode("utf-8"))
            except urllib.error.HTTPError as e:
                body = e.read().decode("utf-8", "ignore")[:300]
                if e.code == 429 or e.code >= 500:
                    wait = 3 * (attempt + 1)
                    log(f"  HTTP {e.code}, {wait} san gözləyirəm...")
                    time.sleep(wait)
                    continue
                if e.code == 404:
                    return None
                if e.code in (401, 403):
                    raise AuthError(f"HTTP {e.code}: icazə yoxdur")
                raise RuntimeError(f"HTTP {e.code}: {body}")
            except urllib.error.URLError as e:
                if attempt == tries - 1:
                    raise RuntimeError(f"Şəbəkə xətası: {e.reason}")
                time.sleep(3)
        raise RuntimeError("API cavab vermədi")


METRICS_OK = [True]


def author_metrics(api, sid):
    """h-index Author API-dən. İcazə yoxdursa (universitet şəbəkəsindən kənar) None qaytarır —
    onda h-index məqalələrin istinadlarından hesablanır."""
    out = {"h_index": None, "citations": None, "documents": None}
    if not METRICS_OK[0]:
        return out
    try:
        js = api.get(f"/author/author_id/{sid}", {"view": "METRICS"})
    except AuthError:
        log("  Author API (h-index) üçün icazə yoxdur — h-index məqalələrdən hesablanacaq")
        METRICS_OK[0] = False
        return out
    if not js:
        return out
    r = js.get("author-retrieval-response")
    if isinstance(r, list):
        r = r[0] if r else {}
    r = r or {}
    out["h_index"] = _int(r.get("h-index"))
    core = r.get("coredata") or {}
    out["citations"] = _int(core.get("citation-count"))
    out["documents"] = _int(core.get("document-count"))
    return out


COMPLETE_OK = [True]


def _search(api, sid, start):
    params = {"query": f"AU-ID({sid})", "count": 25, "start": start, "sort": "-coverDate"}
    if COMPLETE_OK[0]:
        try:
            return api.get("/search/scopus", dict(params, view="COMPLETE"))
        except RuntimeError as e:
            log(f"  COMPLETE görünüşü alınmadı, STANDARD istifadə olunur ({e})")
            COMPLETE_OK[0] = False
    return api.get("/search/scopus", params)


def _afids(v):
    if not v:
        return []
    if isinstance(v, dict):
        v = [v]
    return [str(x.get("$") if isinstance(x, dict) else x) for x in v if x]


def _affils(e):
    """Məqalənin təşkilatları: [{id, name, city, country}] (STANDARD görünüşündə də gəlir)."""
    out = []
    af = e.get("affiliation") or []
    if isinstance(af, dict):
        af = [af]
    for a in af:
        out.append({"id": str(a.get("afid") or ""), "name": a.get("affilname") or "",
                    "city": a.get("affiliation-city") or "", "country": a.get("affiliation-country") or ""})
    return out


def _authors(e):
    out = []
    au = e.get("author") or []
    if isinstance(au, dict):
        au = [au]
    for a in au:
        name = a.get("authname") or " ".join(x for x in (a.get("surname"), a.get("given-name")) if x)
        if name:
            out.append({"name": name, "id": str(a.get("authid") or ""), "afids": _afids(a.get("afid"))})
    if not out and e.get("dc:creator"):
        out.append({"name": e["dc:creator"], "id": ""})
    return out


def author_articles(api, sid):
    arts, start = [], 0
    while True:
        js = _search(api, sid, start)
        sr = (js or {}).get("search-results", {})
        total = _int(sr.get("opensearch:totalResults")) or 0
        entries = sr.get("entry") or []
        for e in entries:
            if "error" in e:
                continue
            arts.append({
                "eid": e.get("eid"),
                "title": e.get("dc:title") or "",
                "journal": e.get("prism:publicationName") or "",
                "issn": _clean_issn(e.get("prism:issn")),
                "eissn": _clean_issn(e.get("prism:eIssn")),
                "year": _int((e.get("prism:coverDate") or "")[:4]),
                "date": (e.get("prism:coverDate") or "")[:10],
                "cited": _int(e.get("citedby-count")) or 0,
                "type": e.get("subtypeDescription") or "",
                "doi": e.get("prism:doi") or "",
                "first_author": e.get("dc:creator") or "",
                "authors": _authors(e),
                "affils": _affils(e),
                "abstract": e.get("dc:description") or "",
                "keywords": e.get("authkeywords") or "",
                "volume": e.get("prism:volume") or "",
                "pages": e.get("prism:pageRange") or "",
            })
        start += 25
        if start >= total or not entries or start >= 5000:
            break
    return arts


def _int(v):
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return None


def _clean_issn(v):
    if not v:
        return ""
    return re.sub(r"[^0-9Xx]", "", str(v)).upper()


# ----------------------------------------------------------------- cache (sqlite)
def db():
    con = sqlite3.connect(os.path.join(DATA, "monitor.db"))
    con.executescript("""
    CREATE TABLE IF NOT EXISTS journal_cs (issn TEXT PRIMARY KEY, json TEXT, fetched TEXT);
    CREATE TABLE IF NOT EXISTS seen (eid TEXT PRIMARY KEY, first_seen TEXT);
    CREATE TABLE IF NOT EXISTS runs (run_at TEXT PRIMARY KEY, summary TEXT);
    CREATE TABLE IF NOT EXISTS abstracts (eid TEXT PRIMARY KEY, abstract TEXT, keywords TEXT, fetched TEXT);
    """)
    return con


# ----------------------------------------------------------------- quartiles
def pct_to_q(p):
    if p is None:
        return None
    p = float(p)
    return "Q1" if p >= 75 else "Q2" if p >= 50 else "Q3" if p >= 25 else "Q4"


def citescore_by_year(api, con, issn):
    """{il: 'Q1'..} CiteScore ən yaxşı faizilə görə. 30 gün keşlənir."""
    row = con.execute("SELECT json, fetched FROM journal_cs WHERE issn=?", (issn,)).fetchone()
    if row and (dt.datetime.now() - dt.datetime.fromisoformat(row[1])).days < 30:
        return {int(k): v for k, v in json.loads(row[0]).items()}
    res = {}
    try:
        js = api.get(f"/serial/title/issn/{issn}", {"view": "CITESCORE"})
        entries = (((js or {}).get("serial-metadata-response") or {}).get("entry")) or []
        for ent in entries:
            info = (ent.get("citeScoreYearInfoList") or {}).get("citeScoreYearInfo") or []
            if isinstance(info, dict):
                info = [info]
            for y in info:
                year = _int(y.get("@year"))
                best = None
                for cil in y.get("citeScoreInformationList") or []:
                    for ci in cil.get("citeScoreInfo") or []:
                        ranks = ci.get("citeScoreSubjectRank") or []
                        if isinstance(ranks, dict):
                            ranks = [ranks]
                        for rk in ranks:
                            p = _int(rk.get("percentile"))
                            if p is not None and (best is None or p > best):
                                best = p
                if year and best is not None:
                    res[year] = pct_to_q(best)
    except RuntimeError as e:
        log(f"  CiteScore alınmadı ({issn}): {e}")
        if isinstance(e, AuthError):
            return {}
        if row:
            return {int(k): v for k, v in json.loads(row[0]).items()}
    con.execute("INSERT OR REPLACE INTO journal_cs VALUES (?,?,?)",
                (issn, json.dumps(res), dt.datetime.now().isoformat()))
    con.commit()
    return res


class Scimago:
    """SJR kvartilləri — scimagojr.com-un illik CSV faylları.
    Əvvəlcə data/ qovluğunda fayl axtarılır (scimago_2024.csv, "scimagojr 2024.csv" və s.),
    yoxdursa saytdan yüklənməyə cəhd edilir. Sayt bloklasa (403), SJR boş qalır və
    CiteScore istifadə olunur."""

    def __init__(self):
        self.years = {}
        self.blocked = False
        self.any_loaded = False

    def _local_file(self, year):
        # data/ qovluğu, proqramın qovluğu, onun üst qovluğu və ~/Downloads
        # (brauzerdən yüklənən "scimagojr 2024.csv" faylları birbaşa tapılır)
        dirs = [DATA, BASE, os.path.dirname(BASE), os.path.expanduser("~/Downloads")]
        seen = set()
        for d in dirs:
            if d in seen or not os.path.isdir(d):
                continue
            seen.add(d)
            try:
                names = os.listdir(d)
            except OSError:
                continue
            for fn in sorted(names):
                low = fn.lower()
                if low.endswith(".csv") and "scimago" in low and str(year) in low:
                    return os.path.join(d, fn)
        return None

    def _download(self, year):
        path = os.path.join(DATA, f"scimago_{year}.csv")
        url = f"https://www.scimagojr.com/journalrank.php?year={year}&out=xls"
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/130.0 Safari/537.36",
            "Accept": "text/csv,application/vnd.ms-excel,*/*",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer": "https://www.scimagojr.com/journalrank.php"})
        with urllib.request.urlopen(req, timeout=120) as r:
            content = r.read()
        if b"Issn" not in content[:3000]:
            raise RuntimeError("gözlənilməz format")
        with open(path, "wb") as f:
            f.write(content)
        log(f"  SCImago {year} yükləndi")
        return path

    def load(self, year):
        if year in self.years:
            return self.years[year]
        path = self._local_file(year)
        if not path and not self.blocked:
            try:
                path = self._download(year)
            except Exception as e:
                self.blocked = True
                log(f"  SCImago saytı avtomatik yükləməyə icazə vermədi ({e}).")
                log("  SJR kvartilləri üçün brauzerdə bu linkləri açın (fayllar Downloads-a yüklənəcək):")
                ty = dt.date.today().year
                for y in range(ty - 1, ty - 11, -1):
                    log(f"    https://www.scimagojr.com/journalrank.php?year={y}&out=xls")
                log("  Hələlik CiteScore kvartilləri istifadə olunur.")
        if not path:
            self.years[year] = None
            return None
        mp = {}
        with open(path, encoding="utf-8", errors="ignore") as f:
            for r in csv.DictReader(f, delimiter=";"):
                q = (r.get("SJR Best Quartile") or r.get("SJR Quartile") or "").strip()
                if q not in ("Q1", "Q2", "Q3", "Q4"):
                    continue
                for issn in (r.get("Issn") or "").split(","):
                    issn = _clean_issn(issn)
                    if issn:
                        mp[issn] = q
        self.years[year] = mp
        self.any_loaded = self.any_loaded or bool(mp)
        return mp

    def quartile(self, issns, year):
        this_year = dt.date.today().year
        base = min(year, this_year - 1)
        # məqalənin ili, yoxdursa ən yaxın mövcud il (əvvəlki, sonra sonrakı illər)
        for y in [base] + [base - k for k in (1, 2, 3)] + [base + k for k in (1, 2, 3) if base + k < this_year]:
            mp = self.load(y)
            if not mp:
                continue
            for i in issns:
                if i and i in mp:
                    return mp[i]
            return None
        return None


def nearest_year(d, year):
    if not d or year is None:
        return None
    if year in d:
        return d[year]
    older = [y for y in d if y <= year]
    y = max(older) if older else min(d)
    return d[y]


# ----------------------------------------------------------------- collect
WORKERS = 4  # eyni anda neçə Scopus sorğusu (Elsevier limiti saniyədə ~9 sorğudur)


def fetch_citescore(api, issn):
    """Serial Title API-dən {il: kvartil}. None → alınmadı (keşdəki köhnə qalır)."""
    res = {}
    try:
        js = api.get(f"/serial/title/issn/{issn}", {"view": "CITESCORE"})
    except AuthError:
        return {}
    except RuntimeError as e:
        log(f"  CiteScore alınmadı ({issn}): {e}")
        return None
    entries = (((js or {}).get("serial-metadata-response") or {}).get("entry")) or []
    for ent in entries:
        info = (ent.get("citeScoreYearInfoList") or {}).get("citeScoreYearInfo") or []
        if isinstance(info, dict):
            info = [info]
        for y in info:
            year = _int(y.get("@year"))
            best = None
            for cil in y.get("citeScoreInformationList") or []:
                for ci in cil.get("citeScoreInfo") or []:
                    ranks = ci.get("citeScoreSubjectRank") or []
                    if isinstance(ranks, dict):
                        ranks = [ranks]
                    for rk in ranks:
                        pc = _int(rk.get("percentile"))
                        if pc is not None and (best is None or pc > best):
                            best = pc
            if year and best is not None:
                res[year] = pct_to_q(best)
    return res


def collect(cfg, depts):
    from concurrent.futures import ThreadPoolExecutor
    api = Api(cfg["api_key"], cfg["request_pause_sec"], cfg.get("insttoken", ""))
    con = db()
    sj = Scimago()
    teachers = [t for d in depts for t in d["teachers"]]
    for t in teachers:
        t["eids"], t["metrics"], t["error"] = [], {}, None

    # 1) h-index: əvvəlcə bir müəllimlə yoxla (icazə yoxdursa, qalanlarını soruşma)
    single = [t for t in teachers if len(t["scopus_ids"]) == 1]
    if single:
        single[0]["metrics"] = author_metrics(api, single[0]["scopus_id"])
        if METRICS_OK[0]:
            with ThreadPoolExecutor(WORKERS) as ex:
                for t, m in zip(single[1:], ex.map(lambda t: author_metrics(api, t["scopus_id"]), single[1:])):
                    t["metrics"] = m

    # 2) məqalələr — hər unikal Scopus ID üçün bir dəfə, paralel
    ids = []
    for t in teachers:
        for sid in t["scopus_ids"]:
            if sid not in ids:
                ids.append(sid)
    if ids:  # COMPLETE görünüşünə icazəni bir dəfə yoxla
        _search(api, ids[0], 0)
    done = [0]

    def fetch(sid):
        try:
            arts = author_articles(api, sid)
            err = None
        except AuthError:
            arts, err = [], ("Scopus API icazə vermədi (401). Proqramı universitetin şəbəkəsində "
                             "işə salın və ya config.json-a insttoken əlavə edin.")
        except RuntimeError as e:
            arts, err = [], str(e)
        done[0] += 1
        log(f"Müəlliflər: {done[0]}/{len(ids)}")
        return sid, arts, err

    by_id, errs = {}, {}
    with ThreadPoolExecutor(WORKERS) as ex:
        for sid, arts, err in ex.map(fetch, ids):
            by_id[sid] = arts
            if err:
                errs[sid] = err

    articles = {}
    for t in teachers:
        for sid in t["scopus_ids"]:
            if sid in errs:
                t["error"] = errs[sid]
            for a in by_id.get(sid, []):
                if not a["eid"]:
                    continue
                articles.setdefault(a["eid"], a)
                if a["eid"] not in t["eids"]:
                    t["eids"].append(a["eid"])

    # 3) jurnal kvartilləri — CiteScore 30 gün keşlənir, qalanı paralel
    log(f"{len(articles)} unikal məqalə · jurnal kvartilləri hesablanır...")
    issns = sorted({a["issn"] or a["eissn"] for a in articles.values() if a["issn"] or a["eissn"]})
    cs_cache, todo = {}, []
    for issn in issns:
        row = con.execute("SELECT json, fetched FROM journal_cs WHERE issn=?", (issn,)).fetchone()
        if row and (dt.datetime.now() - dt.datetime.fromisoformat(row[1])).days < 30:
            cs_cache[issn] = {int(k): v for k, v in json.loads(row[0]).items()}
        else:
            todo.append((issn, row))
    if todo:
        log(f"  {len(todo)} jurnal Scopus-dan soruşulur ({len(issns) - len(todo)} keşdən)")
        with ThreadPoolExecutor(WORKERS) as ex:
            for (issn, row), res in zip(todo, ex.map(lambda x: fetch_citescore(api, x[0]), todo)):
                if res is None:
                    cs_cache[issn] = {int(k): v for k, v in json.loads(row[0]).items()} if row else {}
                    continue
                cs_cache[issn] = res
                con.execute("INSERT OR REPLACE INTO journal_cs VALUES (?,?,?)",
                            (issn, json.dumps(res), dt.datetime.now().isoformat()))
        con.commit()
    for a in articles.values():
        issn = a["issn"] or a["eissn"]
        a["q_sjr"] = sj.quartile([a["issn"], a["eissn"]], a["year"] or dt.date.today().year)
        a["q_cs"] = nearest_year(cs_cache.get(issn, {}), a["year"]) if issn else None
    collect.has_sjr = sj.any_loaded
    remember_abstracts(con, articles)
    apply_sdg(articles)
    return articles, con


def remember_abstracts(con, articles):
    """Universitet şəbəkəsində alınan xülasə və açar sözləri saxla; evdən yeniləyəndə onları istifadə et."""
    now = dt.datetime.now().isoformat()
    new = [(e, a["abstract"], a.get("keywords") or "", now) for e, a in articles.items() if a.get("abstract")]
    if new:
        con.executemany("INSERT OR REPLACE INTO abstracts VALUES (?,?,?,?)", new)
        con.commit()
    missing = [e for e, a in articles.items() if not a.get("abstract")]
    filled = 0
    for i in range(0, len(missing), 500):
        part = missing[i:i + 500]
        q = "SELECT eid, abstract, keywords FROM abstracts WHERE eid IN (%s)" % ",".join("?" * len(part))
        for eid, ab, kw in con.execute(q, part):
            articles[eid]["abstract"] = ab
            if not articles[eid].get("keywords"):
                articles[eid]["keywords"] = kw
            filled += 1
    if filled:
        log(f"  {filled} məqalənin xülasəsi əvvəlki (universitet) yeniləməsindən götürüldü")


def sdg_query_file():
    here = os.path.dirname(os.path.abspath(__file__))
    for d in (globals().get("RES"), here, BASE):
        if d and os.path.exists(os.path.join(d, "sdg_queries_2023.json.gz")):
            return os.path.join(d, "sdg_queries_2023.json.gz")
    return None


def apply_sdg(articles):
    """Elsevier 2023 SDG sorğularını başlıq, xülasə və açar sözlərə tətbiq edir."""
    path = sdg_query_file()
    apply_sdg.basis = None
    if not path:
        log("  SDG sorğu faylı tapılmadı — SDG hesablanmır")
        return
    try:
        import sdg
        res = sdg.classify(articles, path)
        for eid, nums in res.items():
            articles[eid]["sdg"] = nums
        with_abs = sum(1 for a in articles.values() if a.get("abstract"))
        apply_sdg.basis = "full" if with_abs > len(articles) * 0.5 else "title"
        log(f"  SDG: {sum(1 for v in res.values() if v)} məqalə ({'başlıq+xülasə+açar söz' if apply_sdg.basis == 'full' else 'yalnız başlıq'})")
    except Exception as e:
        log(f"  SDG hesablanmadı: {e}")


def demo_collect(depts):
    rnd = random.Random(7)
    journals = [("Journal of Molecular Liquids", "01677322"), ("Fuel", "00162361"),
                ("Chemical Problems", "22214666"), ("Voprosy Khimii i Khimicheskoi Tekhnologii", "03214095"),
                ("Processes", "22279717"), ("SOCAR Proceedings", "22186867"),
                ("Petroleum Chemistry", "09655441"), ("Azerbaijan Chemical Journal", "00052531")]
    articles, pool = {}, []
    this = dt.date.today().year
    n = 0
    for d in depts:
        for t in d["teachers"]:
            t["eids"], t["metrics"], t["error"] = [], {}, None
            if not t["scopus_id"]:
                continue
            k = rnd.randint(4, 30)
            for _ in range(k):
                if pool and rnd.random() < 0.15:
                    eid = rnd.choice(pool)
                else:
                    n += 1
                    eid = f"2-s2.0-DEMO{n:05d}"
                    j = rnd.choice(journals)
                    q = rnd.choice(["Q1", "Q2", "Q2", "Q3", "Q4", None])
                    articles[eid] = {"eid": eid, "title": f"Nümunə məqalə #{n}", "journal": j[0],
                                     "issn": j[1], "eissn": "", "year": (yy := rnd.randint(this - 10, this)), "date": f"{yy}-{rnd.randint(1, 12 if yy < this else 11):02d}-{rnd.randint(1, 28):02d}",
                                     "cited": rnd.randint(0, 40), "type": "Article", "doi": f"10.0000/demo.{n}",
                                     "first_author": t["name"].split()[0],
                                     "authors": [{"name": t["name"], "id": t["scopus_id"]}] + [{"name": rnd.choice(["Aliyev R.", "Mammadova S.", "Huseynov T.", "Ivanov A.", "Smith J."]), "id": ""} for _ in range(rnd.randint(1, 4))],
                                     "volume": str(rnd.randint(1, 300)), "pages": "",
                                     "affils": [{"id": "1", "name": "Azerbaijan State Oil and Industry University", "city": "Baku", "country": "Azerbaijan"}]
                                               + ([{"id": "2", "name": rnd.choice(["Kazan Federal University", "Istanbul Technical University", "University of Tehran", "Tsinghua University", "Gubkin University"]),
                                                    "city": "", "country": rnd.choice(["Russian Federation", "Turkey", "Iran", "China", "Russian Federation"])}] if rnd.random() < 0.35 else []),
                                     "q_sjr": q, "q_cs": q if rnd.random() < .7 else rnd.choice(["Q1", "Q2", "Q3", None])}
                    pool.append(eid)
                if eid not in t["eids"]:
                    t["eids"].append(eid)
            t["metrics"] = {"h_index": rnd.randint(1, 15), "citations": None, "documents": k}
    for a in articles.values():
        a["sdg"] = sorted(rnd.sample([6, 7, 9, 12, 13, 3, 14], rnd.choice([0, 0, 0, 1, 1, 2])))
    apply_sdg.basis = "full"
    return articles, db_demo()


def db_demo():
    con = sqlite3.connect(":memory:")
    con.executescript("""CREATE TABLE seen (eid TEXT PRIMARY KEY, first_seen TEXT);
                         CREATE TABLE runs (run_at TEXT PRIMARY KEY, summary TEXT);""")
    return con


# ----------------------------------------------------------------- aggregate
def qcounts(arts, key):
    c = {"Q1": 0, "Q2": 0, "Q3": 0, "Q4": 0, "none": 0}
    for a in arts:
        c[a.get(key) or "none"] += 1
    return c


def h_index(cites):
    cites = sorted(cites, reverse=True)
    return sum(1 for i, c in enumerate(cites, 1) if c >= i)


HOME_COUNTRY = "Azerbaijan"


def intl_info(a):
    """Xarici həmmüəlliflik: xarici təşkilatlar, ölkələr və (məlumdursa) xarici müəlliflər."""
    affs = a.get("affils") or []
    foreign = [x for x in affs if x.get("country") and x["country"].strip().lower() != HOME_COUNTRY.lower()]
    fids = {x["id"] for x in foreign if x.get("id")}
    # müəllifin təşkilatlarından biri xaricidirsə — xarici müəllif (yalnız tam müəllif siyahısı gələndə)
    fau = [au["name"] for au in a.get("authors", []) if fids & set(au.get("afids") or [])]
    return {
        "intl": bool(foreign),
        "countries": sorted({x["country"] for x in foreign}),
        "forg": [[x["name"], x["country"]] for x in foreign],
        "fau": fau,
    }


def summarize(arts, new_set, this_year):
    return {
        "articles": len(arts),
        "citations": sum(a["cited"] for a in arts),
        "this_year": sum(1 for a in arts if a["year"] == this_year),
        "last_year": sum(1 for a in arts if a["year"] == this_year - 1),
        "sjr": qcounts(arts, "q_sjr"),
        "cs": qcounts(arts, "q_cs"),
        "new": sum(1 for a in arts if a["eid"] in new_set),
    }


def build(cfg, depts, articles, con, demo=False):
    now = dt.datetime.now()
    this_year = now.year
    # yeni məqalələr (ilk işə salınmada heç nə "yeni" sayılmır)
    # uğursuz (0 məqaləli) yoxlamalar müqayisədə nəzərə alınmır
    good_runs = [json.loads(r[0]) for r in con.execute("SELECT summary FROM runs ORDER BY run_at DESC")]
    good_runs = [r for r in good_runs if r.get("articles")]
    first_run = not good_runs
    seen = {r[0] for r in con.execute("SELECT eid FROM seen")}
    new_set = set() if first_run else (set(articles) - seen)
    con.executemany("INSERT OR IGNORE INTO seen VALUES (?,?)",
                    [(e, now.isoformat()) for e in articles])
    prev = good_runs[0] if good_runs else None

    years = list(range(this_year - cfg["years_in_chart"] + 1, this_year + 1))

    def by_year(arts):
        return [sum(1 for a in arts if a["year"] == y) for y in years]

    # eyni Scopus ID fərqli adlı müəllimlərdə yazılıbsa — xəbərdarlıq (Excel-də səhv ola bilər)
    owners = {}
    for d in depts:
        for t in d["teachers"]:
            for sid in t.get("scopus_ids") or []:
                owners.setdefault(sid, set()).add(t["name"])
    out_depts = []
    all_teachers = 0
    with_id = 0
    uniq = set()  # bir müəllim iki kafedrada yazılıbsa, fakültədə bir dəfə sayılır
    for d in depts:
        eids = []
        teachers = []
        for t in d["teachers"]:
            key = t["name"].lower()
            if key not in uniq:
                uniq.add(key)
                all_teachers += 1
                if t["scopus_id"]:
                    with_id += 1
            ta = [articles[e] for e in t["eids"] if e in articles]
            s = summarize(ta, new_set, this_year)
            teachers.append({
                "name": t["name"], "position": t["position"], "scopus_id": t["scopus_id"],
                "scopus_ids": t.get("scopus_ids") or ([t["scopus_id"]] if t["scopus_id"] else []),
                "raw_id": t["raw_id"], "orcid": t["orcid"], "error": t.get("error"),
                "h_index": (t.get("metrics") or {}).get("h_index") if (t.get("metrics") or {}).get("h_index") is not None
                           else (h_index([a["cited"] for a in ta]) if ta else None),
                "eids": t["eids"], **s,
                "warning": "; ".join(f"Scopus ID {sid} başqa müəllimdə də yazılıb: " + ", ".join(sorted(owners[sid] - {t['name']}))
                                     for sid in (t.get("scopus_ids") or []) if len(owners.get(sid, ())) > 1) or None,
            })
            for e in t["eids"]:
                if e not in eids:
                    eids.append(e)
        da = [articles[e] for e in eids]
        out_depts.append({"name": d["name"], "teachers": teachers, "eids": eids,
                          "by_year": by_year(da), **summarize(da, new_set, this_year)})

    fa = list(articles.values())
    fac = summarize(fa, new_set, this_year)
    fac["by_year"] = by_year(fa)
    fac["teachers"] = all_teachers
    fac["with_id"] = with_id
    fac["delta_articles"] = (fac["articles"] - prev["articles"]) if prev else None
    fac["delta_citations"] = (fac["citations"] - prev["citations"]) if prev else None

    if fac["articles"]:
        con.execute("INSERT OR REPLACE INTO runs VALUES (?,?)",
                    (now.isoformat(), json.dumps({"articles": fac["articles"], "citations": fac["citations"]})))
    con.commit()

    fac_authors = {}
    for d in out_depts:
        for t in d["teachers"]:
            for e in t["eids"]:
                if t["name"] not in fac_authors.setdefault(e, []):
                    fac_authors[e].append(t["name"])

    payload = {
        "updated": now.strftime("%Y-%m-%d %H:%M"),
        "demo": demo,
        "has_sjr": True if demo else getattr(collect, "has_sjr", True),
        "sdg_basis": getattr(apply_sdg, "basis", None),
        "faculty_short": cfg["faculty_short"],
        "faculty_name": cfg["faculty_name"],
        "university": cfg.get("university_name", "Azərbaycan Dövlət Neft və Sənaye Universiteti"),
        "years": years,
        "this_year": this_year,
        "faculty": fac,
        "departments": out_depts,
        "articles": {e: dict({k: a.get(k) for k in ("title", "journal", "year", "date", "cited", "q_sjr", "q_cs",
                                                     "doi", "type", "volume", "pages")},
                             authors=[x["name"] for x in a.get("authors", [])],
                             **intl_info(a),
                             sdg=a.get("sdg"),
                             ours=fac_authors.get(e, []),
                             new=e in new_set) for e, a in articles.items()},
    }
    # günlük snapshot (təqvim üçün): data/snapshots/YYYY-MM-DD.json — günün son yeniləməsi
    snap_dir = os.path.join(DATA, "snapshots" if not demo else "snapshots_demo")
    os.makedirs(snap_dir, exist_ok=True)
    if fac["articles"]:
        with open(os.path.join(snap_dir, now.strftime("%Y-%m-%d") + ".json"), "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
    payload["snapshots"] = list_snapshots(demo)
    out_path = write_html(cfg, payload, demo)
    # Excel/CSV ixracı
    export_csv(out_depts, articles)
    export_articles(articles, fac_authors)
    return out_path, fac


def write_html(cfg, payload, demo=False):
    with open(os.path.join(RES, "dashboard_template.html"), encoding="utf-8") as f:
        tpl = f.read()
    if "__LOGO_B64__" in tpl:  # loqo ayrıca faylda (logo.png) — şablona yerləşdir
        import base64
        lp = os.path.join(RES, "logo.png")
        b64 = base64.b64encode(open(lp, "rb").read()).decode() if os.path.exists(lp) else ""
        tpl = tpl.replace("__LOGO_B64__", b64)
    html_out = tpl.replace("/*__DATA__*/null", json.dumps(payload, ensure_ascii=False).replace("</", "<\\/"))
    out = cfg["output_html"] if not demo else "dashboard_demo.html"
    out_path = out if os.path.isabs(out) else os.path.join(BASE, out)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_out)
    return out_path


OLD_DEPT_NAMES = {
    "Proses təhlükəsizliyi müh.": "Proseslərin Təhlükəsizliyi Mühəndisliyi",
    "Proseslərin Təhlükəsizliyi": "Proseslərin Təhlükəsizliyi Mühəndisliyi",
}


def load_snapshot(fp):
    with open(fp, encoding="utf-8") as f:
        payload = json.load(f)
    names = dict(OLD_DEPT_NAMES, **DEPT_DISPLAY_NAMES)
    for d in payload.get("departments") or []:
        d["name"] = names.get(d.get("name"), d.get("name"))
    txt = json.dumps(payload, ensure_ascii=False)
    for old, new in names.items():
        txt = txt.replace(json.dumps(old, ensure_ascii=False), json.dumps(new, ensure_ascii=False))
    return json.loads(txt)


def rerender_latest(cfg, demo=False):
    """Scopus-a müraciət etmədən son saxlanmış günün məlumatını cari şablonla yenidən qurur."""
    snaps = list_snapshots(demo)
    if not snaps:
        return False
    d = os.path.join(DATA, "snapshots" if not demo else "snapshots_demo")
    payload = load_snapshot(os.path.join(d, snaps[-1] + ".json"))
    payload["snapshots"] = snaps
    write_html(cfg, payload, demo)
    return True


def list_snapshots(demo=False):
    d = os.path.join(DATA, "snapshots" if not demo else "snapshots_demo")
    if not os.path.isdir(d):
        return []
    return sorted(f[:-5] for f in os.listdir(d) if re.fullmatch(r"\d{4}-\d{2}-\d{2}\.json", f))


def export_csv(depts, articles):
    path = os.path.join(BASE, "muellim_statistikasi.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Kafedra", "Müəllim", "Vəzifə", "Scopus ID", "h-index", "Məqalə", "İstinad",
                    "Q1 (SJR)", "Q2 (SJR)", "Q3 (SJR)", "Q4 (SJR)", "Kvartilsiz", "Bu il", "Yeni", "Qeyd"])
        for d in depts:
            for t in d["teachers"]:
                w.writerow([d["name"], t["name"], t["position"], ", ".join(t.get("scopus_ids") or []), t["h_index"] or "",
                            t["articles"], t["citations"], t["sjr"]["Q1"], t["sjr"]["Q2"], t["sjr"]["Q3"],
                            t["sjr"]["Q4"], t["sjr"]["none"], t["this_year"], t["new"],
                            t["error"] or ("Scopus ID yoxdur" if not t["scopus_id"] else "")])


def export_articles(articles, fac_authors):
    path = os.path.join(BASE, "meqaleler.csv")
    rows = sorted(articles.values(), key=lambda a: (-(a["year"] or 0), a["title"]))
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Başlıq", "Müəlliflər", "Fakültə müəllimləri", "Jurnal", "İl", "Cild", "Səhifələr",
                    "DOI", "Kvartil (SJR)", "Kvartil (CiteScore)", "İstinad", "Növ", "Scopus EID", "SDG"])
        for a in rows:
            w.writerow([a["title"], "; ".join(x["name"] for x in a.get("authors", [])),
                        "; ".join(fac_authors.get(a["eid"], [])), a["journal"], a["year"] or "",
                        a.get("volume", ""), a.get("pages", ""), a["doi"], a.get("q_sjr") or "",
                        a.get("q_cs") or "", a["cited"], a["type"], a["eid"],
                        ", ".join(f"SDG {n}" for n in a.get("sdg") or [])])


# ----------------------------------------------------------------- main
def run_once(args):
    cfg = load_config()
    depts = read_teachers(cfg)
    n = sum(len(d["teachers"]) for d in depts)
    log(f"Exceldən {len(depts)} kafedra, {n} müəllim oxundu")
    if args.demo:
        articles, con = demo_collect(depts)
    else:
        articles, con = collect(cfg, depts)
    out, fac = build(cfg, depts, articles, con, demo=args.demo)
    log(f"Hazırdır: {out}")
    log(f"Məqalə: {fac['articles']} · İstinad: {fac['citations']} · bu il: {fac['this_year']}")
    return out


LOADING_HTML = """<!DOCTYPE html><html lang="az"><head><meta charset="utf-8"><title>Yüklənir…</title>
<style>body{font:15px system-ui,sans-serif;background:#f6f5f2;color:#1d1d1b;display:flex;align-items:center;
justify-content:center;height:90vh}div{text-align:center}small{color:#6f6d68}</style></head>
<body><div><h2>Scopus-dan məlumat yüklənir…</h2><p>Ən son rəqəmlər hazırlanır. Hazır olanda səhifə özü açılacaq.</p><p><small>Loading data from Scopus… · Загрузка данных из Scopus…</small></p><p id="s"><small>Bir neçə dəqiqə çəkə bilər.</small></p></div>
<script>setInterval(async()=>{try{const s=await (await fetch('/api/status',{cache:'no-store'})).json();
document.getElementById('s').innerHTML='<small>'+(s.last||'')+'</small>';if(!s.running)location.reload();}catch(e){}},1500);</script>
</body></html>"""


def free_port(port):
    """Bu portda işləyən köhnə nüsxəni dayandırır (Mac/Linux: lsof, Windows: netstat + taskkill)."""
    import subprocess
    pids = set()
    try:
        if os.name == "nt":
            out = subprocess.run(["netstat", "-ano"], capture_output=True, text=True,
                                 creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
            for line in out.splitlines():
                parts = line.split()
                if len(parts) >= 5 and parts[1].endswith(f":{port}") and parts[3].upper() == "LISTENING":
                    pids.add(parts[4])
        else:
            out = subprocess.run(["lsof", "-ti", f":{port}"], capture_output=True, text=True).stdout
            pids.update(x for x in out.split() if x.strip())
    except Exception:
        return False
    pids.discard(str(os.getpid()))
    for pid in pids:
        log(f"Köhnə nüsxə dayandırılır (PID {pid})...")
        try:
            if os.name == "nt":
                subprocess.run(["taskkill", "/PID", pid, "/F"], capture_output=True,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            else:
                os.kill(int(pid), 15)
        except Exception:
            pass
    if pids:
        time.sleep(1.5)
    return bool(pids)


def serve(args, port, background=False):
    """Lokal server: dashboard + 'Yenilə' düyməsi üçün /api/refresh və /api/status."""
    import threading
    from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

    lock = threading.Lock()

    def refresh():
        if not lock.acquire(blocking=False):
            return False
        STATUS.update(running=True, error=None, last="Başlayır...")

        def work():
            try:
                run_once(args)
                STATUS["finished"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
            except Exception as e:
                STATUS["error"] = str(e)
                log(f"XƏTA: {e}")
            finally:
                STATUS["running"] = False
                lock.release()
        threading.Thread(target=work, daemon=True).start()
        return True

    cfg = load_config()
    page = os.path.basename(cfg["output_html"]) if not args.demo else "dashboard_demo.html"

    class H(SimpleHTTPRequestHandler):
        def __init__(self, *a, **k):
            super().__init__(*a, directory=BASE, **k)

        def log_message(self, *a):
            pass

        def _json(self, obj, code=200):
            body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            if self.path.startswith("/api/status"):
                return self._json(STATUS)
            if self.path.startswith("/api/snapshot"):
                q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
                day = (q.get("date") or [""])[0]
                if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", day):
                    return self._json({"error": "tarix səhvdir"}, 400)
                fp = os.path.join(DATA, "snapshots" if not args.demo else "snapshots_demo", day + ".json")
                if not os.path.exists(fp):
                    return self._json({"error": "bu gün üçün məlumat yoxdur"}, 404)
                snap = load_snapshot(fp)
                snap["snapshots"] = list_snapshots(args.demo)
                return self._json(snap)
            if self.path.split("?")[0] in ("/", "/index.html"):
                self.path = "/" + page
            # yalnız dashboard fayllarını göstər (config.json-dakı API açarı açılmasın)
            if not (self.path.split("?")[0].endswith(".html") or self.path.split("?")[0].endswith(".csv")):
                return self._json({"error": "not found"}, 404)
            return super().do_GET()

        def end_headers(self):
            self.send_header("Cache-Control", "no-store")
            super().end_headers()

        def do_POST(self):
            if self.path.startswith("/api/refresh"):
                started = refresh()
                return self._json({"started": started, "running": True})
            return self._json({"error": "not found"}, 404)

    httpd = None
    try:
        httpd = ThreadingHTTPServer(("127.0.0.1", port), H)
    except OSError:
        if background:
            # tətbiq artıq açıqdırsa — yeni nüsxə köhnəni öldürmür, sadəcə ona qoşulur
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/status", timeout=2) as r:
                    if "running" in json.loads(r.read().decode("utf-8")):
                        return f"http://127.0.0.1:{port}/"
            except Exception:
                pass
        # port tutulub — köhnə nüsxəni dayandır və yenidən cəhd et, alınmasa növbəti boş port
        free_port(port)
        for p in range(port, port + 20):
            try:
                httpd = ThreadingHTTPServer(("127.0.0.1", p), H)
                port = p
                break
            except OSError:
                continue
    if httpd is None:
        log("Boş port tapılmadı.")
        return None
    # səhifəni həmişə son şablonla qur; saxlanmış məlumat yoxdursa Scopus-dan çək
    # Açılış:
    #  - Excel son yeniləmədən sonra dəyişməyibsə → son məlumat dərhal göstərilir, yeniləmə arxada gedir
    #    (səhifədə "yenilənir…" zolağı, bitəndə səhifə özü yeni rəqəmlərə keçir)
    #  - Excel dəyişibsə və ya məlumat yoxdursa → "yüklənir" səhifəsi, köhnə rəqəmlər göstərilmir
    #  - --no-refresh → yalnız son saxlanmış məlumat
    xl = cfg["excel_file"] if os.path.isabs(cfg["excel_file"]) else os.path.join(BASE, cfg["excel_file"])
    snaps = list_snapshots(args.demo)
    snap_dir = os.path.join(DATA, "snapshots" if not args.demo else "snapshots_demo")
    fresh = bool(snaps) and not (os.path.exists(xl) and
                                 os.path.getmtime(xl) > os.path.getmtime(os.path.join(snap_dir, snaps[-1] + ".json")))
    if fresh and rerender_latest(cfg, args.demo):
        if args.no_refresh:
            log("Son saxlanmış məlumat göstərilir (--no-refresh)")
        else:
            log("Son məlumat göstərilir, Scopus-dan yeniləmə arxada gedir...")
            refresh()
    else:
        if snaps:
            log("Excel dəyişib — yeni müəllim siyahısı ilə yenilənir...")
        with open(os.path.join(BASE, page), "w", encoding="utf-8") as f:
            f.write(LOADING_HTML)
        log("Scopus-dan məlumat çəkilir — hazır olanda dashboard özü açılacaq...")
        refresh()
    url = f"http://127.0.0.1:{port}/"
    log(f"Dashboard: {url}  (bağlamaq üçün bu pəncərədə Ctrl+C)")
    if not args.no_open:
        webbrowser.open(url)
    if args.loop:
        def scheduler():
            while True:
                time.sleep(args.loop * 3600)
                log("Planlı avtomatik yeniləmə")
                refresh()
        threading.Thread(target=scheduler, daemon=True).start()
    if background:
        threading.Thread(target=httpd.serve_forever, daemon=True).start()
        return url
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        log("Server dayandırıldı")


def main():
    p = argparse.ArgumentParser(description="KTF Scopus monitorinqi")
    p.add_argument("--serve", action="store_true", help="dashboard-u 'Yenilə' düyməsi ilə lokal serverdə aç")
    p.add_argument("--port", type=int, default=8765)
    p.add_argument("--no-refresh", action="store_true", help="açılışda Scopus-dan yeniləmə, son saxlanmış məlumatı göstər")
    p.add_argument("--loop", type=float, metavar="SAAT", help="hər N saatdan bir avtomatik yenilə")
    p.add_argument("--demo", action="store_true", help="nümunə məlumatla (internetsiz)")
    p.add_argument("--no-open", action="store_true", help="brauzerdə açma")
    if len(sys.argv) == 1:
        # arqumentsiz işə salınma (python3 scopus_monitor.py): dashboard + hər 24 saatdan bir yeniləmə
        sys.argv += ["--serve", "--loop", "24"]
    args = p.parse_args()
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    if args.serve:
        return serve(args, args.port)
    opened = False
    while True:
        try:
            out = run_once(args)
            if not args.no_open and not opened:
                import pathlib
                webbrowser.open(pathlib.Path(out).resolve().as_uri())
                opened = True
        except Exception as e:
            log(f"XƏTA: {e}")
        if not args.loop:
            break
        log(f"Növbəti yeniləmə {args.loop} saatdan sonra (dayandırmaq üçün Ctrl+C)")
        time.sleep(args.loop * 3600)


if __name__ == "__main__":
    main()
