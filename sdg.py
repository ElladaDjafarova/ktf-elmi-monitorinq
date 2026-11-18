# -*- coding: utf-8 -*-
"""
Elsevier 2023 SDG sorğularını (Scopus sorğu dili) məqalələrə lokal tətbiq edir.

Mənbə: Bedard-Vallee, James, Roberge (2023) "Elsevier 2023 Sustainable Development Goals (SDGs) Mapping",
Mendeley Data, doi:10.17632/y2zyy9vwzy.1 — sdg_queries_2023.json.gz faylında.

Dəstəklənir: TITLE-ABS-KEY, TITLE-ABS, TITLE, ABS, AUTHKEY, KEY, SRCTITLE, SUBJAREA (məlumat yoxdur → boş),
AND / OR / AND NOT, W/n, PRE/n, "ifadə", {dəqiq}, * və ? simvolları, mötərizələr.
Qeyd: Scopus-un maşın öyrənməsi hissəsi açıq deyil — nəticə yalnız sorğu hissəsinə uyğundur.
"""
import bisect
import gzip
import json
import os
import re

WORD_RE = re.compile(r"[0-9a-zà-ÿāăąćčďđēėęěğīįıķļľłńňōőœŕřśşšţťūůűųźżžəß]+", re.I)


def tokenize(text):
    return [w.lower() for w in WORD_RE.findall(text or "")]


# ---------------------------------------------------------------- sorğu sintaksisi
FIELDS = {"TITLE-ABS-KEY", "TITLE-ABS", "TITLE", "ABS", "AUTHKEY", "KEY", "SRCTITLE", "SUBJAREA",
          "INDEXTERMS", "ALL"}
TOK_RE = re.compile(r'\s*(?:(?P<lp>\()|(?P<rp>\))|(?P<q>"[^"]*")|(?P<ex>\{[^}]*\})'
                    r'|(?P<prox>(?:W|PRE)/\d+)(?=[\s("{]|$)|(?P<word>[^\s()"{}]+))', re.I)


def lex(q):
    out, pos = [], 0
    while pos < len(q):
        m = TOK_RE.match(q, pos)
        if not m or m.end() == pos:
            if q[pos:].strip() == "":
                break
            raise ValueError(f"lex xətası: {q[pos:pos + 40]!r}")
        pos = m.end()
        k = m.lastgroup
        v = m.group(k)
        if k == "word":
            u = v.upper()
            if u in ("AND", "OR", "NOT"):
                out.append(("op", u))
                continue
            if u in FIELDS and re.match(r"\s*\(", q[pos:]):
                out.append(("field", u))
                continue
        out.append((k, v))
    # "AND NOT" birləşdir
    res = []
    for t in out:
        if t == ("op", "NOT") and res and res[-1] == ("op", "AND"):
            res[-1] = ("op", "ANDNOT")
        else:
            res.append(t)
    return res


class Parser:
    """Scopus prioriteti (yuxarıdan aşağı): OR, W/n və PRE/n, AND, AND NOT."""

    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def parse(self):
        n = self.andnot()
        while self.i < len(self.t):  # artıq bağlanan mötərizələr və s.
            if self.peek()[0] == "rp":
                self.i += 1
                continue
            n = ("and", n, self.andnot())
        return n

    def starts_primary(self):
        k = self.peek()[0]
        return k in ("lp", "q", "ex", "word", "field") or (k == "op" and self.peek()[1] == "NOT")

    def andnot(self):
        n = self.and_()
        while self.peek() == ("op", "ANDNOT"):
            self.take()
            n = ("not", n, self.and_())
        return n

    def and_(self):
        n = self.prox()
        while True:
            if self.peek() == ("op", "AND"):
                self.take()
                n = ("and", n, self.prox())
            elif self.starts_primary() and self.peek() != ("op", "NOT"):
                n = ("and", n, self.prox())  # ardıcıl sözlər — gizli AND
            else:
                return n

    def prox(self):
        # Scopus prioriteti: OR → W/n, PRE/n → AND → AND NOT
        n = self.or_()
        while self.peek()[0] == "prox":
            kind, num = self.take()[1].upper().split("/")
            n = ("prox", kind, int(num), n, self.or_())
        return n

    def or_(self):
        n = self.primary()
        while self.peek() == ("op", "OR"):
            self.take()
            n = ("or", n, self.primary())
        return n

    def primary(self):
        k, v = self.take()
        if k == "field":
            if self.peek()[0] != "lp":
                raise ValueError("sahədən sonra ( gözlənilirdi")
            self.take()
            inner = self.andnot()
            if self.peek()[0] == "rp":
                self.take()
            return ("field", v, inner)
        if k == "lp":
            inner = self.andnot()
            if self.peek()[0] == "rp":
                self.take()
            return inner
        if k == "q":
            return ("term", v[1:-1], False)
        if k == "ex":
            return ("term", v[1:-1], True)
        if k == "word":
            return ("term", v, False)
        if k == "op" and v == "NOT":  # tək NOT — nadirdir, "AND NOT" kimi qəbul et
            return ("neg", self.primary())
        raise ValueError(f"gözlənilməz: {k} {v}")


# ---------------------------------------------------------------- sənəd indeksi
# sahə → sənədin hansı hissələrində axtarılır
FIELD_PARTS = {
    "TITLE-ABS-KEY": ("title", "abs", "key"), "ALL": ("title", "abs", "key", "src"),
    "TITLE-ABS": ("title", "abs"), "TITLE": ("title",), "ABS": ("abs",),
    "AUTHKEY": ("key",), "KEY": ("key",), "INDEXTERMS": ("key",), "SRCTITLE": ("src",), "SUBJAREA": (),
}


class Corpus:
    def __init__(self, docs):
        """docs: {doc_id: {"title":..., "abs":..., "key": [..] or str, "src":...}}"""
        self.ids = list(docs)
        self.all = set(self.ids)
        # inverted index: part → token → [(doc, seg, pos)]
        self.inv = {p: {} for p in ("title", "abs", "key", "src")}
        for d, doc in docs.items():
            for p in self.inv:
                val = doc.get(p) or ""
                segs = val if isinstance(val, list) else [val]
                for si, seg in enumerate(segs):
                    for pos, w in enumerate(tokenize(seg)):
                        self.inv[p].setdefault(w, []).append((d, si, pos))
        self.vocab = {p: sorted(self.inv[p]) for p in self.inv}
        self._pat = {}

    def words(self, part, pat):
        """Sözə uyğun gələn lüğət sözləri (* və ? ilə, sadə cəm formaları ilə)."""
        key = (part, pat)
        if key in self._pat:
            return self._pat[key]
        vocab, inv = self.vocab[part], self.inv[part]
        if "*" not in pat and "?" not in pat:
            cands = {pat, pat + "s", pat + "es"}
            if pat.endswith("y"):
                cands.add(pat[:-1] + "ies")
            res = [w for w in cands if w in inv]
        else:
            prefix = re.split(r"[*?]", pat, 1)[0]
            rx = re.compile("^" + re.escape(pat).replace(r"\*", ".*").replace(r"\?", ".") + "$")
            lo = bisect.bisect_left(vocab, prefix)
            res = []
            for w in vocab[lo:]:
                if not w.startswith(prefix):
                    break
                if rx.match(w):
                    res.append(w)
        self._pat[key] = res
        return res

    def phrase_spans(self, field, text):
        """İfadənin tapıldığı yerlər: {doc: [(part, seg, start, end)]}"""
        words = [w.lower() for w in re.findall(r"[^\s\-_/.,;:()'’]+", text)]
        words = [w for w in (re.sub(r"[^\w*?ə]", "", w, flags=re.U) for w in words) if w]
        out = {}
        if not words:
            return out
        for part in FIELD_PARTS.get(field, ()):
            inv = self.inv[part]
            sets = [self.words(part, w) for w in words]
            if any(not s for s in sets):
                continue
            # birinci sözün mövqelərindən başla
            nxt = [{(d, si, pos) for ww in s for (d, si, pos) in inv[ww]} for s in sets[1:]]
            for ww in sets[0]:
                for (d, si, pos) in inv[ww]:
                    if all((d, si, pos + k + 1) in nxt[k] for k in range(len(nxt))):
                        out.setdefault(d, []).append((part, si, pos, pos + len(words) - 1))
        return out


# ---------------------------------------------------------------- qiymətləndirmə
def spans(node, c, field):
    """Məsafə operatorları üçün mövqelər."""
    t = node[0]
    if t == "term":
        return c.phrase_spans(field, node[1])
    if t == "field":
        return spans(node[2], c, node[1])
    if t == "or":
        a, b = spans(node[1], c, field), spans(node[2], c, field)
        for d, v in b.items():
            a.setdefault(d, []).extend(v)
        return a
    if t == "and":
        a, b = spans(node[1], c, field), spans(node[2], c, field)
        return {d: a[d] + b[d] for d in a.keys() & b.keys()}
    if t == "prox":
        return prox(node, c, field)
    if t == "not":
        a = spans(node[1], c, field)
        bad = evaluate(node[2], c, field)
        return {d: v for d, v in a.items() if d not in bad}
    return {}


def prox(node, c, field):
    _, kind, n, l, r = node
    a, b = spans(l, c, field), spans(r, c, field)
    out = {}
    for d in a.keys() & b.keys():
        for (p1, s1, x1, y1) in a[d]:
            for (p2, s2, x2, y2) in b[d]:
                if (p1, s1) != (p2, s2):
                    continue
                if kind == "PRE":
                    ok = x2 > y1 and x2 - y1 - 1 <= n
                else:
                    gap = x2 - y1 - 1 if x2 > y1 else x1 - y2 - 1
                    ok = gap <= n
                if ok:
                    out.setdefault(d, []).append((p1, s1, min(x1, x2), max(y1, y2)))
    return out


def evaluate(node, c, field="TITLE-ABS-KEY"):
    t = node[0]
    if t == "term":
        return set(c.phrase_spans(field, node[1]))
    if t == "field":
        return evaluate(node[2], c, node[1])
    if t == "or":
        return evaluate(node[1], c, field) | evaluate(node[2], c, field)
    if t == "and":
        a = evaluate(node[1], c, field)
        return a & evaluate(node[2], c, field) if a else set()
    if t == "not":
        a = evaluate(node[1], c, field)
        return a - evaluate(node[2], c, field) if a else set()
    if t == "neg":
        return c.all - evaluate(node[1], c, field)
    if t == "prox":
        return set(prox(node, c, field))
    return set()


# ---------------------------------------------------------------- xarici interfeys
_PARSED = None


def load_queries(path):
    global _PARSED
    if _PARSED is None:
        with gzip.open(path, "rt", encoding="utf-8") as f:
            raw = json.load(f)
        _PARSED = {}
        import sys
        lim = sys.getrecursionlimit()
        sys.setrecursionlimit(max(lim, 20000))
        for k, q in raw.items():
            try:
                _PARSED[int(k)] = Parser(lex(q)).parse()
            except Exception as e:  # bir SDG oxunmasa, qalanları işləsin
                print(f"SDG {k} sorğusu oxunmadı: {e}", flush=True)
    return _PARSED


def classify(articles, query_path):
    """articles: {eid: {title, abstract, keywords, journal}} → {eid: [sdg nömrələri]}"""
    import sys
    sys.setrecursionlimit(max(sys.getrecursionlimit(), 20000))
    docs = {}
    for eid, a in articles.items():
        kw = a.get("keywords") or ""
        if isinstance(kw, str):
            kw = [k.strip() for k in kw.split("|") if k.strip()]
        docs[eid] = {"title": a.get("title") or "", "abs": a.get("abstract") or "",
                     "key": kw, "src": a.get("journal") or ""}
    c = Corpus(docs)
    res = {eid: [] for eid in docs}
    for n, tree in sorted(load_queries(query_path).items()):
        for eid in evaluate(tree, c):
            res[eid].append(n)
    return res
