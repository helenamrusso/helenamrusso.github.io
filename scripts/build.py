#!/usr/bin/env python3
"""Build the website into _site/.

    python scripts/build.py            # build
    python scripts/build.py --serve    # build and preview at http://localhost:8000
    python scripts/build.py --serve --compare-layouts   # + one page per publication layout
    python scripts/build.py --cv-files # also refresh cv/*.docx and cv/*.pdf in the repo

Pages come from pages/*.html wrapped in templates/layout.html. The publications
page is generated from the .ris files in publications/papers/ and
publications/preprints/. Only the Python standard library is used.
"""
import argparse
import datetime as dt
import html
import json
import re
import shutil
import sys
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "_site"
PUBS = ROOT / "publications"
CONFIG = json.loads((ROOT / "site.json").read_text(encoding="utf-8"))
MARKERS = "†*"
COLLAPSE_OVER = 20  # author lists longer than this get a "show all" button

# ---------------------------------------------------------------- RIS parsing


def parse_ris(text):
    """Return a list of {tag: [values]} dicts, one per record in the file."""
    records, cur, last = [], None, None
    for raw in text.lstrip("﻿").splitlines():
        m = re.match(r"^([A-Z][A-Z0-9])  -(?: (.*))?$", raw.rstrip())
        if not m:
            if cur is not None and last and raw.strip():  # wrapped continuation line
                cur[last][-1] += " " + raw.strip()
            continue
        tag, val = m.group(1), (m.group(2) or "").strip()
        if tag == "TY":
            cur, last = {"TY": [val]}, None
        elif tag == "ER":
            if cur:
                records.append(cur)
            cur, last = None, None
        elif cur is not None:
            cur.setdefault(tag, []).append(val)
            last = tag
    if cur:  # file without a final ER line
        records.append(cur)
    return records


def first(rec, *tags):
    for t in tags:
        for v in rec.get(t, []):
            if v:
                return v
    return ""


def clean_doi(doi):
    return re.sub(r"^(https?://)?(dx\.)?(www\.)?doi\.org/|^doi:\s*", "", doi.strip(), flags=re.I)


def norm(text):
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9, ]", " ", text)).strip()


def split_author(raw):
    """'Last, First †' -> ('First Last', '†', 'last, first')."""
    raw = raw.strip()
    marks = ""
    while raw and raw[-1] in MARKERS + " ":
        if raw[-1] in MARKERS:
            marks = raw[-1] + marks
        raw = raw[:-1]
    if raw.isupper():  # some publishers send "VALLI, MARILIA"
        raw = re.sub(r"[^\W\d_]+", lambda m: m.group(0).capitalize(), raw)
    if "," in raw:
        last, given = [p.strip() for p in raw.split(",", 1)]
        display = f"{given} {last}".strip()
    else:
        display = raw
    return display, "".join(sorted(set(marks), key=MARKERS.index)), raw


ME = [re.compile(p) for p in CONFIG["my_name_patterns"]]


def is_me(raw_name):
    n = norm(raw_name)
    return any(p.search(n) for p in ME)


def to_record(rec, kind, source):
    doi = clean_doi(first(rec, "DO"))
    title = first(rec, "TI", "T1", "CT", "BT")
    date = first(rec, "DA", "PY", "Y1", "Y2")
    year_m = re.search(r"(19|20)\d{2}", date)
    date_key = tuple(int(x) for x in re.findall(r"\d+", date)[:3]) if date else ()
    venue = html.unescape(first(rec, "T2", "JO", "JF", "JA", "J2"))
    if not venue and kind == "preprint":
        venue = "bioRxiv" if doi.startswith(("10.1101/", "10.64898/")) else first(rec, "PB") or "Preprint"
    notes = [n.split(":", 1)[1].strip() for n in rec.get("N1", []) if n.lower().startswith("status:")]
    url = f"https://doi.org/{doi}" if doi else first(rec, "UR", "L1", "L2")
    authors = []
    for a in rec.get("AU", []) + rec.get("A1", []):
        display, marks, raw = split_author(a)
        authors.append({"name": html.unescape(display), "marks": marks, "me": is_me(raw)})
    return {
        "title": title, "authors": authors, "venue": venue, "year": int(year_m.group(0)) if year_m else 0,
        "date_key": date_key, "doi": doi, "url": url, "kind": kind, "status": "; ".join(notes),
        "volume": first(rec, "VL"), "issue": first(rec, "IS"),
        "source": source,
    }


def load_publications():
    pubs, seen = [], {}
    for kind, folder in (("paper", "papers"), ("preprint", "preprints")):
        for path in sorted((PUBS / folder).glob("*.ris")):
            for rec in parse_ris(path.read_text(encoding="utf-8", errors="replace")):
                p = to_record(rec, kind, path.relative_to(ROOT).as_posix())
                if not p["title"]:
                    print(f"  warning: skipped a record without a title in {p['source']}")
                    continue
                key = p["doi"].lower() or norm(strip_tags(p["title"]))
                if key in seen:
                    print(f"  warning: {p['source']} duplicates {seen[key]} — keeping the first")
                    continue
                seen[key] = p["source"]
                pubs.append(p)
    pubs.sort(key=lambda p: (is_in_press(p), p["year"], p["date_key"], norm(p["title"])), reverse=True)
    for i, p in enumerate(pubs):
        p["number"] = len(pubs) - i
    return pubs


def is_in_press(p):
    """A paper that is accepted but not yet published: status says "in press" or "accepted"."""
    status = p["status"].lower()
    return p["kind"] == "paper" and ("press" in status or "accepted" in status)


def load_highlighted(pubs):
    path = PUBS / "highlighted.txt"
    if not path.exists():
        return []
    by_doi = {p["doi"].lower(): p for p in pubs if p["doi"]}
    by_title = {norm(strip_tags(p["title"])): p for p in pubs}
    out = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        p = by_doi.get(clean_doi(line).lower()) or by_title.get(norm(line))
        if p:
            out.append(p)
        else:
            print(f"  warning: highlighted.txt entry not found among the .ris files: {line}")
    return out


# ---------------------------------------------------------------- HTML


def strip_tags(s):
    return re.sub(r"<[^>]+>", "", s)


def safe_title(title):
    """Escape a title but keep the italic/sub/superscript markup publishers use."""
    t = html.escape(html.unescape(title), quote=False)
    t = re.sub(r"&lt;(/?)(i|em|sub|sup|b)&gt;", r"<\1\2>", t, flags=re.I)
    t = re.sub(r"&lt;/?[a-z:]+[^&]*?&gt;", "", t, flags=re.I)  # drop <scp>, <mml:...> and the like
    return t


def render_authors(authors, collapse_over=COLLAPSE_OVER, head=10, tail=3):
    n = len(authors)
    collapse = n > collapse_over
    keep = set(range(head)) | set(range(n - tail, n)) if collapse else set(range(n))
    keep |= {i for i, a in enumerate(authors) if a["me"] or a["marks"]}
    parts, gap = [], False
    for i, a in enumerate(authors):
        hidden = i not in keep
        if hidden and not gap:
            parts.append('<span class="gap">…</span>')
        gap = hidden
        cls = ["au"] + (["me"] if a["me"] else []) + (["extra"] if hidden else [])
        marks = f'<sup class="mark">{html.escape(a["marks"])}</sup>' if a["marks"] else ""
        parts.append(f'<span class="{" ".join(cls)}"><span class="nm">{html.escape(a["name"])}</span>{marks}</span>')
    toggle = f'<button type="button" class="more" aria-expanded="false">Show all {n} authors</button>' if collapse else ""
    return f'<p class="authors{" collapsed" if collapse else ""}">{"".join(parts)}</p>{toggle}'


def title_link(p):
    title = safe_title(p["title"])
    return f'<a href="{html.escape(p["url"])}" target="_blank" rel="noopener">{title}</a>' if p["url"] else title


def render_pub(p, number):
    """Numbered list entry (the "list" layout and the home page)."""
    badge = '<span class="badge">Preprint</span>' if p["kind"] == "preprint" else ""
    status = f', {html.escape(p["status"])}' if p["status"] else ""
    return (
        f'<li class="pub" data-kind="{p["kind"]}">'
        f'<span class="num">{number}.</span><div class="body">'
        f'<h3 class="pub-title">{title_link(p)}</h3>'
        f'{render_authors(p["authors"])}'
        f'<p class="venue"><strong>{html.escape(p["venue"])}</strong>{status}, {p["year"] or "n.d."} {badge}</p>'
        f"</div></li>"
    )


def render_entry(p, show_year=True, compact=False, tag="li"):
    """Journal-first entry used by the "timeline" and "cards" layouts."""
    meta = [f'<span class="journal">{html.escape(p["venue"])}</span>']
    if p["status"]:
        meta.append(f'<span class="status">{html.escape(p["status"])}</span>')
    if show_year:
        meta.append(f'<span class="yr">{p["year"] or "n.d."}</span>')
    if p["kind"] == "preprint":
        meta.append('<span class="badge">Preprint</span>')
    authors = render_authors(p["authors"], 6, 3, 1) if compact else render_authors(p["authors"])
    return (
        f'<{tag} class="pub entry-pub" data-kind="{p["kind"]}">'
        f'<p class="pub-meta">{"".join(meta)}</p>'
        f'<h3 class="pub-title">{title_link(p)}</h3>{authors}</{tag}>'
    )


SEARCH_CONTROLS = """
      <input type="search" id="pub-search" placeholder="Search title, author, journal…" aria-label="Search publications">
      <div class="filters" role="group" aria-label="Filter">
        <button type="button" class="chip on" data-filter="all">All</button>
        <button type="button" class="chip" data-filter="paper">Papers</button>
        <button type="button" class="chip" data-filter="preprint">Preprints</button>
      </div>"""
LEGEND = '<p class="legend"><span><b class="mark">†</b> co-first authors</span><span><b class="mark">*</b> corresponding author</span><span><b class="me-swatch">Name</b> me</span></p>'


def publications_page(pubs, highlighted, layout=None):
    layout = layout or CONFIG.get("publication_layout", "list")
    years = sorted({p["year"] for p in pubs}, reverse=True)
    n_pre = sum(p["kind"] == "preprint" for p in pubs)
    n_first = sum(1 for p in pubs if p["authors"] and (p["authors"][0]["me"] or any(a["me"] and "†" in a["marks"] for a in p["authors"])))
    head = f"""
<section class="page-head">
  <p class="eyebrow">Research output</p>
  <h1>Publications</h1>
  <p class="lede">{len(pubs) - n_pre} peer-reviewed papers and {n_pre} preprints, {n_first} as first or co-first author.
  For citation metrics, see my <a href="https://scholar.google.com/citations?user=-HEC0H8AAAAJ&amp;hl=en" target="_blank" rel="noopener">Google Scholar profile</a>.</p>
</section>"""
    empty = '<p id="no-results" class="card" hidden>No publications match your search.</p>'
    by_year = {y: [p for p in pubs if p["year"] == y] for y in years}

    if layout == "timeline":
        tabs = ['<a href="#highlighted">Highlighted</a>'] if highlighted else []
        tabs += [f'<a href="#year-{y}">{y or "n.d."}</a>' for y in years]
        featured = "".join(render_entry(p, compact=True, tag="article") for p in highlighted)
        groups = "".join(
            f'<section id="year-{y}" class="year-group tl-group"><div class="tl-year"><span>{y or "n.d."}</span>'
            f'<small>{len(items)} {"paper" if len(items) == 1 else "papers"}</small></div>'
            f'<ol class="pubs tl-list">{"".join(render_entry(p, show_year=False) for p in items)}</ol></section>'
            for y, items in by_year.items()
        )
        return f"""{head}
<div class="tl-bar">
  <div class="tl-controls">{SEARCH_CONTROLS}</div>
  <nav class="year-nav tabs" aria-label="Years">{"".join(tabs)}</nav>
</div>
{LEGEND}
{f'<section id="highlighted" class="pub-group featured-wrap"><h2>Highlighted</h2><div class="featured">{featured}</div></section>' if highlighted else ""}
{empty}
<div class="timeline-pubs">{groups}</div>"""

    if layout == "cards":
        options = "".join(f'<option value="year-{y}">{y or "n.d."} ({len(items)})</option>' for y, items in by_year.items())
        featured = "".join(render_entry(p, compact=True, tag="article") for p in highlighted)
        groups = "".join(
            f'<section id="year-{y}" class="year-group cards-group"><h2 class="group-bar"><span>{y or "n.d."}</span>'
            f'<small>{len(items)}</small></h2>'
            f'<ol class="pubs card-grid">{"".join(render_entry(p, show_year=False, compact=True) for p in items)}</ol></section>'
            for y, items in by_year.items()
        )
        return f"""{head}
<div class="toolbar card cards-toolbar">{SEARCH_CONTROLS}
  <label class="year-select">Year
    <select id="year-select"><option value="">All years</option>{options}</select>
  </label>
  {LEGEND}
</div>
{f'<section id="highlighted" class="pub-group featured-wrap"><h2>Highlighted</h2><div class="featured">{featured}</div></section>' if highlighted else ""}
{empty}
{groups}"""

    # default: "list" — numbered lists with a sticky year sidebar
    in_press = [p for p in pubs if is_in_press(p)]
    year_items = {y: [p for p in items if not is_in_press(p)] for y, items in by_year.items()}
    year_items = {y: items for y, items in year_items.items() if items}
    jump = ['<a class="pill active" href="#highlighted">Highlighted</a>'] if highlighted else []
    jump += ['<a class="pill" href="#in-press">Accepted</a>'] if in_press else []
    jump += [f'<a class="pill" href="#year-{y}">{y or "n.d."}</a>' for y in year_items]
    sections = []
    if highlighted:
        featured = "".join(render_entry(p, compact=True, tag="article") for p in highlighted)
        sections.append(f'<section id="highlighted" class="pub-group featured-wrap"><h2>Highlighted papers</h2><div class="featured">{featured}</div></section>')
    sections.append('<h2 class="all-heading">All publications</h2>')
    if in_press:
        sections.append(f'<section id="in-press" class="card pub-group year-group"><h2>Accepted for publication</h2><ol class="pubs">{"".join(render_pub(p, p["number"]) for p in in_press)}</ol></section>')
    for y, items in year_items.items():
        sections.append(f'<section id="year-{y}" class="card pub-group year-group"><h2>{y or "Undated"}</h2><ol class="pubs">{"".join(render_pub(p, p["number"]) for p in items)}</ol></section>')
    return f"""{head}
<div class="pub-layout">
  <aside class="jump">
    <p class="eyebrow">Jump to</p>
    <nav class="year-nav">{"".join(jump)}</nav>
  </aside>
  <div class="pub-main">
    <div class="toolbar card">{SEARCH_CONTROLS}
      {LEGEND}
    </div>
    {empty}
    {"".join(sections)}
  </div>
</div>"""


def selected_list(highlighted):
    items = "".join(render_pub(p, i + 1) for i, p in enumerate(highlighted[: CONFIG.get("selected_on_home", 5)]))
    return f'<ol class="pubs">{items}</ol>'


# ---------------------------------------------------------------- presentations

TALK_LABELS = {"oral": "Oral", "poster": "Poster", "invited": "Invited talk", "workshop": "Workshop"}


def load_presentations():
    """Parse data/presentations.txt: blank-line separated blocks of 'key: value' lines."""
    path = ROOT / "data" / "presentations.txt"
    talks = []
    for block in re.split(r"\n\s*\n", path.read_text(encoding="utf-8")):
        fields = {}
        for line in block.splitlines():
            if line.lstrip().startswith("#") or ":" not in line:
                continue
            key, val = line.split(":", 1)
            fields[key.strip().lower()] = val.strip()
        if not fields:
            continue
        if not fields.get("title") or not re.fullmatch(r"\d{4}", fields.get("year", "")):
            print(f"  warning: presentation skipped (needs year and title): {fields}")
            continue
        types = [t.strip().lower() for t in fields.get("type", "").split(",") if t.strip()]
        for t in types:
            if t not in TALK_LABELS:
                print(f"  warning: unknown presentation type '{t}' in: {fields['title']}")
        location = fields.get("location", "")
        if fields.get("mode", "").lower() == "online":
            location = f"{location} (online)" if location else "Online"
        talks.append({"year": int(fields["year"]), "types": types, "title": fields["title"],
                      "event": fields.get("event", ""), "location": location, "audience": fields.get("audience", "")})
    talks.sort(key=lambda t: t["year"], reverse=True)  # stable: keeps file order within a year
    return talks


def presentations_page(talks):
    years = sorted({t["year"] for t in talks}, reverse=True)
    out = ['<section class="page-head">\n  <p class="eyebrow">Talks, posters &amp; workshops</p>\n  <h1>Presentations</h1>\n</section>\n']
    for y in years:
        merged = {}  # same title within a year -> one entry listing every event
        for t in (t for t in talks if t["year"] == y):
            entry = merged.setdefault(t["title"], {"events": [], "types": []})
            where = ", ".join(x for x in (t["event"], t["location"]) if x)
            entry["events"].append(where + (f". Estimated audience: {t['audience']}" if t["audience"] else ""))
            entry["types"] += [x for x in t["types"] if x not in entry["types"]]
        items = "".join(
            f'<li><h3>{html.escape(title)}</h3>{"".join(f"<p>{html.escape(e)}.</p>" for e in e_["events"])}'
            f'<p class="tags">{"".join(f"<span class=kind>{TALK_LABELS.get(k, k.title())}</span>" for k in e_["types"])}</p></li>'
            for title, e_ in merged.items()
        )
        out.append(f'<h2 class="group-title" id="y{y}">{y}</h2>\n<div class="card list-card">\n  <ul>{items}</ul>\n</div>\n')
    return "\n".join(out)


# ---------------------------------------------------------------- CV


def cv_pub(p):
    authors = []
    for a in p["authors"]:
        name = html.escape(a["name"])
        authors.append((f"<b>{name}</b>" if a["me"] else name) + (f"<sup>{html.escape(a['marks'])}</sup>" if a["marks"] else ""))
    journal = [f"<i>{html.escape(p['venue'])}</i>"]
    if p["volume"]:
        journal.append(f"Vol. {html.escape(p['volume'])}")
    if p["issue"]:
        journal.append(f"No. {html.escape(p['issue'])}")
    if p["status"]:
        journal.append(html.escape(p["status"]))
    journal.append(str(p["year"] or "n.d."))
    link = f' <a href="https://doi.org/{html.escape(p["doi"])}">doi.org/{html.escape(p["doi"])}</a>' if p["doi"] else ""
    return (f'<li><span class="t">{safe_title(p["title"])}</span><br>'
            f'<span class="a">{", ".join(authors)}</span><br>'
            f'<span class="j">{", ".join(journal)}.</span>{link}</li>')


def is_first_or_cofirst(p):
    return bool(p["authors"]) and (p["authors"][0]["me"] or any(a["me"] and "†" in a["marks"] for a in p["authors"]))


def cv_values(pubs, talks):
    accepted = [p for p in pubs if is_in_press(p)]
    published = [p for p in pubs if p["kind"] == "paper" and p not in accepted]
    preprints = [p for p in pubs if p["kind"] == "preprint"]

    def pub_list(items):
        return f'<ol class="cv-pubs" reversed>{"".join(cv_pub(p) for p in items)}</ol>'

    def section(title, items):
        return f"<section><h2>{title}</h2>{pub_list(items)}</section>" if items else ""

    def talk_rows(items):
        rows = []
        for t in items:
            where = ", ".join(x for x in (t["event"], t["location"]) if x)
            detail = where + (f". Estimated audience: {t['audience']}." if t["audience"] else "")
            rows.append(f'<div class="item row"><div><h3>{html.escape(t["title"])}</h3>'
                        f'<p>{html.escape(detail)}</p></div><span class="when">{t["year"]}</span></div>')
        return "".join(rows)

    def talk_group(title, kind):
        items = [t for t in talks if kind in t["types"]]
        if not items:
            return ""
        return f'<h3 class="sub">{title}</h3><div class="cv-talks">{talk_rows(items)}</div>'

    n_first = sum(is_first_or_cofirst(p) for p in published)
    return {
        "cv_pub_counts": f"{len(published)} published articles, {n_first} as first or co-first author.",
        "cv_publications": pub_list(published),
        "cv_accepted": section("Manuscripts Accepted for Publication", accepted),
        "cv_preprints": section("Preprints", preprints),
        "cv_invited": talk_group("Invited talks", "invited"),
        "cv_oral": talk_group("Oral presentations", "oral"),
        "cv_posters": talk_group("Poster presentations", "poster"),
        "cv_workshops": f'<div class="cv-talks">{talk_rows([t for t in talks if "workshop" in t["types"]])}</div>',
        "cv_updated": dt.date.today().strftime("%B %d, %Y"),
    }


def find_chrome():
    import os
    candidates = [os.environ.get("CHROME", ""), "google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
                  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                  "/Applications/Chromium.app/Contents/MacOS/Chromium"]
    for c in candidates:
        if c and (shutil.which(c) or Path(c).exists()):
            return shutil.which(c) or c
    return None


def make_pdf(html_path, pdf_path):
    import subprocess
    chrome = find_chrome()
    if not chrome:
        print("  warning: Chrome/Chromium not found, cv.pdf was not created (it will be on GitHub)")
        return False
    if pdf_path.exists():
        pdf_path.unlink()
    cmd = [chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={pdf_path}"]
    if sys.platform.startswith("linux"):
        cmd.append("--no-sandbox")  # needed on GitHub's runners
    try:
        subprocess.run(cmd + [html_path.resolve().as_uri()], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
    except subprocess.TimeoutExpired:
        print("  warning: Chrome timed out while printing the CV")
    if not pdf_path.exists():
        print("  warning: Chrome ran but cv.pdf was not created")
        return False
    return True


def update_repo_cv(stable_html, cv_html):
    """Keep cv/<name>_CV.docx and .pdf in the repository in sync with the website CV.

    The files are rewritten only when the CV content changed (the "last updated"
    date is ignored), so the GitHub Action does not commit a new copy every run.
    """
    import hashlib
    folder = ROOT / "cv"
    base = CONFIG["name"].replace(" ", "_") + "_CV"
    docx_path, pdf_path, stamp = folder / f"{base}.docx", folder / f"{base}.pdf", folder / ".source-hash"
    digest = hashlib.sha256(stable_html.encode("utf-8")).hexdigest()
    up_to_date = stamp.exists() and stamp.read_text().strip() == digest and docx_path.exists() and pdf_path.exists()
    if up_to_date:
        print("  repository CV files already up to date")
        return
    try:
        from cv_docx import html_to_docx
    except ImportError:
        print("  warning: python-docx is not installed (pip install python-docx); CV .docx not updated")
        return
    folder.mkdir(exist_ok=True)
    html_to_docx(cv_html, docx_path, title=f"{CONFIG['name']} — Curriculum Vitae")
    if (OUT / "cv.pdf").exists():
        shutil.copyfile(OUT / "cv.pdf", pdf_path)
    stamp.write_text(digest + "\n")
    print(f"  updated {docx_path.relative_to(ROOT)} and {pdf_path.name}")


def profile_links():
    return "".join(
        f'<a class="profile" href="{html.escape(url)}" target="_blank" rel="noopener">'
        f'<img src="assets/img/icons/{icon}.png" alt="" width="20" height="20">{html.escape(label)}</a>'
        for label, url, icon in CONFIG["links"]
    )


def build(compare_layouts=False, cv_files=False):
    print("Building site…")
    pubs = load_publications()
    highlighted = load_highlighted(pubs)
    print(f"  {len(pubs)} publications ({sum(p['kind'] == 'preprint' for p in pubs)} preprints), {len(highlighted)} highlighted")

    if OUT.exists():
        shutil.rmtree(OUT)
    shutil.copytree(ROOT / "assets", OUT / "assets")
    layout = (ROOT / "templates" / "layout.html").read_text(encoding="utf-8")
    today = dt.date.today()
    values = {
        "name": html.escape(CONFIG["name"]),
        "tagline": html.escape(CONFIG["tagline"]),
        "profile_links": profile_links(),
        "selected_publications": selected_list(highlighted),
        "updated": today.strftime("%B %Y"),
        "year": str(today.year),
    }

    def fill(text):
        return re.sub(r"\{\{\s*(\w+)\s*\}\}", lambda m: values.get(m.group(1), m.group(0)), text)

    def write(page, title, body, current=None):
        nav = "".join(
            f'<a href="{href}"{" aria-current=page" if href == (current or page) else ""}>{label}</a>' for href, label in CONFIG["nav"]
        )
        full_title = CONFIG["name"] if page == "index.html" else f"{title} · {CONFIG['name']}"
        out = fill(layout.replace("{{content}}", fill(body)).replace("{{nav}}", nav).replace("{{title}}", html.escape(full_title)))
        (OUT / page).write_text(out, encoding="utf-8")

    talks = load_presentations()
    generated = {"publications.html": lambda: publications_page(pubs, highlighted), "presentations.html": lambda: presentations_page(talks)}
    for page, title in CONFIG["pages"].items():
        body = generated[page]() if page in generated else (ROOT / "pages" / page).read_text(encoding="utf-8")
        write(page, title, body)

    cv_fill = {**values, **cv_values(pubs, talks)}
    cv_html = re.sub(r"\{\{\s*(\w+)\s*\}\}", lambda m: cv_fill.get(m.group(1), m.group(0)),
                     (ROOT / "templates" / "cv.html").read_text(encoding="utf-8"))
    (OUT / "cv.html").write_text(cv_html, encoding="utf-8")
    if make_pdf(OUT / "cv.html", OUT / "cv.pdf"):
        print("  wrote CV: cv.html and cv.pdf")
    if cv_files:
        update_repo_cv(cv_html.replace(cv_fill["cv_updated"], ""), cv_html)
    if compare_layouts:  # side-by-side previews: publications-list.html, -timeline.html, -cards.html
        for name in ("list", "timeline", "cards"):
            write(f"publications-{name}.html", "Publications", publications_page(pubs, highlighted, name), "publications.html")
        print("  layout previews: publications-list.html, publications-timeline.html, publications-cards.html")
    (OUT / ".nojekyll").write_text("")
    base = CONFIG.get("base_url", "").rstrip("/")
    if base:  # sitemap + robots.txt so search engines find every page
        urls = [f"{base}/" if p == "index.html" else f"{base}/{p}" for p in CONFIG["pages"]] + [f"{base}/cv.html"]
        entries = "".join(f"  <url><loc>{u}</loc><lastmod>{today.isoformat()}</lastmod></url>\n" for u in urls)
        (OUT / "sitemap.xml").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
            + entries + "</urlset>\n", encoding="utf-8")
        (OUT / "robots.txt").write_text(f"User-agent: *\nAllow: /\nSitemap: {base}/sitemap.xml\n", encoding="utf-8")
    print(f"  wrote {len(CONFIG['pages'])} pages to {OUT.relative_to(ROOT)}/")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--serve", action="store_true", help="preview at http://localhost:8000 after building")
    ap.add_argument("--compare-layouts", action="store_true", help="also write one preview page per publication layout")
    ap.add_argument("--cv-files", action="store_true", help="also update cv/*.docx and cv/*.pdf in the repository")
    args = ap.parse_args()
    build(args.compare_layouts, args.cv_files)
    if args.serve:
        import functools
        import http.server

        handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(OUT))
        print("Preview at http://localhost:8000  (Ctrl+C to stop)")
        http.server.ThreadingHTTPServer(("127.0.0.1", 8000), handler).serve_forever()
