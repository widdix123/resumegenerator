"""Designed two-column resume with photo, for emailing recruiters and direct sharing.

Content comes from resume.md. For job portals, upload the ATS
version instead: sidebars and photos can scramble ATS parsing.
"""
from pathlib import Path
import base64
import html
import subprocess

from create_ats_resume import CHROME
from resume_content import (
    CERTIFICATIONS, CONTACT, CORE_COMPETENCIES, EDUCATION, EXPERIENCE,
    HEADLINE, HIGHLIGHTS, NAME, SIDEBAR_SKILLS, SKILLS, SUMMARY, TAGLINE,
)


def split_lead(bullet):
    """'Lead — detail' -> ('Lead', 'detail'); bullets without a lead return ('', bullet)."""
    lead, sep, rest = bullet.partition(" — ")
    return (lead, rest) if sep else ("", bullet)

ROOT = Path(__file__).resolve().parent.parent  # the resume folder (scripts/ lives inside it)
PHOTO = ROOT / "AbhishekGupta.jpg"
PDF_OUT = ROOT / "Abhishek_Gupta_Resume_Designed.pdf"
HTML_OUT = ROOT / "preview" / "designed_resume.html"

ICONS = {
    "location": '<path d="M12 2a7 7 0 0 0-7 7c0 5 7 13 7 13s7-8 7-13a7 7 0 0 0-7-7zm0 9.5A2.5 2.5 0 1 1 12 6.5a2.5 2.5 0 0 1 0 5z"/>',
    "phone": '<path d="M6.6 10.8a15 15 0 0 0 6.6 6.6l2.2-2.2a1 1 0 0 1 1-.25 11.4 11.4 0 0 0 3.6.57 1 1 0 0 1 1 1V20a1 1 0 0 1-1 1A17 17 0 0 1 3 4a1 1 0 0 1 1-1h3.5a1 1 0 0 1 1 1c0 1.25.2 2.45.57 3.57a1 1 0 0 1-.25 1z"/>',
    "email": '<path d="M20 4H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V6a2 2 0 0 0-2-2zm0 4-8 5-8-5V6l8 5 8-5z"/>',
    "linkedin": '<path d="M19 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zM8.3 18H5.7V9.7h2.6zM7 8.6a1.5 1.5 0 1 1 0-3 1.5 1.5 0 0 1 0 3zM18.3 18h-2.6v-4c0-1-.02-2.2-1.35-2.2-1.35 0-1.55 1.05-1.55 2.13V18h-2.6V9.7h2.5v1.1h.03a2.7 2.7 0 0 1 2.45-1.35c2.6 0 3.1 1.7 3.1 3.95z"/>',
}

CSS = """
@page { size: A4; margin: 0; }
* { box-sizing: border-box; }
:root { --amber: #F2A541; --navy: #1B3A5C; --navy-2: #24496F; --accent: #3D8BD4; --ink: #1E2530; --muted: #5B6675; --line: #D9E1EA; }
body { margin: 0; font-family: "Avenir Next", "Segoe UI", Helvetica, Arial, sans-serif; font-size: 9.3pt;
       line-height: 1.33; color: var(--ink); -webkit-print-color-adjust: exact; print-color-adjust: exact; }
/* Fixed elements repeat on every printed page, so the sidebar runs edge to edge. */
.side-bg { position: fixed; left: 0; top: 0; bottom: 0; width: 2.55in; background: var(--navy); }
/* clone repeats the top/bottom padding on each page fragment, acting as page margins. */
aside, main { padding-top: 0.42in !important; padding-bottom: 0.42in !important;
              box-decoration-break: clone; -webkit-box-decoration-break: clone; }
.layout { display: grid; grid-template-columns: 2.55in 1fr; position: relative; }
aside { color: #E8EEF5; padding: 0 0.26in 0 0.3in; }
main { padding: 0 0.42in 0 0.34in; }
.photo { width: 1.55in; height: 1.55in; border-radius: 50%; border: 4px solid #fff; object-fit: cover; object-position: 50% 15%;
         display: block; margin: 0.05in auto 0.2in; }
aside h2 { font-size: 10.5pt; letter-spacing: .14em; text-transform: uppercase; color: #fff; margin: 16px 0 8px;
           padding-bottom: 5px; border-bottom: 2px solid var(--amber); }
aside p, aside li { font-size: 8.8pt; }
.contact { list-style: none; padding: 0; margin: 0; }
.contact li { display: flex; gap: 8px; align-items: flex-start; margin-bottom: 7px; word-break: break-word; }
.contact svg { width: 13px; height: 13px; fill: #9CC6EE; flex: none; margin-top: 2px; }
.cert { margin: 0 0 7px; padding-left: 10px; border-left: 2px solid var(--amber); }
.edu b { color: #fff; display: block; }
.skill { margin-bottom: 7px; }
.skill b { display: block; color: #9CC6EE; font-size: 8.4pt; letter-spacing: .04em; text-transform: uppercase; }
.chips { display: flex; flex-wrap: wrap; gap: 4px; }
.chips span { font-size: 8.2pt; border: 1px solid rgba(255,255,255,.35); border-radius: 10px; padding: 1px 7px; }
header h1 { font-size: 27pt; line-height: 1.05; margin: 0.08in 0 4px; color: var(--navy); letter-spacing: .02em; }
header h1 span { font-weight: 400; }
header .role { font-size: 10.5pt; color: var(--muted); letter-spacing: .03em; margin: 0; padding-bottom: 10px;
               border-bottom: 2px solid var(--navy); display: inline-block; }
main h2 { font-size: 11pt; letter-spacing: .14em; text-transform: uppercase; color: var(--navy); margin: 13px 0 7px;
          padding-bottom: 4px; border-bottom: 2px solid var(--amber); break-after: avoid; }
.summary { margin: 0; text-align: justify; }
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-top: 12px; }
.stat { background: var(--navy); border-radius: 6px; padding: 7px 8px; text-align: center; }
.stat b { display: block; font-size: 15pt; color: var(--amber); line-height: 1.1; }
.stat span { font-size: 7.8pt; color: #E8EEF5; }
.tagline { font-style: italic; color: var(--accent); font-size: 10.5pt; margin: 6px 0 0; }
.job li b { color: var(--navy); }
.timeline { position: relative; padding-left: 16px; }
.job { position: relative; margin-bottom: 8px; }
.job-head, .org { break-after: avoid; }
.job li { break-inside: avoid; }
.job::before { content: ""; position: absolute; left: -16px; top: 5px; width: 8px; height: 8px; border-radius: 50%;
               background: var(--navy); }
.job::after { content: ""; position: absolute; left: -12.5px; top: 15px; bottom: -11px; width: 1px; background: var(--line); }
.job:last-child::after { display: none; }
.job-head { display: flex; justify-content: space-between; gap: 10px; }
.job h3 { font-size: 10.3pt; margin: 0; color: var(--ink); }
.dates { font-size: 8.6pt; color: var(--muted); white-space: nowrap; font-weight: 600; }
.org { color: var(--accent); font-weight: 600; font-size: 9pt; margin: 0 0 3px; }
.job ul { margin: 0; padding-left: 14px; }
.job li { margin-bottom: 1.5px; }
.job li::marker { color: var(--accent); }
.competencies { display: flex; flex-wrap: wrap; gap: 5px; }
.competencies span { font-size: 8.4pt; background: #EEF4FA; color: var(--navy); border-radius: 10px; padding: 2px 8px; }
"""


def build_html():
    e = html.escape
    photo = base64.b64encode(PHOTO.read_bytes()).decode()
    first, _, last = NAME.partition(" ")

    contact = "".join(
        f'<li><svg viewBox="0 0 24 24">{ICONS[k]}</svg><span>{e(v)}</span></li>'
        for k, v in CONTACT.items() if v
    )
    certs = "".join(f'<p class="cert">{e(c)}</p>' for c in CERTIFICATIONS)
    edu = "".join(
        f'<p class="edu"><b>{e(d)}</b>{e(s)}<br>{e(y)}</p>' for d, s, y in EDUCATION
    )
    skills = "".join(
        f'<div class="skill"><b>{e(label)}</b><div class="chips">'
        + "".join(f"<span>{e(t)}</span>" for t in items.split(" · "))
        + "</div></div>"
        for label, items in SIDEBAR_SKILLS
    )
    stats = "".join(f'<div class="stat"><b>{e(n)}</b><span>{e(t)}</span></div>' for n, t in HIGHLIGHTS)
    comps = "".join(f"<span>{e(c)}</span>" for c in CORE_COMPETENCIES)
    jobs = "".join(
        f'<div class="job"><div class="job-head"><h3>{e(j["title"])}</h3>'
        f'<span class="dates">{e(j["dates"])}</span></div>'
        f'<p class="org">{e(j["company"])} · {e(j["location"])}</p><ul>'
        + "".join(f"<li><b>{e(l)}</b>{' — ' if l else ''}{e(r)}</li>" for l, r in map(split_lead, j["bullets"]))
        + "</ul></div>"
        for j in EXPERIENCE
    )

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<title>{e(NAME)} – Resume</title><style>{CSS}</style></head><body>
<div class="side-bg"></div>
<div class="layout">
<aside>
  <img class="photo" src="data:image/jpeg;base64,{photo}" alt="{e(NAME)}">
  <h2>Contact</h2><ul class="contact">{contact}</ul>
  <h2>Certifications</h2>{certs}
  <h2>Education</h2>{edu}
  <h2>Skills</h2>{skills}
  <h2>Expertise</h2><div class="chips">{comps}</div>
</aside>
<main>
  <header><h1>{e(first)} <span>{e(last)}</span></h1><p class="role">{e(HEADLINE)}</p><p class="tagline">{e(TAGLINE)}</p></header>
  <h2>Profile</h2><p class="summary">{e(SUMMARY)}</p>
  <div class="stats">{stats}</div>
  <h2>Work Experience</h2><div class="timeline">{jobs}</div>
</main>
</div></body></html>"""


def build_pdf():
    HTML_OUT.parent.mkdir(exist_ok=True)
    HTML_OUT.write_text(build_html(), encoding="utf-8")
    subprocess.run(
        [CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
         f"--print-to-pdf={PDF_OUT}", HTML_OUT.as_uri()],
        check=True, capture_output=True,
    )


if __name__ == "__main__":
    build_pdf()
    print(f"Wrote {PDF_OUT.name}")
