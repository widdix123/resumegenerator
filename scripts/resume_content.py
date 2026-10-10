"""Read resume.md into the constants the resume builders use.

resume.md is the only file to edit for content. See the comment at its top for
the format. Errors name the line number so a typo is easy to find.
"""
from pathlib import Path
import re

SOURCE = Path(__file__).resolve().parent.parent / "resume.md"


class ResumeFormatError(ValueError):
    pass


def _split(text, sep, n, lineno):
    parts = [p.strip() for p in text.split(sep)]
    if len(parts) != n:
        raise ResumeFormatError(f"resume.md line {lineno}: expected {n} parts separated by '{sep.strip()}', got {len(parts)}: {text}")
    return parts


def parse(path=SOURCE):
    text = re.sub(r"<!--.*?-->", lambda m: "\n" * m.group().count("\n"), path.read_text(encoding="utf-8"), flags=re.S)
    data = {"meta": {}, "sections": {}}
    section, job = None, None
    for lineno, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()
        if not line:
            continue
        if line.startswith("### "):
            if section != "experience":
                raise ResumeFormatError(f"resume.md line {lineno}: '###' jobs belong under '## Experience'")
            title, company, location, dates = _split(line[4:], " | ", 4, lineno)
            job = {"title": title, "company": company, "location": location, "dates": dates, "bullets": []}
            data["sections"]["experience"].append(job)
        elif line.startswith("## "):
            section = line[3:].strip().lower()
            data["sections"][section] = []
            job = None
        elif line.startswith("# "):
            data["name"] = line[2:].strip()
        elif section is None:
            key, sep, value = line.partition(":")
            if not sep:
                raise ResumeFormatError(f"resume.md line {lineno}: expected 'key: value' before the first section: {line}")
            data["meta"][key.strip().lower()] = value.strip()
        elif line.startswith("- "):
            item = (line[2:].strip(), lineno)
            (job["bullets"] if job else data["sections"][section]).append(item[0] if job else item)
        else:  # paragraph text (Summary)
            data["sections"][section].append((line, lineno))
    return data


def _pairs(items, lineno_label=":"):
    out = []
    for text, lineno in items:
        label, sep, value = text.partition(lineno_label)
        if not sep:
            raise ResumeFormatError(f"resume.md line {lineno}: expected 'Label: items': {text}")
        out.append((label.strip(), value.strip()))
    return out


def _require(sections, name):
    if name not in sections:
        raise ResumeFormatError(f"resume.md: missing section '## {name.title()}'")
    return sections[name]


_d = parse()
_s = _d["sections"]

NAME = _d["name"]
HEADLINE = _d["meta"].get("headline", "")
TAGLINE = _d["meta"].get("tagline", "")
CONTACT = {k.lower(): v for k, v in _pairs(_require(_s, "contact"))}
SUMMARY = " ".join(t for t, _ in _require(_s, "summary"))
HIGHLIGHTS = [tuple(_split(t, " | ", 2, n)) for t, n in _require(_s, "highlights")]
CERTIFICATIONS = [t for t, _ in _require(_s, "certifications")]
EDUCATION = [tuple(_split(t, " | ", 3, n)) for t, n in _require(_s, "education")]
SIDEBAR_SKILLS = _pairs(_require(_s, "sidebar skills"))
SKILLS = _pairs(_require(_s, "technical skills"))
CORE_COMPETENCIES = [t for t, _ in _require(_s, "core competencies")]
EXPERIENCE = _require(_s, "experience")
