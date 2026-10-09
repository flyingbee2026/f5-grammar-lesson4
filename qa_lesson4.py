#!/usr/bin/env python3
"""QA for Lesson 4: structure counts, dedupe scan, sample word/relative counts, PDF pages."""
import re
import subprocess
import sys

import docx
from docx.oxml.ns import qn

FILE = "Lesson 4 - Cleft sentences (Teacher's).docx"
d = docx.Document(FILE)

# ---- collect text in order
texts = []
for p in d.paragraphs:
    if p.text.strip():
        texts.append(p.text)
cells = []
for t in d.tables:
    for row in t.rows:
        for c in row.cells:
            if c.text.strip():
                cells.append(c.text)
allt = texts + cells

# ---- 1. sentences dedupe scan (>25 chars)
def sentences(s):
    s = re.sub(r"\s+", " ", s)
    return [x.strip() for x in re.split(r"(?<=[.!?])\s+", s) if len(x.strip()) > 25]

seen, dupes = {}, []
for chunk in allt:
    for s in sentences(chunk):
        key = s.lower()
        if key in seen and s != seen[key]:
            pass
        if key in seen:
            dupes.append(s)
        else:
            seen[key] = s
print(f"[dedupe] {len(seen)} unique sentences >25 chars; {len(dupes)} exact repeats")
for x in dupes:
    print("   REPEAT:", x)

# ---- 2. warm-up answers / quiz answers present
print(f"[counts] tables={len(d.tables)} paragraphs={len(d.paragraphs)}")
warm = d.tables[2]
print(f"[warm-up] rows={len(warm.rows)} answers={'/'.join(warm.cell(i,1).text for i in range(1,21))}")
quiz = d.tables[0]
print(f"[quiz] rows={len(quiz.rows)}")
ans = d.tables[1]
print(f"[quiz key] rows={len(ans.rows)}")

# ---- 3. samples: word count, relative pronouns, clefts
samples = [p.text for p in d.paragraphs if p.text.startswith("Sample answer")]
REL = r"\b(which|who|whom|whose|that)\b"
for s in samples:
    label = s.split(")")[0] + ")"
    body = s.split(") ", 1)[1]
    words = len(body.split())
    rels = re.findall(REL, body)
    clefts = re.findall(r"\b(It is|It was|What .{1,60}? is|All .{1,60}? is)\b", body)
    print(f"[sample] {label}: {words} words | that/who/which = {rels} | cleft markers={clefts}")

# ---- 4. quiz: two errors each (manual review aid)
print("[quiz items]")
for i, s in enumerate([r.cells[0].text for r in quiz.rows], 1):
    print(f"  {i}. {s}")

# ---- 5. set tables: answer vs plain
print("[sets]")
for h, t in zip([p.text for p in d.paragraphs if p.text.startswith("Example Set")],
                [tb for tb in d.tables if len(tb.rows) == 3]):
    print(f"  {h}\n    plain : {t.cell(0,1).text}\n    prompt: {t.cell(1,1).text}")
    print(f"    answer: {t.cell(2,1).text}")

# ---- 6. PDF render + per-page text
subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", FILE],
               check=True, capture_output=True, timeout=300)
pdf = FILE.replace(".docx", ".pdf")
out = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout
pages = out.split("\f")
print(f"[pdf] {pdf}: {len([p for p in pages if p.strip()])} non-empty pages of {len(pages)-1} sheets")
for i, p in enumerate(pages, 1):
    lines = [l for l in p.splitlines() if l.strip()]
    if not lines:
        print(f"  page {i}: <blank>")
    else:
        print(f"  page {i}: {len(lines)} lines | first: {lines[0][:70]!r} | last: {lines[-1][:70]!r}")
