#!/usr/bin/env python3
"""QA for Lesson 4 (teacher + student): structure, content correctness, PDF pages.

usage: python3 qa_lesson4.py [teacher|student]
"""
import re
import subprocess
import sys
import unicodedata

import docx
from docx.oxml.ns import qn

WHO = sys.argv[1] if len(sys.argv) > 1 else "teacher"
FILE = {"teacher": "Lesson 4 - Cleft sentences (Teacher's).docx",
        "student": "Lesson 4 - Cleft sentences (Student's).docx"}[WHO]
d = docx.Document(FILE)
fail = []


def chk(ok, msg):
    print(("  OK   " if ok else "  FAIL ") + msg)
    if not ok:
        fail.append(msg)


paras = [p.text for p in d.paragraphs if p.text.strip()]
tables = d.tables
body = [c.text.strip() for t in tables for row in t.rows for c in row.cells]

# ---------------------------------------------------------------- 1. dedupe scan
def sentences(s):
    s = re.sub(r"\s+", " ", s)
    return [x.strip() for x in re.split(r"(?<=[.!?])\s+", s) if len(x.strip()) > 25]


seen, dupes = {}, []
for chunk in paras + body:
    for s in sentences(chunk):
        if s.lower() in seen:
            dupes.append(s)
        else:
            seen[s.lower()] = s
print(f"\n[1] dedupe: {len(seen)} unique sentences >25 chars, {len(dupes)} repeats")
for x in dupes:
    print("    repeat:", x)

# ---------------------------------------------------------------- 2. Part 1 quiz
quiz = tables[0]
items = [quiz.cell(i, 0).text for i in range(len(quiz.rows))]
print(f"\n[2] Part 1 quiz: {len(items)} sentences")
chk(len(items) == 6, "6 quiz sentences")
PAIRS = {"convenient": "convenience", "important": "importance", "significant": "significance",
         "authentic": "authenticity", "lose": "loss", "confident": "confidence"}
if WHO == "teacher":
    chk(len(tables[1].rows) == 7, "answers table = header + 6 rows")
for i, s in enumerate(items, 1):
    wrong = [k for k in PAIRS if re.search(rf"\b{k}\b", s, re.I)
             and not re.search(rf"\b{PAIRS[k]}\b", s, re.I)]
    inv = re.search(r"\b(Not only|Only by|Never)\b", s)
    chk(len(wrong) == 1 and bool(inv),
        f"item {i}: wrong form {wrong} + inversion {inv.group(0) if inv else None}")
    if WHO == "teacher":
        key = tables[1].cell(i, 1).text
        chk(wrong and PAIRS[wrong[0]] in key, f"key item {i} fixes the word form")
        chk("invert" in key or "Move" in key or "Add" in key,
            f"key item {i} explains the inversion")
for p in ("Not only", "Only by", "Never"):
    n = sum(1 for s in items if p in s)
    chk(n == 2, f"'{p}' appears twice (found {n})")

# ---------------------------------------------------------------- 3. Part 2 warm-up
warm = next(t for t in tables if len(t.rows) > 3 and t.cell(0, 0).text == "Sentence")
BANK = ["it", "is", "was", "that", "who", "what", "until", "when"]
print(f"\n[3] Part 2 warm-up: {len(warm.rows)} rows")
ans = []
for r in range(1, len(warm.rows)):
    sent, a = warm.cell(r, 0).text, warm.cell(r, 1).text
    chk(sent.count("___") == 1, f"row {r}: exactly one gap")
    if WHO == "teacher":
        chk(a.lower() in BANK, f"row {r}: answer '{a}' in word bank")
        ans.append(a.lower())
if WHO == "teacher":
    chk(sorted(set(ans)) == sorted(BANK), f"all bank words used {sorted(set(ans))}")
chk(warm.cell(0, 1).text == "Missing word", "warm-up header col2")
sent0 = [warm.cell(r, 0).text for r in range(1, len(warm.rows))]
chk(len(set(sent0)) == len(sent0), "no duplicate warm-up sentences")

# ---------------------------------------------------------------- 4. Examples / openers
print("\n[4] Part 3 examples + openers")
pair_tables = [t for t in tables if len(t.rows) == 2 and t.cell(0, 0).text == "Plain:"]
chk(len(pair_tables) == 4, f"4 conversion tables (found {len(pair_tables)})")
for t in pair_tables:
    chk(t.cell(0, 0).text == "Plain:" and t.cell(1, 0).text == "More formal:",
        f"conversion labels in order ({t.cell(0,0).text} / {t.cell(1,0).text})")
openers = [t for t in tables if len(t.rows) == 4]
chk(len(openers) == 1, "one 4-row op-eners table")
demos = [p for p in paras if p.startswith(("It is anonymity", "While many people blame",
                                           "What the school should do", "All the school has"))]
chk(len(demos) == 4, f"4 demo paragraphs ({len(demos)})")
for t, demo in zip(pair_tables, demos):
    chk(demo.startswith(t.cell(1, 1).text), f"demo opens with its formal sentence: {demo[:40]}")

# ---------------------------------------------------------------- 5. Practice sets
sets = [t for t in tables if len(t.rows) == 3 and t.cell(0, 0).text == "Simple sentence:"]
print(f"\n[5] Part 4 sets: {len(sets)}")
chk(len(sets) == 8, "8 practice sets")
for i, t in enumerate(sets, 1):
    plain, prompt, answer = t.cell(0, 1).text, t.cell(1, 1).text, t.cell(2, 1).text
    if WHO == "student":
        chk(answer.strip() == "" and t.cell(2, 0).text.strip() == "",
            f"set {i}: answer row blank + merged writing cell")
        continue
    m = re.search(r"Begin with \u2018(.+?)\u2026", prompt)
    chk(t.cell(2, 0).text == "Teacher\u2019s Answer:", f"set {i}: teacher label present")
    chk(bool(m) and answer.startswith(m.group(1).rstrip()), f"set {i}: answer follows prompt opener")
    chk(bool(re.search(r"(it (is|was)\b|what .+?\bis\b|all .+?\bis\b)", answer.lower())),
        f"set {i}: answer is a cleft")
    chk(answer != plain, f"set {i}: rewritten version differs from plain")
    chk(all(x in answer or x in plain for x in [answer.split()[0]]), f"set {i}: wording sane")

# ---------------------------------------------------------------- 6. Samples
print("\n[6] Part 5 samples")
samples = [p for p in paras if p.startswith("Sample answer")]
if WHO == "student":
    chk(not samples, "student copy has no sample answers")
else:
    chk(len(samples) == 3, "3 sample answers")
    for s in samples:
        label, bodytxt = s.split(") ", 1)
        words = len(bodytxt.split())
        rels = re.findall(r"\b(which|who|whom|whose)\b", bodytxt)
        cleft = re.findall(r"\b(It is|It was|What .{1,60}?\bis\b|All .{1,60}?\bis\b)", bodytxt)
        chk(76 <= words <= 83, f"{label}) {words} words (target ~80)")
        chk(len(rels) == 1, f"{label}) exactly one relative pronoun {rels}")
        chk(len(cleft) == 1, f"{label}) exactly one cleft {cleft}")

# ---------------------------------------------------------------- 7. no teacher text
if WHO == "student":
    bad = [x for x in paras + body
           if re.search(r"Teacher\u2019s answers|Teacher\u2019s Answer|Sample answer|teacher note",
                        x, re.I)]
    chk(not bad, f"no teacher-only text leaks ({bad[:2]})")
    lines = [p for p in paras if set(p) == {"_"}]
    chk(len(lines) == 7, f"7 writing lines ({len(lines)})")

# ---------------------------------------------------------------- 8. PDF
subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", FILE],
               check=True, capture_output=True, timeout=300)
pdf = FILE.replace(".docx", ".pdf")
out = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout
pages = out.split("\f")
real = [p for p in pages if p.strip()]
print(f"\n[7] PDF {pdf}: {len(real)} pages")
chk(len(real) == len(pages) - 1, "no blank page inside/at the end")

print("\n=== " + (f"ALL CHECKS PASSED ({WHO})" if not fail else f"{len(fail)} FAILURES ({WHO})") + " ===")
for f in fail:
    print("  -", f)
