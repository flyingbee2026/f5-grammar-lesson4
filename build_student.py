#!/usr/bin/env python3
"""Derive the STUDENT copy from the finished TEACHER docx.

Mirrors Helen's Lesson 3 student file exactly:
- drop the "Teacher's answers" label + its grey table (+ one blank paragraph)
- blank the answer column of the Part 2 warm-up table
- drop the Part 2 teacher's note
- turn each practice set's answer row into ONE merged full-width writing cell
  (two empty paragraphs, gridSpan 2, tcW 9350)
- drop the sample answers + teacher note in Part 5 and add 7 writing lines
"""
import copy

import docx
from docx.oxml.ns import qn
from lxml import etree

SRC = "Lesson 4 - Cleft sentences (Teacher's).docx"
OUT = "Lesson 4 - Cleft sentences (Student's).docx"
P = qn("w:p")
T = qn("w:tbl")


def ptext(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def drop(body, el):
    body.remove(el)


d = docx.Document(SRC)
body = d.element.body
kids = list(body.iterchildren())

# ---------------------------------------------------------------- 1. quiz answers
label = next(el for el in kids if el.tag == P and ptext(el).strip() == "Teacher\u2019s answers")
i_label = kids.index(label)
tbl_ans = kids[i_label + 1]
assert tbl_ans.tag == T
after = list(body.iterchildren())
gap = after[after.index(tbl_ans) + 1]
drop(body, label)
drop(body, tbl_ans)
if gap.tag == P and not ptext(gap).strip():
    drop(body, gap)
print("removed quiz answers block")

# ---------------------------------------------------------------- 2. helper: tables by content
def find_tables(pred):
    out = []
    for tb in d.tables:
        try:
            if pred(tb):
                out.append(tb)
        except IndexError:
            pass
    return out


# ---------------------------------------------------------------- 3. warm-up answers
warm = find_tables(lambda t: len(t.rows) > 3 and t.cell(0, 0).text == "Sentence")[0]
for r in range(1, len(warm.rows)):
    for para in warm.cell(r, 1)._tc.findall(P):
        for run in para.findall(qn("w:r")):
            para.remove(run)
print(f"blanked {len(warm.rows) - 1} warm-up answer cells")

# ---------------------------------------------------------------- 4. teacher note (Part 2)
note = next(el for el in list(body.iterchildren())
            if el.tag == P and ptext(el).startswith("Teacher\u2019s note: after"))
drop(body, note)
print("removed Part 2 teacher note")

# ---------------------------------------------------------------- 5. practice sets -> writing cells
sets = find_tables(lambda t: len(t.rows) == 3 and t.cell(0, 0).text == "Simple sentence:")
assert len(sets) == 8, len(sets)
for t in sets:
    merged = t.cell(2, 0).merge(t.cell(2, 1))
    for para in list(merged.paragraphs):
        para._p.getparent().remove(para._p)
    for _ in range(2):
        merged._tc.append(etree.Element(P))
    tcPr = merged._tc.find(qn("w:tcPr"))
    tcPr.find(qn("w:tcW")).set(qn("w:w"), "9350")
    tcPr.find(qn("w:tcW")).set(qn("w:type"), "dxa")
print(f"merged answer rows in {len(sets)} sets")

# ---------------------------------------------------------------- 6. Part 5: samples -> writing lines
paras = list(body.iterchildren())
bullet_i = next(i for i, el in enumerate(paras)
                if el.tag == P and ptext(el).startswith(" Use ONE cleft"))
blank_i = bullet_i + 1
assert paras[blank_i].tag == P and not ptext(paras[blank_i]).strip(), "expected blank after box"
lines = []
for _ in range(7):
    p = etree.Element(P)
    pPr = etree.SubElement(p, qn("w:pPr"))
    sp = etree.SubElement(pPr, qn("w:spacing"))
    sp.set(qn("w:line"), "360")
    sp.set(qn("w:lineRule"), "auto")
    r = etree.SubElement(p, qn("w:r"))
    rPr = etree.SubElement(r, qn("w:rPr"))
    rf = etree.SubElement(rPr, qn("w:rFonts"))
    rf.set(qn("w:cs"), "Times New Roman")
    t = etree.SubElement(r, qn("w:t"))
    t.text = "_" * 80
    lines.append(p)
for i, p in enumerate(lines):
    body.insert(blank_i + 1 + i, p)
# drop everything from the first sample to the end of the body (keeps sectPr)
for el in [e for e in list(body.iterchildren())
           if e.tag == P and ptext(e).startswith(("Sample answer", "Each sample uses ONE cleft"))]:
    drop(body, el)
print("added 7 writing lines, removed 3 samples + note")

d.save(OUT)
print("wrote", OUT)
