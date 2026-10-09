#!/usr/bin/env python3
"""Build F5 Grammar Lesson 4: Cleft sentences (teacher version).

Template = Helen's finalised Lesson 3 TEACHER docx (keeps styles.xml: Heading 1/2/3,
Table Grid, header fonts, numbering). Every element is a deep copy of the matching
element in her Lesson 3 file, so borders / shading / indents are identical.
"""
import copy
import sys

import docx
from docx.oxml.ns import qn
from lxml import etree

MODEL = "lesson3-teacher.docx"
OUT = {"teacher": "Lesson 4 - Cleft sentences (Teacher's).docx"}
XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

# ---------------------------------------------------------------- model indices
# (from the Lesson 3 teacher file dump)
H_SCHOOL, H_SUBJ, H_TITLE, H_TOPIC = 0, 1, 2, 3
P_H2 = 4            # Heading 2 part heading
P_PLAIN = 5         # plain Normal body paragraph
T_QUIZ = 6          # 5x1 numbered quiz table
P_BLANK = 7         # empty Normal paragraph
P_ANS_LABEL = 8     # bold 'Teacher's answers'
T_ANSWERS = 9       # 6x2 F2F2F2 answers table
P_H2B = 11
P_INSTR = 12        # plain instruction line
T_WORDS = 13        # 4-col warm-up table
P_H3 = 17           # Heading 3 example heading (single run)
T_PAIR = 18         # 2x2 DEEBF7 Plain / More formal table
P_BOX = 19          # bordered demo paragraph
P_H3_LINK = 30      # 'Nominalisation as a linking device'
P_PLAIN2 = 31
T_LINK = 32         # 2x2 DEEBF7 linking table
P_H1 = 35           # Heading 1 'Part 3: Practice sets'
P_SETINTRO = 36
P_SETTITLE = 37     # Heading 3 set title (single run)
T_SET = 38          # 3x2 set table
P_H1W = 53          # Heading 1 'Part 4: Writing Task'
P_INSTR_B = 54      # lavender bold 'Instructions:'
P_INSTR_1 = 55      # lavender task line
P_INSTR_2 = 56      # lavender bullet line (wingdings sym + br)
P_SAMPLE = 58       # grey sample box: bold label run + body run
P_NOTE = 61         # grey teacher note paragraph

# ---------------------------------------------------------------- content
SCHOOL = "St. Joseph\u2019s Anglo-Chinese School"
SUBJ = "F.5 English Language"
TITLE = "Grammar \u2013 Lesson 4: Cleft sentences"
TOPIC = "Cleft sentences"

QUIZ_INTRO = ("Each sentence contains TWO errors: one in the form of a word (noun or "
              "adjective) and one in an inverted structure. Proofread the sentences.")
QUIZ = [
    "Not only the new canteen is convenient for students, but the convenient of its "
    "online ordering system also attracts busy teachers.",
    "Only by eating breakfast students can keep their energy up, yet many teenagers do not "
    "realise the important of a proper diet.",
    "Never the school has displayed its old photographs in the hall, although the "
    "significant of that history is obvious to former students.",
    "Only by showing customers how their products are made shops can prove the authentic "
    "of their products.",
    "Not only losing sleep affects students\u2019 test results, but the lose of "
    "concentration in class is also a serious problem.",
    "Never students should feel afraid of making mistakes when they speak English, because "
    "the confident to speak comes from practice.",
]
QUIZ_KEY = [
    ("Not only is the new canteen convenient for students, but the convenience of its "
     "online ordering system also attracts busy teachers. (Move is before the subject "
     "\u2014 \u2018Not only\u2019 inverts the subject and the verb; convenient \u2192 "
     "convenience \u2014 a noun is needed after \u2018the\u2019.)"),
    ("Only by eating breakfast can students keep their energy up, yet many teenagers do not "
     "realise the importance of a proper diet. (Move can before students \u2014 \u2018Only "
     "by\u2019 inverts; important \u2192 importance \u2014 the noun follows \u2018the\u2019.)"),
    ("Never has the school displayed its old photographs in the hall, although the "
     "significance of that history is obvious to former students. (Move has before the "
     "school \u2014 \u2018Never\u2019 inverts; significant \u2192 significance \u2014 a noun "
     "is needed after \u2018the\u2019.)"),
    ("Only by showing customers how their products are made can shops prove the "
     "authenticity of their products. (Move can before shops \u2014 \u2018Only by\u2019 "
     "inverts; authentic \u2192 authenticity \u2014 the noun follows \u2018the\u2019.)"),
    ("Not only does losing sleep affect students\u2019 test results, but the loss of "
     "concentration in class is also a serious problem. (Add does before the subject and use "
     "the base verb affect \u2014 \u2018Not only\u2019 inverts; lose \u2192 loss \u2014 a noun "
     "is needed after \u2018the\u2019.)"),
    ("Never should students feel afraid of making mistakes when they speak English, because "
     "the confidence to speak comes from practice. (Move should before students \u2014 "
     "\u2018Never\u2019 inverts; confident \u2192 confidence \u2014 the noun follows "
     "\u2018the\u2019.)"),
]

WARM_INTRO = ("Complete each cleft sentence with ONE word "
              "(it / is / was / that / who / what / until / when).")
WARM = [
    ("___ is peer pressure that pushes teenagers to join in online bullying.", "It"),
    ("It ___ the promise of a quick laugh that makes cruel posts so tempting to share.", "is"),
    ("It was the school newspaper ___ first reported the story.", "that"),
    ("___ drives a teenager to keep scrolling is the fear of missing out.", "What"),
    ("It is teachers ___ first notice the change in a student\u2019s behaviour.", "who"),
    ("All the school needs to do ___ train staff to spot the warning signs.", "is"),
    ("___ many parents fail to realise is how quickly cruel posts spread.", "What"),
    ("It ___ the lack of clear rules that allowed the rumours to spread.", "was"),
    ("It was not ___ parents complained to the principal that the bullying was investigated.",
     "until"),
    ("What worries counsellors most ___ the silence of the students who see it happen.", "is"),
    ("It is the victims of online bullying ___ often suffer in silence.", "who"),
    ("___ is not the phone itself, but the endless notifications, that distract students.", "It"),
    ("All it takes ___ five minutes of checking before you share a post.", "is"),
    ("It was only ___ the school posted the facts that the rumours stopped.", "when"),
    ("___ struck the teachers most was how fast a single post can reach the whole year group.",
     "What"),
    ("It is careless sharing, rather than social media itself, ___ spreads rumours so fast.",
     "that"),
    ("What the panel recommends ___ a clearer policy on reporting abuse.", "is"),
    ("It is one trusted adult, more than any new policy, ___ makes the biggest "
     "difference to a victim.", "that"),
    ("___ matters most is that victims know someone is on their side.", "What"),
    ("It is the way a school responds ___ decides whether victims ever report again.", "that"),
]
WARM_NOTE = ("Teacher\u2019s note: after \u2018It is / It was\u2019, use \u2018that\u2019 for things "
             "and \u2018who\u2019 for people; \u2018which\u2019 is possible for things but less "
             "common. Accept either word from students in items 3, 5, 11, 16, 18 and 20.")

# Part 3 -- intro (mixed runs: bold+underline on the two examples, like Lesson 3)
P3_INTRO = [
    ("A cleft sentence takes one idea and puts it into two parts, so that the most "
     "important information stands alone: ", "plain"),
    ("Peer pressure causes online bullying", "bold_u"),
    (" becomes ", "plain"),
    ("It is peer pressure that causes online bullying", "bold_u"),
    (" or ", "plain"),
    ("What causes online bullying is peer pressure", "bold_u"),
    (". The word order is broken on purpose, so the emphasis lands exactly where you want "
     "it. In formal writing, one or two clefts show the marker that you can control sentence "
     "patterns beyond the ordinary subject\u2013verb order, and they can also carry an idea "
     "from one sentence into the next.", "plain"),
]

EXAMPLES = [
    ("Example 1: It-cleft \u2013 focusing on the cause",
     "People post cruel comments because they are anonymous.",
     "It is anonymity that encourages people to post cruel comments.",
     "It is anonymity that encourages people to post cruel comments. Behind a screen, nobody "
     "has to face the person they are hurting, so the usual restraint of a face-to-face "
     "conversation disappears. Asking students to use their own names on the school forum "
     "would remove that cover entirely."),
    ("Example 2: It-cleft \u2013 setting up a contrast",
     "Many people blame violent games, but the silence of bystanders does the real damage.",
     "While many people blame violent games, it is the silence of bystanders that does the "
     "real damage.",
     "While many people blame violent games, it is the silence of bystanders that does the "
     "real damage. A cruel post can stay online for days because nobody bothers to report "
     "it. Teaching students to speak up matters more than banning one more game."),
    ("Example 3: Wh-cleft \u2013 presenting a recommendation",
     "The school should make it easier for victims to report bullying.",
     "What the school should do is make it easier for victims to report bullying.",
     "What the school should do is make it easier for victims to report bullying. One "
     "anonymous form on the school website would take an afternoon to set up. Once students "
     "know that a report will be read, the silence that protects bullies begins to "
     "disappear."),
    ("Example 4: All-cleft \u2013 naming the only thing needed",
     "The school only has to post one clear rule about online behaviour.",
     "All the school has to do is post one clear rule about online behaviour.",
     "All the school has to do is post one clear rule about online behaviour. A single "
     "sentence in the student handbook would tell everyone what is expected of them. Such a "
     "small step would show victims that the school takes the problem seriously, and more of "
     "them would come forward."),
]

OPENERS_TITLE = "Clefts as paragraph openers"
OPENERS_INTRO = ("Clefts are also useful at the start of a paragraph, an article or a speech: "
                 "they decide what the reader notices first. Put the it-cleft in your thesis "
                 "statement, the wh-cleft in your recommendation, and the all-cleft in your "
                 "closing line.")
OPENERS = [
    ("Plain:", "Social media affects teenagers\u2019 self-esteem very strongly."),
    ("Cleft opener (What \u2026 is \u2026):",
     "What many adults fail to realise is how strongly social media affects teenagers\u2019 "
     "self-esteem."),
    ("Plain:", "Most teenagers only need one adult who will listen without judging them."),
    ("Cleft opener (All \u2026 is \u2026):",
     "All most teenagers need is one adult who will listen without judging them."),
]

SETS_INTRO = ("Rewrite each simple sentence as a cleft sentence. The sentences are about life "
              "online, so keep the register formal. The later sets (4\u20138) are more "
              "challenging.")
SETS = [
    ("Example Set 1 (It-cleft \u2013 the cause)",
     "Teenagers copy the cruel comments of their friends because they are afraid of being "
     "left out.",
     "Begin with \u2018It is the fear of being left out \u2026\u2019.",
     "It is the fear of being left out that leads teenagers to copy the cruel comments of "
     "their friends."),
    ("Example Set 2 (Wh-cleft \u2013 what is needed)",
     "Victims of online bullying need an adult they can trust.",
     "Begin with \u2018What victims of online bullying need \u2026\u2019.",
     "What victims of online bullying need is an adult they can trust."),
    ("Example Set 3 (It-cleft \u2013 the person responsible)",
     "School counsellors, not the discipline team, should handle bullying cases.",
     "Begin with \u2018It is school counsellors \u2026\u2019 and use \u2018who\u2019.",
     "It is school counsellors, not the discipline team, who should handle bullying cases."),
    ("Example Set 4 (It-cleft \u2013 the contrast)",
     "Teachers blame social media for the bullying, but the school\u2019s slow response makes "
     "it worse.",
     "Begin with \u2018While teachers blame social media for the bullying \u2026\u2019.",
     "While teachers blame social media for the bullying, it is the school\u2019s slow "
     "response that makes it worse."),
    ("Example Set 5 (Wh-cleft \u2013 the recommendation)",
     "The government should make platforms remove cruel posts within 24 hours.",
     "Begin with \u2018What the government should do \u2026\u2019.",
     "What the government should do is make platforms remove cruel posts within 24 hours."),
    ("Example Set 6 (All-cleft \u2013 the only thing needed)",
     "To keep victims safe, the school only has to provide a safe place for them at break "
     "time.",
     "Begin with \u2018All the school needs to do \u2026\u2019.",
     "All the school needs to do is provide a safe place for victims at break time."),
    ("Example Set 7 (Wh-cleft \u2013 the hook)",
     "Adults say teenagers should just ignore cruel comments, but that is impossible.",
     "Begin with \u2018What adults fail to understand \u2026\u2019.",
     "What adults fail to understand is that a cruel comment can follow a teenager "
     "everywhere."),
    ("Example Set 8 (It-cleft \u2013 the turning point)",
     "The school only understood how serious the bullying was after a student stayed away.",
     "Begin with \u2018It was not until \u2026\u2019.",
     "It was not until a student stayed away that the school understood how serious the "
     "bullying was."),
]

TASK_TITLE = "Part 5: Writing Task"
TASK_LINE = ("1. Write ONE paragraph (about 80 words) explaining why online bullying is hard "
             "for teenagers to escape, and suggesting ONE thing schools could do to help.")
TASK_BULLETS = [
    ("\u2022", None),
    (" Use ", "plain"),
    ("ONE cleft sentence", "bold_u"),
    (" (It is \u2026 that \u2026 / What \u2026 is \u2026 / All \u2026 needs to do is \u2026).",
     "plain"),
    (None, "br"),
    ("\u2022", None),
    (" Use a relative clause or an inverted pattern e.g. Not only \u2026 / Never \u2026 / "
     "Rarely \u2026", "plain"),
]
SAMPLES = [
    ("Sample answer 1 (It-cleft) ",
     "Online bullying is difficult to escape because it follows teenagers home: cruel "
     "messages appear on their phones long after the school gates have closed. It is this "
     "constant presence that makes even a short comment so damaging, as there is no safe "
     "hour in the day. Schools should teach students how to block and report cruel messages, "
     "which protects every student in the school. A form teacher could cover the basics in "
     "a single lesson, and the advice would last for years."),
    ("Sample answer 2 (Wh-cleft) ",
     "What many victims of online bullying fear most is that adults will simply tell them to "
     "switch off their phones. That advice, which is usually meant kindly, leaves them alone "
     "with the abuse. Schools should therefore run regular lessons on blocking and reporting "
     "cruel messages, and make sure every student knows exactly where to turn for help. "
     "Silence, in the end, is a bully\u2019s best protection, and a short lesson in form time "
     "can begin to break it."),
    ("Sample answer 3 (All-cleft) ",
     "Escaping online bullying is hard because the abuse continues after school, and many "
     "victims never tell anyone about it. All schools really need to do is give students one "
     "trusted adult and one simple way to report cruel messages. This small change, which "
     "costs almost nothing, would help thousands of teenagers currently suffering in "
     "silence. A tutor-group discussion once a month would keep the topic alive, and victims "
     "would know that the school is on their side."),
]
SAMPLE_NOTE = ("Each sample uses ONE cleft: Sample 1 an it-cleft (It is \u2026 that \u2026), "
               "Sample 2 a wh-cleft (What \u2026 is \u2026) and Sample 3 an all-cleft (All "
               "\u2026 need to do is \u2026). Each sample also contains ONE relative clause "
               "(\u2018which\u2019 in every case) \u2014 point these out to students before "
               "they write.")


# ---------------------------------------------------------------- helpers
class Builder:
    def __init__(self):
        self.model = docx.Document(MODEL)
        self.M = list(self.model.element.body.iterchildren())
        self.doc = docx.Document(MODEL)
        body = self.doc.element.body
        self.sectPr = body.find(qn("w:sectPr"))
        for child in list(body.iterchildren()):
            if child is not self.sectPr:
                body.remove(child)
        self.body = body

    def clone(self, idx):
        return copy.deepcopy(self.M[idx])

    def push(self, el):
        self.body.insert(list(self.body).index(self.sectPr), el)
        return el

    # -- run helpers ---------------------------------------------------
    @staticmethod
    def _rpr(fmt):
        rPr = etree.Element(qn("w:rPr"))
        rf = etree.SubElement(rPr, qn("w:rFonts"))
        rf.set(qn("w:cs"), "Times New Roman")
        if fmt in ("bold", "bold_u"):
            etree.SubElement(rPr, qn("w:b"))
            etree.SubElement(rPr, qn("w:bCs"))
        if fmt == "italic":
            etree.SubElement(rPr, qn("w:i"))
        if fmt == "bold_u":
            u = etree.SubElement(rPr, qn("w:u"))
            u.set(qn("w:val"), "single")
        return rPr

    def set_text(self, p, text):
        """Replace a paragraph's text, keeping its FIRST run's formatting."""
        runs = p.findall(qn("w:r"))
        assert runs, "paragraph has no run to copy"
        for r in runs[1:]:
            p.remove(r)
        r = runs[0]
        for child in list(r):
            if child.tag != qn("w:rPr"):
                r.remove(child)
        t = etree.SubElement(r, qn("w:t"))
        t.text = text
        t.set(XMLSPACE, "preserve")

    def set_runs(self, p, parts):
        """Rebuild a paragraph's runs from (text, fmt|'br'|None-symbol) tuples."""
        for r in p.findall(qn("w:r")):
            p.remove(r)
        for text, fmt in parts:
            r = etree.SubElement(p, qn("w:r"))
            r.append(self._rpr("plain" if fmt is None or fmt == "br" else fmt))
            if fmt == "br":
                etree.SubElement(r, qn("w:br"))
                continue
            if fmt is None:      # wingdings bullet
                sym = etree.SubElement(r, qn("w:sym"))
                sym.set(qn("w:font"), "Wingdings")
                sym.set(qn("w:char"), "F0E0")
                continue
            t = etree.SubElement(r, qn("w:t"))
            t.text = text
            t.set(XMLSPACE, "preserve")
        return p

    # -- table helpers -------------------------------------------------
    def cell(self, tbl, r, c):
        tr = tbl.findall(qn("w:tr"))[r]
        tc = tr.findall(qn("w:tc"))[c]
        return tc.find(qn("w:p"))

    def set_cell(self, tbl, r, c, text):
        self.set_text(self.cell(tbl, r, c), text)

    def shade(self, tbl, r, c, fill):
        tc = tbl.findall(qn("w:tr"))[r].findall(qn("w:tc"))[c]
        tcPr = tc.find(qn("w:tcPr"))
        for old in tcPr.findall(qn("w:shd")):
            tcPr.remove(old)
        shd = etree.SubElement(tcPr, qn("w:shd"))
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)

    @staticmethod
    def widths(tbl, widths):
        grid = tbl.find(qn("w:tblGrid"))
        gcs = grid.findall(qn("w:gridCol"))
        for gc, w in zip(gcs, widths):
            gc.set(qn("w:w"), str(w))
        for tr in tbl.findall(qn("w:tr")):
            for tc, w in zip(tr.findall(qn("w:tc")), widths):
                tcPr = tc.find(qn("w:tcPr"))
                tcW = tcPr.find(qn("w:tcW"))
                if tcW is None:
                    tcW = etree.SubElement(tcPr, qn("w:tcW"))
                tcW.set(qn("w:w"), str(w))
                tcW.set(qn("w:type"), "dxa")

    def add_row(self, tbl, model_row_idx):
        tbl.append(copy.deepcopy(tbl.findall(qn("w:tr"))[model_row_idx]))


def build(variant="teacher"):
    if variant != "teacher":
        raise SystemExit("student copy is derived from Helen's finalised teacher docx, "
                         "not from this builder (see skill: f5-grammar-lessons)")
    b = Builder()

    # -- header block
    for idx, text in ((H_SCHOOL, SCHOOL), (H_SUBJ, SUBJ), (H_TITLE, TITLE), (H_TOPIC, TOPIC)):
        el = b.clone(idx)
        b.set_text(el, text)
        b.push(el)

    # -- Part 1
    el = b.clone(P_H2); b.set_text(el, "Part 1: Proofreading quiz"); b.push(el)
    el = b.clone(P_PLAIN); b.set_text(el, QUIZ_INTRO); b.push(el)
    qt = b.clone(T_QUIZ)
    while len(qt.findall(qn("w:tr"))) < len(QUIZ):
        b.add_row(qt, 0)
    for i, s in enumerate(QUIZ):
        b.set_cell(qt, i, 0, s)
    for r in qt.findall(qn("w:tr"))[len(QUIZ):]:
        qt.remove(r)
    b.push(qt)
    b.push(b.clone(P_BLANK))
    el = b.clone(P_ANS_LABEL); b.set_text(el, "Teacher\u2019s answers"); b.push(el)
    at = b.clone(T_ANSWERS)
    while len(at.findall(qn("w:tr"))) < len(QUIZ_KEY) + 1:
        b.add_row(at, 1)
    for i, corr in enumerate(QUIZ_KEY, start=1):
        b.set_cell(at, i, 0, str(i))
        b.set_cell(at, i, 1, corr)
    b.push(at)
    b.push(b.clone(P_BLANK))

    # -- Part 2
    el = b.clone(P_H2B); b.set_text(el, "Part 2: Building a cleft sentence"); b.push(el)
    el = b.clone(P_INSTR); b.set_text(el, WARM_INTRO); b.push(el)
    wt = b.clone(T_WORDS)
    rows = wt.findall(qn("w:tr"))
    for tr in rows:                      # 4 columns -> 2 columns
        tcs = tr.findall(qn("w:tc"))
        tr.remove(tcs[3]); tr.remove(tcs[2])
    grid = wt.find(qn("w:tblGrid"))
    gcs = grid.findall(qn("w:gridCol"))
    grid.remove(gcs[3]); grid.remove(gcs[2])
    base = len(rows)                     # header + 10 data rows
    for _ in range(20 - (base - 1)):
        b.add_row(wt, 1)
    for r in wt.findall(qn("w:tr"))[21:]:
        wt.remove(r)
    b.widths(wt, (7350, 2000))
    b.set_cell(wt, 0, 0, "Sentence")
    b.set_cell(wt, 0, 1, "Missing word")
    for i, (sent, ans) in enumerate(WARM, start=1):
        b.set_cell(wt, i, 0, f"{i}. {sent}")
        b.set_cell(wt, i, 1, ans)
    b.push(wt)
    el = b.clone(P_NOTE); b.set_text(el, WARM_NOTE); b.push(el)

    # -- Part 3
    el = b.clone(P_H2B); b.set_text(el, "Part 3: Cleft sentences in formal writing"); b.push(el)
    el = b.clone(P_PLAIN2); b.set_runs(el, P3_INTRO); b.push(el)
    for title, plain, formal, demo in EXAMPLES:
        el = b.clone(P_H3); b.set_text(el, title); b.push(el)
        t = b.clone(T_PAIR)
        b.set_cell(t, 0, 0, "Plain:"); b.set_cell(t, 0, 1, plain)
        b.set_cell(t, 1, 0, "More formal:"); b.set_cell(t, 1, 1, formal)
        b.push(t)
        el = b.clone(P_BOX); b.set_text(el, demo); b.push(el)
    el = b.clone(P_H3_LINK); b.set_text(el, OPENERS_TITLE); b.push(el)
    el = b.clone(P_PLAIN); b.set_text(el, OPENERS_INTRO); b.push(el)
    ot = b.clone(T_LINK)
    b.add_row(ot, 0); b.add_row(ot, 1)
    for i, (label, text) in enumerate(OPENERS):
        b.set_cell(ot, i, 0, label); b.set_cell(ot, i, 1, text)
    b.push(ot)
    b.push(b.clone(P_BLANK)); b.push(b.clone(P_BLANK))

    # -- Part 4
    el = b.clone(P_H1); b.set_text(el, "Part 4: Practice sets"); b.push(el)
    el = b.clone(P_SETINTRO); b.set_text(el, SETS_INTRO); b.push(el)
    for title, simple, prompt, answer in SETS:
        el = b.clone(P_SETTITLE); b.set_text(el, title); b.push(el)
        t = b.clone(T_SET)
        b.set_cell(t, 0, 1, simple)
        b.set_cell(t, 1, 1, prompt)
        if variant == "teacher":
            b.set_cell(t, 2, 1, answer)
        else:
            b.set_cell(t, 2, 0, "")
            b.set_cell(t, 2, 1, "")
        b.push(t)

    # -- Part 5
    el = b.clone(P_H1W); b.set_text(el, TASK_TITLE); b.push(el)
    el = b.clone(P_INSTR_B); b.set_text(el, "Instructions:"); b.push(el)
    el = b.clone(P_INSTR_1); b.set_text(el, TASK_LINE); b.push(el)
    el = b.clone(P_INSTR_2); b.set_runs(el, TASK_BULLETS); b.push(el)
    if variant == "teacher":
        b.push(b.clone(P_BLANK))
        for label, body in SAMPLES:
            el = b.clone(P_SAMPLE)
            b.set_runs(el, [(label, "bold"), (body, "plain")])
            b.push(el)
        el = b.clone(P_NOTE); b.set_text(el, SAMPLE_NOTE); b.push(el)

    out = OUT[variant]
    b.doc.save(out)
    print("wrote", out)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "teacher")
