#!/usr/bin/env python3
"""Build teacher version from Helen's student docx — robust content-based detection."""
import copy
import docx
from docx.oxml.ns import qn
from lxml import etree

SRC = "/root/.hermes/cache/documents/native-inputs/7adb7363de13e0671058cae55f4b0e2ae92c2f7f09bccfdde0078e4b65ec82c8/doc_77224d23d005_Lesson 4 - Cleft sentences (Student's).docx"
OUT = "/root/f5-grammar-lesson4/Lesson 4 - Cleft sentences (Teacher's).docx"

d = docx.Document(SRC)
body = d.element.body

def ptext(el):
    if el.tag == qn("w:p"):
        return "".join(t.text or "" for t in el.iter(qn("w:t")))
    return ""

def is_underscore_para(el):
    if el.tag != qn("w:p"):
        return False
    text = ptext(el).strip()
    return len(text) > 5 and all(c == "_" or c.isspace() for c in text)

def is_empty_para(el):
    if el.tag != qn("w:p"):
        return False
    return ptext(el).strip() == ""

def make_para(text, bold=False):
    p = etree.Element(qn("w:p"))
    r = etree.SubElement(p, qn("w:r"))
    rPr = etree.SubElement(r, qn("w:rPr"))
    rf = etree.SubElement(rPr, qn("w:rFonts"))
    rf.set(qn("w:cs"), "Times New Roman")
    if bold:
        etree.SubElement(rPr, qn("w:b"))
        etree.SubElement(rPr, qn("w:bCs"))
    t = etree.SubElement(r, qn("w:t"))
    t.text = text
    t.set(qn("xml:space"), "preserve")
    return p

# Answers
quiz_key = [
    ("1A", "Not only do teenagers order food instead of cooking... (Move 'do' before the subject — 'Not only' inverts.)"),
    ("2A", "They sit for hours on apps, which were designed to save time... ('that' cannot be used after a comma in a non-restrictive clause; use 'which'.)"),
    ("3A", "The convenience of instant delivery has destroyed any reason to walk to the shops. ('convenient' → 'convenience' — a noun is needed after 'the'.)"),
    ("4A", "Rarely do parents notice that their children have forgotten how to hold a proper conversation. (Move 'do' before the subject — 'Rarely' inverts.)"),
    ("5A", "Only by encouraging teenagers to put down their phones can the family have a real conversation. (Move 'can' before the subject — 'Only by' inverts.)"),
    ("1B", "...for whom technology was not available. ('who' → 'whom' — the object form of the relative pronoun is needed after a preposition.)"),
    ("2B", "The belief that real work must involve physical sweat ignores the fact... ('believe' → 'belief' — a noun is needed after 'the'.)"),
    ("3B", "Seldom do critics acknowledge that a teenager editing videos... (Move 'do' before the subject — 'Seldom' inverts.)"),
    ("4B", "The reason why teenagers seem distracted is not laziness but the creative work... ('lazy' → 'laziness' — a noun is needed after 'not'.)"),
    ("5B", "Only when adults look at the results of their work can they understand... (Move 'can' before the subject — 'Only when' inverts.)"),
]

part_a_answers = [
    "Teacher's Answer: It is the proliferation of smartphones that has altered interpersonal communication patterns.",
    "Teacher's Answer: It is profit that tech companies prioritize over user privacy in the digital age.",
    "Teacher's Answer: It is short-form video apps that Gen Z teenagers spend most of their leisure hours on.",
    "Teacher's Answer: It was generative AI tools that created widespread concern among secondary educators last year.",
    "Teacher's Answer: It is online anonymity that encourages cyberbullies to post malicious comments.",
]

part_b_answers = [
    "Teacher's Answer: What damages sleep quality and mental well-being is excessive screen time.",
    "Teacher's Answer: What the government should do is regulate misleading targeted advertisements on social media.",
    "Teacher's Answer: Those who desire instant validation through likes and comments are modern teenagers.",
    "Teacher's Answer: All (that) social media platforms aim to do is keep users scrolling as long as possible.",
    "Teacher's Answer: What cyberbullies do is hide behind anonymous profile pictures to cause harm.",
]

part_c_answers = [
    "Sample transformation: It is standard messaging apps that lack genuine emotional warmth.",
    "Sample transformation: What schools need is critical digital literacy education, not total phone bans.",
    "Sample transformation: It is financial profit that social media algorithms exploit human psychology for.",
    "Sample transformation: All users require is a conscious boundaries setting to regain control over their time.",
]

popular_culture = [
    "",
    "Teacher's reference: Popular Culture Examples",
    "",
    "It-cleft examples from popular culture:",
    "• \"It is not the strength of the body, but the strength of the spirit that counts.\" (The Lord of the Rings)",
    "• \"It is not our abilities that show what we truly are. It is our choices.\" (Harry Potter and the Chamber of Secrets)",
    "• \"It was the best of times, it was the worst of times.\" (Charles Dickens, A Tale of Two Cities)",
    "",
    "Wh-cleft examples from popular culture:",
    "• \"What doesn't kill you makes you stronger.\" (Kelly Clarkson / Nietzsche)",
    "• \"What we do in life echoes in eternity.\" (Gladiator)",
    "• \"What the world needs now is love, sweet love.\" (Burt Bacharach / Jackie DeShannon)",
    "",
    "All-cleft examples from popular culture:",
    "• \"All I want for Christmas is you.\" (Mariah Carey) — All [that] I want for Christmas is you.",
    "• \"All we have to decide is what to do with the time that is given us.\" (Gandalf, The Lord of the Rings)",
    "• \"All that glitters is not gold.\" (Shakespeare / J.R.R. Tolkien)",
    "",
]

# ========================================================================
# Step 1: Insert teacher answers after Part 1 table
# ========================================================================
children = list(body.iterchildren())
part1_table_idx = None
for i, el in enumerate(children):
    if el.tag == qn("w:tbl"):
        # Check if previous paragraph mentions proofreading
        if i > 0:
            prev_text = ptext(children[i-1])
            if "errors" in prev_text.lower() and "proofread" in prev_text.lower():
                part1_table_idx = i
                break

if part1_table_idx is not None:
    ref = children[part1_table_idx]
    ref_pos = list(body).index(ref)
    body.insert(ref_pos + 1, make_para(""))
    body.insert(ref_pos + 2, make_para("Teacher's answers", bold=True))
    offset = 3
    for num, corr in quiz_key:
        body.insert(ref_pos + offset, make_para(f"{num}. {corr}"))
        offset += 1
    print(f"Inserted Part 1 answers after table at index {part1_table_idx}")
else:
    print("WARNING: Could not find Part 1 table")

# ========================================================================
# Step 2: Insert popular culture after Part 2 structure table
# ========================================================================
children = list(body.iterchildren())
structure_table_idx = None
for i, el in enumerate(children):
    if el.tag == qn("w:tbl") and i > part1_table_idx:
        # Check first cell for "Type"
        try:
            rows = el.findall(qn("w:tr"))
            if rows:
                cells = rows[0].findall(qn("w:tc"))
                if cells:
                    first_cell_text = ptext(cells[0].find(qn("w:p")))
                    if "Type" in first_cell_text:
                        structure_table_idx = i
                        break
        except Exception:
            pass

if structure_table_idx is not None:
    ref = children[structure_table_idx]
    ref_pos = list(body).index(ref)
    for line in reversed(popular_culture):
        body.insert(ref_pos + 1, make_para(line, bold=("Teacher's reference" in line)))
    print(f"Inserted popular culture after structure table at index {structure_table_idx}")
else:
    print("WARNING: Could not find Part 2 structure table")

# ========================================================================
# Step 3: Replace blank pairs in Part 3A with answers
# ========================================================================
children = list(body.iterchildren())
part_a_start = None
part_b_start = None
for i, el in enumerate(children):
    text = ptext(el)
    if "Practice Sets - Part A" in text or ("It-Clefts" in text and "Rewrite" in text):
        part_a_start = i
    elif "Part B: Form Wh-Clefts" in text or ("Wh-Clefts" in text and "All-Clefts" in text):
        part_b_start = i

if part_a_start and part_b_start:
    # Collect blank pairs
    blanks = []
    for i in range(part_a_start, part_b_start):
        if is_underscore_para(children[i]):
            blanks.append(i)
    
    # Group into consecutive pairs
    pairs = []
    i = 0
    while i < len(blanks):
        if i + 1 < len(blanks) and blanks[i+1] == blanks[i] + 1:
            pairs.append((blanks[i], blanks[i+1]))
            i += 2
        else:
            i += 1
    
    print(f"Part A: found {len(pairs)} blank pairs")
    
    # Replace from highest index to lowest
    for idx in range(len(pairs) - 1, -1, -1):
        if idx < len(part_a_answers):
            b1, b2 = pairs[idx]
            c = list(body.iterchildren())
            el2 = c[b2]
            el1 = c[b1]
            body.remove(el2)
            c = list(body.iterchildren())
            el1 = c[b1]
            body.remove(el1)
            c = list(body.iterchildren())
            ref_pos = b1
            if ref_pos < len(c):
                ref = c[ref_pos]
                insert_pos = list(body).index(ref)
            else:
                insert_pos = len(list(body.iterchildren())) - 1
            body.insert(insert_pos, make_para(part_a_answers[idx]))

# ========================================================================
# Step 4: Replace blank pairs in Part 3B with answers
# ========================================================================
children = list(body.iterchildren())
part_b_start = None
part_c_start = None
for i, el in enumerate(children):
    text = ptext(el)
    if "Part B: Form Wh-Clefts" in text or ("Wh-Clefts" in text and "All-Clefts" in text):
        part_b_start = i
    elif "Part C: DSE Proofreading" in text or ("DSE Proofreading" in text and "Sentence Transformation" in text):
        part_c_start = i

if part_b_start and part_c_start:
    blanks = []
    for i in range(part_b_start, part_c_start):
        if is_underscore_para(children[i]):
            blanks.append(i)
    
    pairs = []
    i = 0
    while i < len(blanks):
        if i + 1 < len(blanks) and blanks[i+1] == blanks[i] + 1:
            pairs.append((blanks[i], blanks[i+1]))
            i += 2
        else:
            i += 1
    
    print(f"Part B: found {len(pairs)} blank pairs")
    
    for idx in range(len(pairs) - 1, -1, -1):
        if idx < len(part_b_answers):
            b1, b2 = pairs[idx]
            c = list(body.iterchildren())
            el2 = c[b2]
            el1 = c[b1]
            body.remove(el2)
            c = list(body.iterchildren())
            el1 = c[b1]
            body.remove(el1)
            c = list(body.iterchildren())
            ref_pos = b1
            if ref_pos < len(c):
                ref = c[ref_pos]
                insert_pos = list(body).index(ref)
            else:
                insert_pos = len(list(body.iterchildren())) - 1
            body.insert(insert_pos, make_para(part_b_answers[idx]))

# ========================================================================
# Step 5: Part 3C — remove blanks, add sample transformations after table
# ========================================================================
children = list(body.iterchildren())
part_c_start = None
for i, el in enumerate(children):
    text = ptext(el)
    if "Part C: DSE Proofreading" in text or ("DSE Proofreading" in text and "Sentence Transformation" in text):
        part_c_start = i
        break

if part_c_start:
    # Find the table in Part C (the draft paragraph table)
    draft_table_idx = None
    for i in range(part_c_start, min(part_c_start + 5, len(children))):
        if children[i].tag == qn("w:tbl"):
            draft_table_idx = i
            break
    
    if draft_table_idx:
        # Remove all blanks after the table
        c = list(body.iterchildren())
        to_remove = []
        for i in range(draft_table_idx + 1, len(c)):
            if is_underscore_para(c[i]) or is_empty_para(c[i]):
                to_remove.append(c[i])
        
        for el in reversed(to_remove):
            body.remove(el)
        
        # Insert sample transformations after the table
        c = list(body.iterchildren())
        ref = c[draft_table_idx]
        ref_pos = list(body).index(ref)
        body.insert(ref_pos + 1, make_para(""))
        body.insert(ref_pos + 2, make_para("Sample transformations:", bold=True))
        offset = 3
        for ans in part_c_answers:
            body.insert(ref_pos + offset, make_para(ans))
            offset += 1
        
        print(f"Part C: inserted {len(part_c_answers)} sample transformations")

# Save
d.save(OUT)
print(f"Saved teacher version to {OUT}")
