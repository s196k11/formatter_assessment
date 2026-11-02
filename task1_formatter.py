#!/usr/bin/env python3
"""
task1_formatter.py - turn a JSON file of Task 1 questions + sample answers into a Word document
that uses YOUR Task 1 template (new_task_1_template_version_10_7.docx) and sample-pack colours.

    python task1_formatter.py questions.json                       # -> questions.docx
    python task1_formatter.py questions.json -o Task1.docx -t template.docx
    python task1_formatter.py questions.json --check                # validate the JSON only
    python task1_formatter.py --example > example.json              # print a full example input

Needs only:  pip install python-docx lxml

COLOUR RULES (taken from the FWPCOT2275 sample pack)
    * pre-filled text (printed for the student)  -> BLACK text on PINK cell (D9A1B7)
    * sample answer  (what the student writes)   -> RED text (EE0000), NO shading

JSON SHAPE
{
  "unit_code": "FWPCOT2275",
  "unit_title": "Fell trees manually (basic)",
  "title": "optional heading text",                     # default: "Task 1: Knowledge Assessment - <code> <title>"
  "questions": [ <question>, <question>, ... ]
}

Every question has  "type"  plus  "chapter" / "section"  (for the "Refer to Chapter x, section y" line):

 "mcq"          question, options {"A": "...", "B": "..."}, answer "B"
 "match"        question, pairs [[left, right], ...], extra_options [..optional distractors..],
                column_a_title, column_b_title (optional), shuffle_seed (optional)
 "rearrange"    question, steps [ ...in the CORRECT order... ]  (the script scrambles and letters them)
 "fill_blank"   question (optional), word_bank [..], sentences [{"text": "A ______ is ...", "answer": "hazard"}]
 "written"      question, sample_answer, word_limit_text (e.g. "Answer must be within 50 words.")
 "tabular"      question, columns [header, header, ...], rows [[cell, cell, ...], ...],
                n_rows / n_columns (optional - if given they MUST match the data),
                prefilled: {"columns": [0] or ["Header name"], "rows": [..row numbers..], "cells": [[row, col], ...]}
                   -> anything listed there is BLACK on pink; every other cell is a student answer (RED, no shading).
                   Row/column numbers start at 0 and count data rows only (the header row is always pre-filled).
                A cell may also be {"text": "...", "filled_by": "prefilled" | "student"} to override the rule for one cell.
                Any number of columns is supported (1-8); every row gets its own S / NS tick boxes.
 "scenario"     scenario (text, or list of paragraphs), instruction (optional),
                questions [ <mcq|match|rearrange|fill_blank|written|tabular>, ... ]   -> numbered 1a, 1b, ...
 "benchmark_statement"  text    (the template's "Benchmarking Statement" block, placed after the previous question)

Optional on any question:  "answer_topic" (finishes "...based on their knowledge of <topic>"),
                           "answer_note" (replaces the whole red instruction line above the sample answer),
                           "number" (force a question number/label).
"""
from __future__ import annotations

import argparse
import json
import random
import string
import sys
from copy import deepcopy
from pathlib import Path

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PINK, STUDENT_RED, ANSWER_RED, BLACK = "D9A1B7", "EE0000", "FF0000", "000000"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
TYPES = {"mcq", "match", "rearrange", "fill_blank", "written", "tabular", "scenario", "benchmark_statement"}
CHILD_TYPES = TYPES - {"scenario", "benchmark_statement"}

# (question-table index, answer-table index) of each prototype block inside the template body
PROTO = {"mcq": (2, 4), "rearrange": (8, 10), "match": (15, 17), "fill_blank": (21, 23), "tabular": (26, 28),
         "written": (41, 43), "scen_single_mcq": (50, 52), "scen_single_written": (33, 43), "scen_parent": (56, None),
         "child_mcq": (59, 61)}
EXPECT_HEADINGS = {1: "MCQ", 6: "Rearrange", 13: "Mix and Match", 19: "Fill in the blanks", 25: "Tabular",
                   31: "Scenario based Single Question", 37: "Scenario based multiple Questions",
                   49: "Scenario based single MCQ", 54: "Scenario based multiple MCQs", 68: "Benchmarking"}


# ============================================================================ OOXML helpers
def ptext(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def paras(tc):
    return tc.findall(qn("w:p"))


def rows_of(tbl):
    return tbl.findall(qn("w:tr"))


def cells_of(tr):
    return tr.findall(qn("w:tc"))


def nested(tc):
    return tc.findall(qn("w:tbl"))


def _flag(rpr, tag, on):
    if on is None:
        return
    for t in (tag, tag + "Cs"):
        el = rpr.find(qn(t))
        if on and el is None:
            rpr.insert(0, OxmlElement(t))
        elif not on and el is not None:
            rpr.remove(el)


def set_runs(p, segs):
    """Replace the runs of paragraph p. segs: [(text, bold, italic, color)], None = keep the template run's formatting."""
    runs = p.findall(qn("w:r"))
    base = next((r for r in runs if r.find(qn("w:t")) is not None), runs[0] if runs else None)
    base_rpr = deepcopy(base.find(qn("w:rPr"))) if base is not None and base.find(qn("w:rPr")) is not None else None
    for r in runs:
        p.remove(r)
    for seg in segs:
        text, bold, ital, color = (tuple(seg) + (None, None, None))[:4]
        r = OxmlElement("w:r")
        rpr = deepcopy(base_rpr) if base_rpr is not None else OxmlElement("w:rPr")
        _flag(rpr, "w:b", bold)
        _flag(rpr, "w:i", ital)
        if color:
            c = rpr.find(qn("w:color"))
            if c is None:
                c = OxmlElement("w:color")
                rpr.append(c)
            c.set(qn("w:val"), color)
            for k in ("w:themeColor", "w:themeShade", "w:themeTint"):
                if c.get(qn(k)) is not None:
                    del c.attrib[qn(k)]
        if len(rpr):
            r.append(rpr)
        t = OxmlElement("w:t")
        t.text = text
        t.set(XML_SPACE, "preserve")
        r.append(t)
        p.append(r)


def clone_after(p, segs):
    n = deepcopy(p)
    p.addnext(n)
    set_runs(n, segs)
    return n


def set_fill(tc, fill):
    """fill = hex string (shaded) or None (no shading, i.e. white like the sample's student-answer cells)."""
    tcpr = tc.find(qn("w:tcPr"))
    if tcpr is None:
        tcpr = OxmlElement("w:tcPr")
        tc.insert(0, tcpr)
    for old in tcpr.findall(qn("w:shd")):
        tcpr.remove(old)
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        tcpr.append(shd)


def set_cell(tc, text, *, by="prefilled", bold=None):
    """Write text into a cell using the colour/shading rule for who fills it."""
    colour, fill = (BLACK, PINK) if by == "prefilled" else (STUDENT_RED, None)
    ps = paras(tc)
    lines = (text or "").split("\n")
    set_runs(ps[0], [(lines[0], bold, None, colour)])
    last = ps[0]
    for ln in lines[1:]:
        last = clone_after(last, [(ln, bold, None, colour)])
    for extra in ps[1:]:                       # drop leftover empty template paragraphs
        if extra.getparent() is tc and not ptext(extra).strip() and extra is not last:
            tc.remove(extra)
    set_fill(tc, fill)


def clear_cell(tc):
    for p in paras(tc):
        for r in p.findall(qn("w:r")):
            p.remove(r)


def reset_rows(tbl, n_data, header_rows=1):
    rs = rows_of(tbl)
    proto = deepcopy(rs[header_rows])
    for r in rs[header_rows:]:
        tbl.remove(r)
    out = []
    for _ in range(max(n_data, 0)):
        r = deepcopy(proto)
        tbl.append(r)
        out.append(r)
    return out


def set_content_columns(nt, n):
    """Make the nested answer table have exactly n content columns (the last 2 columns are always the S / NS boxes)."""
    grid = nt.find(qn("w:tblGrid"))
    gcols = grid.findall(qn("w:gridCol"))
    cur = len(gcols) - 2
    total = sum(int(c.get(qn("w:w"))) for c in gcols[:cur])
    while cur > n:                                       # remove from the right-most content column
        idx = cur - 1
        grid.remove(grid.findall(qn("w:gridCol"))[idx])
        for tr in rows_of(nt):
            tr.remove(cells_of(tr)[idx])
        cur -= 1
    while cur < n:                                       # add by cloning the right-most content column
        idx = cur - 1
        g = grid.findall(qn("w:gridCol"))[idx]
        g.addnext(deepcopy(g))
        for tr in rows_of(nt):
            c = cells_of(tr)[idx]
            c.addnext(deepcopy(c))
        cur += 1
    w = str(total // n)
    for g in grid.findall(qn("w:gridCol"))[:n]:
        g.set(qn("w:w"), w)
    for tr in rows_of(nt):
        for c in cells_of(tr)[:n]:
            tcpr = c.find(qn("w:tcPr"))
            tcw = tcpr.find(qn("w:tcW")) if tcpr is not None else None
            if tcw is not None:
                tcw.set(qn("w:w"), w)
                tcw.set(qn("w:type"), "dxa")


def no_split_rows(tbl):
    for tr in tbl.iter(qn("w:tr")):
        trpr = tr.find(qn("w:trPr"))
        if trpr is None:
            trpr = OxmlElement("w:trPr")
            tr.insert(1 if tr.find(qn("w:tblPrEx")) is not None else 0, trpr)
        if trpr.find(qn("w:cantSplit")) is None:
            trpr.append(OxmlElement("w:cantSplit"))


# ============================================================================ JSON validation
class SpecError(Exception):
    pass


def _label(q, path):
    return f"{path} ({q.get('type', '?')})"


def validate(spec: dict) -> list[str]:
    """Return a list of human-readable problems (empty list = OK)."""
    errs: list[str] = []
    if not isinstance(spec, dict) or not isinstance(spec.get("questions"), list) or not spec["questions"]:
        return ['top level must be an object with a non-empty "questions" list']

    def need(q, key, path, kind=None):
        v = q.get(key)
        if v in (None, "", [], {}):
            errs.append(f'{_label(q, path)}: missing "{key}"')
            return None
        if kind and not isinstance(v, kind):
            errs.append(f'{_label(q, path)}: "{key}" must be {kind.__name__}')
            return None
        return v

    def check(q, path, child=False):
        if not isinstance(q, dict):
            errs.append(f"{path}: must be an object")
            return
        t = q.get("type")
        if t not in TYPES:
            errs.append(f'{path}: unknown "type" {t!r}; use one of {sorted(TYPES)}')
            return
        if child and t not in CHILD_TYPES:
            errs.append(f"{path}: a {t} cannot be nested inside a scenario")
            return
        if t == "benchmark_statement":
            need(q, "text", path)
        elif t == "scenario":
            need(q, "scenario", path)
            kids = need(q, "questions", path, list)
            for i, k in enumerate(kids or [], 1):
                check(k, f"{path}.{i}", child=True)
        elif t == "mcq":
            need(q, "question", path)
            o = need(q, "options", path, dict)
            a = need(q, "answer", path)
            if o and len(o) < 2:
                errs.append(f"{_label(q, path)}: needs at least 2 options")
            if o and a and a not in o:
                errs.append(f'{_label(q, path)}: answer "{a}" is not one of the options {list(o)}')
        elif t == "match":
            need(q, "question", path)
            pairs = need(q, "pairs", path, list)
            if pairs and (len(pairs) < 2 or any(not isinstance(p, (list, tuple)) or len(p) != 2 for p in pairs)):
                errs.append(f"{_label(q, path)}: pairs must be a list of [left, right] with at least 2 pairs")
        elif t == "rearrange":
            need(q, "question", path)
            s = need(q, "steps", path, list)
            if s and len(s) < 3:
                errs.append(f"{_label(q, path)}: needs at least 3 steps")
        elif t == "fill_blank":
            bank = need(q, "word_bank", path, list)
            sents = need(q, "sentences", path, list)
            low = {str(b).strip().lower() for b in (bank or [])}
            for i, s in enumerate(sents or [], 1):
                if not isinstance(s, dict) or "text" not in s or "answer" not in s:
                    errs.append(f'{_label(q, path)}: sentence {i} needs "text" and "answer"')
                    continue
                if "__" not in s["text"]:
                    errs.append(f'{_label(q, path)}: sentence {i} has no blank (use ______)')
                if str(s["answer"]).strip().lower() not in low:
                    errs.append(f'{_label(q, path)}: sentence {i} answer "{s["answer"]}" is not in word_bank')
        elif t == "written":
            need(q, "question", path)
            need(q, "sample_answer", path)
        elif t == "tabular":
            need(q, "question", path)
            cols = need(q, "columns", path, list)
            rws = need(q, "rows", path, list)
            if cols and not 1 <= len(cols) <= 8:
                errs.append(f"{_label(q, path)}: 1-8 columns supported, got {len(cols)}")
            if q.get("n_columns") is not None and cols and q["n_columns"] != len(cols):
                errs.append(f'{_label(q, path)}: n_columns={q["n_columns"]} but "columns" has {len(cols)}')
            if q.get("n_rows") is not None and rws is not None and q["n_rows"] != len(rws):
                errs.append(f'{_label(q, path)}: n_rows={q["n_rows"]} but "rows" has {len(rws)} data rows')
            for r, row in enumerate(rws or []):
                if not isinstance(row, list) or (cols and len(row) != len(cols)):
                    errs.append(f"{_label(q, path)}: row {r} must be a list of {len(cols or [])} cells")
            pf = q.get("prefilled", {}) or {}
            for c in pf.get("columns", []):
                if isinstance(c, str) and cols and c not in cols:
                    errs.append(f'{_label(q, path)}: prefilled column "{c}" is not in columns')
                if isinstance(c, int) and cols and not 0 <= c < len(cols):
                    errs.append(f"{_label(q, path)}: prefilled column {c} out of range")
            for r in pf.get("rows", []):
                if rws is not None and not 0 <= r < len(rws):
                    errs.append(f"{_label(q, path)}: prefilled row {r} out of range")
            for rc in pf.get("cells", []):
                if rws is not None and cols and not (isinstance(rc, list) and len(rc) == 2 and 0 <= rc[0] < len(rws) and 0 <= rc[1] < len(cols)):
                    errs.append(f"{_label(q, path)}: prefilled cell {rc} out of range")

    for i, q in enumerate(spec["questions"], 1):
        check(q, f"question #{i}")
    return errs


# ============================================================================ rendering
def _letters(n):
    return list(string.ascii_uppercase[:n])


def _lines(text):
    return str(text).split("\n")


def _ref(q):
    return f"Refer to Chapter {q.get('chapter') or '____'}, section {q.get('section') or '____'}"


def _set_ref(tc, q):
    refp = next((p for p in paras(tc) if ptext(p).strip().lower().startswith("refer")), None)
    if refp is not None:
        set_runs(refp, [(_ref(q), None, None, None)])


def _question_cell(qtbl, label, text, q, extra_lines=(), word_limit=None):
    tc = cells_of(rows_of(qtbl)[0])[0]
    ps = paras(tc)
    ls = _lines(text)
    set_runs(ps[0], [(label, True, None, None), (" " + ls[0], False, None, None)])
    last = ps[0]
    for ln in list(ls[1:]) + list(extra_lines):
        last = clone_after(last, [(ln, False, None, None)])
    _set_ref(tc, q)
    wl = next((p for p in paras(tc) if "within" in ptext(p).lower() and "words" in ptext(p).lower()), None)
    if wl is not None:
        if word_limit:
            set_runs(wl, [(word_limit, None, None, None)])
        else:
            tc.remove(wl)
    return tc


def _scenario_cell(qtbl, label, intro, scenario, question, q, extra_lines=()):
    tc = cells_of(rows_of(qtbl)[0])[0]
    ps = paras(tc)
    set_runs(ps[0], [(label, True, None, None), (" " + intro, False, None, None)])
    idx = [i for i, p in enumerate(ps) if ptext(p).strip() in ("Scenario", "Question")]
    sc_i, q_i = idx[0], idx[1]
    sc = scenario if isinstance(scenario, list) else _lines(scenario)
    set_runs(ps[sc_i + 1], [(sc[0], False, None, None)])
    last = ps[sc_i + 1]
    for ln in sc[1:]:
        last = clone_after(last, [(ln, False, None, None)])
    ql = _lines(question)
    set_runs(ps[q_i + 1], [(ql[0], False, None, None)])
    last = ps[q_i + 1]
    for ln in ql[1:] + list(extra_lines):
        last = clone_after(last, [(ln, False, None, None)])
    _set_ref(tc, q)


def _answer_head(atbl, instruction, sample=None):
    """Red instruction line; optional sample-answer text (red) on the third paragraph."""
    tc = cells_of(rows_of(atbl)[0])[0]
    ps = paras(tc)
    set_runs(ps[0], [(instruction, None, None, None)])
    if sample is not None and len(ps) >= 3:
        ls = _lines(sample)
        set_runs(ps[2], [(ls[0], False, False, ANSWER_RED)])
        last = ps[2]
        for ln in ls[1:]:
            last = clone_after(last, [(ln, False, False, ANSWER_RED)])
    return tc


def _instruction(q, default):
    if q.get("answer_note"):
        return q["answer_note"]
    return default.replace("{topic}", q.get("answer_topic") or "this unit")


def _cell_spec(cell):
    if isinstance(cell, dict):
        return str(cell.get("text", "")), cell.get("filled_by")
    return ("" if cell is None else str(cell)), None


def _tabular_owner(q, r, c, explicit):
    if explicit in ("prefilled", "student"):
        return explicit
    pf = q.get("prefilled")
    if pf is None:                                        # no rule given -> first column is the printed prompt column
        return "prefilled" if c == 0 else "student"
    cols = q["columns"]
    pcols = {cols.index(x) if isinstance(x, str) else x for x in pf.get("columns", [])}
    if c in pcols or r in set(pf.get("rows", [])) or [r, c] in [list(x) for x in pf.get("cells", [])]:
        return "prefilled"
    return "student"


class Renderer:
    def __init__(self, proto):
        self.proto = proto
        self.out = []

    # ---- low level
    def _pair(self, kind, fill_q, fill_a=None):
        qi, ai = PROTO[kind]
        q = deepcopy(self.proto[qi])
        fill_q(q)
        self.out += [q, deepcopy(self.proto["gap"])]
        if ai is not None and fill_a is not None:
            a = deepcopy(self.proto[ai])
            fill_a(a)
            self.out += [a, deepcopy(self.proto["gap"])]

    # ---- one question of any non-scenario type
    def item(self, q, label, scenario=None, in_group=False):
        t = q["type"]
        if t == "mcq":
            opts = [f"{k}. {v}" for k, v in q["options"].items()]
            if scenario is not None:
                kind = "scen_single_mcq"
                fq = lambda tb: _scenario_cell(tb, label, "Read the scenario below and answer the question:", scenario, q["question"], q, opts)
            else:
                kind = "child_mcq" if in_group else "mcq"
                fq = lambda tb: _question_cell(tb, label, q["question"], q, opts)
            self._pair(kind, fq, lambda a: _answer_head(a, _instruction(q, "Students must select the correct option."), q["answer"] + "."))

        elif t == "written":
            wl = q.get("word_limit_text")
            if scenario is not None:
                kind = "scen_single_written"
                fq = lambda tb: _scenario_cell(tb, label, "Read the scenario below and answer the question:", scenario, q["question"], q)
            else:
                kind = "written"
                fq = lambda tb: _question_cell(tb, label, q["question"], q, word_limit=wl)
            self._pair(kind, fq, lambda a: _answer_head(a, _instruction(q, "Answers may vary. Students must answer the above question based on their knowledge of {topic}."), q["sample_answer"]))

        elif t == "rearrange":
            steps = q["steps"]
            order = list(range(len(steps)))
            rnd = random.Random(q.get("shuffle_seed", 11))
            for _ in range(30):
                rnd.shuffle(order)
                if order != sorted(order):
                    break
            letters = _letters(len(steps))
            pos_letter = {order[i]: letters[i] for i in range(len(order))}
            correct_seq = [pos_letter[k] for k in range(len(steps))]

            def fa(a):
                tc = _answer_head(a, _instruction(q, f"Students must rearrange the steps in the correct sequence (1 - {len(steps)})."), None)
                nt = nested(tc)[0]
                for i, r in enumerate(reset_rows(nt, len(steps))):
                    cs = cells_of(r)
                    set_cell(cs[0], f"{letters[i]}. {steps[order[i]]}")
                    set_cell(cs[1], correct_seq[i], by="student")
            self._pair("rearrange", lambda tb: _question_cell(tb, label, q["question"], q), fa)

        elif t == "match":
            lefts = [p[0] for p in q["pairs"]]
            options = [p[1] for p in q["pairs"]] + list(q.get("extra_options", []))
            idx = list(range(len(options)))
            random.Random(q.get("shuffle_seed", 7)).shuffle(idx)
            letters = _letters(len(options))
            letter_of = {idx[k]: letters[k] for k in range(len(options))}
            n_rows = max(len(lefts), len(options))
            ta, tb_ = q.get("column_a_title", "Column A"), q.get("column_b_title", "Column B")

            def fa(a):
                tc = _answer_head(a, _instruction(q, f"Students must correctly match each item in {ta} with that of {tb_}, based on their knowledge of " + "{topic}."), None)
                nt = nested(tc)[0]
                hdr = cells_of(rows_of(nt)[0])
                set_cell(hdr[0], ta, bold=True)
                set_cell(hdr[1], tb_, bold=True)
                set_cell(hdr[2], "Correct option", bold=True)
                for i, r in enumerate(reset_rows(nt, n_rows)):
                    cs = cells_of(r)
                    set_cell(cs[0], f"{i + 1}. {lefts[i]}" if i < len(lefts) else "")
                    set_cell(cs[1], f"{letters[i]}. {options[idx[i]]}" if i < len(options) else "")
                    set_cell(cs[2], letter_of[i] if i < len(lefts) else "", by="student")
                    if i >= len(lefts):
                        clear_cell(cs[3]); clear_cell(cs[4])
            self._pair("match", lambda tb: _question_cell(tb, label, q["question"], q), fa)

        elif t == "fill_blank":
            bank, sents = q["word_bank"], q["sentences"]
            stem = q.get("question") or "Complete the following sentences by filling in the blanks with the correct words from the list provided below."

            def fq(tb):
                tc = _question_cell(tb, label, stem, q)
                nts = nested(tc)
                if nts:
                    nt = nts[0]
                    per = len(cells_of(rows_of(nt)[0]))
                    for ri, r in enumerate(reset_rows(nt, -(-len(bank) // per), header_rows=0)):
                        for ci, c in enumerate(cells_of(r)[:per]):
                            k = ri * per + ci
                            set_cell(c, str(bank[k]) if k < len(bank) else "")

            def fa(a):
                tc = _answer_head(a, _instruction(q, "Students must correctly complete the following sentences by filling in the blanks with the correct words from the list provided below."), None)
                nt = nested(tc)[0]
                hdr = cells_of(rows_of(nt)[0])
                for i, r in enumerate(reset_rows(nt, len(sents))):
                    cs = cells_of(r)
                    set_cell(cs[0], str(i + 1)); set_cell(cs[1], sents[i]["text"]); set_cell(cs[2], str(sents[i]["answer"]), by="student", bold=True)
            self._pair("fill_blank", fq, fa)

        elif t == "tabular":
            cols, rws = q["columns"], q["rows"]

            def fa(a):
                tc = _answer_head(a, _instruction(q, "Answers may vary. Students must answer the above question based on their knowledge of {topic}."), None)
                nt = nested(tc)[0]
                set_content_columns(nt, len(cols))
                hdr = cells_of(rows_of(nt)[0])
                for i, name in enumerate(cols):
                    set_cell(hdr[i], str(name), bold=True)                     # header row: always printed (black on pink)
                for r, row_el in enumerate(reset_rows(nt, len(rws))):
                    cs = cells_of(row_el)
                    for c in range(len(cols)):
                        text, explicit = _cell_spec(rws[r][c])
                        set_cell(cs[c], text, by=_tabular_owner(q, r, c, explicit))
            self._pair("tabular", lambda tb: _question_cell(tb, label, q["question"], q), fa)

    # ---- scenario group
    def scenario(self, g, n):
        kids = g["questions"]
        sc = g["scenario"] if isinstance(g["scenario"], list) else _lines(g["scenario"])
        if len(kids) == 1 and kids[0]["type"] in ("mcq", "written"):
            self.item(kids[0], f"Question {n}.", scenario=sc)
            return
        parent = deepcopy(self.proto[PROTO["scen_parent"][0]])
        tc = cells_of(rows_of(parent)[0])[0]
        ps = paras(tc)
        set_runs(ps[0], [(f"Question {n}.", True, None, None), (" " + g.get("instruction", "Read the scenario below and answer the questions"), False, None, None)])
        sc_i = next(i for i, p in enumerate(ps) if ptext(p).strip() == "Scenario")
        set_runs(ps[sc_i + 1], [(sc[0], False, None, None)])
        last = ps[sc_i + 1]
        for ln in sc[1:]:
            last = clone_after(last, [(ln, False, None, None)])
        self.out += [parent, deepcopy(self.proto["gap"])]
        for k, kid in enumerate(kids):
            self.item(kid, kid.get("number") or f"Question {n}{string.ascii_lowercase[k]}.", in_group=True)

    def benchmark(self, text):
        p1, p2 = deepcopy(self.proto["bench_quote"]), deepcopy(self.proto["bench_instr"])
        set_runs(p1, [("“" + text + "”", None, None, None)])
        self.out += [p1, p2, deepcopy(self.proto["gap"])]


def check_template(doc):
    body = list(doc.element.body)
    for idx, name in EXPECT_HEADINGS.items():
        if idx >= len(body) or name.lower() not in ptext(body[idx]).lower():
            raise SystemExit(f'Template layout differs from the expected one: heading "{name}" not found at position {idx}. '
                             "Use new_task_1_template_version_10_7.docx (or update PROTO / EXPECT_HEADINGS).")


def build_document(spec: dict, template: Path, out: Path, with_title: bool = True) -> Path:
    d = docx.Document(str(template))
    check_template(d)
    body = d.element.body
    kids = list(body)
    proto = {i: deepcopy(k) for i, k in enumerate(kids)}
    proto["gap"] = deepcopy(kids[3])
    proto["bench_quote"], proto["bench_instr"] = deepcopy(kids[69]), deepcopy(kids[70])
    heading, sect = deepcopy(kids[1]), kids[-1]
    for k in kids[:-1]:
        body.remove(k)
    if with_title:
        title = spec.get("title") or f"Task 1: Knowledge Assessment - {spec.get('unit_code', '')} {spec.get('unit_title', '')}".strip(" -")
        set_runs(heading, [(title, None, None, None)])
        sect.addprevious(heading)

    r = Renderer(proto)
    n = 0
    for q in spec["questions"]:
        if q["type"] == "benchmark_statement":
            r.benchmark(q["text"])
            continue
        n += 1
        label = q.get("number")
        if q["type"] == "scenario":
            r.scenario(q, label or n)
        else:
            r.item(q, f"Question {label or n}.")
    for el in r.out:
        if el.tag == qn("w:tbl"):
            no_split_rows(el)
        sect.addprevious(el)
    out.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(out))
    return out


def verify(out: Path, spec: dict) -> list[str]:
    """Re-open the file and make sure every question and every sample answer actually landed in it."""
    import re
    d = docx.Document(str(out))
    text = re.sub(r"\s+", " ", " ".join(t.text or "" for t in d.element.body.iter(qn("w:t"))))
    miss = []

    def need(s, what):
        s = re.sub(r"\s+", " ", str(s)).strip()
        if s and s not in text:
            miss.append(f"{what}: {s[:60]}")

    def walk(q):
        t = q["type"]
        if t == "scenario":
            for ln in (q["scenario"] if isinstance(q["scenario"], list) else _lines(q["scenario"])):
                need(ln, "scenario")
            for k in q["questions"]:
                walk(k)
        elif t == "benchmark_statement":
            need(q["text"], "benchmark")
        else:
            if q.get("question"):
                need(q["question"].split("\n")[0], "question")
            if t == "mcq":
                [need(f"{k}. {v}", "option") for k, v in q["options"].items()]
            elif t == "written":
                need(q["sample_answer"].split("\n")[0], "sample answer")
            elif t == "tabular":
                for row in q["rows"]:
                    for c in row:
                        need(_cell_spec(c)[0], "table cell")
            elif t == "fill_blank":
                [need(s["text"], "sentence") for s in q["sentences"]]
            elif t == "rearrange":
                [need(s, "step") for s in q["steps"]]
            elif t == "match":
                [need(p[0], "match item") for p in q["pairs"]]
    for q in spec["questions"]:
        walk(q)
    return miss


EXAMPLE = {
    "unit_code": "FWPCOT2275", "unit_title": "Fell trees manually (basic)",
    "questions": [
        {"type": "mcq", "chapter": "4", "section": "4.1", "question": "What is the main purpose of the scarf cut?",
         "options": {"A": "To sharpen the chain", "B": "To set the direction of fall", "C": "To remove bark", "D": "To cool the bar"}, "answer": "B"},
        {"type": "tabular", "chapter": "1", "section": "1.2", "n_rows": 2, "n_columns": 3,
         "question": "Complete the table. For each source, identify one environmental protection requirement and one practice used to meet it.",
         "answer_topic": "environmental protection requirements",
         "columns": ["Source of requirement", "Requirement", "Practice used to meet it"],
         "rows": [["Legal obligations", "Prevent environmental harm", "Identify protected areas before felling"],
                  ["Industry requirements", "Follow accepted forestry practice", "Minimise damage to surrounding vegetation"]],
         "prefilled": {"columns": [0]}},
    ],
}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Format Task 1 questions (JSON) into the Task 1 Word template.")
    ap.add_argument("json_file", nargs="?", help="input JSON (see the docstring at the top of this file)")
    ap.add_argument("-t", "--template", default=None, help="Task 1 template .docx (default: next to this script)")
    ap.add_argument("-o", "--output", default=None, help="output .docx (default: <json name>.docx)")
    ap.add_argument("--check", action="store_true", help="only validate the JSON")
    ap.add_argument("--no-title", action="store_true", help="do not add the heading at the top")
    ap.add_argument("--example", action="store_true", help="print a small example JSON and exit")
    a = ap.parse_args(argv)
    if a.example:
        print(json.dumps(EXAMPLE, indent=2, ensure_ascii=False))
        return 0
    if not a.json_file:
        ap.error("json_file is required")
    spec = json.loads(Path(a.json_file).read_text(encoding="utf-8"))
    errs = validate(spec)
    if errs:
        print(f"{len(errs)} problem(s) in {a.json_file}:")
        for e in errs:
            print("  -", e)
        return 1
    if a.check:
        print(f"OK: {len(spec['questions'])} top-level entries.")
        return 0
    here = Path(__file__).resolve().parent
    template = Path(a.template) if a.template else here / "new_task_1_template_version_10_7.docx"
    if not template.exists():
        print(f"Template not found: {template}  (use -t)")
        return 2
    out = Path(a.output) if a.output else Path(a.json_file).with_suffix(".docx")
    build_document(spec, template, out, with_title=not a.no_title)
    miss = verify(out, spec)
    if miss:
        print("WARNING - these items were not found in the output:")
        for m in miss:
            print("  -", m)
        return 3
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
