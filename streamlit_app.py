"""
PBS Support Tool
Tab 1: Upload a PBSP → Support Reference Card + ABC Recording Form.
Tab 2: Enter client behaviours → AI strategy recommendations + PDF report.
"""

import json
from datetime import date
import streamlit as st
import anthropic
from io import BytesIO
from docx import Document as DocxDocument
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.colors import HexColor, white
from reportlab.pdfbase.pdfmetrics import stringWidth

st.set_page_config(page_title="PBS Support Tool", page_icon="📋", layout="centered")

DEEP_BLUE  = HexColor('#0D4F6E')
TEAL       = HexColor('#1A9B8A')
BLUE2      = HexColor('#1A5276')
GREEN      = HexColor('#27AE60')
RED        = HexColor('#C0392B')
AMBER      = HexColor('#D4700A')
LIGHT_TEAL = HexColor('#E0F6F3')
LIGHT_BLUE = HexColor('#DCEEF8')
LIGHT_GRN  = HexColor('#DFFAE9')
LIGHT_RED  = HexColor('#FADBD8')
LIGHT_AMBR = HexColor('#FDEBD0')
MID_GREY   = HexColor('#CCCCCC')
DARK_TEXT  = HexColor('#1A2B35')
MED_TEXT   = HexColor('#4A7C8E')

W, H = A4
LM, RM = 20, 20
CW = W - LM - RM


# ══════════════════════════════════════════════════════════════════════════════
# TEXT EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = DocxDocument(BytesIO(file_bytes))
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                parts.append("  |  ".join(dict.fromkeys(cells)))
    return "\n".join(parts)

def extract_text_from_pdf(file_bytes: bytes) -> str:
    reader = PdfReader(BytesIO(file_bytes))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


# ══════════════════════════════════════════════════════════════════════════════
# CLAUDE — PBSP EXTRACTION (Tab 1)
# ══════════════════════════════════════════════════════════════════════════════

EXTRACTION_PROMPT = """
You are a Positive Behaviour Support (PBS) assistant.
Read the PBSP text below and extract key information into a JSON object.

Rules:
- Keep each bullet point under 90 characters — written for support workers to scan quickly
- Use plain language (no clinical jargon where possible)
- Focus on what is most operationally useful during a shift
- If a field cannot be found, use an empty string or empty list

Return ONLY valid JSON — no commentary, no markdown fences.

JSON structure:
{
  "name": "full name",
  "preferred": "preferred/short name",
  "pronouns": "pronouns",
  "age_info": "Age XX  |  [Diagnosis]  |  [Organisation / Location]",
  "about": ["up to 5 key points a support worker must know about this person"],
  "warning_signs": ["up to 5 observable early warning signs that behaviour is building"],
  "triggers": ["up to 6 known setting events and immediate triggers"],
  "proactive": ["up to 5 proactive strategies — things to DO to prevent behaviour"],
  "reactive": ["up to 5 reactive strategies — what to DO when behaviour occurs"],
  "do_not": ["up to 4 things NOT to do — known escalators"],
  "behaviours": [
    {
      "label": "short behaviour name (e.g. Verbal Outbursts)",
      "descriptors": ["up to 5 observable descriptors of this behaviour"]
    }
  ],
  "setting_events_checklist": ["up to 6 setting events as short checkbox labels"],
  "antecedents_checklist": ["up to 8 common antecedents as short checkbox labels"],
  "staff_responses_checklist": ["up to 7 common staff responses as short checkbox labels"],
  "review_date": "review year or date",
  "practitioner": "name and title",
  "contact": "email or contact"
}

PBSP TEXT:
"""

def extract_client_data(pbsp_text: str, api_key: str) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-opus-4-5", max_tokens=2000,
        messages=[{"role": "user", "content": EXTRACTION_PROMPT + pbsp_text[:40000]}]
    )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"): raw = "\n".join(raw.split("\n")[1:])
    if raw.endswith("```"):   raw = "\n".join(raw.split("\n")[:-1])
    return json.loads(raw)


# ══════════════════════════════════════════════════════════════════════════════
# CLAUDE — STRATEGY RECOMMENDER (Tab 2)
# ══════════════════════════════════════════════════════════════════════════════

STRATEGY_PROMPT = """\
You are an experienced Positive Behaviour Support (PBS) Practitioner.
Based on the client profile and behaviours described, recommend practical PBS strategies for support workers.

{library_instruction}

Return ONLY valid JSON — no commentary, no markdown fences.

JSON structure:
{{
  "client_summary": "2-3 sentence summary of this person's likely support needs and overall PBS approach",
  "general_strategies": ["Up to 5 general/environmental strategies that apply across all behaviours"],
  "behaviours": [
    {{
      "behaviour": "behaviour name from input",
      "likely_function": "Primary function (Escape/Avoidance | Access | Sensory | Attention) — one sentence rationale",
      "proactive": ["Up to 5 proactive/antecedent strategies to prevent this behaviour"],
      "reactive": ["Up to 5 reactive de-escalation strategies for when this behaviour occurs"],
      "teach_instead": ["Up to 3 replacement skills or alternative communication strategies to build"],
      "avoid": ["Up to 3 specific things NOT to do — common mistakes that escalate this behaviour"]
    }}
  ]
}}

Rules:
- Keep each item under 90 characters — written for support workers to act on quickly
- Be specific to the triggers and context described — avoid generic advice
- Base strategy recommendations on the likely function of each behaviour
- Use plain language, not clinical jargon
- Replacement skills must serve the same function as the behaviour (functionally equivalent)

CLIENT PROFILE:
{profile}

BEHAVIOURS OF CONCERN:
{behaviours}
{library_section}"""

def recommend_strategies(client_info: dict, behaviours: list, api_key: str,
                          library_text: str = None) -> dict:
    profile_text = (
        f"Name: {client_info['name']}\n"
        f"Age: {client_info['age']}\n"
        f"Diagnosis/Condition: {client_info['diagnosis']}\n"
        f"Communication level: {client_info['comms']}\n"
        f"Additional context: {client_info['other']}"
    )
    behaviours_text = ""
    for i, b in enumerate(behaviours, 1):
        behaviours_text += (
            f"\nBehaviour {i}: {b['name']}\n"
            f"  What it looks like: {b['description']}\n"
            f"  Known triggers / when it occurs: {b['triggers']}\n"
        )
    if library_text:
        lib_instruction = (
            "IMPORTANT: Your primary task is to SELECT strategies FROM THE STRATEGY LIBRARY "
            "provided at the end of this prompt. Quote or closely paraphrase library strategies. "
            "Where the library has no relevant strategy for a specific need, you may suggest an "
            "evidence-based alternative and append '(not in library)' to that item."
        )
        lib_section = f"\nSTRATEGY LIBRARY:\n{library_text[:30000]}"
    else:
        lib_instruction = (
            "Generate evidence-based PBS strategies based on the likely function of each behaviour."
        )
        lib_section = ""
    prompt = STRATEGY_PROMPT.format(
        profile=profile_text, behaviours=behaviours_text,
        library_instruction=lib_instruction, library_section=lib_section,
    )
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-opus-4-5", max_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"): raw = "\n".join(raw.split("\n")[1:])
    if raw.endswith("```"):   raw = "\n".join(raw.split("\n")[:-1])
    return json.loads(raw)


def pbsp_to_sr_format(data: dict) -> tuple:
    """Convert extract_client_data output → (client_info dict, behaviours list) for Tab 2."""
    age_info = data.get("age_info", "")
    parts    = [p.strip() for p in age_info.split("|")]
    age      = parts[0].replace("Age", "").strip() if parts else "not specified"
    diagnosis = parts[1] if len(parts) > 1 else "not specified"
    client_info = {
        "name":      data.get("name") or data.get("preferred") or "Unknown",
        "age":       age or "not specified",
        "diagnosis": diagnosis or "not specified",
        "comms":     "refer to PBSP",
        "other":     "; ".join(data.get("about", [])) or "none provided",
    }
    triggers_str = "; ".join(data.get("triggers", []))
    behaviours = [
        {
            "name":        b.get("label", "Behaviour"),
            "description": "; ".join(b.get("descriptors", [])),
            "triggers":    triggers_str,
        }
        for b in data.get("behaviours", [])
    ]
    return client_info, behaviours


# ── Free-text behaviour description → strategies ──────────────────────────────

FREETEXT_STRATEGY_PROMPT = """\
You are an experienced Positive Behaviour Support (PBS) Practitioner.
A practitioner has described a client and their behaviours in plain language below.

Your tasks:
1. Read the description and identify each distinct behaviour of concern
2. Assign each a clear clinical label (e.g. "Physical aggression", "Self-injurious behaviour",
   "Property destruction", "Verbal aggression", "Elopement")
3. Infer the likely triggers and context from what is described
4. Determine the likely function of each behaviour
5. Recommend practical PBS strategies

{library_instruction}

Return ONLY valid JSON — no commentary, no markdown fences.

JSON structure:
{{
  "client_summary": "2-3 sentence clinical summary based on the description",
  "general_strategies": ["Up to 5 general/environmental strategies across all behaviours"],
  "behaviours": [
    {{
      "behaviour": "Clinical label you have assigned",
      "likely_function": "Primary function (Escape/Avoidance | Access | Sensory | Attention) — one sentence rationale drawn from the description",
      "proactive": ["Up to 5 proactive strategies to prevent this behaviour"],
      "reactive": ["Up to 5 reactive de-escalation strategies when this behaviour occurs"],
      "teach_instead": ["Up to 3 replacement skills or communication strategies to build"],
      "avoid": ["Up to 3 things NOT to do — common mistakes that escalate this behaviour"]
    }}
  ]
}}

Rules:
- Keep each item under 90 characters — written for support workers to act on quickly
- Use the description to infer triggers and context — be specific, not generic
- Assign clinically appropriate behaviour labels
- Use plain language in strategy recommendations
- Replacement skills must serve the same function as the behaviour

CLIENT PROFILE:
{profile}

PRACTITIONER'S DESCRIPTION:
{freetext}
{library_section}"""


def recommend_from_freetext(client_info: dict, freetext: str, api_key: str,
                              library_text: str = None) -> dict:
    profile_text = (
        f"Name: {client_info['name']}\n"
        f"Age: {client_info['age']}\n"
        f"Diagnosis/Condition: {client_info['diagnosis']}\n"
        f"Communication level: {client_info['comms']}\n"
        f"Additional context: {client_info['other']}"
    )
    if library_text:
        lib_instruction = (
            "IMPORTANT: Select strategies FROM THE STRATEGY LIBRARY provided at the end of this "
            "prompt. Quote or closely paraphrase library strategies. Where the library has no "
            "relevant strategy, suggest an evidence-based alternative and note '(not in library)'."
        )
        lib_section = f"\nSTRATEGY LIBRARY:\n{library_text[:30000]}"
    else:
        lib_instruction = "Generate evidence-based PBS strategies based on function and context."
        lib_section = ""
    prompt = FREETEXT_STRATEGY_PROMPT.format(
        profile=profile_text, freetext=freetext,
        library_instruction=lib_instruction, library_section=lib_section,
    )
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-opus-4-5", max_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = msg.content[0].text.strip()
    if raw.startswith("```"): raw = "\n".join(raw.split("\n")[1:])
    if raw.endswith("```"):   raw = "\n".join(raw.split("\n")[:-1])
    return json.loads(raw)


# ══════════════════════════════════════════════════════════════════════════════
# PDF HELPERS (shared)
# ══════════════════════════════════════════════════════════════════════════════

def drect(c, x, y, w, h, fill=None, stroke=None, lw=0.75):
    c.saveState()
    if fill:   c.setFillColor(fill)
    if stroke: c.setStrokeColor(stroke); c.setLineWidth(lw)
    c.rect(x, y, w, h, fill=1 if fill else 0, stroke=1 if stroke else 0)
    c.restoreState()

def dlbl(c, x, y, s, size, bold=False, italic=False, color=DARK_TEXT, align='left'):
    c.saveState()
    c.setFillColor(color)
    face = ('Helvetica-BoldOblique' if bold and italic else
            'Helvetica-Bold'        if bold           else
            'Helvetica-Oblique'     if italic         else 'Helvetica')
    c.setFont(face, size)
    if align == 'center': c.drawCentredString(x, y, s)
    elif align == 'right': c.drawRightString(x, y, s)
    else: c.drawString(x, y, s)
    c.restoreState()

def sec(c, y, title, bg, ht=20):
    drect(c, LM, y - ht, CW, ht, fill=bg)
    dlbl(c, LM + 8, y - ht + 6, title, 11, bold=True, color=white)
    return y - ht

def checkbox(c, x, y, size=8):
    c.saveState(); c.setStrokeColor(MID_GREY); c.setLineWidth(0.75)
    c.rect(x, y, size, size, fill=0, stroke=1); c.restoreState()

def write_line(c, x, y, w):
    c.saveState(); c.setStrokeColor(MID_GREY); c.setLineWidth(0.5)
    c.line(x, y, x + w, y); c.restoreState()

def cb_item(c, x, y, text, size=9, color=DARK_TEXT, bold=False):
    checkbox(c, x, y, 8)
    dlbl(c, x + 13, y + 1, text, size, bold=bold, color=color)

def wrap_text(text: str, font_name: str, font_size: float, max_w: float) -> list:
    words = text.split()
    lines, line = [], ""
    for word in words:
        test = (line + " " + word).strip()
        if stringWidth(test, font_name, font_size) <= max_w:
            line = test
        else:
            if line: lines.append(line)
            line = word
    if line: lines.append(line)
    return lines or [""]


# ══════════════════════════════════════════════════════════════════════════════
# SUPPORT REFERENCE CARD
# ══════════════════════════════════════════════════════════════════════════════

def generate_support_card(d: dict) -> BytesIO:
    buf = BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    RH, SH, GAP = 22, 20, 7

    ABOUT_H = SH + len(d["about"])         * RH
    WARN_H  = SH + len(d["warning_signs"]) * RH
    SIDE_H  = SH + max(len(d["triggers"]), len(d["proactive"])) * RH
    REACT_H = SH + len(d["reactive"])      * RH
    DONT_H  = SH + len(d["do_not"])        * RH
    HDR_H, FTR_H, BOT = 72, 38, 12
    y = H

    drect(c, 0, y-HDR_H, W, HDR_H, fill=DEEP_BLUE)
    drect(c, 0, y-HDR_H, 8, HDR_H, fill=TEAL)
    px, py = W-LM-58, y-HDR_H+6
    drect(c, px, py, 52, 58, fill=BLUE2, stroke=HexColor('#C8E6F5'), lw=1)
    dlbl(c, px+26, py+28, "PHOTO",      7, color=HexColor('#C8E6F5'), align='center')
    dlbl(c, px+26, py+18, "(optional)", 7, italic=True, color=HexColor('#8FC8DC'), align='center')
    dlbl(c, LM+6, y-22, d["name"],           22, bold=True, color=white)
    dlbl(c, LM+6, y-40, d.get("pronouns",""), 11, color=HexColor('#C8E6F5'))
    dlbl(c, LM+6, y-56, d.get("age_info",""), 10, color=HexColor('#8FC8DC'))
    dlbl(c, LM+6, y-70, "Positive Behaviour Support  —  Behaviour Support Reference Card",
         8, italic=True, color=HexColor('#5B9FC0'))
    y -= HDR_H + GAP

    y = sec(c, y, f"About {d['preferred']}", TEAL)
    drect(c, LM, y-(ABOUT_H-SH), CW, ABOUT_H-SH, fill=LIGHT_TEAL)
    for i, item in enumerate(d["about"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=TEAL)
        dlbl(c,  LM+21, y- 7-i*RH, item, 9.5, color=DARK_TEXT)
    y -= (ABOUT_H-SH) + GAP

    y = sec(c, y, "Early warning signs — watch for these", AMBER)
    drect(c, LM, y-(WARN_H-SH), CW, WARN_H-SH, fill=LIGHT_AMBR)
    for i, item in enumerate(d["warning_signs"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=AMBER)
        dlbl(c,  LM+21, y- 7-i*RH, item, 9.5, color=DARK_TEXT)
    y -= (WARN_H-SH) + GAP

    half, cont = (CW-5)/2, SIDE_H-SH
    drect(c, LM,       y-SH, half, SH, fill=BLUE2)
    dlbl(c, LM+8,      y-13, "Known triggers",       11, bold=True, color=white)
    drect(c, LM,       y-SIDE_H, half, cont, fill=LIGHT_BLUE)
    for i, item in enumerate(d["triggers"]):
        drect(c, LM+10, y-SH-10-i*RH, 7, 7, fill=BLUE2)
        dlbl(c,  LM+21, y-SH- 5-i*RH, item, 9.5, color=DARK_TEXT)
    rx = LM+half+5
    drect(c, rx,       y-SH, half, SH, fill=GREEN)
    dlbl(c, rx+8,      y-13, "Proactive strategies",  11, bold=True, color=white)
    drect(c, rx,       y-SIDE_H, half, cont, fill=LIGHT_GRN)
    for i, item in enumerate(d["proactive"]):
        drect(c, rx+10, y-SH-10-i*RH, 7, 7, fill=GREEN)
        dlbl(c,  rx+21, y-SH- 5-i*RH, item, 9.5, color=DARK_TEXT)
    y -= SIDE_H + GAP

    y = sec(c, y, "When behaviour occurs — do this", TEAL)
    drect(c, LM, y-(REACT_H-SH), CW, REACT_H-SH, fill=LIGHT_TEAL)
    for i, item in enumerate(d["reactive"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=TEAL)
        dlbl(c,  LM+21, y- 7-i*RH, item, 9.5, color=DARK_TEXT)
    y -= (REACT_H-SH) + GAP

    y = sec(c, y, f"DO NOT — things that escalate behaviour for {d['preferred']}", RED)
    drect(c, LM, y-(DONT_H-SH), CW, DONT_H-SH, fill=LIGHT_RED)
    for i, item in enumerate(d["do_not"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=RED)
        dlbl(c,  LM+21, y- 7-i*RH, item, 9.5, bold=True, color=RED)
    y -= (DONT_H-SH) + GAP

    drect(c, 0, BOT, W, FTR_H, fill=DEEP_BLUE)
    drect(c, 0, BOT, 8, FTR_H, fill=TEAL)
    ft = BOT + FTR_H
    dlbl(c, LM+6, ft-16,
         f"Plan review: {d.get('review_date','')}  |  {d.get('practitioner','')}", 9, bold=True, color=white)
    dlbl(c, LM+6, ft-29, d.get("contact",""), 9, italic=True, color=HexColor('#C8E6F5'))
    dlbl(c, W-LM, ft-29, "CONFIDENTIAL — handle in line with your privacy policy",
         8, italic=True, color=HexColor('#5B9FC0'), align='right')
    c.save(); buf.seek(0); return buf


# ══════════════════════════════════════════════════════════════════════════════
# ABC RECORDING FORM
# ══════════════════════════════════════════════════════════════════════════════

def generate_abc_form(d: dict) -> BytesIO:
    buf = BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    y   = H
    half_cw = (CW - 10) / 2

    drect(c, 0, y-58, W, 58, fill=DEEP_BLUE)
    drect(c, 0, y-58, 7, 58, fill=TEAL)
    dlbl(c, LM+4, y-20, "ABC Behaviour Recording Form", 16, bold=True, color=white)
    dlbl(c, LM+4, y-36,
         f"{d['name']}  |  {d.get('pronouns','')}  |  {d.get('age_info','')}", 9, color=HexColor('#C8E6F5'))
    dlbl(c, LM+4, y-52, "Complete as soon as safely possible after any behaviour of concern",
         8, italic=True, color=HexColor('#8FC8DC'))
    dlbl(c, W-LM, y-52, "CONFIDENTIAL", 8, bold=True, color=HexColor('#5B9FC0'), align='right')
    y -= 63

    drect(c, LM, y-22, CW, 22, fill=LIGHT_BLUE)
    dlbl(c, LM+6,   y-14, "Date:",           9, bold=True, color=BLUE2)
    write_line(c, LM+32,  y-15, 68)
    dlbl(c, LM+110, y-14, "Time:",           9, bold=True, color=BLUE2)
    write_line(c, LM+134, y-15, 30)
    dlbl(c, LM+170, y-14, "to",              9, color=MED_TEXT)
    write_line(c, LM+181, y-15, 30)
    dlbl(c, LM+220, y-14, "Support worker:", 9, bold=True, color=BLUE2)
    write_line(c, LM+297, y-15, 90)
    dlbl(c, LM+398, y-14, "Shift:",          9, bold=True, color=BLUE2)
    for si, sh in enumerate(["Day","Aft","Night"]):
        sx = LM+425+si*38; checkbox(c, sx, y-16); dlbl(c, sx+11, y-14, sh, 8.5, color=DARK_TEXT)
    y -= 27

    drect(c, LM, y-20, 20, 20, fill=AMBER)
    dlbl(c, LM+10, y-13, "S", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_AMBR)
    dlbl(c, LM+28, y-13, "Setting events today", 11, bold=True, color=AMBER)
    y -= 20
    se = d.get("setting_events_checklist", [])
    se_h = 10 + ((len(se)+1)//2)*18 + 18
    drect(c, LM, y-se_h, CW, se_h, fill=LIGHT_AMBR)
    for i, item in enumerate(se):
        col, row = i%2, i//2
        cb_item(c, LM+6+col*(half_cw+10), y-14-row*18, item)
    dlbl(c, LM+6, y-se_h+6, "Other:", 9, color=MED_TEXT)
    write_line(c, LM+38, y-se_h+5, CW-44)
    y -= se_h + 5

    drect(c, LM, y-20, 20, 20, fill=BLUE2)
    dlbl(c, LM+10, y-13, "A", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_BLUE)
    dlbl(c, LM+28, y-13, "Antecedent — what happened immediately BEFORE?", 11, bold=True, color=BLUE2)
    y -= 20
    ant = d.get("antecedents_checklist", [])
    ant_h = 10 + ((len(ant)+1)//2)*18 + 18
    drect(c, LM, y-ant_h, CW, ant_h, fill=LIGHT_BLUE)
    for i, item in enumerate(ant):
        col, row = i%2, i//2
        cb_item(c, LM+6+col*(half_cw+10), y-14-row*18, item)
    dlbl(c, LM+6, y-ant_h+6, "Other / describe:", 9, color=MED_TEXT)
    write_line(c, LM+96, y-ant_h+5, CW-102)
    y -= ant_h + 5

    drect(c, LM, y-20, 20, 20, fill=TEAL)
    dlbl(c, LM+10, y-13, "B", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_TEAL)
    dlbl(c, LM+28, y-13, "Behaviour — what did you observe? (facts only, no interpretations)",
         11, bold=True, color=TEAL)
    y -= 20
    behs = d.get("behaviours", [])
    n_beh = len(behs)
    beh_col_w = (CW-8) / max(n_beh, 1)
    max_desc = max((len(b.get("descriptors",[])) for b in behs), default=3)
    beh_h = 18 + max_desc*16 + 24
    drect(c, LM, y-beh_h, CW, beh_h, fill=LIGHT_TEAL)
    beh_colors = [TEAL, BLUE2, GREEN, AMBER]
    for bi, beh in enumerate(behs):
        bx = LM+4+bi*(beh_col_w+4); bw = beh_col_w-4; bc = beh_colors[bi%len(beh_colors)]
        drect(c, bx, y-18, bw, 18, fill=bc)
        dlbl(c, bx+5, y-12, beh.get("label",""), 9, bold=True, color=white)
        for di, desc in enumerate(beh.get("descriptors",[])):
            cb_item(c, bx+2, y-26-di*16, desc, size=8.5)
    ir = y-beh_h+22
    drect(c, LM+4, ir, CW-8, 16, fill=HexColor('#DFF0EE'))
    dlbl(c, LM+8, ir+5, "Intensity:", 9, bold=True, color=TEAL)
    for ii, lvl in enumerate(["Mild","Moderate","Severe"]):
        sx = LM+54+ii*58; checkbox(c, sx, ir+4); dlbl(c, sx+11, ir+5, lvl, 9, color=DARK_TEXT)
    dlbl(c, LM+232, ir+5, "Duration:", 9, bold=True, color=TEAL)
    write_line(c, LM+272, ir+4, 28)
    dlbl(c, LM+305, ir+5, "mins", 9, color=MED_TEXT)
    dlbl(c, LM+345, ir+5, "Location:", 9, bold=True, color=TEAL)
    write_line(c, LM+385, ir+4, CW-203)
    dlbl(c, LM+6, y-beh_h+8, "Describe what you saw (objective language):", 9, italic=True, color=MED_TEXT)
    write_line(c, LM+6, y-beh_h+2, CW-12)
    y -= beh_h + 5

    drect(c, LM, y-20, 20, 20, fill=GREEN)
    dlbl(c, LM+10, y-13, "C", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_GRN)
    dlbl(c, LM+28, y-13, "Consequence — what happened after? What did you do?",
         11, bold=True, color=GREEN)
    y -= 20
    resp = d.get("staff_responses_checklist", [])
    resp_h = 10 + ((len(resp)+1)//2)*17 + 52
    drect(c, LM, y-resp_h, CW, resp_h, fill=LIGHT_GRN)
    dlbl(c, LM+6, y-10, "What did you do?", 9, bold=True, color=GREEN)
    for i, item in enumerate(resp):
        col, row = i%2, i//2
        cb_item(c, LM+6+col*(half_cw+10), y-22-row*17, item)
    ry = y-resp_h+48
    dlbl(c, LM+6, ry, "How did they respond?", 9, bold=True, color=GREEN)
    for ri, r in enumerate(["De-escalated < 5 mins","De-escalated < 30 mins",
                             "Continued > 30 mins","Escalated further"]):
        rx2 = LM+6+ri*134; checkbox(c, rx2, ry-14); dlbl(c, rx2+11, ry-13, r, 8.5, color=DARK_TEXT)
    dlbl(c, LM+6, ry-28, "Other:", 9, color=MED_TEXT)
    write_line(c, LM+38, ry-29, CW-44)
    y -= resp_h + 5

    drect(c, LM, y-18, CW, 18, fill=LIGHT_BLUE)
    dlbl(c, LM+6, y-12, "Additional notes:", 10, bold=True, color=BLUE2)
    y -= 18
    drect(c, LM, y-44, CW, 44, fill=HexColor('#FAFAFA'), stroke=MID_GREY, lw=0.5)
    for li in range(3): write_line(c, LM+6, y-16-li*13, CW-12)
    y -= 48

    drect(c, LM, y-28, CW, 28, fill=LIGHT_TEAL)
    for i, ch in enumerate(["Incident report completed","Handed over to next shift",
                             "Staff debrief completed","Behaviour data entered"]):
        ix = LM+6+i*(CW/4); checkbox(c, ix, y-18); dlbl(c, ix+11, y-16, ch, 8.5, color=DARK_TEXT)
    y -= 32

    drect(c, LM, y-24, CW, 24, fill=DEEP_BLUE)
    dlbl(c, LM+6,   y-10, "Signature:",   9, bold=True, color=white)
    write_line(c, LM+55,  y-11, 110)
    dlbl(c, LM+175, y-10, "Print name:",  9, bold=True, color=white)
    write_line(c, LM+227, y-11, 110)
    dlbl(c, LM+347, y-10, "Date / Time:", 9, bold=True, color=white)
    write_line(c, LM+407, y-11, CW-230)
    dlbl(c, W/2, y-20, f"{d.get('practitioner','')}  |  {d.get('contact','')}",
         7, italic=True, color=HexColor('#8FC8DC'), align='center')

    c.save(); buf.seek(0); return buf


# ══════════════════════════════════════════════════════════════════════════════
# STRATEGY REPORT PDF (Tab 2)
# ══════════════════════════════════════════════════════════════════════════════

def generate_strategy_report(result: dict, client_name: str,
                              practitioner: str = "", contact: str = "") -> BytesIO:
    buf = BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)

    BOT, FTR_H, HDR_H  = 12, 38, 68
    MINI_HDR_H          = 28
    MIN_Y               = BOT + FTR_H + 20
    RH, SH, GAP         = 17, 20, 8
    page_num            = [1]

    def draw_footer():
        drect(c, 0, BOT, W, FTR_H, fill=DEEP_BLUE)
        drect(c, 0, BOT, 8, FTR_H, fill=TEAL)
        ft = BOT + FTR_H
        dlbl(c, LM+6, ft-14, practitioner or "Positive Behaviour Support", 9, bold=True, color=white)
        if contact:
            dlbl(c, LM+6, ft-28, contact, 9, italic=True, color=HexColor('#C8E6F5'))
        dlbl(c, W-LM, ft-14, f"Page {page_num[0]}", 9, color=HexColor('#8FC8DC'), align='right')
        dlbl(c, W-LM, ft-28, "CONFIDENTIAL — handle in line with your privacy policy",
             8, italic=True, color=HexColor('#5B9FC0'), align='right')

    def new_page():
        draw_footer(); c.showPage(); page_num[0] += 1
        drect(c, 0, H-MINI_HDR_H, W, MINI_HDR_H, fill=DEEP_BLUE)
        drect(c, 0, H-MINI_HDR_H, 8, MINI_HDR_H, fill=TEAL)
        dlbl(c, LM+6, H-MINI_HDR_H+9,
             f"Behaviour Strategy Recommendations — {client_name}", 10, bold=True, color=white)
        return H - MINI_HDR_H - GAP

    def ensure(y, needed):
        return new_page() if y - needed < MIN_Y else y

    y = H
    drect(c, 0, y-HDR_H, W, HDR_H, fill=DEEP_BLUE)
    drect(c, 0, y-HDR_H, 8, HDR_H, fill=TEAL)
    dlbl(c, LM+6, y-22, "Behaviour Strategy Recommendations", 17, bold=True, color=white)
    dlbl(c, LM+6, y-42, client_name, 13, color=HexColor('#C8E6F5'))
    dlbl(c, LM+6, y-58, "Positive Behaviour Support — Strategy Guide for Support Workers",
         9, italic=True, color=HexColor('#5B9FC0'))
    dlbl(c, W-LM, y-58, f"Generated: {date.today().strftime('%d %B %Y')}",
         9, italic=True, color=HexColor('#5B9FC0'), align='right')
    y -= HDR_H + GAP

    summary = result.get("client_summary", "")
    if summary:
        lines = wrap_text(summary, "Helvetica-Oblique", 9.5, CW - 24)
        sum_h = len(lines) * 15 + 10
        y = ensure(y, sum_h + SH)
        y = sec(c, y, "Clinical Summary", TEAL)
        drect(c, LM, y-sum_h, CW, sum_h, fill=LIGHT_TEAL)
        for li, ln in enumerate(lines):
            dlbl(c, LM+10, y-13-li*15, ln, 9.5, italic=True, color=DARK_TEXT)
        y -= sum_h + GAP

    gen = result.get("general_strategies", [])
    if gen:
        gen_h = len(gen) * RH
        y = ensure(y, gen_h + SH)
        y = sec(c, y, "General Strategies — apply across all behaviours", BLUE2)
        drect(c, LM, y-gen_h, CW, gen_h, fill=LIGHT_BLUE)
        for i, item in enumerate(gen):
            drect(c, LM+10, y-9-i*RH, 7, 7, fill=BLUE2)
            dlbl(c,  LM+21, y-4-i*RH, item, 9.5, color=DARK_TEXT)
        y -= gen_h + GAP

    for beh in result.get("behaviours", []):
        b_name  = beh.get("behaviour", "Behaviour")
        b_func  = beh.get("likely_function", "")
        pros    = beh.get("proactive", [])
        reacts  = beh.get("reactive", [])
        teaches = beh.get("teach_instead", [])
        avoids  = beh.get("avoid", [])

        y = ensure(y, 26)
        drect(c, LM, y-26, CW, 26, fill=DEEP_BLUE)
        drect(c, LM, y-26,  6, 26, fill=TEAL)
        dlbl(c, LM+14, y-17, f"Behaviour: {b_name}", 12, bold=True, color=white)
        y -= 26

        y = ensure(y, 22)
        drect(c, LM, y-22, CW, 22, fill=LIGHT_AMBR)
        dlbl(c, LM+8,  y-14, "Likely function:", 9.5, bold=True, color=AMBER)
        func_lines = wrap_text(b_func, "Helvetica-Oblique", 9.5, CW - 112)
        dlbl(c, LM+106, y-14, func_lines[0] if func_lines else "", 9.5, italic=True, color=DARK_TEXT)
        y -= 26

        def draw_section(items, title, bg, light_bg, dot_col):
            nonlocal y
            if not items: return
            h = len(items) * RH
            y = ensure(y, h + SH)
            y = sec(c, y, title, bg)
            drect(c, LM, y-h, CW, h, fill=light_bg)
            for i, item in enumerate(items):
                drect(c, LM+10, y-9-i*RH, 7, 7, fill=dot_col)
                dlbl(c,  LM+21, y-4-i*RH, item, 9.5,
                     bold=(dot_col == RED), color=(dot_col if dot_col == RED else DARK_TEXT))
            y -= h + 4

        draw_section(pros,    "Proactive strategies — prevent this behaviour",         GREEN, LIGHT_GRN,  GREEN)
        draw_section(reacts,  "Reactive strategies — when this behaviour occurs",       TEAL,  LIGHT_TEAL, TEAL)
        draw_section(teaches, "Skills to build — teach as a replacement behaviour",     BLUE2, LIGHT_BLUE, BLUE2)
        draw_section(avoids,  "DO NOT — avoid these with this behaviour",               RED,   LIGHT_RED,  RED)
        y -= GAP

    draw_footer()
    c.save(); buf.seek(0); return buf


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

st.title("📋 PBS Support Tool")

# API key — sidebar (visible from both tabs)
try:
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
except Exception:
    api_key = ""

if not api_key:
    with st.sidebar:
        st.markdown("### 🔑 API Key")
        api_key = st.text_input(
            "Anthropic API key",
            type="password",
            help="Get a key at console.anthropic.com — a few cents per document",
            key="api_key_input",
        )
        st.caption("Your key is never stored.")

tab1, tab2 = st.tabs(["📄 Generate from PBSP", "🧠 Strategy Recommender"])


# ── TAB 1 ─────────────────────────────────────────────────────────────────────
with tab1:
    st.markdown(
        "Upload a client's **Positive Behaviour Support Plan** (PDF or Word) and this tool "
        "will generate a **support reference card** and an **ABC recording form** — "
        "both pre-populated with the client's specific information."
    )
    st.divider()

    uploaded = st.file_uploader("Upload PBSP (PDF or Word .docx)", type=["pdf","docx"])
    gen_btn  = st.button("Generate documents", type="primary",
                         disabled=not (uploaded and api_key), key="gen_tab1")

    if gen_btn and uploaded and api_key:
        with st.spinner("Reading plan…"):
            fb = uploaded.read()
            pbsp_text = extract_text_from_docx(fb) if uploaded.name.lower().endswith(".docx") \
                        else extract_text_from_pdf(fb)

        with st.spinner("Extracting client information…"):
            try:   data = extract_client_data(pbsp_text, api_key)
            except Exception as e: st.error(f"Extraction failed: {e}"); st.stop()

        with st.spinner("Generating PDFs…"):
            try:
                card_buf = generate_support_card(data)
                abc_buf  = generate_abc_form(data)
            except Exception as e: st.error(f"PDF error: {e}"); st.stop()

        st.session_state["t1_data"] = data   # share with Tab 2
        st.success(f"✅  Documents generated for **{data.get('name','client')}**")
        st.divider()
        safe = data.get("name","client").replace(" ","_")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 📄 Support Reference Card")
            st.markdown("One-page quick reference — print and laminate.")
            st.download_button("Download support card (PDF)", card_buf,
                               f"{safe}_Support_Card.pdf", "application/pdf",
                               use_container_width=True)
        with c2:
            st.markdown("#### 📋 ABC Recording Form")
            st.markdown("Pre-populated with this client's behaviours.")
            st.download_button("Download ABC form (PDF)", abc_buf,
                               f"{safe}_ABC_Form.pdf", "application/pdf",
                               use_container_width=True)
        with st.expander("Review extracted information"):
            st.json(data)

    st.divider()
    st.caption("Always review generated documents before distributing to staff. "
               "Handle all client documents in line with your organisation's privacy policy.")


# ── TAB 2 ─────────────────────────────────────────────────────────────────────
with tab2:
    st.markdown(
        "Upload a client's PBSP and your **strategy library** — the tool will match the most "
        "appropriate strategies from your library to this client's behaviours and generate a "
        "**printable strategy report**."
    )
    st.divider()

    # ── STEP 1: Client source ─────────────────────────────────────────────────
    st.markdown("#### Step 1 — Client information")

    source_opts = ["📄 Upload a PBSP (auto-extract)", "✏️ Enter manually"]
    if "t1_data" in st.session_state:
        source_opts.insert(0,
            f"♻️ Use client from Tab 1 ({st.session_state['t1_data'].get('name','')})")
    src = st.radio("Where is the client information coming from?",
                   source_opts, key="sr_src", horizontal=True)

    client_info    = None   # resolved below
    valid_behs     = []     # resolved below
    freetext_value = ""     # set only in free-text manual mode

    src = src or source_opts[0]
    if src.startswith("♻️"):
        # ── reuse Tab 1 extraction ──
        t1 = st.session_state["t1_data"]
        client_info, valid_behs = pbsp_to_sr_format(t1)
        st.success(f"✅  Using: **{t1.get('name','')}** — {t1.get('age_info','')}")
        with st.expander("View extracted profile and behaviours"):
            cols = st.columns(2)
            with cols[0]:
                st.markdown(f"**Name:** {t1.get('name','')}")
                st.markdown(f"**Profile:** {t1.get('age_info','')}")
                st.markdown("**About:**")
                for a in t1.get("about",[]): st.markdown(f"- {a}")
            with cols[1]:
                st.markdown("**Behaviours of concern:**")
                for b in t1.get("behaviours",[]): st.markdown(f"- **{b.get('label','')}:** " +
                    ", ".join(b.get("descriptors",[])))
                st.markdown("**Triggers:**")
                for t in t1.get("triggers",[]): st.markdown(f"- {t}")

    elif src.startswith("📄"):
        # ── upload PBSP for Tab 2 ──
        sr_pbsp_file = st.file_uploader("Upload client's PBSP (PDF or Word)",
                                         type=["pdf","docx"], key="sr_pbsp_upload")
        if sr_pbsp_file:
            if st.button("Extract client information", type="secondary", key="sr_pbsp_extract"):
                if not api_key:
                    st.error("API key required — enter it in the sidebar.")
                else:
                    with st.spinner("Extracting from PBSP…"):
                        fb = sr_pbsp_file.read()
                        txt = extract_text_from_docx(fb) if sr_pbsp_file.name.lower().endswith(".docx") \
                              else extract_text_from_pdf(fb)
                        try:
                            extracted = extract_client_data(txt, api_key)
                            st.session_state["sr_pbsp_data"] = extracted
                            st.success(f"✅  Extracted: {extracted.get('name','')}")
                        except Exception as e:
                            st.error(f"Extraction failed: {e}")

        if "sr_pbsp_data" in st.session_state:
            pd_ = st.session_state["sr_pbsp_data"]
            client_info, valid_behs = pbsp_to_sr_format(pd_)
            st.info(f"Using extracted data for **{pd_.get('name','')}** — "
                    f"{len(valid_behs)} behaviour(s) found.")
            with st.expander("View extracted behaviours"):
                for b in pd_.get("behaviours",[]): st.markdown(
                    f"- **{b.get('label','')}:** " + ", ".join(b.get("descriptors",[])))
            if st.button("🗑 Clear extracted data", key="sr_pbsp_clear"):
                del st.session_state["sr_pbsp_data"]; st.rerun()

    else:
        # ── manual entry ──
        ca, cb = st.columns(2)
        with ca:
            sr_name = st.text_input("Client name *", key="sr_name",
                                     placeholder="e.g. Alex Thompson")
            sr_age  = st.text_input("Age", key="sr_age", placeholder="e.g. 24")
        with cb:
            sr_diag  = st.text_input("Diagnosis / condition", key="sr_diag",
                                      placeholder="e.g. Autism Spectrum Disorder, ABI")
            sr_comms = st.selectbox("Communication level", key="sr_comms",
                                     options=["Verbal","Limited verbal","Non-verbal","Uses AAC device"])
        sr_other = st.text_area("Other relevant context (optional)", key="sr_other", height=75,
                                 placeholder="e.g. Routines are very important, history of trauma…")

        st.markdown("**Behaviours of concern**")
        entry_style = st.radio(
            "How would you like to describe the behaviours?",
            ["📝 Describe in your own words", "📋 Structured entry"],
            key="sr_entry_style", horizontal=True,
        )

        if (entry_style or "").startswith("📝"):
            # ── free-text mode ──
            sr_freetext = st.text_area(
                "Describe what you've observed — write naturally",
                key="sr_freetext",
                height=160,
                placeholder=(
                    "Write however feels natural — describe what you see, when it happens, "
                    "and what seems to set it off. The tool will identify the behaviours, "
                    "assign clinical labels, and develop strategies.\n\n"
                    "e.g. 'Alex hits out at staff when demands are placed, especially during "
                    "transitions. He also bangs his head on surfaces when he's frustrated or "
                    "overwhelmed. Sometimes he shouts and swears when things don't go his way.'"
                ),
            )
            freetext_value = (sr_freetext or "").strip()
            # Use a sentinel so the generation block knows which path to take
            valid_behs = [{"_freetext": True}] if freetext_value else []
            if freetext_value:
                client_info = {
                    "name":      (sr_name or "").strip(),
                    "age":       (sr_age  or "").strip() or "not specified",
                    "diagnosis": (sr_diag or "").strip() or "not specified",
                    "comms":     sr_comms or "Verbal",
                    "other":     (sr_other or "").strip() or "none provided",
                }

        else:
            # ── structured entry ──
            freetext_value = ""
            if "sr_n" not in st.session_state:
                st.session_state.sr_n = 1

            beh_raw = []
            for i in range(st.session_state.sr_n):
                with st.expander(f"Behaviour {i+1}", expanded=True):
                    bn = st.text_input("Behaviour name / label", key=f"sr_bn_{i}",
                                        placeholder="e.g. Physical aggression, Self-injurious behaviour")
                    bd = st.text_area("Describe what it looks like", key=f"sr_bd_{i}", height=70,
                                       placeholder="e.g. Hitting out, throwing objects — lasts 2–5 minutes")
                    bt = st.text_area("Known triggers / when it tends to occur", key=f"sr_bt_{i}", height=70,
                                       placeholder="e.g. When demands are placed, during transitions")
                    beh_raw.append({"name": bn or "", "description": bd or "", "triggers": bt or ""})

            caddbtn, crmbtn = st.columns(2)
            with caddbtn:
                if st.session_state.sr_n < 5 and st.button("＋ Add another behaviour", key="sr_add"):
                    st.session_state.sr_n += 1; st.rerun()
            with crmbtn:
                if st.session_state.sr_n > 1 and st.button("− Remove last", key="sr_rm"):
                    st.session_state.sr_n -= 1; st.rerun()

            valid_behs = [b for b in beh_raw if b["name"].strip()]
            if valid_behs:
                client_info = {
                    "name":      (sr_name or "").strip(),
                    "age":       (sr_age  or "").strip() or "not specified",
                    "diagnosis": (sr_diag or "").strip() or "not specified",
                    "comms":     sr_comms or "Verbal",
                    "other":     (sr_other or "").strip() or "none provided",
                }

    st.divider()

    # ── STEP 2: Strategy library ──────────────────────────────────────────────
    st.markdown("#### Step 2 — Your strategy library")
    st.markdown(
        "Upload your organisation's approved strategy library or PBS manual. "
        "The tool will select strategies **from your library** to match this client's behaviours."
    )

    sr_lib_file = st.file_uploader(
        "Upload strategy library (PDF or Word)", type=["pdf","docx"], key="sr_lib_upload",
        help="Strategies, clinical guidelines, PBS manuals, or any document containing your approved strategies",
    )
    if sr_lib_file:
        fb = sr_lib_file.read()
        lib_text = extract_text_from_docx(fb) if sr_lib_file.name.lower().endswith(".docx") \
                   else extract_text_from_pdf(fb)
        st.session_state["sr_lib_text"] = lib_text
        st.session_state["sr_lib_name"] = sr_lib_file.name

    if "sr_lib_text" in st.session_state:
        word_count = len(st.session_state["sr_lib_text"].split())
        st.success(
            f"✅  Library loaded: **{st.session_state['sr_lib_name']}** "
            f"({word_count:,} words) — strategies will be selected from this document."
        )
        if st.button("🗑 Clear library", key="sr_lib_clear"):
            del st.session_state["sr_lib_text"]
            del st.session_state["sr_lib_name"]
            st.rerun()
    else:
        st.caption(
            "No library uploaded — Claude will generate evidence-based strategies instead. "
            "Upload your library to get strategies from your own clinical material."
        )

    st.divider()

    # ── STEP 3: Report footer ─────────────────────────────────────────────────
    st.markdown("#### Step 3 — Report footer (optional)")
    cp, cc = st.columns(2)
    with cp: sr_prac    = st.text_input("Your name / title", key="sr_prac",
                                         placeholder="e.g. Janine Hogg — PBS Practitioner")
    with cc: sr_contact = st.text_input("Contact email", key="sr_contact",
                                         placeholder="e.g. janine@org.com.au")

    st.divider()

    # ── GENERATE ──────────────────────────────────────────────────────────────
    can_go  = bool(client_info and valid_behs and api_key)
    rec_btn = st.button("Generate strategy recommendations", type="primary",
                         disabled=not can_go, key="sr_gen")

    if not api_key:
        st.info("Enter your Anthropic API key in the sidebar to enable this tool.")
    elif not client_info or not valid_behs:
        st.info("Complete Step 1 to provide client information before generating.")

    if rec_btn and can_go:
        library_text  = st.session_state.get("sr_lib_text", None)
        lib_label     = st.session_state.get("sr_lib_name", None)
        client_label  = client_info.get("name", "client")
        is_freetext   = bool(valid_behs and valid_behs[0].get("_freetext"))

        if is_freetext:
            spinner_msg = (
                f"Reading your description and matching strategies for {client_label}…"
                if library_text else
                f"Reading your description and developing strategies for {client_label}…"
            )
        else:
            spinner_msg = (
                f"Matching strategies from your library for {client_label}…"
                if library_text else
                f"Generating evidence-based strategies for {client_label}…"
            )

        with st.spinner(spinner_msg):
            try:
                if is_freetext:
                    result = recommend_from_freetext(
                        client_info, freetext_value, api_key, library_text)
                else:
                    result = recommend_strategies(client_info, valid_behs, api_key, library_text)
            except Exception as e:
                st.error(f"Could not generate recommendations: {e}"); st.stop()

        source_note = f" (matched from **{lib_label}**)" if lib_label else ""
        st.success(f"✅  Recommendations generated for **{client_label}**{source_note}")
        st.divider()

        # ── On-screen results ──
        if result.get("client_summary"):
            st.info(f"**Clinical summary:** {result['client_summary']}")

        gen_s = result.get("general_strategies", [])
        if gen_s:
            st.markdown("##### General strategies — apply across all behaviours")
            for s in gen_s: st.markdown(f"- {s}")

        for beh in result.get("behaviours", []):
            st.divider()
            st.markdown(f"### {beh.get('behaviour','Behaviour')}")
            if beh.get("likely_function"):
                st.markdown(
                    f"<div style='background:#FDEBD0;padding:8px 12px;border-radius:6px;"
                    f"border-left:4px solid #D4700A;margin-bottom:10px'>"
                    f"<strong style='color:#D4700A'>Likely function:</strong> "
                    f"<span style='color:#1A2B35'>{beh['likely_function']}</span></div>",
                    unsafe_allow_html=True)
            cl, cr = st.columns(2)
            with cl:
                if beh.get("proactive"):
                    st.markdown("<div style='background:#DFFAE9;padding:6px 10px;border-radius:4px;"
                                "margin-bottom:6px'><strong style='color:#27AE60'>Proactive strategies"
                                "</strong></div>", unsafe_allow_html=True)
                    for s in beh["proactive"]: st.markdown(f"- {s}")
                if beh.get("teach_instead"):
                    st.markdown("<div style='background:#DCEEF8;padding:6px 10px;border-radius:4px;"
                                "margin:10px 0 6px'><strong style='color:#1A5276'>Skills to build"
                                "</strong></div>", unsafe_allow_html=True)
                    for s in beh["teach_instead"]: st.markdown(f"- {s}")
            with cr:
                if beh.get("reactive"):
                    st.markdown("<div style='background:#E0F6F3;padding:6px 10px;border-radius:4px;"
                                "margin-bottom:6px'><strong style='color:#1A9B8A'>Reactive strategies"
                                "</strong></div>", unsafe_allow_html=True)
                    for s in beh["reactive"]: st.markdown(f"- {s}")
                if beh.get("avoid"):
                    st.markdown("<div style='background:#FADBD8;padding:6px 10px;border-radius:4px;"
                                "margin:10px 0 6px'><strong style='color:#C0392B'>DO NOT"
                                "</strong></div>", unsafe_allow_html=True)
                    for s in beh["avoid"]: st.markdown(f"- {s}")

        st.divider()
        with st.spinner("Building PDF report…"):
            try:
                report_buf = generate_strategy_report(
                    result, client_label,
                    (sr_prac or "").strip(), (sr_contact or "").strip())
            except Exception as e: st.error(f"PDF error: {e}"); st.stop()

        st.download_button(
            "📥 Download strategy report (PDF)", report_buf,
            f"{client_label.replace(' ','_')}_Strategy_Report.pdf",
            "application/pdf", use_container_width=True)

    st.divider()
    st.caption(
        "Recommendations are AI-generated. Always review before implementing. "
        "When a strategy library is uploaded, Claude selects from that document — "
        "verify that selected strategies are appropriate for this individual. "
        "This tool complements but does not replace a formal Behaviour Support Plan."
    )
