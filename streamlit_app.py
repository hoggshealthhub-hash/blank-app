"""
PBS Support Tool
Upload a client's PBSP → generate a Support Reference Card + ABC Recording Form.
"""

import json
import streamlit as st
import anthropic
from io import BytesIO
from docx import Document as DocxDocument
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.colors import HexColor, white

# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PBS Support Tool",
    page_icon="📋",
    layout="centered",
)

# ── colour palette ────────────────────────────────────────────────────────────
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
    return "\n".join(
        page.extract_text() or "" for page in reader.pages
    )


# ══════════════════════════════════════════════════════════════════════════════
# CLAUDE EXTRACTION
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
  "about": [
    "up to 5 key points a support worker must know about this person"
  ],
  "warning_signs": [
    "up to 5 observable early warning signs that behaviour is building"
  ],
  "triggers": [
    "up to 6 known setting events and immediate triggers"
  ],
  "proactive": [
    "up to 5 proactive strategies — things to DO to prevent behaviour"
  ],
  "reactive": [
    "up to 5 reactive strategies — what to DO when behaviour occurs"
  ],
  "do_not": [
    "up to 4 things NOT to do — known escalators"
  ],
  "behaviours": [
    {
      "label": "short behaviour name (e.g. Verbal Outbursts)",
      "descriptors": ["up to 5 observable descriptors of this behaviour"]
    }
  ],
  "setting_events_checklist": [
    "up to 6 setting events as short checkbox labels for the ABC form"
  ],
  "antecedents_checklist": [
    "up to 8 common antecedents as short checkbox labels for the ABC form"
  ],
  "staff_responses_checklist": [
    "up to 7 common staff responses as short checkbox labels for the ABC form"
  ],
  "review_date": "review year or date",
  "practitioner": "name and title",
  "contact": "email or contact"
}

PBSP TEXT:
"""

def extract_client_data(pbsp_text: str, api_key: str) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        messages=[{
            "role": "user",
            "content": EXTRACTION_PROMPT + pbsp_text[:40000]
        }]
    )
    raw = message.content[0].text.strip()
    # Strip markdown fences if present
    if raw.startswith("```"):
        raw = "\n".join(raw.split("\n")[1:])
    if raw.endswith("```"):
        raw = "\n".join(raw.split("\n")[:-1])
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

def dlbl(c, x, y, s, size, bold=False, italic=False,
         color=DARK_TEXT, align='left'):
    c.saveState()
    c.setFillColor(color)
    face = ('Helvetica-BoldOblique' if bold and italic
            else 'Helvetica-Bold'    if bold
            else 'Helvetica-Oblique' if italic
            else 'Helvetica')
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
    c.saveState()
    c.setStrokeColor(MID_GREY)
    c.setLineWidth(0.75)
    c.rect(x, y, size, size, fill=0, stroke=1)
    c.restoreState()

def write_line(c, x, y, w):
    c.saveState()
    c.setStrokeColor(MID_GREY)
    c.setLineWidth(0.5)
    c.line(x, y, x + w, y)
    c.restoreState()

def cb_item(c, x, y, text, size=9, color=DARK_TEXT, bold=False):
    checkbox(c, x, y, 8)
    dlbl(c, x + 13, y + 1, text, size, bold=bold, color=color)


# ══════════════════════════════════════════════════════════════════════════════
# SUPPORT REFERENCE CARD
# ══════════════════════════════════════════════════════════════════════════════

def generate_support_card(d: dict) -> BytesIO:
    buf = BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    RH  = 22
    SH  = 20
    GAP = 7

    ABOUT_H = SH + len(d["about"])         * RH
    WARN_H  = SH + len(d["warning_signs"]) * RH
    SIDE_H  = SH + max(len(d["triggers"]), len(d["proactive"])) * RH
    REACT_H = SH + len(d["reactive"])      * RH
    DONT_H  = SH + len(d["do_not"])        * RH
    HDR_H   = 72
    FTR_H   = 38
    BOT     = 12

    y = H

    # Header
    drect(c, 0, y - HDR_H, W, HDR_H, fill=DEEP_BLUE)
    drect(c, 0, y - HDR_H, 8, HDR_H, fill=TEAL)
    px, py = W - LM - 58, y - HDR_H + 6
    drect(c, px, py, 52, 58, fill=BLUE2, stroke=HexColor('#C8E6F5'), lw=1)
    dlbl(c, px+26, py+28, "PHOTO",     7, color=HexColor('#C8E6F5'), align='center')
    dlbl(c, px+26, py+18, "(optional)", 7, italic=True,
         color=HexColor('#8FC8DC'), align='center')
    dlbl(c, LM+6, y-22, d["name"],     22, bold=True, color=white)
    dlbl(c, LM+6, y-40, d.get("pronouns",""), 11, color=HexColor('#C8E6F5'))
    dlbl(c, LM+6, y-56, d.get("age_info",""), 10, color=HexColor('#8FC8DC'))
    dlbl(c, LM+6, y-70,
         "Positive Behaviour Support  —  Behaviour Support Reference Card",
         8, italic=True, color=HexColor('#5B9FC0'))
    y -= HDR_H + GAP

    # About
    y = sec(c, y, f"About {d['preferred']}", TEAL)
    drect(c, LM, y-(ABOUT_H-SH), CW, ABOUT_H-SH, fill=LIGHT_TEAL)
    for i, item in enumerate(d["about"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=TEAL)
        dlbl(c, LM+21,  y- 7-i*RH, item, 9.5, color=DARK_TEXT)
    y -= (ABOUT_H-SH) + GAP

    # Warning signs
    y = sec(c, y, "Early warning signs — watch for these", AMBER)
    drect(c, LM, y-(WARN_H-SH), CW, WARN_H-SH, fill=LIGHT_AMBR)
    for i, item in enumerate(d["warning_signs"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=AMBER)
        dlbl(c, LM+21,  y- 7-i*RH, item, 9.5, color=DARK_TEXT)
    y -= (WARN_H-SH) + GAP

    # Triggers | Proactive (side by side)
    half = (CW - 5) / 2
    cont = SIDE_H - SH
    drect(c, LM, y-SH, half, SH, fill=BLUE2)
    dlbl(c, LM+8, y-13, "Known triggers", 11, bold=True, color=white)
    drect(c, LM, y-SIDE_H, half, cont, fill=LIGHT_BLUE)
    for i, item in enumerate(d["triggers"]):
        drect(c, LM+10, y-SH-10-i*RH, 7, 7, fill=BLUE2)
        dlbl(c, LM+21,  y-SH- 5-i*RH, item, 9.5, color=DARK_TEXT)
    rx = LM + half + 5
    drect(c, rx, y-SH, half, SH, fill=GREEN)
    dlbl(c, rx+8, y-13, "Proactive strategies", 11, bold=True, color=white)
    drect(c, rx, y-SIDE_H, half, cont, fill=LIGHT_GRN)
    for i, item in enumerate(d["proactive"]):
        drect(c, rx+10, y-SH-10-i*RH, 7, 7, fill=GREEN)
        dlbl(c, rx+21,  y-SH- 5-i*RH, item, 9.5, color=DARK_TEXT)
    y -= SIDE_H + GAP

    # Reactive
    y = sec(c, y, "When behaviour occurs — do this", TEAL)
    drect(c, LM, y-(REACT_H-SH), CW, REACT_H-SH, fill=LIGHT_TEAL)
    for i, item in enumerate(d["reactive"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=TEAL)
        dlbl(c, LM+21,  y- 7-i*RH, item, 9.5, color=DARK_TEXT)
    y -= (REACT_H-SH) + GAP

    # Do not
    y = sec(c, y, f"DO NOT — things that escalate behaviour for {d['preferred']}", RED)
    drect(c, LM, y-(DONT_H-SH), CW, DONT_H-SH, fill=LIGHT_RED)
    for i, item in enumerate(d["do_not"]):
        drect(c, LM+10, y-12-i*RH, 7, 7, fill=RED)
        dlbl(c, LM+21,  y- 7-i*RH, item, 9.5, bold=True, color=RED)
    y -= (DONT_H-SH) + GAP

    # Footer
    drect(c, 0, BOT, W, FTR_H, fill=DEEP_BLUE)
    drect(c, 0, BOT, 8, FTR_H, fill=TEAL)
    ft = BOT + FTR_H
    dlbl(c, LM+6, ft-16,
         f"Plan review: {d.get('review_date','')}  |  {d.get('practitioner','')}",
         9, bold=True, color=white)
    dlbl(c, LM+6, ft-29, d.get("contact",""), 9, italic=True,
         color=HexColor('#C8E6F5'))
    dlbl(c, W-LM, ft-29, "CONFIDENTIAL — handle in line with your privacy policy",
         8, italic=True, color=HexColor('#5B9FC0'), align='right')

    c.save()
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# ABC RECORDING FORM
# ══════════════════════════════════════════════════════════════════════════════

def generate_abc_form(d: dict) -> BytesIO:
    buf = BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    y   = H
    half_cw = (CW - 10) / 2

    # Header
    drect(c, 0, y-58, W, 58, fill=DEEP_BLUE)
    drect(c, 0, y-58, 7, 58, fill=TEAL)
    dlbl(c, LM+4, y-20, "ABC Behaviour Recording Form", 16, bold=True, color=white)
    dlbl(c, LM+4, y-36,
         f"{d['name']}  |  {d.get('pronouns','')}  |  {d.get('age_info','')}",
         9, color=HexColor('#C8E6F5'))
    dlbl(c, LM+4, y-52,
         "Complete as soon as safely possible after any behaviour of concern",
         8, italic=True, color=HexColor('#8FC8DC'))
    dlbl(c, W-LM, y-52, "CONFIDENTIAL", 8, bold=True,
         color=HexColor('#5B9FC0'), align='right')
    y -= 58 + 5

    # Shift row
    drect(c, LM, y-22, CW, 22, fill=LIGHT_BLUE)
    dlbl(c, LM+6,   y-14, "Date:",            9, bold=True, color=BLUE2)
    write_line(c, LM+32,  y-15, 68)
    dlbl(c, LM+110, y-14, "Time:",            9, bold=True, color=BLUE2)
    write_line(c, LM+134, y-15, 30)
    dlbl(c, LM+170, y-14, "to",               9, color=MED_TEXT)
    write_line(c, LM+181, y-15, 30)
    dlbl(c, LM+220, y-14, "Support worker:",  9, bold=True, color=BLUE2)
    write_line(c, LM+297, y-15, 90)
    dlbl(c, LM+398, y-14, "Shift:",           9, bold=True, color=BLUE2)
    for si, sh in enumerate(["Day", "Aft", "Night"]):
        sx = LM + 425 + si * 38
        checkbox(c, sx, y-16)
        dlbl(c, sx+11, y-14, sh, 8.5, color=DARK_TEXT)
    y -= 22 + 5

    # Section S — Setting events
    drect(c, LM, y-20, 20, 20, fill=AMBER)
    dlbl(c, LM+10, y-13, "S", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_AMBR)
    dlbl(c, LM+28, y-13, "Setting events today", 11, bold=True, color=AMBER)
    y -= 20
    se_items = d.get("setting_events_checklist", [])
    se_h = 10 + ((len(se_items)+1)//2) * 18 + 18
    drect(c, LM, y-se_h, CW, se_h, fill=LIGHT_AMBR)
    for i, item in enumerate(se_items):
        col, row = i % 2, i // 2
        cb_item(c, LM+6+col*(half_cw+10), y-14-row*18, item)
    dlbl(c, LM+6, y-se_h+6, "Other:", 9, color=MED_TEXT)
    write_line(c, LM+38, y-se_h+5, CW-44)
    y -= se_h + 5

    # Section A — Antecedent
    drect(c, LM, y-20, 20, 20, fill=BLUE2)
    dlbl(c, LM+10, y-13, "A", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_BLUE)
    dlbl(c, LM+28, y-13,
         "Antecedent — what happened immediately BEFORE?", 11, bold=True, color=BLUE2)
    y -= 20
    ant_items = d.get("antecedents_checklist", [])
    ant_h = 10 + ((len(ant_items)+1)//2) * 18 + 18
    drect(c, LM, y-ant_h, CW, ant_h, fill=LIGHT_BLUE)
    for i, item in enumerate(ant_items):
        col, row = i % 2, i // 2
        cb_item(c, LM+6+col*(half_cw+10), y-14-row*18, item)
    dlbl(c, LM+6, y-ant_h+6, "Other / describe:", 9, color=MED_TEXT)
    write_line(c, LM+96, y-ant_h+5, CW-102)
    y -= ant_h + 5

    # Section B — Behaviour
    drect(c, LM, y-20, 20, 20, fill=TEAL)
    dlbl(c, LM+10, y-13, "B", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_TEAL)
    dlbl(c, LM+28, y-13,
         "Behaviour — what did you observe? (facts only, no interpretations)",
         11, bold=True, color=TEAL)
    y -= 20

    behaviours = d.get("behaviours", [])
    n_beh = len(behaviours)
    beh_col_w = (CW - 8) / max(n_beh, 1)
    max_desc = max((len(b.get("descriptors", [])) for b in behaviours), default=3)
    beh_h = 18 + max_desc * 16 + 24

    drect(c, LM, y-beh_h, CW, beh_h, fill=LIGHT_TEAL)
    beh_colors = [TEAL, BLUE2, GREEN, AMBER]
    for bi, beh in enumerate(behaviours):
        bx = LM + 4 + bi * (beh_col_w + 4)
        bw = beh_col_w - 4
        bc = beh_colors[bi % len(beh_colors)]
        drect(c, bx, y-18, bw, 18, fill=bc)
        dlbl(c, bx+5, y-12, beh.get("label",""), 9, bold=True, color=white)
        for di, desc in enumerate(beh.get("descriptors", [])):
            cb_item(c, bx+2, y-26-di*16, desc, size=8.5)

    # Intensity / Duration row
    ir = y - beh_h + 22
    drect(c, LM+4, ir, CW-8, 16, fill=HexColor('#DFF0EE'))
    dlbl(c, LM+8, ir+5, "Intensity:", 9, bold=True, color=TEAL)
    for ii, lvl in enumerate(["Mild", "Moderate", "Severe"]):
        sx = LM+54 + ii*58
        checkbox(c, sx, ir+4)
        dlbl(c, sx+11, ir+5, lvl, 9, color=DARK_TEXT)
    dlbl(c, LM+232, ir+5, "Duration:", 9, bold=True, color=TEAL)
    write_line(c, LM+272, ir+4, 28)
    dlbl(c, LM+305, ir+5, "mins", 9, color=MED_TEXT)
    dlbl(c, LM+345, ir+5, "Location:", 9, bold=True, color=TEAL)
    write_line(c, LM+385, ir+4, CW-203)

    dlbl(c, LM+6, y-beh_h+8,
         "Describe what you saw (objective language):", 9, italic=True, color=MED_TEXT)
    write_line(c, LM+6, y-beh_h+2, CW-12)
    y -= beh_h + 5

    # Section C — Consequence
    drect(c, LM, y-20, 20, 20, fill=GREEN)
    dlbl(c, LM+10, y-13, "C", 11, bold=True, color=white, align='center')
    drect(c, LM+20, y-20, CW-20, 20, fill=LIGHT_GRN)
    dlbl(c, LM+28, y-13,
         "Consequence — what happened after? What did you do?",
         11, bold=True, color=GREEN)
    y -= 20
    resp_items = d.get("staff_responses_checklist", [])
    resp_h = 10 + ((len(resp_items)+1)//2) * 17 + 52
    drect(c, LM, y-resp_h, CW, resp_h, fill=LIGHT_GRN)
    dlbl(c, LM+6, y-10, "What did you do?", 9, bold=True, color=GREEN)
    for i, item in enumerate(resp_items):
        col, row = i % 2, i // 2
        cb_item(c, LM+6+col*(half_cw+10), y-22-row*17, item)
    ry = y - resp_h + 48
    dlbl(c, LM+6, ry, "How did they respond?", 9, bold=True, color=GREEN)
    for ri, resp in enumerate(["De-escalated < 5 mins",
                                "De-escalated < 30 mins",
                                "Continued > 30 mins",
                                "Escalated further"]):
        rx2 = LM + 6 + ri * 134
        checkbox(c, rx2, ry-14)
        dlbl(c, rx2+11, ry-13, resp, 8.5, color=DARK_TEXT)
    dlbl(c, LM+6, ry-28, "Other:", 9, color=MED_TEXT)
    write_line(c, LM+38, ry-29, CW-44)
    y -= resp_h + 5

    # Notes
    drect(c, LM, y-18, CW, 18, fill=LIGHT_BLUE)
    dlbl(c, LM+6, y-12, "Additional notes:", 10, bold=True, color=BLUE2)
    y -= 18
    drect(c, LM, y-44, CW, 44, fill=HexColor('#FAFAFA'),
          stroke=MID_GREY, lw=0.5)
    for li in range(3):
        write_line(c, LM+6, y-16-li*13, CW-12)
    y -= 44 + 4

    # Footer checklist + signature
    drect(c, LM, y-28, CW, 28, fill=LIGHT_TEAL)
    checks = ["Incident report completed", "Handed over to next shift",
              "Staff debrief completed", "Behaviour data entered"]
    for i, ch in enumerate(checks):
        ix = LM + 6 + i * (CW / 4)
        checkbox(c, ix, y-18)
        dlbl(c, ix+11, y-16, ch, 8.5, color=DARK_TEXT)
    y -= 28 + 4

    drect(c, LM, y-24, CW, 24, fill=DEEP_BLUE)
    dlbl(c, LM+6, y-10, "Signature:", 9, bold=True, color=white)
    write_line(c, LM+55, y-11, 110)
    dlbl(c, LM+175, y-10, "Print name:", 9, bold=True, color=white)
    write_line(c, LM+227, y-11, 110)
    dlbl(c, LM+347, y-10, "Date / Time:", 9, bold=True, color=white)
    write_line(c, LM+407, y-11, CW-230)
    dlbl(c, W/2, y-20,
         f"{d.get('practitioner','')}  |  {d.get('contact','')}",
         7, italic=True, color=HexColor('#8FC8DC'), align='center')

    c.save()
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ══════════════════════════════════════════════════════════════════════════════

st.title("📋 PBS Support Tool")
st.markdown(
    "Upload a client's **Positive Behaviour Support Plan** (PDF or Word) and this tool "
    "will generate a **support reference card** and an **ABC recording form** — "
    "both pre-populated with the client's specific information."
)
st.divider()

# API key — from Streamlit secrets if set, otherwise ask user
api_key = st.secrets.get("ANTHROPIC_API_KEY", "") if hasattr(st, "secrets") else ""
if not api_key:
    api_key = st.text_input(
        "Anthropic API key",
        type="password",
        help="Get a free key at platform.anthropic.com — costs a few cents per document",
    )

uploaded = st.file_uploader(
    "Upload PBSP (PDF or Word .docx)",
    type=["pdf", "docx"],
    help="You can upload the full plan or just the strategies / behaviours section",
)

generate_btn = st.button("Generate documents", type="primary",
                         disabled=not (uploaded and api_key))

if generate_btn and uploaded and api_key:
    with st.spinner("Reading plan…"):
        file_bytes = uploaded.read()
        if uploaded.name.lower().endswith(".docx"):
            pbsp_text = extract_text_from_docx(file_bytes)
        else:
            pbsp_text = extract_text_from_pdf(file_bytes)

    with st.spinner("Extracting client information…"):
        try:
            data = extract_client_data(pbsp_text, api_key)
        except Exception as e:
            st.error(f"Could not extract information from the plan: {e}")
            st.stop()

    with st.spinner("Generating PDFs…"):
        try:
            card_buf = generate_support_card(data)
            abc_buf  = generate_abc_form(data)
        except Exception as e:
            st.error(f"PDF generation error: {e}")
            st.stop()

    st.success(f"✅  Documents generated for **{data.get('name', 'client')}**")
    st.divider()

    col1, col2 = st.columns(2)
    safe_name = data.get("name", "client").replace(" ", "_")

    with col1:
        st.markdown("#### 📄 Support Reference Card")
        st.markdown("One-page quick reference for support workers — print and laminate.")
        st.download_button(
            label="Download support card (PDF)",
            data=card_buf,
            file_name=f"{safe_name}_Support_Card.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with col2:
        st.markdown("#### 📋 ABC Recording Form")
        st.markdown("Pre-populated recording form specific to this client's behaviours.")
        st.download_button(
            label="Download ABC form (PDF)",
            data=abc_buf,
            file_name=f"{safe_name}_ABC_Form.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with st.expander("Review extracted information"):
        st.json(data)

st.divider()
st.caption(
    "This tool uses AI to extract information from behaviour support plans. "
    "Always review generated documents before distributing to staff. "
    "Handle all client documents in line with your organisation's privacy policy."
)
