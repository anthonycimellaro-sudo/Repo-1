#!/usr/bin/env python3
"""
generate_pdf.py
Generates a Product Development Process RACI PDF.
Run: pip install reportlab && python generate_pdf.py
Output: KICKR_Product_Development_RACI.pdf
"""

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, PageBreak)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

PAGE = landscape(A4)
W, H = PAGE

# ── Colors ────────────────────────────────────────────────────────────────────
C_DARK_BG   = colors.HexColor('#12182E')
C_WHITE     = colors.white
C_OFF_WHITE = colors.HexColor('#F5F5F5')
C_LIGHT_ROW = colors.HexColor('#EBEBEB')
C_DARK_GRAY = colors.HexColor('#333333')
C_MED_GRAY  = colors.HexColor('#888888')
C_WAHOO_RED = colors.HexColor('#E8202F')

C_R   = colors.HexColor('#BF1F1F')
C_A   = colors.HexColor('#E07C00')
C_C   = colors.HexColor('#1A65AD')
C_I   = colors.HexColor('#888888')
C_RA  = colors.HexColor('#7A187A')

RACI_COLORS = {'R': C_R, 'A': C_A, 'C': C_C, 'I': C_I, 'R/A': C_RA}

PHASE_COLORS = [
    colors.HexColor('#1B457A'), colors.HexColor('#0D628A'),
    colors.HexColor('#0A7465'), colors.HexColor('#167A3A'),
    colors.HexColor('#5A7200'), colors.HexColor('#7A5000'),
    colors.HexColor('#8B2500'), colors.HexColor('#8B001A'),
    colors.HexColor('#4A004A'),
]

# ── RACI Data ─────────────────────────────────────────────────────────────────
PHASES = [
    {
        "name": "Discovery Phase",
        "color_idx": 0,
        "question": "Is this worth real concept work and deeper investment?",
        "activities": [
            ("Conduct Innovation Sessions",          ["I","I","R","C","C","C","C","A","",""]),
            ("Competitive Benchmarking",             ["I","I","I","R","R","R","I","A","I",""]),
            ("Technical Feasibility Assessments",    ["","A","I","R","R","R","","I","C",""]),
            ("Bench Top Concepts",                   ["","A","I","R","R","R","I","I","",""]),
            ("Platform Assessment",                  ["I","A","I","R","R","R","","C","",""]),
            ("System Mapping",                       ["","A","I","R","R","R","I","","",""]),
            ("Pre-BOM Inputs",                       ["","A","I","R","R","R","","I","","C"]),
            ("Past Problem History",                 ["I","I","I","I","I","I","I","I","R/A","I"]),
            ("Scope Regulations & Compliance",       ["","C","I","C","C","C","","R/A","I","I"]),
            ("Initial Make vs Buy Analysis",         ["I","C","I","C","C","C","","C","","R/A"]),
            ("IP Assessment",                        ["","C","I","C","C","C","C","I","","R/A"]),
            ("GATE: Discovery Exit Review",          ["A","R","R","R","R","R","R","R","R","R"]),
        ],
    },
    {
        "name": "Concept Phase",
        "color_idx": 1,
        "question": "Develop the idea to give the business clarity — should we invest further?",
        "activities": [
            ("ID Refinement",                        ["","C","I","C","C","","R/A","C","",""]),
            ("3D Print Evaluation",                  ["C","A","I","R","","","C","C","",""]),
            ("Initial CAD Assessment",               ["","A","I","R","C","","C","","",""]),
            ("Initial Failure Simulations",          ["","A","I","R","","","","","",""]),
            ("Initial DFM Review",                   ["I","R/A","I","","","","","C","C",""]),
            ("Schematic Outline",                    ["","C","I","I","R/A","C","","","",""]),
            ("Preliminary PCB Layout",               ["","C","I","I","R/A","C","","","",""]),
            ("Initial Firmware Development",         ["","I","I","","","R/A","","","",""]),
            ("Preliminary DFMEA",                    ["","A","I","R","R","R","","C","C",""]),
            ("Risk Register Established",            ["I","C","R/A","","","","C","C","C",""]),
            ("GATE: Concept Exit Review",            ["A","R","R","R","R","R","R","R","R","R"]),
        ],
    },
    {
        "name": "EVT Entry Expectations",
        "color_idx": 2,
        "question": "Do we have the required scope and is the design mature enough to begin EVT?",
        "activities": [
            ("Concept Agreement",                    ["C","R/A","I","C","C","C","R","C","",""]),
            ("BOM Review Complete",                  ["","R/A","I","C","C","C","","I","C","C"]),
            ("EVT Build Plan",                       ["","R/A","I","C","C","C","","","C","C"]),
            ("Concept CAD Frozen",                   ["","A","I","R","C","C","","","",""]),
            ("Schematic Frozen",                     ["","C","I","","R/A","C","","","",""]),
            ("Contract Manufacturing Confirmed",     ["I","C","","","","","","C","","R/A"]),
        ],
    },
    {
        "name": "EVT Phase",
        "color_idx": 3,
        "question": "Progressing design towards full validation and tooling release.",
        "activities": [
            ("Hardware Bring-Up",                    ["","A","I","","R","C","","","",""]),
            ("Firmware Development",                 ["","A","I","","C","R","","","",""]),
            ("Electrical Pre-Scan",                  ["","A","I","","R","","I","I","",""]),
            ("Assembly Verification",                ["","A","I","R","","","","","C",""]),
            ("DFM Finalized",                        ["","A","I","R","","","","","C","I"]),
            ("EoL Development",                      ["","A","I","","C","R","","I","I",""]),
            ("System Architecture & Feature Alloc.", ["","A","I","R","R","","","","",""]),
            ("DFMEA & Test Matrix Updated",          ["","R/A","I","C","C","C","","","C","C"]),
            ("Risk Register Updated",                ["I","C","R/A","","","","","","C","C"]),
            ("CM Reporting Cadence Agreement",       ["","C","R/A","","","","","","C","C"]),
            ("EVT Build",                            ["I","R/A","I","","","","","","C","C"]),
            ("GATE: EVT Exit Review",                ["A","R","R","R","R","R","R","R","R","R"]),
        ],
    },
    {
        "name": "DVT Entry Expectations",
        "color_idx": 4,
        "question": "Is the system architecture mature? Do we have a confirmed test plan?",
        "activities": [
            ("Design Status - Frozen",               ["I","A","R","C","C","C","C","C","C","C"]),
            ("Feature Functionality Complete",       ["","A","I","R","R","R","","C","",""]),
            ("Tooling Agreement",                    ["C","C","R","","","","","","C","A"]),
            ("BOM Review Complete",                  ["","R/A","I","C","C","C","","I","C","C"]),
            ("EVT Test Report Documented",           ["","A","I","R","R","R","","","C",""]),
            ("Certification Plan Documented",        ["I","C","R/A","","","","","I","C","C"]),
            ("Yield Report Reviewed",                ["I","I","I","","","","","","R/A","C"]),
        ],
    },
    {
        "name": "DVT Phase",
        "color_idx": 5,
        "question": "Progressing towards production level builds that can pass all validation and certification.",
        "activities": [
            ("Tooling Validation",                   ["I","C","I","","","","","","R/A","C"]),
            ("Reliability Testing",                  ["","C","I","","","","","","A","R"]),
            ("Process Flow Diagram Development",     ["","I","I","","","","","","R/A",""]),
            ("Establish Work Instructions",          ["","I","I","C","C","C","","","","R/A"]),
            ("Develop PFMEA",                        ["","R","I","C","C","C","","A","C",""]),
            ("Write Control Plan & Production Docs", ["","","I","","","","","","R/A","C"]),
            ("Determine Cosmetic Limit",             ["","C","I","","","","C","","R/A","C"]),
            ("DVT Build",                            ["I","R/A","I","","","","","","C","C"]),
            ("DVT Testing",                          ["I","R/A","I","","","","","","C","C"]),
            ("Regulatory Certification",             ["","A","R","","","","","I","",""]),
            ("Confirm Firmware Stability",           ["","I","I","","","R/A","","","",""]),
            ("Yield Analysis",                       ["","I","I","","","","","","R/A","C"]),
            ("Packaging Testing",                    ["","","I","","","","","","R/A",""]),
            ("GATE: DVT Exit Review",                ["A","R","R","R","R","R","R","R","R","R"]),
        ],
    },
    {
        "name": "DVT to PVT Handover",
        "color_idx": 6,
        "question": "Have we identified all engineering risks and provided support needed for manufacturing?",
        "activities": [
            ("Tooling Approval",                     ["","C","I","","","","","","R/A",""]),
            ("Regulatory Certifications Confirmed",  ["","A","R","","","","","I","",""]),
            ("Reliability Test Report Reviewed",     ["I","R/A","I","","","","","","C","C"]),
            ("Firmware Release Candidate Ready",     ["","I","I","","","R/A","","","C","C"]),
            ("Quality Plan Reviewed",                ["","","I","","","","","","R/A",""]),
            ("Run at Rate Validation",               ["","I","I","","","","","","R/A","C"]),
            ("Target Yield Met",                     ["","I","I","","","","","","R/A","C"]),
            ("GATE: PVT Handover Review",            ["A","R","R","R","R","R","R","R","R","R"]),
        ],
    },
    {
        "name": "PVT Phase",
        "color_idx": 7,
        "question": "Supporting activities: final BOM, certifications, PFMEA, work instructions.",
        "activities": [
            ("Final Component Cost Analysis",        ["","R/A","I","C","C","C","","I","C","C"]),
            ("PVT Build Complete",                   ["I","R/A","I","","","","","","C","C"]),
            ("Yield Root Cause Analysis",            ["","I","I","","","","","","R/A","C"]),
            ("Limit Sample Tuning",                  ["","C","I","","","","","","R/A","C"]),
            ("Maintain Firmware Stability",          ["","I","I","","","R/A","","","C",""]),
            ("Develop Service Manual / Repair Guide",["","R/A","I","","","","C","I","",""]),
            ("Technical Training",                   ["","R/A","I","","","","C","I","",""]),
            ("Complete After Action Review",         ["C","C","R/A","C","C","C","C","C","C","C"]),
        ],
    },
    {
        "name": "MP / Launch Exit Phase",
        "color_idx": 8,
        "question": "Are we moving from growth/launch into sustain/retire mode intentionally?",
        "activities": [
            ("Product Support / Sustainment Plan",   ["A","R","I","","","","","C","C","C"]),
        ],
    },
]

ROLES = ["SVP", "Eng Dir", "PM", "Mech", "Elec", "FW", "ID/UX", "Prod Dev", "Q/Mfg", "Ops"]

# ── Styles ────────────────────────────────────────────────────────────────────
def style(name, **kw):
    return ParagraphStyle(name, **kw)

TITLE_S  = style('title',  fontSize=28, textColor=C_WHITE,   alignment=TA_CENTER, leading=34)
SUB_S    = style('sub',    fontSize=13, textColor=C_MED_GRAY, alignment=TA_CENTER)
PHASE_S  = style('phase',  fontSize=14, textColor=C_WHITE,   alignment=TA_LEFT,  leading=18)
Q_S      = style('q',      fontSize=9,  textColor=colors.HexColor('#DDDDDD'),
                 alignment=TA_LEFT, leading=12)
BODY_S   = style('body',   fontSize=9,  textColor=C_DARK_GRAY)
LEGEND_S = style('legend', fontSize=9,  textColor=C_WHITE,   alignment=TA_CENTER)

# ── Helpers ───────────────────────────────────────────────────────────────────
def p(text, s): return Paragraph(text, s)

def make_raci_table(phase):
    pc = PHASE_COLORS[phase["color_idx"]]
    acts = phase["activities"]

    col_w = [6.5*cm] + [1.65*cm]*10
    header = ["Activity"] + ROLES
    rows = [header]
    for name, raci in acts:
        rows.append([name] + raci)

    style_cmds = [
        ('FONTNAME',    (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE',    (0,0), (-1,-1), 7.5),
        ('FONTNAME',    (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',    (0,0), (-1,0),  7.5),
        ('BACKGROUND',  (0,0), (-1,0),  pc),
        ('TEXTCOLOR',   (0,0), (-1,0),  C_WHITE),
        ('ALIGN',       (0,0), (-1,-1), 'CENTER'),
        ('ALIGN',       (0,0), (0,-1),  'LEFT'),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [C_OFF_WHITE, C_LIGHT_ROW]),
        ('GRID',        (0,0), (-1,-1), 0.25, colors.HexColor('#CCCCCC')),
        ('TOPPADDING',  (0,0), (-1,-1), 3),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3),
        ('LEFTPADDING', (0,0), (0,-1),  5),
    ]

    for ri, (name, raci) in enumerate(acts):
        row = ri + 1
        is_gate = name.startswith("GATE:")
        if is_gate:
            style_cmds.append(('BACKGROUND', (0,row), (-1,row), C_DARK_BG))
            style_cmds.append(('TEXTCOLOR',  (0,row), (-1,row), C_WHITE))
            style_cmds.append(('FONTNAME',   (0,row), (-1,row), 'Helvetica-Bold'))
        for ci, val in enumerate(raci):
            if val in RACI_COLORS:
                style_cmds.append(('BACKGROUND', (ci+1,row), (ci+1,row), RACI_COLORS[val]))
                style_cmds.append(('TEXTCOLOR',  (ci+1,row), (ci+1,row), C_WHITE))
                style_cmds.append(('FONTNAME',   (ci+1,row), (ci+1,row), 'Helvetica-Bold'))

    return Table(rows, colWidths=col_w, repeatRows=1,
                 style=TableStyle(style_cmds))


def make_overview_table():
    names = ["Discovery", "Concept", "EVT Entry", "EVT", "DVT Entry",
             "DVT", "PVT Handover", "PVT", "MP/Launch"]
    questions = [
        "Worth deeper investment?",       "Should we invest further?",
        "Design mature for EVT?",          "Freeze arch, push to DVT?",
        "Arch mature? Test plan set?",    "Pass validation & certs?",
        "Factory ready for PVT?",          "Final BOM, certs, training",
        "Sustain or retire mode?",
    ]
    col_w = [(W - 4*cm) / 9] * 9
    rows = [names, questions]
    style_cmds = [
        ('FONTNAME',    (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE',    (0,0), (-1,0),  9),
        ('FONTNAME',    (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',    (0,1), (-1,1),  7.5),
        ('ALIGN',       (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',  (0,0), (-1,-1), 6),
        ('BOTTOMPADDING',(0,0),(-1,-1), 6),
        ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#444444')),
    ]
    for i, col in enumerate(PHASE_COLORS):
        style_cmds.append(('BACKGROUND', (i,0), (i,0), col))
        style_cmds.append(('TEXTCOLOR',  (i,0), (i,0), C_WHITE))
        style_cmds.append(('TEXTCOLOR',  (i,1), (i,1), colors.HexColor('#DDDDDD')))
        style_cmds.append(('BACKGROUND', (i,1), (i,1), colors.HexColor('#1E1E2E')))
    return Table(rows, colWidths=col_w, style=TableStyle(style_cmds))


def make_legend_table():
    items = [("R  Responsible", C_R), ("A  Accountable", C_A),
             ("C  Consulted",   C_C), ("I  Informed",    C_I),
             ("R/A  Both",      C_RA)]
    col_w = [3.8*cm] * 5
    rows = [[lbl for lbl, _ in items]]
    style_cmds = [
        ('FONTNAME',    (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',    (0,0), (-1,-1), 9),
        ('ALIGN',       (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING',  (0,0), (-1,-1), 6),
        ('BOTTOMPADDING',(0,0),(-1,-1), 6),
    ]
    for i, (_, col) in enumerate(items):
        style_cmds.append(('BACKGROUND', (i,0), (i,0), col))
        style_cmds.append(('TEXTCOLOR',  (i,0), (i,0), C_WHITE))
    return Table(rows, colWidths=col_w, style=TableStyle(style_cmds))


# ── Page background ───────────────────────────────────────────────────────────
def dark_bg(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C_DARK_BG)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    canvas.restoreState()

def light_bg(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C_OFF_WHITE)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    canvas.restoreState()


# ── Build ─────────────────────────────────────────────────────────────────────
def main():
    out = "KICKR_Product_Development_RACI.pdf"
    doc = SimpleDocTemplate(out, pagesize=PAGE,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.2*cm, bottomMargin=1.2*cm)

    story = []

    # ── Title page ──
    story.append(Spacer(1, 3*cm))
    story.append(p("Product Development Process", TITLE_S))
    story.append(Spacer(1, 0.4*cm))
    story.append(p("RACI Responsibility Matrix  |  Discovery through MP Launch", SUB_S))
    story.append(Spacer(1, 1.2*cm))
    story.append(make_legend_table())
    story.append(PageBreak())

    # ── Overview page ──
    story.append(p("Development Phases at a Glance", PHASE_S))
    story.append(Spacer(1, 0.4*cm))
    story.append(make_overview_table())
    story.append(Spacer(1, 0.5*cm))
    story.append(make_legend_table())
    story.append(PageBreak())

    # ── Phase pages ──
    for phase in PHASES:
        pc = PHASE_COLORS[phase["color_idx"]]
        hdr_style = style(f'hdr_{phase["color_idx"]}',
                          fontSize=14, textColor=C_WHITE,
                          backColor=pc, alignment=TA_LEFT, leading=18)
        q_style   = style(f'q_{phase["color_idx"]}',
                          fontSize=9, textColor=colors.HexColor('#DDDDDD'),
                          backColor=pc, alignment=TA_LEFT, leading=13)

        story.append(p(phase["name"], hdr_style))
        story.append(p(phase["question"], q_style))
        story.append(Spacer(1, 0.25*cm))
        story.append(make_raci_table(phase))
        story.append(PageBreak())

    doc.build(story)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
