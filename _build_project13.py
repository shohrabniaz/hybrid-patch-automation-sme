"""Build Project 13 submission documents."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.util import Inches, Pt as PPt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Final_Submission"
DOCS = ROOT / "01_Design"
EVAL = ROOT / "02_Evaluation"
NAVY = RGBColor(0x1F, 0x4E, 0x79)
TODAY = date.today().isoformat()


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = NAVY


def build_docs():
    DOCS.mkdir(parents=True, exist_ok=True)
    d = Document()
    d.add_heading("Patch current-state assessment", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Project 13 | {TODAY}")
    heading(d, "As-is", 1)
    d.add_paragraph(
        "The fictional SME patches Windows and third-party apps manually. There is no "
        "single inventory, no weekly upgrade ring, and no Essential Eight-aligned report."
    )
    heading(d, "Lab scope", 1)
    d.add_paragraph(
        "This pilot uses one Windows lab host with winget and PowerShell. It does not "
        "claim Intune fleet control. SaaS vendor patching stays with each vendor SLA."
    )
    d.save(DOCS / "Patch_Current_State.docx")

    d = Document()
    d.add_heading("Automated patching design and pilot", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Project 13 | {TODAY}")
    heading(d, "Design", 1)
    for x in [
        "Inventory installed software (registry + winget).",
        "List upgrade candidates with winget upgrade.",
        "Apply on a lab ring with -Apply only after dry run.",
        "Write JSON evidence and an Essential Eight patch report model.",
    ]:
        d.add_paragraph(x, style="List Number")
    heading(d, "Cycle", 1)
    d.add_paragraph(
        "Critical/exploited: target 48 hours after a smoke test. Routine third-party: "
        "weekly. OS quality updates: monthly after a pilot ring. Always keep rollback."
    )
    d.save(DOCS / "Automated_Patching_Design.docx")

    d = Document()
    d.add_heading("Evaluation and rollout plan", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Project 13 | {TODAY}")
    heading(d, "Measures", 1)
    d.add_paragraph("Coverage: share of apps inventoriable. Timeliness: days from release to apply on the lab ring. Failure rate: failed winget upgrades.")
    heading(d, "Rollout", 1)
    d.add_paragraph("Lab host → IT ring → early adopters → broad. Intune/WSUS can replace winget later without changing the reporting questions.")
    d.save(DOCS / "Evaluation_Rollout_Plan.docx")


def build_workbook():
    EVAL.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    fill = PatternFill("solid", fgColor="1F4E79")
    white = Font(color="FFFFFF", bold=True)
    ws = wb.active
    ws.title = "E8_PatchReporting"
    ws.append(["Control", "Lab indicator", "Pilot claim", "Not claimed"])
    for c in ws[1]:
        c.fill = fill
        c.font = white
    ws.append(["Patch applications", "winget upgrade list on lab host", "Third-party apps inventoriable/upgradable on one host", "Fleet maturity attestation"])
    ws.append(["Patch operating systems", "Windows Update channel note", "OS path identified", "48-hour exploited CVE proof for whole estate"])
    ws.append(["Timeliness", "JSON timestamps on inventory/upgrade", "Evidence files dated", "Production SLA compliance"])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 36
    wb.save(EVAL / "EssentialEight_Patch_Reporting.xlsx")


def build_report():
    OUT.mkdir(parents=True, exist_ok=True)
    doc = Document()
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(
        "Automating patch management for a hybrid Windows/SaaS SME\n"
        "A free-tier winget pilot with Essential Eight-aligned reporting"
    )
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = NAVY

    sections = [
        (
            "Abstract",
            "I designed a low-cost patch pilot for a fictional SME with Windows endpoints and many third-party apps. The laboratory uses PowerShell inventory, winget upgrade discovery, an optional apply step, and an Essential Eight-style report. The pilot improves visibility and creates a repeatable cycle without paid Intune. It does not attest fleet maturity.",
        ),
        (
            "1. Introduction",
            "Unpatched apps remain a common SME weakness. Manual patching does not keep pace. Project 13 asks how to automate OS and third-party patching affordably, what cycle balances speed and stability, and how to measure coverage and timeliness for Essential Eight.",
        ),
        (
            "2. Method",
            "Design-science pilot on one Windows lab host: inventory, find outdated packages, dry-run then optional apply, write JSON evidence, and map claims to Essential Eight patch applications/OS. SaaS patching stays with vendor SLAs and is reported as out of local automation scope.",
        ),
        (
            "3. Design",
            "Weekly third-party winget ring. Critical/exploited target 48 hours after smoke test. Monthly OS quality after a pilot ring. Rollback via restore point or previous package. Scripts refuse to pretend Intune coverage they do not have.",
        ),
        (
            "4. Results and limits",
            "The deliverable is a working inventory/upgrade/report path and a reporting workbook. Live apply is optional (-Apply). Limits: one host, no Intune, no production fleet, SaaS not locally patched. Next step is an Intune or WSUS ring using the same measures.",
        ),
        (
            "5. Conclusion",
            "An SME can start automated third-party patch visibility with winget and PowerShell before buying endpoint-management licences. Essential Eight reporting should state what was measured on which hosts, not claim maturity the lab did not earn.",
        ),
        (
            "References",
            "Australian Signals Directorate. (n.d.). Essential Eight. Microsoft. (n.d.). winget tool. NIST SP 800-40. ISO/IEC 27001:2022 A.8.8.",
        ),
    ]
    bw = 0
    for title, body in sections:
        heading(doc, title, 1)
        doc.add_paragraph(body)
        if title != "References":
            bw += len(body.split())
    meta = doc.paragraphs[0]
    # insert meta after title
    p = doc.add_paragraph()
    # rebuild meta as second paragraph by editing first block - add at top via clear approach
    m = Document()
    tp = m.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(
        "Automating patch management for a hybrid Windows/SaaS SME\n"
        "A free-tier winget pilot with Essential Eight-aligned reporting"
    )
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.color.rgb = NAVY
    mp = m.add_paragraph()
    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = mp.add_run(
        f"Md Shohrab Hossen Niaz\nProject 13 — Patch Management\n"
        f"Fictional SME case | Laboratory host only\n"
        f"Word count (excluding title block and reference list): {bw}"
    )
    mr.font.size = Pt(11)
    for title, body in sections:
        heading(m, title, 1)
        m.add_paragraph(body)
    path = OUT / "01_Project13_Report.docx"
    m.save(path)
    print("BODY_WORDS", bw)
    return path


def build_slides():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    navy = PRgb(0x1F, 0x4E, 0x79)

    def add(title, bullets, notes):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        b = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(1))
        r = b.text_frame.paragraphs[0].add_run()
        r.text = title
        r.font.size = PPt(28)
        r.font.bold = True
        r.font.color.rgb = navy
        body = s.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12), Inches(5))
        tf = body.text_frame
        tf.word_wrap = True
        first = True
        for line in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run = p.add_run()
            run.text = line
            run.font.size = PPt(18)
            p.space_after = PPt(10)
        s.notes_slide.notes_text_frame.text = notes

    add("Automated patching for an SME", ["Free pilot: PowerShell + winget", "Essential Eight-aligned reporting", "One lab host, not a fake fleet claim"], "Project 13.")
    add("The problem", ["Manual OS and third-party patching", "No inventory, no weekly ring", "Cannot evidence E8 patch maturity"], "Speed without visibility is not control.")
    add("The pilot", ["Inventory → outdated → dry-run → optional apply", "JSON evidence each step", "Weekly third-party; 48h critical after test"], "SaaS stays with vendor SLAs.")
    add("What I claim", ["Visibility and a repeatable cycle", "Reporting that states host scope", "Not Intune maturity for the whole firm"], "Questions?")
    path = OUT / "02_Project13_Presentation.pptx"
    prs.save(path)
    return path


def assemble():
    build_docs()
    build_workbook()
    build_report()
    build_slides()
    app_a = OUT / "Appendix" / "A_Evaluation"
    app_b = OUT / "Appendix" / "B_Design"
    app_c = OUT / "Appendix" / "C_Scripts"
    app_d = OUT / "Appendix" / "D_Evidence"
    for p in (app_a, app_b, app_c, app_d):
        p.mkdir(parents=True, exist_ok=True)
    shutil.copy2(EVAL / "EssentialEight_Patch_Reporting.xlsx", app_a / "EssentialEight_Patch_Reporting.xlsx")
    for f in DOCS.glob("*.docx"):
        shutil.copy2(f, app_b / f.name)
    for f in ROOT.glob("*.ps1"):
        if not f.name.startswith("_"):
            shutil.copy2(f, app_c / f.name)
    evid = ROOT / "evidence"
    if evid.exists():
        for f in evid.glob("*.json"):
            shutil.copy2(f, app_d / f.name)
    readme = Document()
    readme.add_heading("Project submission", 0)
    readme.add_paragraph(
        "Md Shohrab Hossen Niaz\nProject 13 — patch automation\nLaboratory host winget pilot."
    )
    readme.save(OUT / "00_Read_Me_First.docx")
    print("assembled", OUT)


if __name__ == "__main__":
    assemble()
