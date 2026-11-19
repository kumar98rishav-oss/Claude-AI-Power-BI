"""Build the two setup-guide PDFs from docs/SETUP_GUIDE.md.

Usage (Windows, from the repo root):  python docs/build_setup_pdfs.py
Needs: pip install markdown pymdown-extensions, and Microsoft Edge (prints the PDFs).
Writes docs/Setup_Guide_Part_A_IT_Team.pdf and docs/Setup_Guide_Part_B_Participants.pdf.
"""
import re
import subprocess
import tempfile
from pathlib import Path

import markdown

DOCS = Path(__file__).resolve().parent
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

CSS = """
@page{size:A4;margin:18mm 16mm 20mm 16mm;
 @bottom-right{content:"Page " counter(page) " of " counter(pages);font:9pt 'Segoe UI',Arial,sans-serif;color:#666}
 @bottom-left{content:"%FOOTER%";font:9pt 'Segoe UI',Arial,sans-serif;color:#999}}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:10.5pt;line-height:1.45;color:#1b1f24}
h1{font-size:20pt;margin:0 0 8px;color:#14171d}
h2{font-size:14pt;border-bottom:2px solid #0f766e;padding-bottom:3px;margin-top:22px;color:#0f766e;page-break-after:avoid}
h3{font-size:12pt;page-break-after:avoid}hr{display:none}
pre{background:#14171d;color:#e8e8e8;padding:8px 12px;border-radius:6px;font-size:9pt;white-space:pre-wrap;page-break-inside:avoid}
code{font-family:Consolas,monospace;font-size:9pt}p code,li code,td code{background:#eef0f2;padding:1px 4px;border-radius:3px}
pre code{background:none;color:inherit;padding:0}
table{border-collapse:collapse;width:100%;font-size:9.3pt;margin:8px 0}
th,td{border:1px solid #cfd4da;padding:4px 7px;text-align:left;vertical-align:top}th{background:#14171d;color:#fff}
tr{page-break-inside:avoid}ul.task-list{list-style:none;padding-left:6px}li.task-list-item input{margin-right:7px}
a{color:#0f766e}em{color:#555}
"""


def render(md_text, footer, out_pdf):
    body = markdown.markdown(md_text, extensions=["tables", "pymdownx.superfences", "pymdownx.tasklist"])
    body = body.replace("\\|", "|")
    html = f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS.replace("%FOOTER%", footer)}</style></head><body>{body}</body></html>'
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f:
        f.write(html)
    subprocess.run([EDGE, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={out_pdf}", Path(f.name).as_uri()], check=True, capture_output=True)
    print("wrote", out_pdf.name)


def main():
    md = (DOCS / "SETUP_GUIDE.md").read_text(encoding="utf-8")
    checked = re.search(r"^\*Checked on.*?\*$", md, re.M).group(0)
    part_a = md[md.index("# Part A"):md.index("# Part B")].rstrip().rstrip("-").rstrip()
    part_b = md[md.index("# Part B"):]

    # Point cross-references at the other PDF instead of "Part A/B".
    part_a = part_a.replace("(see Part B step 3)", "(see the participant guide, step 3)")
    part_a = part_a.replace("(Part B step 3)", "(participant guide, step 3)")
    part_b = part_b.replace("ask your IT team (Part A)", "ask your IT team").replace("(Part A, A3)", "(IT team guide, section A3)").replace("(Part A)", "(IT team guide)")

    intro_a = ("# Setup guide for the IT team\n\n**Claude + Power BI MCP in VS Code.** Please install and check everything "
               "below on each participant's Windows machine before the next hands-on session. Participants then follow "
               "the separate *participant guide* to sign in and connect.\n\n" + checked + "\n\n")
    intro_b = ("# Setup guide for participants\n\n**Claude + Power BI MCP in VS Code.** Follow these steps in order, "
               "after your IT team has installed the software in the separate *IT team guide*. "
               "Each step ends with a **Check**: don't move on until it passes.\n\n" + checked + "\n\n")

    render(intro_a + re.sub(r"^# Part A: For the IT team\n", "", part_a), "Setup guide for the IT team",
           DOCS / "Setup_Guide_Part_A_IT_Team.pdf")
    render(intro_b + re.sub(r"^# Part B: For participants\n+Do these steps.*?\n", "", part_b, flags=re.S),
           "Setup guide for participants", DOCS / "Setup_Guide_Part_B_Participants.pdf")


if __name__ == "__main__":
    main()
