from pathlib import Path
import re
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, KeepTogether

ROOT = Path(r"D:\CompUse2nd")
src = ROOT / "output/pdf/2026년상시01_컴활2급_상세해설.md"
dst = ROOT / "output/pdf/2026년상시01_컴활2급_상세해설.pdf"

font = Path(r"C:\Windows\Fonts\malgun.ttf")
bold = Path(r"C:\Windows\Fonts\malgunbd.ttf")
pdfmetrics.registerFont(TTFont("Malgun", str(font)))
pdfmetrics.registerFont(TTFont("MalgunB", str(bold)))

styles = getSampleStyleSheet()
base = ParagraphStyle("KBody", parent=styles["BodyText"], fontName="Malgun", fontSize=9.2,
                      leading=14.2, textColor=colors.HexColor("#27313a"), spaceAfter=4*mm)
h1 = ParagraphStyle("KH1", parent=base, fontName="MalgunB", fontSize=19, leading=26,
                    textColor=colors.HexColor("#163a5f"), alignment=TA_CENTER, spaceAfter=8*mm)
h2 = ParagraphStyle("KH2", parent=base, fontName="MalgunB", fontSize=14, leading=20,
                    textColor=colors.white, backColor=colors.HexColor("#24577f"),
                    borderPadding=(6,8,6,8), spaceBefore=5*mm, spaceAfter=5*mm)
h3 = ParagraphStyle("KH3", parent=base, fontName="MalgunB", fontSize=11.5, leading=17,
                    textColor=colors.HexColor("#153f63"), spaceBefore=3*mm, spaceAfter=2*mm,
                    keepWithNext=True)
quote = ParagraphStyle("Quote", parent=base, backColor=colors.HexColor("#eef4f7"),
                       borderColor=colors.HexColor("#9cb5c5"), borderWidth=0.5,
                       borderPadding=7, leftIndent=4*mm, rightIndent=4*mm)
bullet = ParagraphStyle("Bullet", parent=base, leftIndent=5*mm, firstLineIndent=-3*mm,
                        bulletIndent=1*mm, spaceAfter=1.3*mm)

def inline(s):
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = re.sub(r"`([^`]+)`", r'<font name="MalgunB" color="#7a3e00">\1</font>', s)
    s = re.sub(r"\*\*([^*]+)\*\*", r'<font name="MalgunB">\1</font>', s)
    return s

story = []
lines = src.read_text(encoding="utf-8").splitlines()
i = 0
while i < len(lines):
    line = lines[i].strip()
    if not line or line == "---":
        i += 1; continue
    if line.startswith("# "):
        story += [Spacer(1, 12*mm), Paragraph(inline(line[2:]), h1)]
    elif line.startswith("## "):
        if "제1과목" in line or "제2과목" in line:
            story.append(PageBreak())
        story.append(Paragraph(inline(line[3:]), h2))
    elif line.startswith("### "):
        story.append(Paragraph(inline(line[4:]), h3))
    elif line.startswith("> "):
        story.append(Paragraph(inline(line[2:]), quote))
    elif line.startswith("- "):
        story.append(Paragraph(inline(line[2:]), bullet, bulletText="•"))
    elif line.startswith("|"):
        rows=[]
        while i < len(lines) and lines[i].strip().startswith("|"):
            cells=[c.strip() for c in lines[i].strip().strip("|").split("|")]
            if not all(re.fullmatch(r"[-:]+", c or "-") for c in cells):
                rows.append([Paragraph(inline(c), base) for c in cells])
            i += 1
        t=Table(rows, colWidths=[15*mm]+[13.5*mm]*10, repeatRows=1, hAlign="CENTER")
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#24577f")),
                               ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                               ("FONTNAME",(0,0),(-1,-1),"Malgun"),
                               ("ALIGN",(0,0),(-1,-1),"CENTER"),
                               ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                               ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#a9b6bf")),
                               ("BACKGROUND",(0,1),(-1,-1),colors.HexColor("#f4f7f9")),
                               ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4)]))
        story += [t, Spacer(1,5*mm)]
        continue
    else:
        story.append(Paragraph(inline(line), base))
    i += 1

def footer(canvas, doc):
    canvas.saveState(); canvas.setFont("Malgun", 8); canvas.setFillColor(colors.HexColor("#6b7780"))
    canvas.drawString(18*mm, 10*mm, "2026년 상시 01 컴퓨터활용능력 2급 - 학습용 상세 해설")
    canvas.drawRightString(A4[0]-18*mm, 10*mm, str(doc.page)); canvas.restoreState()

doc = SimpleDocTemplate(str(dst), pagesize=A4, leftMargin=18*mm, rightMargin=18*mm,
                        topMargin=16*mm, bottomMargin=17*mm, title="2026년 상시 01 컴활 2급 상세 해설",
                        author="OpenAI Codex")
doc.build(story, onFirstPage=footer, onLaterPages=footer)
print(dst)
