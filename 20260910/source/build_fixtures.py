# -*- coding: utf-8 -*-
"""造示例数据：模拟 QA 在 Word 里排好的报告书模板 + 仪器导出的 CSV。

运行一次即可，会在当前目录下生成 templates/ 与 data/。
真实使用时不需要这一步——模板由 QA 用 Word 排好，CSV 由仪器导出。
"""
import subprocess
import sys
from pathlib import Path

MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print(f"缺 python-docx，正在用清华源安装（解释器: {sys.executable}）...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-i", MIRROR, "python-docx"],
                   check=False)
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE = Path(__file__).resolve().parent.parent   # 示例根目录（本脚本在 source/ 里）

# ---------- 模板 ----------
doc = Document()
h = doc.add_heading("检验报告书", level=0)
h.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.add_run("样品编号：").bold = True
p.add_run("{{ 样品编号 }}")
p.add_run("    检品名称：")
p.add_run("{{ 样品名称 }}")
p.add_run("    批号：")
p.add_run("{{ 批号 }}")
for r in p.runs:
    r.font.size = Pt(10.5)

tbl = doc.add_table(rows=4, cols=4)
tbl.style = "Table Grid"
for i, t in enumerate(["检验项目", "标准规定", "检验结果", "单项结论"]):
    c = tbl.rows[0].cells[i]
    c.text = t
    for para in c.paragraphs:
        for r in para.runs:
            r.bold = True
            r.font.size = Pt(9)
# 第 2 行：只放 for
tbl.rows[1].cells[0].text = "{%tr for it in items %}"
# 第 3 行：明细内容行，模板里只有 1 行
for i, key in enumerate(["proj", "spec", "res", "concl"]):
    c = tbl.rows[2].cells[i]
    c.text = "{{ it.%s }}" % key
    for para in c.paragraphs:
        for r in para.runs:
            r.font.size = Pt(9)
# 第 4 行：只放 endfor
tbl.rows[3].cells[0].text = "{%tr endfor %}"

doc.add_paragraph("结论：{{ 结论 }}")
doc.sections[0].header.paragraphs[0].text = "报告编号：{{ 报告编号 }}"

(BASE / "source" / "templates").mkdir(parents=True, exist_ok=True)
doc.save(BASE / "source" / "templates" / "检验报告书模板.docx")

# ---------- 数据 ----------
(BASE / "01_raw_data").mkdir(parents=True, exist_ok=True)
(BASE / "01_raw_data" / "检验结果.csv").write_text(
    "SampleID,SampleName,BatchNo,检验项目,标准规定,检验结果,单项结论\n"
    "S-2026-0910-001,注射用头孢曲松钠,A20260901,性状,白色或类白色结晶性粉末,符合规定,符合\n"
    "S-2026-0910-001,注射用头孢曲松钠,A20260901,含量,含头孢曲松应为 90.0%~110.0%,98.6%,符合\n"
    "S-2026-0910-001,注射用头孢曲松钠,A20260901,有关物质,单个杂质 <0.5% & 总量 ≤2.0%,0.21%,符合\n"
    "S-2026-0910-001,注射用头孢曲松钠,A20260901,水分,不得过 11.0%,9.4%,符合\n"
    "S-2026-0910-002,注射用头孢曲松钠,A20260902,性状,白色或类白色结晶性粉末,符合规定,符合\n"
    "S-2026-0910-002,注射用头孢曲松钠,A20260902,含量,含头孢曲松应为 90.0%~110.0%,99.1%,符合\n"
    "S-2026-0910-002,注射用头孢曲松钠,A20260902,水分,不得过 11.0%,9.7%,符合\n",
    encoding="utf-8-sig",   # 带 BOM，跟仪器导出的一样
)

print("已生成 source/templates/检验报告书模板.docx 与 01_raw_data/检验结果.csv")
