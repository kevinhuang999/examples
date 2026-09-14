# -*- coding: utf-8 -*-
"""重造 20260914 示例的模板与原始数据（检验报告批量生成）。

这个文件上下两半分工：
  上半部分（build_template）—— 用 python-docx 把「检验报告」模板排出来：
    标题、样品编号行、样品信息表（3 行 4 列）、检验结果明细表（表头 + 一行样例行）、结论行、签字行。
    真实场景里这张模板是人工在 Word 里排好的，这里用代码生成只是为了能一条命令复跑。
    样例行里放的是 `{{检项}}` 这类占位符，run.py 拿它复制成 N 行。
  下半部分（write_samples / write_items）—— 造两份 CSV：样品信息（一个样品一行）
    和仪器导出的检项明细（长表，一个检项一行），写完在明细末尾补上仪器导出常见的尾巴。

边界是故意造的，删掉就跑不出坑了（见 README 的边界表）：
  · 样品名里带斜杠（阿莫西林/克拉维酸钾片）——直接拿去建文件会建出子目录
  · 样品编号带前导零（00903）
  · 一个样品的检项极多（12 项），一个样品只有 1 项，还有一个样品 0 项（结果还没录）
  · 检验结果为文字（符合 / 不符合）和带符号的数（<0.5）混在同一列
  · 个别格子是空的（结果还没出）
  · 模板里 `{{样品编号}}` 与 `{{样品名称}}` 故意拆成多个 run —— Word 存盘、拼写检查都会这么干
  · 明细 CSV 末尾留空行 + 一行「合计」+ 一行说明

跨平台：路径全部相对定位，不写死任何本机路径。
"""
import csv
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

# 在 source/ 里，往上一层才是包根
BASE = Path(__file__).resolve().parent.parent
RAW_DIR = BASE / "01_raw_data"
TPL_DIR = BASE / "source" / "templates"

# 样品级信息：一个样品一行
SAMPLES = [
    {"样品编号": "S2026090001", "样品名称": "阿莫西林胶囊", "批号": "250301", "规格": "0.25g",
     "送检部门": "制剂车间", "收样日期": "2026-09-01", "检测日期": "2026-09-05",
     "检验人": "黄工", "复核人": "李工"},
    {"样品编号": "S2026090002", "样品名称": "阿莫西林胶囊", "批号": "250302", "规格": "0.25g",
     "送检部门": "制剂车间", "收样日期": "2026-09-01", "检测日期": "2026-09-05",
     "检验人": "黄工", "复核人": "李工"},
    # 样品编号带前导零
    {"样品编号": "00903", "样品名称": "维生素C片", "批号": "00903", "规格": "100mg",
     "送检部门": "QC 一室", "收样日期": "2026-09-02", "检测日期": "2026-09-06",
     "检验人": "黄工", "复核人": "李工"},
    # 检项最多的那个（12 项，明细表很长）
    {"样品编号": "S2026090004", "样品名称": "复方甘草片", "批号": "250401",
     "规格": "每片含甘草流浸膏粉 112.5mg",
     "送检部门": "QC 二室", "收样日期": "2026-09-03", "检测日期": "2026-09-08",
     "检验人": "黄工", "复核人": "王工"},
    # 样品名里带斜杠；明细里有一项结果还没出
    {"样品编号": "S2026090005", "样品名称": "阿莫西林/克拉维酸钾片", "批号": "250402",
     "规格": "0.457g",
     "送检部门": "制剂车间", "收样日期": "2026-09-03", "检测日期": "2026-09-08",
     "检验人": "黄工", "复核人": "王工"},
    # 明细一行都没有（结果还没录），报告里不该留一行空占位符
    {"样品编号": "S2026090006", "样品名称": "注射用水", "批号": "250501", "规格": "—",
     "送检部门": "公用工程", "收样日期": "2026-09-04", "检测日期": "2026-09-09",
     "检验人": "黄工", "复核人": "李工"},
]

# 检项 → (标准规定, 单位, 基准结果)
ITEMS = {
    "性状":       ("白色或类白色粉末", "", "白色粉末"),
    "鉴别":       ("应与对照品一致", "", "符合"),
    "溶解度":     ("应符合规定", "", "符合"),
    "溶液颜色":   ("≤ Y6", "号", "Y4"),
    "澄清度":     ("≤ 2 号浊度标准液", "号", "1"),
    "pH 值":      ("5.0~7.0", "", "6.2"),
    "有关物质":   ("≤ 0.5%", "%", "0.12"),
    "干燥失重":   ("≤ 2.0%", "%", "0.35"),
    "炽灼残渣":   ("≤ 0.1%", "%", "0.03"),
    "重金属":     ("≤ 20ppm", "ppm", "8"),
    "微生物限度": ("≤ 1000cfu/g", "cfu/g", "120"),
    "含量测定":   ("98.0%~102.0%", "%", "99.6"),
}


# 每个样品做了哪些检项：00903 只有 1 项（极少），S2026090004 有 12 项（极多）
SAMPLE_ITEMS = {
    "S2026090001": ["性状", "鉴别", "有关物质", "含量测定", "干燥失重", "pH 值"],
    "S2026090002": ["性状", "鉴别", "有关物质", "含量测定", "干燥失重", "pH 值"],
    "00903":       ["含量测定"],
    "S2026090004": ["性状", "鉴别", "有关物质", "含量测定", "干燥失重", "溶解度",
                    "溶液颜色", "澄清度", "pH 值", "重金属", "炽灼残渣", "微生物限度"],
    "S2026090005": ["性状", "鉴别", "含量测定"],
    "S2026090006": [],
}

# 故意让一个样品不合格：S2026090001 的有关物质超标
OVER_LIMIT = {("S2026090001", "有关物质"): "0.68"}
# 故意留空：结果还没出（S2026090005 的鉴别）
PENDING = {("S2026090005", "鉴别")}


def build_template() -> Path:
    """上半部分：排出「检验报告」模板，占位符用 {{字段}}。"""
    TPL_DIR.mkdir(parents=True, exist_ok=True)
    out = TPL_DIR / "检验报告_模板.docx"

    doc = Document()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("检 验 报 告")
    run.bold = True
    run.font.size = Pt(18)

    # 这一行的占位符故意拆成三个 run：Word 存盘、拼写检查都会把 {{样品编号}} 拆开
    sid = doc.add_paragraph()
    sid.add_run("样品编号：")
    sid.add_run("{{样品")
    sid.add_run("编号}}")

    info_rows = [
        ("样品名称", "{{样品名称}}", "批号", "{{批号}}"),
        ("规格", "{{规格}}", "送检部门", "{{送检部门}}"),
        ("收样日期", "{{收样日期}}", "检测日期", "{{检测日期}}"),
    ]
    info = doc.add_table(rows=len(info_rows), cols=4)
    info.style = "Table Grid"
    for i, row in enumerate(info_rows):
        for j, text in enumerate(row):
            para = info.cell(i, j).paragraphs[0]
            if text == "{{样品名称}}":        # 这个占位符也拆开，模拟在 Word 里编辑过的模板
                para.add_run("{{样品")
                para.add_run("名称}}")
            else:
                para.add_run(text)

    doc.add_paragraph("检验结果：")
    det = doc.add_table(rows=1, cols=4)
    det.style = "Table Grid"
    for j, head in enumerate(["检项", "标准规定", "检验结果", "单位"]):
        head_run = det.cell(0, j).paragraphs[0].add_run(head)
        head_run.bold = True
    # 样例行：run.py 会把它深拷成 N 行，所以这一行的字号就是所有明细行的字号
    tpl_row = det.add_row()
    for j, placeholder in enumerate(["{{检项}}", "{{标准规定}}", "{{检验结果}}", "{{单位}}"]):
        cell_run = tpl_row.cells[j].paragraphs[0].add_run(placeholder)
        cell_run.font.size = Pt(10.5)

    doc.add_paragraph()
    doc.add_paragraph("结论：{{结论}}")
    doc.add_paragraph()
    doc.add_paragraph("检验人：{{检验人}}　　　　复核人：{{复核人}}")

    # 默认模板的属性里带着 python-docx，改成中性值，免得包一交出去带出本机信息
    doc.core_properties.author = "示例"
    doc.core_properties.last_modified_by = "示例"
    doc.save(str(out))
    return out


def make_items() -> list:
    """下半部分之一：按每个样品的检项清单拼出明细行。

    明细里比报告上多一列「判定」——工作站导出时就已经判好了，
    报告的结论行就是拿这一列算出来的，所以报告表上不用再摆一列。
    """
    rows = []
    for sample in SAMPLES:
        sid = sample["样品编号"]
        for item in SAMPLE_ITEMS[sid]:
            std, unit, base = ITEMS[item]
            key = (sid, item)
            if key in OVER_LIMIT:
                result, verdict = OVER_LIMIT[key], "不符合规定"
            elif key in PENDING:
                result, verdict = "", ""          # 结果还没出，判定也跟着空着
            else:
                result, verdict = base, "符合规定"
            rows.append({
                "样品编号": sid,
                "检项": item,
                "标准规定": std,
                "检验结果": result,
                "单位": unit,
                "判定": verdict,
            })
    return rows



def write_samples() -> Path:
    """下半部分之二：样品信息落成一个 CSV。"""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    out = RAW_DIR / "样品信息.csv"
    cols = list(SAMPLES[0].keys())
    # utf-8-sig：带 BOM，Excel 双击不按 ANSI 解码，中文才不会变乱码
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=cols)
        writer.writeheader()
        writer.writerows(SAMPLES)
    return out


def write_items(rows: list) -> Path:
    """下半部分之三：仪器导出的明细，末尾补上常见的空行 + 合计行 + 说明行。"""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    out = RAW_DIR / "检项明细_仪器导出.csv"
    cols = ["样品编号", "检项", "标准规定", "检验结果", "单位", "判定"]
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=cols)
        writer.writeheader()
        writer.writerows(rows)
        f.write("\n")
        f.write(",,,,,\n")                                   # 空行
        f.write(f"合计,,,,{len(rows)},\n")                    # 合计行：只填了条数，样品编号/检项是空的
        f.write("说明：本文件由仪器工作站导出，仅供内部核对。,,,,,\n")
    return out



if __name__ == "__main__":
    tpl = build_template()
    items = make_items()
    a = write_samples()
    b = write_items(items)
    print(f"模板 -> {tpl.name}")
    print(f"样品 {len(SAMPLES)} 个 -> {a.name}")
    print(f"明细 {len(items)} 行 -> {b.name}")
