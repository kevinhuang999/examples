#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""重造 20260925 示例的模板与原始数据。

上半个文件建「报告发放登记表」模板（标题行 + 表头行 + 一行带边框的空行）；
下半个文件造三份 CSV（报告台账 / 客户命名规则 / 客户收件人）与报告原件 PDF。

跑一次就把 01_raw_data 与 source/templates 恢复成初始状态，方便反复试跑：

    python source/build_fixtures.py
"""

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "01_raw_data"
TPL_DIR = ROOT / "source" / "templates"
ORIG_DIR = RAW_DIR / "报告原件"

REPORT_CSV = RAW_DIR / "报告台账.csv"
NAMING_CSV = RAW_DIR / "客户命名规则.csv"
RECIPIENT_CSV = RAW_DIR / "客户收件人.csv"
TPL_XLSX = TPL_DIR / "报告发放登记表.xlsx"

HEADER = ["报告编号", "产品名称", "报告数量（份）", "领取单位", "领取人", "联系电话",
          "发放方式", "发放日期", "文件名", "经办人", "备注"]
WIDTHS = [16, 22, 12, 26, 10, 14, 10, 12, 40, 10, 26]

THIN = Side(style="thin", color="8C8C8C")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 原件占位 PDF：只为占位，不是真报告。别让读者以为里面有内容。
PDF_TEXT = """%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]/Contents 4 0 R>>endobj
4 0 obj<</Length 68>>stream
BT /F1 12 Tf 60 780 Td (Placeholder report file: {no}) Tj ET
endstream
endobj
trailer<</Root 1 0 R>>
%%EOF
"""

# ----------------------------------------------------------------- 上半：模板

REPORTS = [
    # 报告编号, 委托单号, 客户名称, 产品名称, 报告份数, 报告日期, 状态
    ("WX260915001", "WT26091512", "浙江宏远纺织有限公司", "铝合金/塑钢型材", "2", "2026-09-15", "已出具"),
    ("WX260915002", "WT26091513", "浙江宏远纺织有限公司", "棉麻混纺面料", "1", "2026.09.15", "已出具"),
    ("QH260916003", "WT26091620", "上海启海化工有限公司", "工业级环氧树脂", "3", "2026-09-16", "已出具"),
    ("QH260916004", "WT26091621", "上海启海化工有限公司", "工业级环氧树脂", "1", "2026/09/16", "已出具"),
    ("JD260917005", "WT26091731", "江苏金鼎建材有限公司", "建筑用钢化玻璃", "1", "2026-09-17", "已出具"),
    ("JD260917006", "WT26091732", "江苏金鼎建材有限公司", "建筑用钢化玻璃", "1", "2026-09-17", "已出具"),
    ("AZ260918007", "WT26091840", "安徽新洲新材料有限公司", "碳酸钙粉体", "1", "2026-09-18", "已出具"),
    ("AZ260918008", "WT26091841", "安徽新洲新材料有限公司", "滑石粉", "1", "2026-09-18", "已出具"),
    ("WX260918009", "WT26091850", "浙江宏远纺织有限公司", "涤纶长丝", "1", "2026-09-18", "待审核"),
    ("WX260918010", "WT26091851", "浙江宏远纺织有限公司", "涤纶长丝", "1", "", "已作废"),
    ("QH260918011", "WT26091860", "上海启海化工有限公司", "工业级环氧树脂", "1", "2026-09-18", "已出具"),
]

# 金鼎的模板故意只用「产品名 + 日期」，不含编号 —— 客户就是这么要求的，
# 两家报告同产品同日期，算出来的文件名一模一样，撞名两份都得挂起。
NAMING = [
    # 客户名称, 文件名里客户写法, 命名模板, 日期写法
    ("浙江宏远纺织有限公司", "", "{报告编号}_{产品名称}_{报告日期}", "%Y%m%d"),
    ("上海启海化工有限公司", "启海化工", "{客户}-{委托单号}-{报告编号}", "%Y-%m-%d"),
    ("江苏金鼎建材有限公司", "金鼎", "{产品名称}_{报告日期}", "%Y.%m.%d"),
    ("安徽新洲新材料有限公司", "", "{报告编号}_{客户}", "%Y%m%d"),
]

# 新洲这家在收件人表里查不到 —— 报告出来了不知道该发给谁，整家挂起。
RECIPIENTS = [
    # 客户名称, 收件人, 联系电话, 邮箱, 发放方式
    ("浙江宏远纺织有限公司", "周敏", "13800000001", "zhoumin@example.com", "电子"),
    # 这一行客户名尾巴上带个全角空格（从 Excel 粘过来的常态）——脚本要认得出是同一家。
    ("上海启海化工有限公司\u3000", "陆志远", "13900000002", "luzhiyuan@example.com", "电子"),
    ("上海启海化工有限公司", "李慧", "13900000003", "lihui@example.com", "纸质"),
    ("江苏金鼎建材有限公司", "王建国", "13900000004", "wangjianguo@example.com", "电子"),
]

# 原件目录：0000122.pdf 是去年老系统留下的编号，台账里没有 —— 单列一页，不丢也不发。
# QH260918011 台账里有、原件目录里没有 —— 挂起「原件缺失」。
ORIGINALS = ["WX260915001", "WX260915002", "QH260916003", "QH260916004",
             "JD260917005", "JD260917006", "AZ260918007", "AZ260918008",
             "WX260918010", "0000122"]


def build_template():
    TPL_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "发放登记"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADER))
    title = ws.cell(row=1, column=1, value="检测报告发放登记表")
    title.font = Font(name="等线", size=14, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    for col, (name, width) in enumerate(zip(HEADER, WIDTHS), start=1):
        cell = ws.cell(row=2, column=col, value=name)
        cell.font = Font(name="等线", size=10, bold=True)
        cell.fill = PatternFill("solid", fgColor="E8EEF7")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = width
    sample = ws.cell(row=3, column=1, value=None)
    for col in range(1, len(HEADER) + 1):
        ws.cell(row=3, column=col).border = BORDER
    del sample
    wb.save(TPL_XLSX)
    return TPL_XLSX


# ----------------------------------------------------------------- 下半：造数据


def write_csv(path, header, rows, tail=None):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
        for line in (tail or []):
            w.writerow(line)


def build_csvs():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    write_csv(REPORT_CSV, ["报告编号", "委托单号", "客户名称", "产品名称", "报告份数", "报告日期", "状态"],
              REPORTS,
              # 表尾说明行的第一列留空 —— 脚本靠"第一列是否为空"跳过它，
              # 说明写的数字全往后面几列放，不然会被当成一份报告。
              [[], ["", "说明：报告日期以最终审批日期为准；作废报告的电子件需从分发目录一并回收。",
                    "", "", "", "", ""]])
    write_csv(NAMING_CSV, ["客户名称", "文件名里客户写法", "命名模板", "日期写法"], NAMING)
    write_csv(RECIPIENT_CSV, ["客户名称", "收件人", "联系电话", "邮箱", "发放方式"], RECIPIENTS)


def build_originals():
    ORIG_DIR.mkdir(parents=True, exist_ok=True)
    for old in ORIG_DIR.glob("*"):
        old.unlink()
    for no in ORIGINALS:
        (ORIG_DIR / f"{no}.pdf").write_text(PDF_TEXT.format(no=no), encoding="ascii")


if __name__ == "__main__":
    print("模板：", build_template())
    build_csvs()
    print("数据：", REPORT_CSV, NAMING_CSV, RECIPIENT_CSV)
    build_originals()
    print(f"原件：{ORIG_DIR}（{len(ORIGINALS)} 个占位 PDF）")
