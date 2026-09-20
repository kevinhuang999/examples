# -*- coding: utf-8 -*-
"""重造本示例的模板与数据（跑 run.py 之前先跑它一次）。

这个文件上半个干什么、下半个干什么：

    上半：建三个中心的报告模板 xlsx（source/templates/）
          —— 三个模板故意做成三种布局：字段起始行不同、标签写法不同、
             字段集合不同、表头位置不同，好让 run.py 只能按"字段标签"找位置。
    下半：造两份输入 CSV（01_raw_data/）
          —— 仪器导出的检验结果明细 + 中心模板配置；边界值都埋在这里。

用法：python source/build_fixtures.py
"""
import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = Path(__file__).resolve().parent.parent
TPL_DIR = BASE / "source" / "templates"
RAW_DIR = BASE / "01_raw_data"

THIN = Side(style="thin", color="808080")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True)
LABEL_FONT = Font(name="微软雅黑", size=10, bold=True)
BODY_FONT = Font(name="微软雅黑", size=10)
HEAD_FILL = PatternFill("solid", fgColor="DCE6F1")
TITLE_FILL = PatternFill("solid", fgColor="F2F2F2")


def _put_pair(ws, row, label, both_side=False):
    """写一对「标签 / 空值格」：标签在 A 列，值格在 B 列。"""
    ws.cell(row=row, column=1, value=label).font = LABEL_FONT
    cell = ws.cell(row=row, column=2)
    cell.font = BODY_FONT
    cell.border = BOX
    return cell


def _put_detail(ws, head_row, head_at=1, rows=12):
    """写明细区：表头一行 + 若干带边框的空行（空行撑住 max_row，脚本才知道能填几行）。"""
    heads = ["检项", "结果", "单位", "参考范围"]
    for i, text in enumerate(heads):
        c = ws.cell(row=head_row, column=head_at + i, value=text)
        c.font = LABEL_FONT
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center")
    for r in range(head_row + 1, head_row + 1 + rows):
        for i in range(len(heads)):
            c = ws.cell(row=r, column=head_at + i)
            c.font = BODY_FONT
            c.border = BOX
    return head_row + 1, rows


def make_template_a():
    """中心 A：标题占两行、字段从第 3 行起、标签是「XXX：」，国际单位，缺项留空。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "报告"
    ws.merge_cells("A1:D1")
    ws["A1"] = "样本检验结果报告"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].fill = TITLE_FILL
    for k, label in enumerate(["样本编号：", "受试者编号：", "中心编号：", "访视点：", "检测日期："]):
        _put_pair(ws, 3 + k, label)
    _put_detail(ws, 9)
    ws.column_dimensions["A"].width = 18
    for col in "BCD":
        ws.column_dimensions[col].width = 16
    wb.save(TPL_DIR / "中心A_报告模板.xlsx")


def make_template_b():
    """中心 B：字段从第 2 行起、标签带英文括号、没有「访视点」、有「中心名称」。

    值格在 B 列，标签写「样本编号 (Sample ID)：」这种双语写法。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Lab Report"
    ws.merge_cells("A1:D1")
    ws["A1"] = "Sample Test Report"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].fill = TITLE_FILL
    labels = ["样本编号 (Sample ID)：", "受试者编号 (Subject No.)：", "中心名称：", "检测日期："]
    for k, label in enumerate(labels):
        _put_pair(ws, 2 + k, label)
    _put_detail(ws, 7)
    ws.column_dimensions["A"].width = 30
    for col in "BCD":
        ws.column_dimensions[col].width = 16
    wb.save(TPL_DIR / "中心B_报告模板.xlsx")


def make_template_c():
    """中心 C：字段从第 2 行起、但**多一个「报告编号」**（数据里没有，得按规则拼），
    标签里夹了全角空格（「样本编号　：」），明细表头在第 9 行。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "报告单"
    ws.merge_cells("A1:D1")
    ws["A1"] = "检验结果报告单"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].fill = TITLE_FILL
    labels = ["报告编号：", "样本编号\u3000：", "受试者编号：", "中心编号：", "访视点：", "检测日期："]
    for k, label in enumerate(labels):
        _put_pair(ws, 2 + k, label)
    _put_detail(ws, 9)
    ws.column_dimensions["A"].width = 20
    for col in "BCD":
        ws.column_dimensions[col].width = 16
    wb.save(TPL_DIR / "中心C_报告模板.xlsx")


RESULT_COLS = ["样本编号", "受试者编号", "中心编号", "访视点", "检项",
               "结果", "单位", "参考下限", "参考上限", "检测日期"]

# 中心 A / B / C 的结果明细。结果留空 = 该项没出结果（模板要按各中心规矩写缺项）。
ROWS = [
    # —— 中心 A（国际单位，缺项留空）——
    ["0091", "0103", "A01", "V2", "葡萄糖", "5.4", "mmol/L", "3.9", "6.1", "2026-09-08"],
    ["0091", "0103", "A01", "V2", "总胆固醇", "4.9", "mmol/L", "0", "5.2", "2026-09-08"],
    ["0091", "0103", "A01", "V2", "甘油三酯", "1.6", "mmol/L", "0", "1.7", "2026-09-08"],
    ["0091", "0103", "A01", "V2", "丙氨酸氨基转移酶", "32", "U/L", "0", "40", "2026-09-08"],
    ["0091", "0103", "A01", "V2", "总胆红素", "", "μmol/L", "0", "21", "2026-09-08"],
    # 0092 一共 13 项 —— 模板 A 的明细区只有 12 行，这份要挂起，不能硬塞
    ["0092", "0104", "A01", "V3", "葡萄糖", "5.8", "mmol/L", "3.9", "6.1", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "总胆固醇", "5.6", "mmol/L", "0", "5.2", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "甘油三酯", "2.4", "mmol/L", "0", "1.7", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "高密度脂蛋白胆固醇", "1.1", "mmol/L", "1.0", "2.0", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "低密度脂蛋白胆固醇", "3.4", "mmol/L", "0", "3.4", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "丙氨酸氨基转移酶", "45", "U/L", "0", "40", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "天门冬氨酸氨基转移酶", "38", "U/L", "0", "40", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "总胆红素", "12", "μmol/L", "0", "21", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "直接胆红素", "3.1", "μmol/L", "0", "7", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "总蛋白", "72", "g/L", "65", "85", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "白蛋白", "44", "g/L", "40", "55", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "尿素", "5.1", "mmol/L", "2.9", "8.2", "2026-09-09"],
    ["0092", "0104", "A01", "V3", "肌酐", "88", "μmol/L", "57", "97", "2026-09-09"],
    # —— 中心 B（常规单位，缺项写「未做」，模板里没有「访视点」字段）——
    ["B-1007", "0205", "B02", "V1", "葡萄糖", "5.1", "mmol/L", "3.9", "6.1", "2026-09-10"],
    ["B-1007", "0205", "B02", "V1", "总胆固醇", "5.2", "mmol/L", "0", "5.2", "2026-09-10"],
    ["B-1007", "0205", "B02", "V1", "甘油三酯", "1.8", "mmol/L", "0", "1.7", "2026-09-10"],
    ["B-1007", "0205", "B02", "V1", "钾", "4.2", "mmol/L", "3.5", "5.3", "2026-09-10"],
    # B-1008 的葡萄糖单位已经是 mg/dL —— 再乘一次系数就是 1946，绝不能换
    ["B-1008", "0206", "B02", "V1", "葡萄糖", "108", "mg/dL", "70", "110", "2026-09-10"],
    ["B-1008", "0206", "B02", "V1", "甘油三酯", "1.5", "mmol/L", "0", "1.7", "2026-09-10"],
    ["B-1008", "0206", "B02", "V1", "总胆红素", "", "μmol/L", "0", "21", "2026-09-10"],
    # —— 中心 C（国际单位，缺项写「—」，模板多要一个「报告编号」）——
    ["0093", "0110", "C03", "V1", "葡萄糖", "4.7", "mmol/L", "3.9", "6.1", "2026-09-11"],
    ["0093", "0110", "C03", "V1", "总胆固醇", "4.4", "mmol/L", "0", "5.2", "2026-09-11"],
    ["0093", "0110", "C03", "V1", "甘油三酯", "0.9", "mmol/L", "0", "1.7", "2026-09-11"],
    ["0093", "0110", "C03", "V1", "总胆红素", "", "μmol/L", "0", "21", "2026-09-11"],
    ["0094", "0111", "C03", "V1", "葡萄糖", "5.0", "mmol/L", "3.9", "6.1", "2026-09-11"],
    # —— 配置表里没有的中心 D9：不猜模板，单列出来等人工处理 ——
    ["0095", "0199", "D9", "V1", "葡萄糖", "5.5", "mmol/L", "3.9", "6.1", "2026-09-12"],
    ["0095", "0199", "D9", "V1", "总蛋白质", "70", "g/L", "65", "85", "2026-09-12"],
]

CONFIG_COLS = ["中心编号", "中心名称", "模板文件", "单位制", "缺项写法", "报告编号前缀"]

CONFIG = [
    ["A01", "华东中心", "中心A_报告模板.xlsx", "国际单位", "", "A"],
    ["B02", "North Site", "中心B_报告模板.xlsx", "常规单位", "未做", "B"],
    ["C03", "华南中心", "中心C_报告模板.xlsx", "国际单位", "—", "C"],
]


def write_inputs():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    path = RAW_DIR / "检验结果_仪器导出.csv"
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(RESULT_COLS)
        w.writerows(ROWS)
        # 表尾两行：判空列是「样本编号」，所以它们不填样本编号；
        # 文字分别放在「检项」和「中心编号」列，不会被当字段又不会让整行被误读。
        w.writerow(["", "", "", "", f"合计 {len(ROWS)} 条", "", "", "", "", ""])
        w.writerow(["", "", "说明：结果由仪器软件导出，未做的项目结果列留空。", "", "", "", "", "", "", ""])

    path = RAW_DIR / "中心模板配置.csv"
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(CONFIG_COLS)
        w.writerows(CONFIG)
    print(f"已重造：{len(ROWS)} 行结果 + {len(CONFIG)} 个中心配置，以及 3 个模板")


if __name__ == "__main__":
    TPL_DIR.mkdir(parents=True, exist_ok=True)
    make_template_a()
    make_template_b()
    make_template_c()
    write_inputs()
