# -*- coding: utf-8 -*-
"""造示例数据与模板。

上半个文件建模板（source/templates/客户检测数据汇总报表模板.xlsx），
下半个文件造三张 CSV（01_raw_data/）：检测结果明细、客户报表口径表、判定标准表。

故意留的边界：结果值带「<」号、结果值是文字结论、整格空；
同一样品同一项目复检重出两条；已作废行日期更新（不能被当成最新）；
客户名尾巴带半角空格与全角空格；报告还没出；判定标准表里故意漏一项；
客户没进口径表；三张表尾都带一行说明行（用来验证判空列）。

用法：在包根目录执行 python source/build_fixtures.py
"""

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = Path(__file__).resolve().parent.parent
RAW = BASE / "01_raw_data"
TPL = Path(__file__).resolve().parent / "templates" / "客户检测数据汇总报表模板.xlsx"

THIN = Side(style="thin", color="9AA0A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", start_color="DDEBF7")
TITLE_FONT = Font(bold=True, size=14)
BOLD = Font(bold=True)

INDICATORS = ["送检样品数", "检测项数", "已判定项数", "合格项数",
              "不合格项数", "未完成项数", "待确认项数", "合格率"]
GROUP_HEAD = ["检测项数", "合格", "不合格", "合格率"]
BAD_HEAD = ["样品编号", "检测项目", "结果值", "判定标准", "报告日期"]


# ---------------------------------------------------------------- 模板
def build_template():
    """建模板：一张「汇总」索引页 + 一张单客户「报表」页。

    报表页的明细空行全部画了边框——不画边框的话，openpyxl 读回来这些行根本不存在
    （max_row 只到表头），脚本会以为一行都填不了。
    """
    wb = Workbook()
    idx = wb.active
    idx.title = "汇总"
    head = ["客户名称", "送检样品数", "检测项数", "合格项数", "不合格项数", "合格率", "备注"]
    for j, name in enumerate(head):
        cell = idx.cell(1, 1 + j, name)
        cell.font = BOLD
        cell.fill = HEAD_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center")
    idx.column_dimensions["A"].width = 26
    idx.column_dimensions["G"].width = 38
    for col in "BCDEF":
        idx.column_dimensions[col].width = 13

    rep = wb.create_sheet("报表")
    rep["A1"] = "客户检测数据汇总报表"
    rep["A1"].font = TITLE_FONT
    rep.merge_cells("A1:E1")
    rep["A2"] = "客户名称："
    rep["A3"] = "统计期间："

    rep["A5"] = "一、指标汇总"
    rep["A5"].font = BOLD
    for i, name in enumerate(INDICATORS):
        rep.cell(6 + i, 1, name)

    rep["A15"] = "二、分组统计"
    rep["A15"].font = BOLD
    rep.cell(16, 1).fill = HEAD_FILL
    rep.cell(16, 1).border = BOX
    for j, name in enumerate(GROUP_HEAD):
        cell = rep.cell(16, 2 + j, name)
        cell.font = BOLD
        cell.fill = HEAD_FILL
        cell.border = BOX
    for r in range(17, 27):
        for c in range(1, 6):
            rep.cell(r, c).border = BOX
        rep.cell(r, 1).number_format = "@"

    rep["A28"] = "三、不合格项明细"
    rep["A28"].font = BOLD
    for j, name in enumerate(BAD_HEAD):
        cell = rep.cell(29, 1 + j, name)
        cell.font = BOLD
        cell.fill = HEAD_FILL
        cell.border = BOX
    for r in range(30, 36):
        for c in range(1, 6):
            rep.cell(r, c).border = BOX
        rep.cell(r, 1).number_format = "@"

    rep["A38"] = "说明："
    rep["A38"].font = BOLD
    rep["A39"] = ("本表按客户报表口径表生成；报告还没出的（未完成）与判不了标准的（待确认）"
                  "都不计入合格率，只在指标区单独计数。")
    rep.column_dimensions["A"].width = 22
    for col in "BCDE":
        rep.column_dimensions[col].width = 14

    TPL.parent.mkdir(parents=True, exist_ok=True)
    wb.save(TPL)
    return TPL


# ---------------------------------------------------------------- 数据
RESULT_HEAD = ["委托单号", "客户名称", "样品编号", "送样日期", "检测类别",
               "检测项目", "结果值", "单位", "报告状态", "报告日期"]

RESULT_ROWS = [
    # 华泰：一家客户、两个委托单，含复检重出与已作废
    ["WT2026-0000701", "华泰新材料有限公司", "0000101", "2026-09-03", "水质", "pH值", "7.2", "", "已出报告", "2026-09-08"],
    ["WT2026-0000701", "华泰新材料有限公司", "0000101", "2026-09-03", "水质", "铅", "<0.01", "mg/L", "已出报告", "2026-09-08"],
    ["WT2026-0000701", "华泰新材料有限公司", "0000102", "2026-09-03", "水质", "化学需氧量", "128", "mg/L", "已出报告", "2026-09-08"],
    ["WT2026-0000701", "华泰新材料有限公司", "0000102", "2026-09-03", "水质", "化学需氧量", "96", "mg/L", "已出报告", "2026-09-20"],
    ["WT2026-0000702", "华泰新材料有限公司", "0000103", "2026-09-10", "土壤", "铅", "156", "mg/kg", "已出报告", "2026-09-15"],
    ["WT2026-0000702", "华泰新材料有限公司", "0000103", "2026-09-10", "土壤", "镉", "0.9", "mg/kg", "已出报告", "2026-09-15"],
    ["WT2026-0000702", "华泰新材料有限公司", "0000103", "2026-09-14", "土壤", "镉", "0.3", "mg/kg", "已作废", "2026-09-16"],
    ["WT2026-0000703", "华泰新材料有限公司", "0000104", "2026-09-20", "水质", "pH值", "4.1", "", "审核中", ""],
    ["WT2026-0000703", "华泰新材料有限公司", "0000104", "2026-09-20", "水质", "铅", "0.032", "mg/L", "审核中", ""],

    # 同安：客户名尾巴带半角空格与全角空格，数量写法不改；判定标准表里漏了「总硬度」
    ["WT2026-0000704", "同安环境检测有限公司 ", "0000105", "2026-09-05", "水质", "铅", "<0.01", "mg/L", "已出报告", "2026-09-09"],
    ["WT2026-0000704", "同安环境检测有限公司\u3000", "0000105", "2026-09-05", "水质", "总硬度", "210", "mg/L", "已出报告", "2026-09-09"],
    ["WT2026-0000704", "同安环境检测有限公司", "0000106", "2026-09-05", "水质", "pH值", "7.8", "", "已出报告", "2026-09-09"],
    ["WT2026-0000705", "同安环境检测有限公司", "0000107", "2026-09-12", "土壤", "镉", "0.45", "mg/kg", "已出报告", "2026-09-17"],
    ["WT2026-0000705", "同安环境检测有限公司", "0000108", "2026-09-12", "土壤", "铅", "182", "mg/kg", "已出报告", "2026-09-17"],

    # 恒基：文字结论判定 + 结果值整格空
    ["WT2026-0000706", "恒基食品科技有限公司", "0000109", "2026-09-06", "食品", "菌落总数", "86000", "CFU/g", "已出报告", "2026-09-10"],
    ["WT2026-0000706", "恒基食品科技有限公司", "0000110", "2026-09-06", "食品", "大肠菌群", "未检出", "", "已出报告", "2026-09-10"],
    ["WT2026-0000707", "恒基食品科技有限公司", "0000111", "2026-09-18", "食品", "菌落总数", "150000", "CFU/g", "已出报告", "2026-09-22"],
    ["WT2026-0000707", "恒基食品科技有限公司", "0000111", "2026-09-18", "食品", "大肠菌群", "检出", "", "已出报告", "2026-09-22"],
    ["WT2026-0000709", "恒基食品科技有限公司", "0000113", "2026-09-21", "食品", "大肠菌群", "", "", "已出报告", "2026-09-23"],

    # 信远：这家客户没进客户报表口径表，整家挂起
    ["WT2026-0000708", "信远电子科技有限公司", "0000112", "2026-09-08", "电子", "铅含量", "860", "mg/kg", "已出报告", "2026-09-12"],

    # 表尾说明行：第一列留空，读的时候要被跳过
    ["", "", "", "", "", "", "", "", "", "合计：20 条（含已作废与报告未出）"],
]

SPEC_HEAD = ["客户名称", "分组维度", "表头组名", "含不合格明细"]
SPEC_ROWS = [
    ["华泰新材料有限公司", "检测类别", "检测类别", "是"],
    ["同安环境检测有限公司", "委托单号", "委托单号", "否"],
    ["恒基食品科技有限公司", "检测类别", "检测项目类别", "是"],
    ["", "", "", "说明：分组维度只能填「检测类别」或「委托单号」；表头组名按客户原话写，会直接印到报表上。"],
]

RULE_HEAD = ["检测类别", "检测项目", "判定方式", "下限", "上限", "合格结论文字"]
RULE_ROWS = [
    ["水质", "pH值", "区间", "6.5", "8.5", ""],
    ["水质", "铅", "上限", "", "0.01", ""],
    ["水质", "化学需氧量", "上限", "", "100", ""],
    ["土壤", "铅", "上限", "", "170", ""],
    ["土壤", "镉", "上限", "", "0.6", ""],
    ["食品", "菌落总数", "上限", "", "100000", ""],
    ["食品", "大肠菌群", "文字", "", "", "未检出"],
    ["电子", "铅含量", "上限", "", "1000", ""],
    ["", "", "", "", "", "说明：水质/总硬度故意不登记，用来验证「标准缺失要挂起、不许猜」。"],
]


def write_csv(path, head, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    # utf-8-sig：带 BOM，Excel 双击打开中文才不是乱码
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(head)
        writer.writerows(rows)
    return path


def main():
    tpl = build_template()
    print(f"模板已生成：{tpl.relative_to(BASE)}")
    for name, head, rows in (("检测结果明细.csv", RESULT_HEAD, RESULT_ROWS),
                             ("客户报表口径表.csv", SPEC_HEAD, SPEC_ROWS),
                             ("判定标准表.csv", RULE_HEAD, RULE_ROWS)):
        path = write_csv(RAW / name, head, rows)
        print(f"数据已生成：{path.relative_to(BASE)}（{len(rows)} 行，含 1 行表尾说明）")


if __name__ == "__main__":
    main()
