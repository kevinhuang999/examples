# -*- coding: utf-8 -*-
"""造 20261002 示例的输入数据。

上半个文件建两个源工作簿（01_raw_data/*.xlsx）——就是各实验室自己维护的台账，
表头行、列宽、文字结论、纵向合并各不一样；
下半个文件写导出规则 CSV（01_raw_data/导出规则.csv）——就是我们自己维护的那张表，
一行一个导出任务：要哪个文件、哪张工作表、横还是纵、文件名怎么拼。

故意埋的边界：批号前导零、文字结果（未检出）、空结果、表尾说明行、只有表头的空白页、
纵向合并的考察点、含 & 的长说明、列宽超一页、命名规则里带 /、工作表名不存在、文件缺失。
"""
import csv
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = Path(__file__).resolve().parent.parent      # source/ → 包根
RAW = BASE / "01_raw_data"

HEAD_FILL = PatternFill("solid", fgColor="DCE6F1")
THIN = Side(style="thin", color="8EA9DB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FONT = Font(name="微软雅黑", size=10, bold=True)
CELL_FONT = Font(name="微软雅黑", size=10)

PRODUCTS = ["对乙酰氨基酚片", "盐酸二甲双胍片", "阿莫西林胶囊", "维生素C片", "布洛芬缓释胶囊", "甲硝唑片"]
ITEMS = [("含量", "%"), ("溶出度", "%"), ("有关物质", "%"), ("干燥失重", "%"), ("性状", "—"), ("微生物限度", "cfu/g")]
ANALYSTS = ["张工", "李工", "王工", "赵工"]

RELEASE_NOTE = "说明：判定栏按内控标准执行，未检出按低于检出限处理，待复检项以复检报告为准。"


def write_head(ws, head, widths):
    """写表头行 + 设列宽（列宽单位是字符数，run.py 靠它换算 PDF 列宽）。"""
    for c, name in enumerate(head, 1):
        cell = ws.cell(1, c, name)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = widths[c - 1]
    ws.freeze_panes = "A2"


def build_release_ledger(ws):
    """放行台账：60 行明细（故意超过一页）+ 表尾一行说明（第一列留空）。"""
    write_head(
        ws,
        ["批号", "品名", "检验项目", "检验结果", "单位", "判定", "检验日期", "检验员"],
        [11, 16, 11, 11, 7, 9, 11, 9],
    )
    for i in range(60):
        r = i + 2
        item, unit = ITEMS[i % 6]
        if i % 13 == 5:
            result = None                       # 空结果
        elif i % 9 == 4:
            result = "未检出"                    # 文字结论
        else:
            result = round(95 + (i % 17) * 0.37, 4)
        if i % 23 == 7:
            verdict = "不合格"
        elif i % 11 == 3:
            verdict = "待复检"
        else:
            verdict = "合格"
        ws.cell(r, 1, "0000%03d" % (710 + i)).number_format = "@"   # 前导零走文本格式
        ws.cell(r, 2, PRODUCTS[i % 6])
        ws.cell(r, 3, item)
        ws.cell(r, 4, result).number_format = "0.0000"
        ws.cell(r, 5, unit)
        ws.cell(r, 6, verdict)
        ws.cell(r, 7, date(2026, 7, 1) + timedelta(days=i)).number_format = "yyyy-mm-dd"
        ws.cell(r, 8, ANALYSTS[i % 4])
        for c in range(1, 9):
            cell = ws.cell(r, c)
            cell.font = CELL_FONT
            cell.border = BOX
            cell.alignment = Alignment(horizontal="center" if c != 2 else "left", vertical="center")
    ws.cell(63, 2, RELEASE_NOTE).font = CELL_FONT      # 第一列（批号）留空 → 属表外说明


def build_reagent_ledger(ws):
    """试剂消耗台账：12 列，合计列宽远超一页 → 迫使脚本缩放。"""
    write_head(
        ws,
        ["领用日期", "试剂名称", "级别", "规格", "批号", "开瓶日期", "效期至",
         "领用量", "单位", "剩余量", "领用人", "备注"],
        [12, 34, 8, 24, 12, 12, 12, 9, 7, 9, 10, 18],
    )
    rows = [
        (date(2026, 9, 2), "甲醇（色谱纯）", "色谱纯", "4L/瓶", "0000712", date(2026, 9, 1), date(2027, 9, 1), 2, "瓶", 1, "李工", ""),
        (date(2026, 9, 5), "磷酸二氢钾", "分析纯", "500g/瓶", "0000713", date(2026, 9, 4), date(2028, 9, 4), 1, "瓶", 3, "王工", "受潮结块已换新"),
        (date(2026, 9, 11), "乙腈（色谱纯）", "色谱纯", "4L/瓶", "0000714", date(2026, 9, 10), date(2027, 3, 10), 3, "瓶", 0, "张工", "库存见底"),
        (date(2026, 9, 18), "标准品（对乙酰氨基酚）", "标准品", "100mg/支", "0000715", None, date(2027, 6, 30), 2, "支", 5, "赵工", ""),
        (date(2026, 9, 23), "氢氧化钠", "分析纯", "500g/瓶", "0000716", date(2026, 9, 22), date(2029, 9, 22), 4, "瓶", 8, "李工", ""),
        (date(2026, 9, 30), "纯化水", "—", "10L/桶", "0000717", date(2026, 9, 30), date(2026, 10, 7), 1, "桶", 0, "王工", "效期临近，作废处理"),
    ]
    for i, row in enumerate(rows):
        r = i + 2
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.font = CELL_FONT
            cell.border = BOX
            if c in (1, 6, 7):
                cell.number_format = "yyyy-mm-dd"
            cell.alignment = Alignment(horizontal="left" if c in (2, 4, 12) else "center", vertical="center")


def build_empty_sheet(ws):
    """只有表头、一行数据都没有的空白页——导出时必须挂起，不能出一张空 PDF。"""
    write_head(ws, ["批号", "培养基名称", "配制日期", "配制人", "有效期至", "判定"],
               [11, 22, 12, 9, 12, 9])


def build_stability(ws):
    """稳定性考察记录：考察点列纵向合并（跨 2 行），导出时靠"值下填"跨页不丢。"""
    write_head(ws, ["考察点", "取样日期", "含量(%)", "有关物质(%)", "性状", "判定"],
               [11, 12, 11, 13, 10, 9])
    rows = [
        ("0 月", date(2026, 7, 1), 99.62, 0.11, "白色片", "符合"),
        ("0 月", date(2026, 7, 8), 99.40, 0.13, "白色片", "符合"),
        ("3 月", date(2026, 10, 1), 99.05, 0.18, "白色片", "符合"),
        ("3 月", date(2026, 10, 8), 98.88, 0.21, "白色片", "符合"),
        ("6 月", date(2027, 1, 1), 97.90, 0.26, "白色片", "待评估"),
        ("6 月", date(2027, 1, 8), 97.55, 0.29, "微黄片", "待评估"),
    ]
    for i, row in enumerate(rows):
        r = i + 2
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.font = CELL_FONT
            cell.border = BOX
            if c == 2:
                cell.number_format = "yyyy-mm-dd"
            if c in (3, 4):
                cell.number_format = "0.00"
            cell.alignment = Alignment(horizontal="left" if c in (1, 5) else "center", vertical="center")
    for rng in ("A2:A3", "A4:A5", "A6:A7"):
        ws.merge_cells(rng)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")


def build_remark(ws):
    """备注说明：说明列有长文本，其中一条故意带 &（Mini-HTML 转义的坑）。"""
    write_head(ws, ["项目", "说明"], [14, 46])
    rows = [
        ("取样安排", "考察点按长期条件 25℃±2℃、RH 60%±5% 执行，每次取样后剩余样品回位存放"),
        ("判定口径", "含量不低于 95.0% 且有关物质不高于 0.5% 判符合；任一超限即启动 OOS 调查（偏差&变更同步走）"),
        ("归档要求", "本表随稳定性考察报告一并归档，保存至有效期后一年"),
    ]
    for i, (k, v) in enumerate(rows):
        r = i + 2
        for c, val in enumerate((k, v), 1):
            cell = ws.cell(r, c, val)
            cell.font = CELL_FONT
            cell.border = BOX
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def write_books():
    """上半个文件：建两个源工作簿。"""
    RAW.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    build_release_ledger(wb.active)
    wb.active.title = "放行台账"
    build_reagent_ledger(wb.create_sheet("试剂消耗台账"))
    build_empty_sheet(wb.create_sheet("培养基空白页"))
    wb.save(RAW / "检验记录台账_2026Q3.xlsx")

    wb2 = Workbook()
    build_stability(wb2.active)
    wb2.active.title = "第 3 批"
    build_remark(wb2.create_sheet("备注说明"))
    wb2.save(RAW / "稳定性考察记录_2026Q3.xlsx")
    print("已重建源工作簿：检验记录台账_2026Q3.xlsx（3 张表）、稳定性考察记录_2026Q3.xlsx（2 张表）")


def write_rules():
    """下半个文件：写导出规则 CSV。

    最后两行是故意造的坑：`报废台账` 这张工作表不存在；`试剂出入库_2026Q3.xlsx`
    这个文件根本没放进 01_raw_data。两者的挂起理由必须能分开说清。
    """
    head = ["文件", "工作表", "方向", "表头行数", "命名规则", "页码", "接收方"]
    rows = [
        ["检验记录台账_2026Q3.xlsx", "放行台账", "纵向", 1, "{源文件}_{工作表}_{导出日}", "是", "QA"],
        ["检验记录台账_2026Q3.xlsx", "试剂消耗台账", "横向", 1, "{源文件}_{工作表}", "是", "仓储"],
        ["检验记录台账_2026Q3.xlsx", "培养基空白页", "纵向", 1, "{源文件}_{工作表}", "是", "QA"],
        ["稳定性考察记录_2026Q3.xlsx", "第 3 批", "纵向", 1, "{源文件}_{工作表}_{导出日}", "是", "QA"],
        ["稳定性考察记录_2026Q3.xlsx", "备注说明", "纵向", 1, "{源文件}/{工作表}", "否", "QA"],
        ["检验记录台账_2026Q3.xlsx", "报废台账", "纵向", 1, "{源文件}_{工作表}", "是", "QA"],
        ["试剂出入库_2026Q3.xlsx", "出入库明细", "纵向", 1, "{源文件}_{工作表}", "是", "仓储"],
        ["检验记录台账_2026Q3.xlsx", "试剂消耗台账", "纵向", 1, "{源文件}_{工作表}", "是", "QA"],
    ]
    path = RAW / "导出规则.csv"
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(head)
        w.writerows(rows)
    print(f"已重建导出规则：{path.name}（{len(rows)} 条任务）")


if __name__ == "__main__":
    write_books()
    write_rules()
