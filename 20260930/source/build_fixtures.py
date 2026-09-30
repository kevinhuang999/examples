# -*- coding: utf-8 -*-
"""重造模板与示例数据（第三方检测「检测周期进度跟踪台账」）。

上半个文件建模板：一页「检测进度台账」表头 + 20 行带边框的空明细行。
下半个文件造数据：检测委托台账 / 检测项目周期表 / 环节节点记录 / 工作日历 四张 CSV。

用法：python source/build_fixtures.py
"""

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = Path(__file__).resolve().parent.parent
RAW_DIR = BASE / "01_raw_data"
TPL_DIR = BASE / "source" / "templates"

HEADERS = [
    "委托单号", "客户名称", "样品编号", "检测项目", "项目数",
    "收样日期", "承诺周期(工作日)", "承诺交期", "当前环节", "已完成环节",
    "已耗工作日", "剩余工作日", "进度", "状态", "下一环节负责岗位", "备注",
]
DETAIL_ROWS = 20          # 模板预留的明细行数
WIDTHS = [20, 14, 12, 14, 8, 12, 14, 12, 12, 12, 11, 11, 9, 11, 16, 24]

THIN = Side(style="thin", color="9E9E9E")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_template():
    """建模板：标题行 + 表头行 + 20 行空明细（必须画边框）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "检测进度台账"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    title = ws.cell(row=1, column=1, value="检测进度台账")
    title.font = Font(size=14, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for c, name in enumerate(HEADERS, start=1):
        cell = ws.cell(row=2, column=c, value=name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDEBF7")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = WIDTHS[c - 1]
    ws.row_dimensions[2].height = 30

    # 明细区空行逐格画边框：不画的话 openpyxl 读回来这些行根本不存在（max_row 只到表头）
    for r in range(3, 3 + DETAIL_ROWS):
        for c in range(1, len(HEADERS) + 1):
            ws.cell(row=r, column=c).border = BORDER
    ws.freeze_panes = "A3"

    TPL_DIR.mkdir(parents=True, exist_ok=True)
    out = TPL_DIR / "检测进度台账模板.xlsx"
    wb.save(out)
    print(f"模板已生成：{out.relative_to(BASE)}（明细 {DETAIL_ROWS} 行，已画边框）")


# ---------------------------------------------------------------------------
# 下面造数据：四张 CSV 都带 BOM（双击 Excel 打开不乱码）
# ---------------------------------------------------------------------------

# 委托台账：一行一个检测项目；一张委托单可以有多个项目（周期取最长那个）
ORDERS = [
    ("WT2026-0000121", "华泰环境", "0000201", "化学需氧量", "否", "2026-09-24"),
    ("WT2026-0000121", "华泰环境", "0000201", "pH", "否", "2026-09-24"),
    ("WT2026-0000122", "同安建材", "0000202", "抗压强度", "否", "2026-09-28"),
    ("WT2026-0000123", "恒基食品", "0000203", "菌落总数", "否", "2026-09-23"),
    # 客户名尾巴粘一个全角空格：不归一化就会在台账上变成两家「华泰环境」
    ("WT2026-0000124", "华泰环境\u3000", "0000204", "六价铬", "否", "2026-09-29"),
    # 分包件：周期要按分包周期算，不能按本所周期
    ("WT2026-0000125", "金鼎材料", "0000205", "挥发性有机物", "是", "2026-09-24"),
    # 检测项目没进周期表 → 算不出交期，挂起
    ("WT2026-0000126", "信远电子", "0000206", "铅含量", "否", "2026-09-25"),
    # 收样日期空着 → 挂起（绝不拿当天日期凑）
    ("WT2026-0000127", "同安建材", "0000207", "抗压强度", "否", ""),
    ("WT2026-0000128", "恒基食品", "0000208", "菌落总数", "否", "2026-09-18"),
    # 一条环节记录都没有 → 挂起
    ("WT2026-0000129", "金鼎材料", "0000209", "硬度", "否", "2026-09-21"),
    ("WT2026-0000130", "同安建材", "0000210", "挥发酚", "否", "2026-09-28"),
    # 环节时间倒挂（上机检测早于样品前处理）→ 挂起，不许按时间排序糊过去
    ("WT2026-0000131", "信远电子", "0000211", "抗压强度", "否", "2026-09-25"),
]

# 检测项目周期：工作日口径；分包项目走「分包周期」那一列
STANDARDS = [
    ("化学需氧量", "5", ""),
    ("pH", "3", ""),
    ("抗压强度", "3", ""),
    ("菌落总数", "5", ""),
    ("六价铬", "5", ""),
    ("挥发性有机物", "5", "7"),
    ("挥发酚", "5", ""),
    ("硬度", "4", ""),
]

# 环节节点记录：委托单号 / 环节 / 完成时间
STEP_ROWS = [
    ("WT2026-0000121", "收样登记", "2026-09-24 09:10"),
    ("WT2026-0000121", "样品前处理", "2026-09-25 14:00"),
    ("WT2026-0000121", "上机检测", "2026-09-26 11:20"),
    ("WT2026-0000121", "数据审核", "2026-09-28 16:40"),

    ("WT2026-0000122", "收样登记", "2026-09-28 10:05"),
    ("WT2026-0000122", "样品前处理", "2026-09-29 15:30"),

    ("WT2026-0000123", "收样登记", "2026-09-23 09:40"),
    ("WT2026-0000123", "样品前处理", "2026-09-24 13:10"),
    ("WT2026-0000123", "上机检测", "2026-09-25 17:00"),

    ("WT2026-0000124", "收样登记", "2026-09-29 11:00"),

    ("WT2026-0000125", "收样登记", "2026-09-24 09:30"),
    ("WT2026-0000125", "样品前处理", "2026-09-25 10:20"),

    ("WT2026-0000126", "收样登记", "2026-09-25 15:00"),

    ("WT2026-0000128", "收样登记", "2026-09-18 09:00"),
    ("WT2026-0000128", "样品前处理", "2026-09-21 10:30"),
    ("WT2026-0000128", "上机检测", "2026-09-22 16:10"),
    ("WT2026-0000128", "数据审核", "2026-09-24 11:40"),
    ("WT2026-0000128", "报告编制", "2026-09-25 15:20"),
    ("WT2026-0000128", "报告审核", "2026-09-28 10:00"),
    ("WT2026-0000128", "报告发出", "2026-09-29 14:30"),

    ("WT2026-0000130", "收样登记", "2026-09-28 09:15"),
    ("WT2026-0000130", "样品前处理", "2026-09-29 14:30"),

    # 倒挂：前处理 09-29 完成，上机检测却记成 09-28 —— 真实转抄时很常见
    ("WT2026-0000131", "收样登记", "2026-09-25 10:00"),
    ("WT2026-0000131", "样品前处理", "2026-09-29 15:00"),
    ("WT2026-0000131", "上机检测", "2026-09-28 17:30"),
]

# 工作日历：9-10 月，含国庆假期与两个调休上班的周六
HOLIDAYS = {f"2026-10-0{d}" for d in range(1, 8)}
EXTRA_WORKDAYS = {"2026-09-26", "2026-10-10"}


def build_calendar():
    from datetime import date, timedelta
    rows = []
    cur = date(2026, 9, 1)
    end = date(2026, 10, 31)
    while cur <= end:
        key = cur.isoformat()
        if key in HOLIDAYS:
            flag, note = "否", "国庆假期"
        elif key in EXTRA_WORKDAYS:
            flag, note = "是", "调休上班"
        elif cur.weekday() < 5:
            flag, note = "是", ""
        else:
            flag, note = "否", "周末"
        rows.append((key, flag, note))
        cur += timedelta(days=1)
    return rows


def write_csv(path, header, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    print(f"数据已生成：{path.relative_to(BASE)}（{len(rows)} 行）")


def build_data():
    write_csv(RAW_DIR / "检测委托台账.csv",
              ["委托单号", "客户名称", "样品编号", "检测项目", "是否分包", "收样日期"],
              ORDERS)
    write_csv(RAW_DIR / "检测项目周期表.csv",
              ["检测项目", "检测周期(工作日)", "分包周期(工作日)"],
              STANDARDS)
    write_csv(RAW_DIR / "环节节点记录.csv",
              ["委托单号", "环节", "完成时间"],
              STEP_ROWS + [("", "说明：环节按 收样登记→样品前处理→上机检测→数据审核→报告编制→报告审核→报告发出 的顺序记完成时间", "")])
    cal = build_calendar()
    write_csv(RAW_DIR / "工作日历.csv", ["日期", "是否工作日", "说明"],
              cal + [("", "", "说明：本表由 build_fixtures.py 生成，口径为国务院节假日安排")])


if __name__ == "__main__":
    build_template()
    build_data()
    print("完成：模板与四张 CSV 已就位，可以跑 run.py 了")
