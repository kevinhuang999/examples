#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""重造本示例的模板与数据（上半建模板，下半造三张表）。

上半：建 `source/templates/报价单模板.xlsx`——单头是「标签 + 取值」两列一组，
明细区固定 12 行、**每行都画了边框**（不画边框，openpyxl 读回时这些行根本不存在），
合计行故意不画边框（画了就会被当成可填的明细行）。

下半：造 `01_raw_data/` 下三张表——
  ① `价目表.xlsx`：长表。同一个项目好几行很正常：不同方法不同价、今年调过价、
     去年的旧价还留着、明年 10 月才生效的调价也已经录进来了。
  ② `询价明细.csv`：客户询价原件，一行一个（样品 × 项目）。故意留了客户名夹全角空格、
     项目名带括号注解、同一个项目没写方法、数量写成「2 件」、同编号同项目重复两行这些边界。
  ③ `历史成交价.csv`：这家客户上次成交的价，故意留了一条记着「待定」的。

用法：python source/build_fixtures.py
"""

from csv import DictWriter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

ROOT = Path(__file__).resolve().parent.parent      # 包根 = build_fixtures.py 的上上层
RAW = ROOT / "01_raw_data"
TPL = ROOT / "source" / "templates" / "报价单模板.xlsx"

# 单头：标签写 A / C 两列，取值写在下一列
HEAD_FIELDS = [("报价单号", 2), ("报价日期", 2), ("客户名称", 3),
               ("项数", 3), ("有效期至", 4), ("折扣档", 4)]
DETAIL_HEAD = ["序号", "检测项目", "检测方法", "客户样品编号", "数量", "单价", "金额"]
DETAIL_ROWS = 12
WIDTHS = {"A": 16, "B": 24, "C": 16, "D": 18, "E": 8, "F": 10, "G": 12}

THIN = Side(style="thin", color="808080")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_template():
    """上半：建空白报价单模板。模板由质量部备案，脚本只能填、不能改结构。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "报价单"

    ws["A1"] = "检 测 服 务 报 价 单"
    ws.merge_cells("A1:G1")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    for label, row in HEAD_FIELDS:
        col = 1 if label in ("报价单号", "客户名称", "有效期至") else 3
        cell = ws.cell(row=row, column=col, value=label)
        cell.font = Font(bold=True)
        ws.cell(row=row, column=col + 1).border = BOX

    ws["A5"] = "计价说明"
    ws["A5"].font = Font(bold=True)
    ws.merge_cells("B5:D5")
    ws["B5"].border = BOX

    head = 6
    for i, name in enumerate(DETAIL_HEAD, 1):
        cell = ws.cell(row=head, column=i, value=name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = BOX

    # 明细区空行必须画边框：不画的话这些行在 openpyxl 眼里"不存在"，
    # 回读时 ws.max_row 只到表头那一行，靠 max_row 反推明细区的写法会一行都找不到。
    for r in range(head + 1, head + 1 + DETAIL_ROWS):
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = BOX

    # 合计行**不画边框**：画了就会连着明细区一路算下去，合计行被当成可填的明细行。
    ws.cell(row=head + 1 + DETAIL_ROWS, column=5, value="合计").font = Font(bold=True)

    for col, width in WIDTHS.items():
        ws.column_dimensions[col].width = width

    TPL.parent.mkdir(parents=True, exist_ok=True)
    wb.save(TPL)
    print(f"模板已生成：{TPL}")


def build_price_list():
    """价目表（长表）。注意铅含量有 3 行、pH值有 2 行——这就是报价匹配难的地方。"""
    head = ["项目编号", "项目名称", "别名", "检测方法", "计价单位", "单价", "生效日期", "失效日期"]
    rows = [
        # 铅含量：两种方法两个价；第三条是 2026-10-01 才生效的调价，今天报价不能用
        ["HJ-001", "铅含量", "Pb、重金属铅、铅", "ICP-OES", "元/项", 220, "2026-01-01", "2026-12-31"],
        ["HJ-001", "铅含量", "Pb、重金属铅、铅", "XRF 荧光光谱", "元/项", 130, "2026-01-01", "2026-12-31"],
        ["HJ-001", "铅含量", "Pb、重金属铅、铅", "ICP-OES", "元/项", 260, "2026-10-01", "2027-10-01"],
        ["HJ-002", "镉含量", "Cd、重金属镉", "ICP-OES", "元/项", 220, "2026-01-01", "2026-12-31"],
        # pH值：第二条是 2025 年的旧价，早失效了，也不能用
        ["HJ-003", "pH值", "PH、酸碱度", "pH计", "元/项", 60, "2026-01-01", "2026-12-31"],
        ["HJ-003", "pH值", "PH、酸碱度", "pH计", "元/项", 50, "2025-01-01", "2025-12-31"],
        ["HJ-004", "甲醛含量", "甲醛", "分光光度法", "元/项", 280, "2026-01-01", "2026-12-31"],
        ["HJ-004", "甲醛含量", "甲醛", "液相色谱法", "元/项", 450, "2026-01-01", "2026-12-31"],
        ["HJ-005", "总磷", "磷", "分光光度法", "元/项", 150, "2026-01-01", "2026-12-31"],
        ["HJ-006", "抗拉强度", "拉伸强度", "万能试验机", "元/项", 350, "2026-01-01", "2026-12-31"],
        ["HJ-007", "断裂伸长率", "延伸率", "万能试验机", "元/项", 350, "2026-01-01", "2026-12-31"],
        ["HJ-008", "剥离强度", "剥离力", "剥离试验机", "元/项", 400, "2026-01-01", "2026-12-31"],
        ["HJ-009", "玻璃化转变温度", "Tg、玻璃化温度", "DSC", "元/项", 600, "2026-01-01", "2026-12-31"],
        ["HJ-010", "耐摩擦色牢度", "色牢度、摩擦色牢度", "摩擦色牢度仪", "元/项", 180, "2026-01-01", "2026-12-31"],
        ["HJ-011", "阻燃性能", "燃烧性能", "垂直燃烧法", "元/项", 900, "2026-01-01", "2026-12-31"],
        ["HJ-012", "化学成分", "成分分析", "光谱法", "元/项", 300, "2026-01-01", "2026-12-31"],
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "价目表"
    ws.append(head)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    for i, width in enumerate([10, 18, 22, 16, 10, 8, 12, 12], start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
    RAW.mkdir(parents=True, exist_ok=True)
    wb.save(RAW / "价目表.xlsx")
    print(f"价目表已生成：{RAW / '价目表.xlsx'}  {len(rows)} 行")


def build_inquiry():
    """客户询价明细：一行一个（样品 × 项目）。六种边界混在四家客户里。"""
    head = ["询价日期", "客户名称", "客户样品编号", "样品名称", "检测项目", "检测方法", "数量"]
    hy, xy, qh, nb, jt = ("浙江宏远纺织有限公司", "苏州新亚电子材料有限公司",
                          "上海启海化工有限公司", "宁波航海紧固件有限公司", "杭州锦泰包装有限公司")
    rows = []

    def add(cust, code, name, item, method="", qty="1", day="2026-09-24"):
        rows.append([day, cust, code, name, item, method, qty])

    # ① 客户A：6 行，其中 1 行方法待确认 → 能报 5 项，正好等于 5 项那一档的门槛
    add(hy, "0000713", "全棉染色布", "甲醛含量", "分光光度法", "3")
    add(hy, "0000713", "全棉染色布", "pH值", "pH计", "3")
    # 客户名尾巴上粘了一个全角空格——不去干净就会跟同名客户拆成两家，出两份报价单
    add(hy + "\u3000", "0000714", "涤纶针织布", "耐摩擦色牢度", "摩擦色牢度仪", "2")
    add(hy, "0000715", "涤棉混纺布", "玻璃化转变温度", "DSC", "1")
    add(hy, "0000716", "家纺面料", "断裂伸长率", "万能试验机", "1")
    # 客户带括号注解、且没写方法 → 归一化能锚到"铅"，但铅有两种方法 → 方法待确认
    add(hy, "0000717", "装饰面料", "铅（Pb）含量测定", "", "1")

    # ② 客户B：4 行全可报 → 4 项，卡在 5 项门槛的另一侧，不打折
    add(xy, "XYA-001", "覆铜板", "剥离强度", "剥离试验机", "2")
    add(xy, "XYA-002", "环氧树脂", "玻璃化转变温度", "DSC", "1")
    add(xy, "XYA-003", "铜箔", "抗拉强度", "万能试验机", "2")
    add(xy, "XYA-003", "铜箔", "断裂伸长率", "万能试验机", "2")

    # ③ 客户C：10 行全可报 → 命中 9 折。铅含量两种方法各走各的价
    add(qh, "W-0924-A", "废水水样 A", "pH值", "pH计", "3")
    add(qh, "W-0924-A", "废水水样 A", "总磷", "分光光度法", "3")
    # 客户写"重金属铅"，价目表里叫"铅含量"——靠别名表才落得上
    add(qh, "W-0924-B", "废水水样 B", "重金属铅", "ICP-OES", "3")
    add(qh, "W-0924-C", "废水水样 C", "镉含量", "ICP-OES", "3")
    add(qh, "W-0924-D", "废水水样 D", "铅含量", "XRF 荧光光谱", "3")
    add(qh, "W-0924-E", "废水水样 E", "pH值", "pH计", "1")
    add(qh, "W-0924-F", "废水水样 F", "总磷", "分光光度法", "1")
    add(qh, "W-0924-G", "废水水样 G", "镉含量", "ICP-OES", "1")
    add(qh, "W-0924-H", "废水水样 H", "铅含量", "ICP-OES", "1")
    add(qh, "W-0924-I", "废水水样 I", "pH值", "pH计", "2")

    # ④ 客户D：三项都报不出来 → 这家一张报价单都不出
    add(nb, "HB-0924-1", "螺栓 M12", "多环芳烃 PAHs", "GC-MS", "1")           # 价目表里没有
    add(nb, "HB-0924-2", "螺母 M10", "抗拉强度", "万能试验机", "2 件")         # 数量不是整数
    add(nb, "HB-0924-3", "垫圈", "铅含量", "液相色谱法", "1")                  # 铅没这个方法

    # ⑤ 客户名空着：前台登记时漏填，整组挂起
    add("", "HT-0924-1", "铝合金型材", "化学成分", "光谱法", "1")
    add("", "HT-0924-2", "铝合金型材", "抗拉强度", "万能试验机", "1")

    # ⑥ 同一编号同一项目出现两行：留先出现的那行，后一行挂起
    add(jt, "JT-0924-1", "瓦楞纸箱", "抗拉强度", "万能试验机", "1")
    add(jt, "JT-0924-1", "瓦楞纸箱", "抗拉强度", "万能试验机", "1")

    # 表尾说明行：第一列（询价日期）留空，读的时候整行跳过。
    # 也正因为这样，说明文字千万别写到第一列——写进去就会被当成一条询价。
    rows.append(["", "说明：本表是客户询价原件，一行一个（样品 × 检测项目）；"
                 "客户没写检测方法、或者数量写的不是整数，一律不改口径，挂起后回去确认。", "", "", "", "", ""])

    with open(RAW / "询价明细.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = DictWriter(f, fieldnames=head)
        w.writeheader()
        for r in rows:
            w.writerow(dict(zip(head, r)))
    print(f"询价明细已生成：{RAW / '询价明细.csv'}  数据 {len(rows) - 1} 行 + 表尾说明 1 行")


def build_history():
    """历史成交价。方法要一起写——同一项目不同方法本来就不是一个价。"""
    head = ["客户名称", "检测项目", "检测方法", "上次单价", "成交日期"]
    rows = [
        ["浙江宏远纺织有限公司", "甲醛含量", "分光光度法", 350, "2026-03-11"],     # 这次 280，降 20%
        ["浙江宏远纺织有限公司", "pH值", "pH计", 60, "2026-03-11"],              # 没变
        ["浙江宏远纺织有限公司", "断裂伸长率", "万能试验机", "待定", "2026-03-11"],  # 记的不是数字
        ["苏州新亚电子材料有限公司", "玻璃化转变温度", "DSC", 600, "2026-05-20"],   # 没变
        ["上海启海化工有限公司", "铅含量", "ICP-OES", 200, "2026-06-08"],          # 涨 10.0%，正好卡线
        ["上海启海化工有限公司", "总磷", "分光光度法", 120, "2026-06-08"],         # 涨 25%
        ["杭州锦泰包装有限公司", "抗拉强度", "万能试验机", 320, "2026-07-15"],     # 涨 9.4%，不到线
        ["", "说明：一行一次成交；同一项目不同方法要分开写，不然比价会比出假信号。", "", ""],
    ]
    with open(RAW / "历史成交价.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = DictWriter(f, fieldnames=head)
        w.writeheader()
        for r in rows:
            w.writerow(dict(zip(head, r)))
    print(f"历史成交价已生成：{RAW / '历史成交价.csv'}  数据 {len(rows) - 1} 行 + 表尾说明 1 行")


if __name__ == "__main__":
    build_template()
    build_price_list()
    build_inquiry()
    build_history()
