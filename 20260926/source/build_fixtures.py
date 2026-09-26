#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""重造 20260926 示例用到的原始数据与台账模板（跑 run.py 之前想重置数据就跑它）。

上半个文件：造 `source/templates/样品登记台账.xlsx`——表头三行 + 一行带边框的空明细行。
下半个文件：造 `01_raw_data/` 下三张 CSV——送样登记原始记录、客户代码表、样品类别代码表。

改数据时有个纪律：**规则表和数据要一起想**。比如给某一行填了「类别＝其他」，
就得接受它整行挂起；把客户名改成一个没建代码的新客户，也是整行挂起。
两边不一致的时候，跑出来一屏挂起，很难分清是脚本错还是数据错。

用法：
    python source/build_fixtures.py
"""

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "01_raw_data"
TPL_DIR = ROOT / "source" / "templates"

TPL_XLSX = TPL_DIR / "样品登记台账.xlsx"
REG_CSV = RAW_DIR / "送样登记原始记录.csv"
CUST_CSV = RAW_DIR / "客户代码表.csv"
CAT_CSV = RAW_DIR / "样品类别代码表.csv"

# 台账栏目：改这里要跟 run.py 的 HEADER 一起改，两边逐字一致才能按栏目名反查。
HEADER = ["样品编号", "受理日期", "委托单号", "客户名称", "客户样品编号",
          "样品名称", "样品类别", "数量", "存放位置", "接收人", "备注"]
WIDTH = [18, 12, 14, 26, 16, 20, 10, 8, 12, 10, 18]

TITLE = "样品登记台账"
SUBTITLE = "收样室 · 示例数据（机构名与客户名均为虚构）"

CUST_ROWS = [
    ["客户名称", "客户代码"],
    ["上海启海食品有限公司", "QH"],
    ["浙江宏远纺织有限公司", "HY"],
    ["江苏金鼎建材有限公司", "JD"],
    ["苏州明达环境科技有限公司", "MD"],
]

CAT_ROWS = [
    ["样品类别", "类别码"],
    ["饮用水", "A"],
    ["废水", "B"],
    ["纺织品", "C"],
    ["建材", "D"],
    ["土壤", "E"],
]

# 送样登记原始记录：一行 = 客户送样单上的一行，不等于一件样品。
# 数量那一列写的是客户原件上的写法，脚本按它把这一行展开成几件。
REG_ROWS = [
    ["受理日期", "委托单号", "客户名称", "客户样品编号", "样品名称", "样品类别", "数量", "存放位置", "备注"],
    ["2026-09-26", "WT26092601", "上海启海食品有限公司", "0000713", "瓶装饮用水", "饮用水", "1", "常温库A区", ""],
    ["2026-09-26", "WT26092601", "上海启海食品有限公司", "0000714", "桶装饮用水", "饮用水", "3", "常温库A区", "加大桶"],
    ["2026-09-26", "WT26092602", "浙江宏远纺织有限公司\u3000", "HY-2026-0088", "棉麻混纺面料", "纺织品", "2", "样品间B架", ""],
    ["2026-09-26", "WT26092603", "江苏金鼎建材有限公司", "JD-0912", "铝合金型材", "建材", "1", "样品间C架", ""],
    ["2026-09-26", "WT26092604", "苏州明达环境科技有限公司", "MD-A-771", "厂界废水", "废水", "2", "冷藏柜1", "需 4℃ 保存"],
    ["2026-09-26", "WT26092605", "上海启海食品有限公司", "0000715", "瓶装饮用水", "饮用水", "2", "常温库A区", "下午补送"],
    ["待定", "WT26092606", "江苏金鼎建材有限公司", "JD-0913", "水泥试块", "建材", "1", "样品间C架", "送样单日期未填"],
    ["2026-09-26", "WT26092607", "", "WX-001", "生活污水", "废水", "1", "冷藏柜1", "客户名称未填"],
    ["2026-09-26", "WT26092608", "无锡华信水务有限公司", "WX-2026-11", "自来水", "饮用水", "1", "常温库A区", "新客户未建代码"],
    ["2026-09-26", "WT26092609", "苏州明达环境科技有限公司", "MD-A-772", "厂界噪声样品", "其他", "1", "样品间D架", "类别写其他"],
    ["2026-09-26", "WT26092610", "浙江宏远纺织有限公司", "HY-2026-0090", "", "纺织品", "1", "样品间B架", "样品名称空"],
    ["2026-09-26", "WT26092611", "江苏金鼎建材有限公司", "JD-0914", "防水卷材", "建材", "2 支", "样品间C架", "数量带单位"],
    ["2026-09-26", "WT26092612", "苏州明达环境科技有限公司", "MD-A-773", "土壤浸出液", "土壤", "1", "样品间D架", ""],
    ["2026-09-26", "WT26092612", "苏州明达环境科技有限公司", "MD-A-773", "土壤浸出液", "土壤", "1", "样品间D架", "同单重复样号"],
    ["2026-09-26", "WT26092613", "苏州明达环境科技有限公司", "MD-A-775", "土壤浸出液", "土壤", "1", "样品间D架", "另一份土壤样"],
    ["2026-09-26", "WT26092614", "上海启海食品有限公司", "0000716", "饮用水", "饮用水", "120", "常温库A区", "大批量委托"],
    # 表尾：合计行与说明行的第一列都留空 —— 脚本靠「受理日期」这一列判空跳过它们。
    # 合计数字千万别写进第一列，写进去就多出一条"编号为空"的假样品。
    ["", "", "合计", "126", "", "", "", "", "", ""],
    ["", "", "本表由收样室导出，数量一列按客户送样单原样记录，与本表不符时以送样单为准。", "", "", "", "", "", "", ""],
]


def build_template():
    """造台账模板：三行抬头 + 表头 + 一行带边框的空明细行。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "样品登记台账"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADER))
    ws.cell(row=1, column=1, value=TITLE).font = Font(name="微软雅黑", size=16, bold=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HEADER))
    ws.cell(row=2, column=1, value=SUBTITLE).font = Font(name="微软雅黑", size=10)
    ws.row_dimensions[1].height = 28

    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, name in enumerate(HEADER, start=1):
        c = ws.cell(row=3, column=i, value=name)
        c.font = Font(name="微软雅黑", size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = WIDTH[i - 1]

    # 明细区第一行留空但**必须画边框**：脚本靠"表头下面还有行"认明细起点，
    # 而 openpyxl 读回时既没内容又没样式的行根本不存在（max_row 只到表头行）。
    # 这一行只用来保证"表头下面有地方可填"，真正的行数按件数在 run.py 里按需创建。
    for i in range(1, len(HEADER) + 1):
        ws.cell(row=4, column=i).border = border
    ws.row_dimensions[4].height = 18

    ws.freeze_panes = "A4"
    TPL_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(TPL_XLSX)
    print(f"[模板] {TPL_XLSX.relative_to(ROOT)}  {len(HEADER)} 列 + 表头行 + 1 行带边框空行")


def write_csv(path, rows):
    """中文 CSV 一律带 BOM（utf-8-sig）——不带的话 Excel 双击打开是乱码，收样室同事第一个就炸。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    print(f"[数据] {path.relative_to(ROOT)}  {len(rows) - 1} 行")


def main():
    build_template()
    write_csv(REG_CSV, REG_ROWS)
    write_csv(CUST_CSV, CUST_ROWS)
    write_csv(CAT_CSV, CAT_ROWS)
    print("\n造数据完成。跑 `python run.py` 看结果。")


if __name__ == "__main__":
    main()
