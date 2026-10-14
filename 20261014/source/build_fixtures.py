# -*- coding: utf-8 -*-
"""重造 20261014 示例的输入数据：上半截建检项清单与记录模板，下半截造两份 CSV。

跑法：
    python source/build_fixtures.py

上半截：
    source/templates/检项清单.csv      —— 品种 → 该做哪些项、限度、数据来源（仪器 / 人工）
    source/templates/检验记录模板.xlsx  —— 受控记录表模板（表头信息区 + 明细区 + 签名区，带打印区域）
下半截：
    01_raw_data/检验记录台账.csv       —— 本批要出哪些记录（带断号与重号）
    01_raw_data/仪器导出_结果明细.csv  —— 仪器导出的结果（带空值、限号值、文字结论、清单外项目）

数据故意造的边界：记录编号断号（缺 0004）与重号（0003 两份）、批号前导零在仪器侧被吃掉、
结果空值、限号值 `<0.05`、文字结论「未检出」、同值重复导出、
仪器出了检项清单里没有的项目、有一份记录的品种不在检项清单里、
有一份记录仪器一条结果都没匹配上、有一份记录复核人漏签、有一份检验人漏填。
"""
import csv
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = Path(__file__).resolve().parent.parent
RAW_DIR = BASE / "01_raw_data"
TPL_DIR = BASE / "source" / "templates"

ITEM_FILE = TPL_DIR / "检项清单.csv"
TPL_FILE = TPL_DIR / "检验记录模板.xlsx"
RECORD_FILE = RAW_DIR / "检验记录台账.csv"
RESULT_FILE = RAW_DIR / "仪器导出_结果明细.csv"


def write_csv(path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    print(f"已写入 {path.relative_to(BASE)}（{len(rows)} 行）")


def build_items():
    """上半截之一：检项清单。

    「该检哪些项」是品种属性，写在这张表里；代码里只认表，不认品种。
    数据来源一列是关键：仪器出的项脚本负责填，人工检查项留空等人填——
    两种项混在一个「缺项」里报，会把该手填的项误报成漏检。
    """
    write_csv(ITEM_FILE, [
        ["品种", "检验项目", "数据来源", "下限", "上限", "单位", "说明"],
        ["甲硝唑片", "性状", "人工", "", "", "", "本品为白色片，无变色"],
        ["甲硝唑片", "含量测定", "仪器", "93.0", "107.0", "%", "高效液相色谱法"],
        ["甲硝唑片", "溶出度", "仪器", "80", "", "%", "取 6 片平均"],
        ["甲硝唑片", "水分", "仪器", "", "0.5", "%", "卡氏法"],
        ["布洛芬缓释胶囊", "性状", "人工", "", "", "", "白色粉末，无结块"],
        ["布洛芬缓释胶囊", "含量测定", "仪器", "90.0", "110.0", "%", "高效液相色谱法"],
        ["布洛芬缓释胶囊", "释放度", "仪器", "75", "", "%", "高效液相色谱法"],
        ["布洛芬缓释胶囊", "有关物质", "仪器", "", "1.0", "%", "高效液相色谱法"],
    ])


def build_template():
    """上半截之二：受控记录模板。

    表头信息区用「标签 + 右边一格留空」，脚本按标签找值格——模板挪位置也不用改代码。
    明细区预置 8 行；预置行不够时脚本直接报错，不自动插行（insert_rows 会把签名区格式顶乱）。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "模板"

    thin = Side(style="thin", color="8EA9C1")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_head = PatternFill("solid", fgColor="DDEBF7")
    center = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:F1")
    ws["A1"] = "检验记录表"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = center

    head = [(2, "A", "记录编号"), (2, "C", "品种"), (2, "E", "检验日期"),
            (3, "A", "批号"), (3, "C", "检验人"), (3, "E", "复核人")]
    for row, col, label in head:
        ws[f"{col}{row}"] = label
        ws[f"{col}{row}"].font = Font(bold=True)
        ws[f"{col}{row}"].fill = fill_head
        ws[f"{col}{row}"].border = box
        value_col = chr(ord(col) + 1)
        ws[f"{value_col}{row}"].border = box

    detail_head = ["序号", "检验项目", "标准规定", "检验结果", "单位", "单项判定"]
    for i, name in enumerate(detail_head, start=1):
        c = ws.cell(row=5, column=i, value=name)
        c.font = Font(bold=True)
        c.fill = fill_head
        c.alignment = center
        c.border = box
    for r in range(6, 14):                      # 预置 8 个明细行
        for i in range(1, 7):
            ws.cell(row=r, column=i).border = box

    for col, label in (("A", "检验人签名"), ("C", "复核人签名")):
        ws[f"{col}15"] = label
        ws[f"{col}15"].font = Font(bold=True)
        ws[f"{chr(ord(col) + 1)}15"].border = box
    ws["A16"] = "备注"
    ws["A16"].font = Font(bold=True)
    ws.merge_cells("B16:F16")
    ws["B16"] = "原始数据见随附的仪器导出文件"

    for col, width in (("A", 8), ("B", 20), ("C", 22), ("D", 14), ("E", 8), ("F", 12)):
        ws.column_dimensions[col].width = width
    ws.print_area = "A1:F16"                    # 一页一份记录，靠它
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    TPL_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(TPL_FILE)
    print(f"已写入 {TPL_FILE.relative_to(BASE)}（打印区域 {ws.print_area}）")


def build_records():
    """下半截之一：记录台账。9 个连号里缺 0004、0003 出现两次，另加一份无仪器数据的和一份清单外品种的。"""
    write_csv(RECORD_FILE, [
        ["记录编号", "品种", "批号", "检验人", "复核人", "检验日期"],
        ["QC-202610-0001", "甲硝唑片", "0231A", "王琳", "李辉", "2026-10-08"],
        ["QC-202610-0002", "甲硝唑片", "0231B", "王琳", "", "2026-10-09"],
        ["QC-202610-0003", "布洛芬缓释胶囊", "0B1207", "陈敏", "李辉", "2026-10-09"],
        ["QC-202610-0003", "布洛芬缓释胶囊", "0B1208", "陈敏", "李辉", "2026-10-09"],
        ["QC-202610-0005", "布洛芬缓释胶囊", "0B1209", "陈敏", "", "2026-10-10"],
        ["QC-202610-0006", "甲硝唑片", "0231C", "王琳", "李辉", "2026-10-10"],
        ["QC-202610-0007", "甲硝唑片", "0231D", "", "李辉", "2026-10-11"],
        ["QC-202610-0008", "布洛芬缓释胶囊", "0B1210", "陈敏", "李辉", "2026-10-12"],
        ["QC-202610-0009", "甲硝唑片", "0231E", "王琳", "李辉", "2026-10-12"],
        ["QC-202610-0010", "甲硝唑片", "0231F", "王琳", "李辉", "2026-10-13"],
        ["QC-202610-0011", "甲硝唑胶囊", "0231G", "王琳", "李辉", "2026-10-13"],
    ])


def build_results():
    """下半截之二：仪器导出的结果明细。表头前压两行导出信息（测表头定位）。"""
    rows = [
        ["仪器结果文件批量导出", "2026-10-13 09:12"],
        ["数据来源：仪器工作站导出"],
        ["批号", "检验项目", "结果", "单位", "仪器编号", "导出时间"],
        ["0231A", "含量测定", "98.6", "%", "HPLC-03", "2026-10-08 16:20"],
        ["0231A", "溶出度", "92", "%", "UV-02", "2026-10-08 16:35"],
        ["0231A", "水分", "0.32", "%", "KF-01", "2026-10-08 17:02"],
        ["0231A", "干燥失重", "0.8", "%", "KF-01", "2026-10-08 17:05"],      # 清单外项目
        ["0231B", "含量测定", "107.4", "%", "HPLC-03", "2026-10-09 15:40"],  # 超上限
        ["0231B", "溶出度", "78", "%", "UV-02", "2026-10-09 15:52"],         # 低于下限
        ["0231B", "水分", "<0.05", "%", "KF-01", "2026-10-09 16:10"],        # 限号值
        ["0231C", "含量测定", "", "%", "HPLC-03", "2026-10-10 11:02"],       # 空值
        ["0231C", "溶出度", "88", "%", "UV-02", "2026-10-10 11:15"],
        ["0231C", "溶出度", "88", "%", "UV-02", "2026-10-10 11:15"],         # 同值重复导出
        ["0231C", "水分", "0.41", "%", "KF-01", "2026-10-10 11:30"],
        ["0231D", "含量测定", "95.2", "%", "HPLC-03", "2026-10-11 14:20"],
        ["0231D", "溶出度", "86", "%", "UV-02", "2026-10-11 14:33"],         # 水分整项没导出
        ["0231E", "含量测定", "101.5", "%", "HPLC-03", "2026-10-12 10:05"],
        ["0231E", "溶出度", "90", "%", "UV-02", "2026-10-12 10:18"],
        ["0231E", "水分", "0.28", "%", "KF-01", "2026-10-12 10:40"],
        ["B1207", "含量测定", "95.8", "%", "HPLC-04", "2026-10-09 14:05"],   # 前导零被吃
        ["B1207", "释放度", "96", "%", "HPLC-04", "2026-10-09 14:22"],
        ["B1207", "有关物质", "未检出", "%", "HPLC-04", "2026-10-09 14:40"], # 文字结论
        ["0B1208", "含量测定", "111.5", "%", "HPLC-04", "2026-10-09 17:02"], # 超上限
        ["0B1208", "释放度", "68", "%", "HPLC-04", "2026-10-09 17:15"],      # 低于下限
        ["0B1208", "有关物质", "0.6", "%", "HPLC-04", "2026-10-09 17:33"],
        ["0B1209", "含量测定", "94.1", "%", "HPLC-04", "2026-10-10 16:12"],
        ["0B1209", "释放度", "85", "%", "HPLC-04", "2026-10-10 16:30"],      # 有关物质没导出
        ["0B1210", "含量测定", "96.0", "%", "HPLC-04", "2026-10-12 15:05"],
        ["0B1210", "释放度", "90", "%", "HPLC-04", "2026-10-12 15:22"],
        ["0B1210", "释放度", "92", "%", "HPLC-04", "2026-10-12 15:41"],      # 同批同项两条不一致
        ["0B1210", "有关物质", "0.3", "%", "HPLC-04", "2026-10-12 15:40"],
    ]
    write_csv(RESULT_FILE, rows)


def main():
    build_items()
    build_template()
    build_records()
    build_results()
    print("数据已重造，接着跑 python run.py --no-pause")
    return 0


if __name__ == "__main__":
    sys.exit(main())
