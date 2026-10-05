#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""重造 20261005 示例的原始数据：三台仪器导出的结果文件 + 两份「问题文件」。

上半部分：定义每台仪器的导出文件长什么样（编码、分隔符、表头叫什么、列名是中文还是英文）。
下半部分：按这些定义写出 01_raw_data/ 下的 6 个文件，其中包括两份故意做坏的文件。

用法：python source/build_fixtures.py
"""

import csv
from datetime import date
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent      # source/build_fixtures.py → 包根
RAW_DIR = BASE / "01_raw_data"

DAY = "2026-10-05"
D = date(2026, 10, 5)

LAB_HEAD = ["样品编号", "样品名称", "批号", "检验项目", "结果", "单位", "分析日期"]

# HPLC 工作站：前两行是仪器信息，表头在第 3 行，制表符分隔，文件是 GBK
HPLC_ROWS = [
    ["0000713", "甲硝唑片", "250801", "含量", "99.6", "%", DAY],
    ["0000713", "甲硝唑片", "250801", "有关物质", "0.12", "%", DAY],
    ["0000714", "甲硝唑片", "250802", "含量", "98.9", "%", DAY],
    ["0000714", "甲硝唑片", "250802", "有关物质", "<0.01", "%", DAY],
]

# GC 工作站：第 1 行是标题、表头在第 2 行，逗号分隔，带 BOM
GC_ROWS = [
    ["00903", "乙醇", "250901", "残留溶剂", "0.03", "%", DAY],
    ["00903", "乙醇", "250901", "甲醇", "未检出", "%", DAY],
    ["00904", "乙醇", "250902", "残留溶剂", "0.31", "%", DAY],
]

# 紫外分光光度计：表头就在第 1 行，列名是英文，日期是真正的日期值
UV_HEAD = ["Sample", "Item", "Result", "Unit", "Date"]
UV_ROWS = [
    ["UV26-001", "吸收度", 0.452, "-", D],
    ["UV26-001", "含量", "12.3 mg/L", None, D],     # 单位列空着，单位跟在结果里
    ["UV26-002", "含量", 98.5, "%", D],
    ["UV26-002", "吸收度", None, "-", D],          # 结果列为空 → 后面挂起
]

# 同一批数据又导了一遍（文件名多带个「(1)」）：一条一模一样，一条结果不一样
UV_COPY_ROWS = [
    ["UV26-001", "吸收度", 0.452, "-", D],
    ["UV26-002", "含量", 97.1, "%", D],
]

BALANCE_TXT = """电子天平导出记录
日期：2026-10-05
操作员：李工
称量结果见纸质原始记录本，本文件仅作交接用。
"""


def write_text(name, text, encoding):
    (RAW_DIR / name).write_text(text, encoding=encoding, newline="\n")


def write_hplc(path):
    lines = ["仪器：HPLC 工作站（序列号 AN-2024-0371）",
             f"导出时间：{DAY} 08:41",
             "\t".join(LAB_HEAD)]
    lines += ["\t".join(row) for row in HPLC_ROWS]
    lines.append("\t".join(["", "合计 4 条", "", "", "", "", ""]))
    path.write_text("\n".join(lines) + "\n", encoding="gbk", newline="\n")


def write_gc(path):
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["气相色谱检测结果（GC 工作站）"])
        writer.writerow(LAB_HEAD)
        writer.writerows(GC_ROWS)
        writer.writerow(["", "说明：n.d. 表示低于检出限", "", "", "", "", ""])


def write_uv(path, rows):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "结果"
    ws.append(UV_HEAD)
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


def main():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    for old in RAW_DIR.iterdir():
        old.unlink()

    write_hplc(RAW_DIR / "HPLC工作站_结果_20261005.txt")
    write_gc(RAW_DIR / "GC工作站_结果_20261005.csv")
    write_uv(RAW_DIR / "紫外分光光度计_结果_20261005.xlsx", UV_ROWS)
    write_uv(RAW_DIR / "紫外分光光度计_结果_20261005 (1).xlsx", UV_COPY_ROWS)
    write_text("天平导出_原始记录_20261005.txt", BALANCE_TXT, "gbk")
    write_text("酶标仪_结果_20261005.csv", "酶标仪检测结果\n" + ",".join(LAB_HEAD) + "\n", "utf-8-sig")

    print(f"已生成 {len(list(RAW_DIR.iterdir()))} 个文件到 {RAW_DIR}")


if __name__ == "__main__":
    main()
