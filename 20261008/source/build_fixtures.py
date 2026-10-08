# -*- coding: utf-8 -*-
"""重造示例用的字段映射表与 ELN 导出数据（会覆盖 01_raw_data 与 source/templates）。

上半部分建模板：字段映射表（标准字段 / 别名 / 是否必需）。
下半部分造数据：4 份 ELN 导出（不同模板版本、不同编码、不同列顺序，含空值与表尾说明行）
+ 1 份混在同一个导出目录里的无关记录（用来演示"认不出的文件要挂起"）。

用法：python source/build_fixtures.py
"""
import csv
from pathlib import Path

from openpyxl import Workbook

BASE = Path(__file__).resolve().parent.parent
RAW = BASE / "01_raw_data"
TPL = BASE / "source" / "templates"

# ---------------------------------------------------------------- 上半：建模板

MAPPING = [
    ("标准字段", "别名（用 | 分隔，同一字段可写多个）", "是否必需"),
    ("实验编号", "实验编号|Experiment ID|实验号", "是"),
    ("批号", "批号|Batch No.|BatchNo", "是"),
    ("原料", "原料|物料|Material|原料名称", "是"),
    ("投料量_g", "投料量(g)|投料量（g）|Charge(g)|投料量", "否"),
    ("温度_C", "反应温度(℃)|反应温度（℃）|Temp(C)|温度/℃", "否"),
    ("时间_h", "反应时间(h)|Time(h)|时间/h", "否"),
    ("收率_pct", "收率(%)|Yield(%)|收率/%", "否"),
    ("实验员", "操作人|Operator|实验员", "否"),
]

NOTE = """ELN 导出记录（示例数据说明）
================================
这些文件模拟同一份 ELN 里按批次分次导出的实验记录，各次导出的模板版本不同：

- ELN_BR-2026-001.csv   早期模板，UTF-8，表头在第 1 行，中文列名
- ELN_BR-2026-002.csv   中期模板，GBK   ，表头上方压了 3 行导出信息，英文列名
- ELN_BR-2026-003.xlsx  最新模板，列顺序与前两版都不同，含空值与表尾说明行
- ELN_BR-2026-004.csv   最新模板，但导出时列顺序被手工调过，其中一行缺实验编号
- 其它_环境记录_20261008.csv  同目录下顺手导出的环境记录，字段与实验记录无关

共同字段按 source/templates/字段映射表.csv 归并。
"""

# ---------------------------------------------------------------- 下半：造数据

BR_001 = """实验编号,批号,物料,投料量(g),反应温度(℃),反应时间(h),收率(%),操作人
0098,BR-2026-001A,苯甲酸,12.5,80,4.0,82.5,李明
0099,BR-2026-001B,苯甲酸,12.5,85,4.0,86.1,李明
"""

BR_002 = """导出系统：ELN v3.2
导出时间：2026-03-18 09:12
实验范围：BR-2026-002 批次
Experiment ID,Batch No.,Material,Charge(g),Temp(C),Time(h),Yield(%),Operator
0100,BR-2026-002A,苯甲酸,12.5,80,5,88.3,Wang
0101,BR-2026-002B,苯甲酸,13.0,80,5,89.0,Wang
"""

BR_003 = [
    ["实验编号", "批号", "原料", "投料量", "温度/℃", "时间/h", "收率/%", "实验员", "备注"],
    ["0102", "BR-2026-003A", "苯甲酸", 12.5, 90, 6, 78.4, "张伟", ""],
    ["0103", "BR-2026-003B", "苯甲酸", 12.5, 90, 6, "未测", "张伟", "样品送检"],
    ["0104", "BR-2026-003C", "苯甲酸", "", 90, 6, "", "张伟", "投料记录待补"],
    ["", "", "本表由 ELN v3.4 导出 2026-03-20", "", "", "", "", "", ""],
]

BR_004 = """Batch No.,Experiment ID,Material,Charge(g),Temp(C),Time(h),Yield(%),Operator
BR-2026-004A,0105,苯甲酸,12.5,85,5,84.2,Li
BR-2026-004B,,苯甲酸,12.5,85,5,85.0,Li
"""

ENV = """记录时间,温度(℃),湿度(%),记录人
2026-03-18 08:00,22.5,45,王芳
2026-03-18 14:00,23.1,44,王芳
"""


def build_templates():
    TPL.mkdir(parents=True, exist_ok=True)
    with (TPL / "字段映射表.csv").open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(MAPPING)
    (RAW / "ELN导出说明_20261008.txt").write_text(NOTE, encoding="utf-8")


def build_data():
    RAW.mkdir(parents=True, exist_ok=True)
    (RAW / "ELN_BR-2026-001.csv").write_text(BR_001, encoding="utf-8")
    # 中期模板这份是 GBK：用 utf-8 读会直接抛 UnicodeDecodeError，不是读出乱码
    (RAW / "ELN_BR-2026-002.csv").write_text(BR_002, encoding="gbk")
    (RAW / "ELN_BR-2026-004.csv").write_text(BR_004, encoding="utf-8")
    (RAW / "其它_环境记录_20261008.csv").write_text(ENV, encoding="utf-8")

    wb = Workbook()
    ws = wb.active
    ws.title = "实验记录"
    for row in BR_003:
        ws.append(row)
    for i in range(2, 5):        # 实验编号列按文本存，前导零才不会被吃掉
        ws.cell(row=i, column=1).number_format = "@"
    wb.save(RAW / "ELN_BR-2026-003.xlsx")


if __name__ == "__main__":
    RAW.mkdir(parents=True, exist_ok=True)
    build_templates()
    build_data()
    print("已重造：source/templates/字段映射表.csv")
    print("已重造：01_raw_data/ 下 5 个数据文件 + 1 个说明文件")
