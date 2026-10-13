# -*- coding: utf-8 -*-
"""重造 20261013 示例的输入数据：上半截建编号规则表，下半截造两份 CSV。

跑法：
    python source/build_fixtures.py

上半截：编号规则表（source/templates/编号规则表.csv）
下半截：
    01_raw_data/编号台账_已用编号.csv   —— 历史已用编号（带重号与丢前导零的脏数据）
    01_raw_data/本批样品清单.csv        —— 本批待编号样品（带缺字段的待确认行）

数据故意造的边界：丢前导零的 9、同段重号、流水位数溢出（999 → 1000）、
补录上月、检测类型不在规则表、样品名称为空、送样日期认不出、某组只有 1 条。
"""
import csv
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
RAW_DIR = BASE / "01_raw_data"
TPL_DIR = BASE / "source" / "templates"

RULE_FILE = TPL_DIR / "编号规则表.csv"
LEDGER_FILE = RAW_DIR / "编号台账_已用编号.csv"
SAMPLES_FILE = RAW_DIR / "本批样品清单.csv"


def write_csv(path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)
    print(f"已写入 {path.relative_to(BASE)}（{len(rows)} 行）")


def build_rules():
    """上半截：编号规则表。位数 3 是故意留的——食品接触材料那组历史已到 998。"""
    write_csv(RULE_FILE, [
        ["检测类型", "编号前缀", "流水位数", "备注"],
        ["水质检测", "WS", "4", "按送样月份分段，同月连编"],
        ["土壤检测", "TR", "4", "按送样月份分段，同月连编"],
        ["食品接触材料", "FCM", "3", "历史已接近位数上限，注意进位"],
    ])


def build_ledger():
    """历史台账：表头前压两行导出信息（测 find_header），并混入两类脏数据。"""
    rows = [
        ["编号台账导出", "2026-10-12 17:20"],
        ["数据来源：收样台账导出"],
        ["编号", "样品名称", "检测类型", "编号生成日"],
        ["WS-202609-0005", "地表水 A", "水质检测", "2026-09-18"],
        ["WS-202609-0006", "地表水 B", "水质检测", "2026-09-25"],
        ["WS-202609-0007", "地下水 C", "水质检测", "2026-09-30"],
        ["WS-202610-0001", "废水 D", "水质检测", "2026-10-08"],
        ["WS-202610-0002", "废水 E", "水质检测", "2026-10-08"],
        ["WS-202610-0003", "废水 F", "水质检测", "2026-10-09"],
        ["WS-202610-0003", "废水 F 复检", "水质检测", "2026-10-09"],
        ["TR-202610-0009", "土壤 S1", "土壤检测", "2026-10-06"],
        ["TR-202610-9", "土壤 S2", "土壤检测", "2026-10-07"],
        ["TR-202610-0010", "土壤 S3", "土壤检测", "2026-10-09"],
        ["FCM-202610-997", "迁移量 P1", "食品接触材料", "2026-10-05"],
        ["FCM-202610-998", "迁移量 P2", "食品接触材料", "2026-10-07"],
    ]
    write_csv(LEDGER_FILE, rows)


def build_samples():
    """本批样品：11 条当月水质 + 1 条补录上月 + 1 条土壤 + 3 条食品接触材料 + 4 条待确认。"""
    rows = [["样品名称", "检测类型", "样品来源", "送样日期", "备注"]]
    for i in range(1, 12):
        rows.append([f"废水样 {i} 号点位", "水质检测", f"某工业园 {i} 号排口",
                     "2026-10-13", "常规九项"])
    rows.append(["土壤 S4", "土壤检测", "某地块修复项目", "2026-10-13", "单点，重金属"])
    for i in range(1, 4):
        rows.append([f"食品接触材料 P{i}", "食品接触材料", "某包装厂",
                     "2026-10-13", "迁移量试验"])
    rows.append(["菌落总数样", "微生物检测", "某食品厂", "2026-10-13", "类型未进规则表"])
    rows.append(["", "水质检测", "某水务公司", "2026-10-13", "名称漏填"])
    rows.append(["空白水样", "", "某水务公司", "2026-10-13", "检测类型漏填"])
    rows.append(["地下水 D", "水质检测", "某牧场", "待定", "送样日期还没定"])
    rows.append(["废水样 12 号点位（补录）", "水质检测", "某工业园 12 号排口",
                 "2026-09-30", "上月漏录，本批补"])
    write_csv(SAMPLES_FILE, rows)


def main():
    build_rules()
    build_ledger()
    build_samples()
    print("数据已重造，接着跑 python run.py --no-pause")
    return 0


if __name__ == "__main__":
    sys.exit(main())
