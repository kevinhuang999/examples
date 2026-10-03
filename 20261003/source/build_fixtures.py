#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""重新生成示例数据。

上半个文件建规则表（source/templates/峰归属规则.csv），
下半个文件造仪器导出的峰表文件（01_raw_data/*.txt）。

跑法：python source/build_fixtures.py
会先清空 01_raw_data/ 与 source/templates/ 再重建，重复跑结果一致。
"""
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent   # 写回示例根目录
RAW = BASE / "01_raw_data"
TPL = BASE / "source" / "templates"

# ---------------------------------------------------------------- 上半：规则表
RULES_CSV = """峰名关键词,参与归一,是否目标峰,备注
溶剂,否,否,溶剂峰/系统峰：不参与面积归一
系统,否,否,系统峰：不参与面积归一
主峰,是,是,待测成分主峰
杂质,是,否,杂质峰：参与归一
"""

# ---------------------------------------------------------------- 下半：峰表数据
# (文件名, 编码, 样品编号, 样品名称, 进样次数, 峰行, 总计面积, 总计峰高)
INJECTIONS = [
    ("HPLC-20260618-001.txt", "utf-8", "0000182", "阿莫西林胶囊", "1",
     [("1", "溶剂峰", "1.215", "12560", "890"),
      ("2", "主峰", "3.482", "1246800", "98450"),
      ("3", "杂质A", "4.901", "18630", "1520")],
     "1277990", "100860"),

    # 中文峰名 + GBK 编码：仪器软件在中文 Windows 上默认吐 ANSI
    ("HPLC-20260618-002.txt", "gbk", "0000183", "复方磺胺甲噁唑片", "1",
     [("1", "溶剂峰", "1.180", "9840", "720"),
      ("2", "磺胺甲噁唑", "3.201", "892400", "71230"),
      ("3", "甲氧苄啶", "4.115", "<1000", "<100")],
     "902240", "71950"),

    # 只有溶剂峰：没有可参与归一的峰
    ("HPLC-20260618-003.txt", "utf-8", "0000184", "空白对照", "1",
     [("1", "溶剂峰", "1.201", "10240", "810")],
     "10240", "810"),

    # 目标峰面积写成 n.d.（未检出）
    ("HPLC-20260618-004.txt", "utf-8", "0000185", "回收率样品", "1",
     [("1", "溶剂峰", "1.190", "8760", "690"),
      ("2", "主峰", "3.455", "n.d.", "—"),
      ("3", "杂质A", "4.880", "21050", "1780")],
     "21050", "1780"),

    # 同一样品两次进样：不能并成一条（面积会翻倍），偏差超限要提示复测
    ("HPLC-20260618-005.txt", "utf-8", "0000186", "阿莫西林胶囊", "1",
     [("1", "溶剂峰", "1.210", "11200", "860"),
      ("2", "主峰", "3.470", "1155200", "91200"),
      ("3", "杂质A", "4.892", "17200", "1400")],
     "1172400", "92600"),

    ("HPLC-20260618-005R.txt", "utf-8", "0000186", "阿莫西林胶囊", "2",
     [("1", "溶剂峰", "1.208", "10870", "840"),
      ("2", "主峰", "3.468", "1100000", "88400"),
      ("3", "杂质A", "4.886", "60000", "5100")],
     "1160000", "93500"),
]

TEMPLATE = (
    "# 仪器: HPLC-2030 (UV 254nm)\n"
    "# 采集日期: 2026-06-18 09:12:03\n"
    "# 导出时间: 2026-06-18 18:02:11\n"
    "样品编号: {no}\n"
    "样品名称: {sample}\n"
    "进样次数: {seq}\n"
    "# " + "-" * 52 + "\n"
    "序号\t峰名\t保留时间/min\t峰面积\t峰高\n"
    "{body}\n"
    "总计\t\t\t{total}\t{htotal}\n"
)


def render(no, sample, seq, peaks, total, htotal):
    body = "\n".join("\t".join(p) for p in peaks)
    return TEMPLATE.format(no=no, sample=sample, seq=seq,
                           body=body, total=total, htotal=htotal)


def main():
    for d in (RAW, TPL):
        if d.exists():
            shutil.rmtree(d)
        d.mkdir(parents=True)

    (TPL / "峰归属规则.csv").write_text(RULES_CSV, encoding="utf-8-sig")
    for name, enc, no, sample, seq, peaks, total, htotal in INJECTIONS:
        text = render(no, sample, seq, peaks, total, htotal)
        (RAW / name).write_bytes(text.encode(enc))
        print(f"  写入 {name}（{enc}）")

    print(f"完成：规则表 1 份、峰表文件 {len(INJECTIONS)} 份 → {RAW}")


if __name__ == "__main__":
    main()
