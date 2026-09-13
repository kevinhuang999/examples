# -*- coding: utf-8 -*-
"""重造 20260913 示例的原始数据（稳定性考察汇总用的仪器导出明细）。

这个文件上下两半分工：
  上半部分（常量区 ~ make_detail）—— 定义稳定性考察的批次、时间点、检项，
    以及每个检项的标准规定，拼出全部明细行；
  下半部分（write_csv 起）—— 把明细写成 01_raw_data/稳定性考察_仪器导出.csv（带 BOM），
    并顺手把"考察方案"页签也写成一个单独的 xlsx，供 run.py 读取时间点顺序。

边界是故意造的，删掉就跑不出坑了（见 README 的边界表）：
  · 时间点文字混排（0 天 / 1 个月 / 3 个月 / 6 个月），不按字典序排
  · 检验结果为文字（符合 / 不符合），混在同一列里
  · 个别格子是空的（结果还没出）
  · 批号带前导零（如 00903）
  · 一个批次缺一个时间点（漏做/待做）
  · 一个批次的检项极多（长毒 12 个检项），一个批次极少（1 个检项）
  · CSV 末尾留空行 + 一行「合计」+ 一行说明（仪器导出常见尾巴）

跨平台：路径全部相对定位，不写死任何本机路径。
"""
import csv
from pathlib import Path

# 在 source/ 里，往上一层才是包根
BASE = Path(__file__).resolve().parent.parent
RAW_DIR = BASE / "01_raw_data"

# 稳定性考察方案：批次 → 考察时间点（文字，故意不按字典序）
PLAN = {
    "240905": ["0 天", "1 个月", "3 个月", "6 个月"],
    "240906": ["0 天", "1 个月", "3 个月", "6 个月"],
    "00903": ["0 天", "1 个月", "3 个月"],           # 批号带前导零；6 个月还没到
    "241012": ["0 天"],                              # 刚开始考察，只有 0 天
    "241115": ["0 天", "1 个月", "3 个月", "6 个月"],
}

# 检项 → (标准规定, 单位, 该检项的基准结果)
ITEMS = {
    "性状":        ("白色或类白色粉末", "", "白色粉末"),
    "鉴别":        ("应与对照品一致", "", "符合"),
    "有关物质":    ("≤ 0.5%", "%", "0.12"),
    "含量测定":    ("98.0%~102.0%", "%", "99.6"),
    "干燥失重":    ("≤ 2.0%", "%", "0.35"),
    "溶解度":      ("应符合规定", "", "符合"),
    "溶液颜色":    ("≤ Y6", "号", "Y4"),
    "澄清度":      ("≤ 2 号浊度标准液", "号", "1"),
    "pH 值":       ("5.0~7.0", "", "6.2"),
    "重金属":      ("≤ 20ppm", "ppm", "8"),
    "炽灼残渣":    ("≤ 0.1%", "%", "0.03"),
    "微生物限度":  ("≤ 1000cfu/g", "cfu/g", "120"),
}

# 各批次抽查的检项：241012 只做 1 项（极少），241115 做 12 项（极多 → 明细表很宽）
BATCH_ITEMS = {
    "240905": ["性状", "鉴别", "有关物质", "含量测定", "干燥失重", "pH 值"],
    "240906": ["性状", "鉴别", "有关物质", "含量测定", "干燥失重", "pH 值"],
    "00903":  ["性状", "鉴别", "有关物质", "含量测定", "干燥失重", "pH 值"],
    "241012": ["含量测定"],
    "241115": ["性状", "鉴别", "有关物质", "含量测定", "干燥失重", "溶解度",
               "溶液颜色", "澄清度", "pH 值", "重金属", "炽灼残渣", "微生物限度"],
}

# 故意让一批数据不合格：241115 的 6 个月有关物质超标
OVER_LIMIT = {("241115", "6 个月", "有关物质"): "0.63"}
# 故意留空：结果还没出
PENDING = {("240906", "6 个月", "干燥失重"), ("241115", "3 个月", "重金属")}


def make_detail() -> list:
    """按考察方案拼出全部明细行。"""
    rows = []
    for batch, points in PLAN.items():
        for point in points:
            for item in BATCH_ITEMS[batch]:
                std, unit, base = ITEMS[item]
                key = (batch, point, item)
                if key in OVER_LIMIT:
                    result = OVER_LIMIT[key]
                elif key in PENDING:
                    result = ""                      # 结果还没出，格子是空的
                else:
                    result = base
                rows.append({
                    "批号": batch,
                    "时间点": point,
                    "检项": item,
                    "标准规定": std,
                    "检验结果": result,
                    "单位": unit,
                })
    return rows


def write_csv(rows: list) -> Path:
    """写成仪器导出风格的 CSV：表头一行 + 数据 + 空行 + 合计行 + 说明行。"""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    out = RAW_DIR / "稳定性考察_仪器导出.csv"
    cols = ["批号", "时间点", "检项", "标准规定", "检验结果", "单位"]
    # utf-8-sig：带 BOM，Excel 双击不按 ANSI 解码，中文才不会变乱码
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
        f.write("\n")
        f.write(",,,,,,\n")                                     # 空行
        f.write(f"合计,{len(rows)},,,,\n")                       # 合计行：条数落在「时间点」列
        f.write("说明：本文件由仪器工作站导出，仅供内部核对。,,,,,\n")
    return out


def write_plan(rows: list) -> Path:
    """把考察时间点的先后顺序单独落一份，让 run.py 不必靠猜来排时间点。"""
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    out = RAW_DIR / "考察方案_时间点顺序.csv"
    # 时间点在方案里的出现顺序就是考察顺序，去重保序
    seen, order = set(), []
    for r in rows:
        if r["时间点"] not in seen:
            seen.add(r["时间点"])
            order.append(r["时间点"])
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["时间点"])
        for p in order:
            w.writerow([p])
    return out


if __name__ == "__main__":
    data = make_detail()
    a = write_csv(data)
    b = write_plan(data)
    print(f"明细 {len(data)} 行 -> {a.name}")
    print(f"时间点 {len(set(r['时间点'] for r in data))} 个 -> {b.name}")
