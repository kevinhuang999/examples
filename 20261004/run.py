#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从系统导出的结果明细生成月度质量报表（口径表驱动）。

用法：
    python run.py              # 跑完暂停，方便双击看结果
    python run.py --no-pause   # 不暂停

依赖：openpyxl，缺了会自动走清华源安装。

输入：01_raw_data/lims_export_结果明细.csv（导出的明细）+ 01_raw_data/报表口径.xlsx（业务口径）
输出：02_output/月度质量报表_2026-09.xlsx（每个中心一个工作表 + 一张未纳入清单）
"""

import csv
import re
import socket
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
TPL_DIR = BASE / "source" / "templates"
LOG_FILE = BASE / "source" / "run_log.txt"

DETAIL_CSV = RAW_DIR / "lims_export_结果明细.csv"
SPEC_XLSX = RAW_DIR / "报表口径.xlsx"
REPORT_TPL = TPL_DIR / "月度质量报表模板.xlsx"

# —— 明细文件的列名：一律按名字取，不按第几列 ——
F_SAMPLE = "样本编号"
F_CENTER = "中心代码"
F_ITEM = "检测项目"
F_RESULT = "结果"
F_UNIT = "单位"
F_JUDGE = "判定"
F_STATUS = "状态"

STATUS_OK = "已出报告"                # 只有这个状态进统计
JUDGE_PASS, JUDGE_FAIL = "合格", "不合格"
REPORT_COLS = ["检测项目", "单位", "检测项数", "合格", "不合格", "未判定", "合格率(%)", "结果均值"]
SHEET_EXCLUDED = "未纳入清单"

EXCLUDE_REASON = {"已作废": "已作废，不进统计", "复检中": "复检中，结果未定"}

# 结果列里带单位时（例：98.6 g/L）拆成数值 + 单位；只给边界的写法（例：<0.01）走另一条
RE_NUM_UNIT = re.compile(r"^([-+]?\d+(?:\.\d+)?)\s+([A-Za-z%/][A-Za-z0-9/%.]*)$")
RE_LIMIT = re.compile(r"^([<>]=?)\s*([-+]?\d+(?:\.\d+)?)$")


def norm(text):
    """去掉所有空白（含全角空格），比较中心名 / 列名 / 项目名之前一律先过它。"""
    return "".join(str(text if text is not None else "").split())


def host_tag():
    """本段日志跑在哪台机器上：主机名（本机对外网卡 IP）。"""
    name = "未知"
    try:
        name = socket.gethostname()
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            sock.connect(("8.8.8.8", 80))
            return f"{name}（{sock.getsockname()[0]}）"
        finally:
            sock.close()
    except OSError:
        try:
            return f"{name}（{socket.gethostbyname(name)}）"
        except OSError:
            return f"{name}（未知）"


def log_line(text):
    with LOG_FILE.open("a", encoding="utf-8") as fh:
        fh.write(text + "\n")


def ensure_deps():
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        print("缺少 openpyxl，改用清华源安装……")
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", "openpyxl"])


def read_rows(path, sheet, key_col):
    """读一张表，返回 [{列名: 值}]。

    key_col 是这张表的判空列：这一列为空的行当表尾说明行跳过。
    每张表都有自己合适的判空列，写死某一列会把别的表整张读空。
    """
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    head = [norm(c) for c in rows[0]]
    if key_col not in head:
        raise ValueError(f"「{sheet}」里没有判空列「{key_col}」，表头是 {head}")
    idx = head.index(key_col)
    out = []
    for row in rows[1:]:
        if len(row) <= idx or norm(row[idx]) == "":
            continue
        out.append(dict(zip(head, row)))
    wb.close()
    return out


def parse_result(raw):
    """结果解析成 (形态, 数值, 结果列里带的单位)。

    形态：num 纯数值 / limit 只给边界（<0.01） / text 文字结论（未检出） / empty 空。
    只有 num 进均值——<0.01 按 0.01 计会把均值拉低，未检出按 0 计更离谱。
    """
    text = (raw or "").strip()
    unit = ""
    if not text:
        return "empty", None, unit
    hit = RE_NUM_UNIT.match(text)
    if hit:
        text, unit = hit.group(1), hit.group(2)      # 有的导出版本把单位写在结果里
    hit = RE_LIMIT.match(text)
    if hit:
        return "limit", float(hit.group(2)), unit
    try:
        return "num", float(text), unit
    except ValueError:
        return "text", None, unit


def read_detail():
    """读结果明细，逐行把结果解析成形态 + 数值。表尾说明行靠第一列判空跳过。"""
    out = []
    with DETAIL_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            if norm(row.get(F_SAMPLE)) == "":
                continue
            kind, value, unit_in_result = parse_result(row.get(F_RESULT))
            out.append({
                F_SAMPLE: norm(row.get(F_SAMPLE)),
                F_CENTER: norm(row.get(F_CENTER)),
                F_ITEM: norm(row.get(F_ITEM)),
                F_UNIT: norm(row.get(F_UNIT)) or unit_in_result,
                F_JUDGE: norm(row.get(F_JUDGE)),
                F_STATUS: norm(row.get(F_STATUS)),
                "形态": kind,
                "数值": value,
            })
    return out


def read_center_dict():
    """中心代码 -> 中心名称。名称里的全角空格在这里就被抹掉，免得工作表名跟着带。"""
    return {norm(r["中心代码"]): norm(r["中心名称"])
            for r in read_rows(SPEC_XLSX, "中心字典", "中心代码")}


def read_item_spec():
    """返回 ({项目: 口径}, [项目顺序])。顺序按口径表排，报表就照这个序出。"""
    spec, order = {}, []
    for row in read_rows(SPEC_XLSX, "项目口径", "检测项目"):
        item = norm(row["检测项目"])
        order.append(item)
        spec[item] = {
            "目标单位": norm(row["目标单位"]),
            "换算系数": float(row["换算系数"] or 1),
            "小数位": int(row["小数位"] or 2),
        }
    return spec, order


def read_report_def():
    """报表定义：每份报表一行，中心列表用竖线分隔。"""
    return read_rows(SPEC_XLSX, "报表定义", "报表名")


def aggregate(center, detail, item_spec, item_order):
    """一个中心一段明细：按项目归堆，出「一个项目一行」的统计。

    只吃状态为「已出报告」的记录；结果先按口径换算到目标单位再求均值。
    """
    rows = []
    for item in item_order:
        recs = [d for d in detail
                if d[F_CENTER] == center and d[F_ITEM] == item and d[F_STATUS] == STATUS_OK]
        if not recs:
            continue
        spec = item_spec.get(item, {})
        nums = [d["数值"] * spec.get("换算系数", 1.0) for d in recs if d["形态"] == "num"]
        passed = sum(1 for d in recs if d[F_JUDGE] == JUDGE_PASS)
        failed = sum(1 for d in recs if d[F_JUDGE] == JUDGE_FAIL)
        denom = passed + failed
        rows.append({
            "检测项目": item,
            "单位": spec.get("目标单位", ""),
            "检测项数": len(recs),
            "合格": passed,
            "不合格": failed,
            "未判定": len(recs) - passed - failed,
            "合格率(%)": round(passed / denom * 100, 1) if denom else None,
            "结果均值": round(sum(nums) / len(nums), spec.get("小数位", 2)) if nums else None,
        })
    return rows


def head_row(ws, label):
    """按第一列的标签反查行号——模板挪行、加行都不用改代码。"""
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, 1).value) == label:
            return r
    raise ValueError(f"模板里找不到标签「{label}」")


def fill_sheet(ws, center_name, rows):
    """把统计结果填进模板：目标行不够就清值留边框，一格不动样式。"""
    head = head_row(ws, "检测项目")
    total = head_row(ws, "合计")
    for r in range(1, head):
        if norm(ws.cell(r, 1).value) == "中心：":
            ws.cell(r, 2).value = center_name
    for i, row_no in enumerate(range(head + 1, total)):
        for j, col in enumerate(REPORT_COLS, 1):
            value = rows[i][col] if i < len(rows) else None
            ws.cell(row_no, j).value = value if value != "" else None
    ws.cell(total, 1).value = "合计"
    for j, col in zip(range(3, 7), ["检测项数", "合格", "不合格", "未判定"]):
        ws.cell(total, j).value = sum(r[col] for r in rows)
    denom = ws.cell(total, 4).value + ws.cell(total, 5).value
    ws.cell(total, 7).value = round(ws.cell(total, 4).value / denom * 100, 1) if denom else None


def add_excluded_sheet(wb, excluded):
    """未纳入清单：状态不是「已出报告」的记录全在这里，写清为什么没进统计。"""
    from openpyxl.styles import Border, Font, PatternFill, Side

    thin = Side(style="thin", color="9BAFC4")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    head = ["样本编号", "中心代码", "检测项目", "状态", "未纳入原因"]
    ws = wb.create_sheet(SHEET_EXCLUDED)
    for j, title in enumerate(head, 1):
        cell = ws.cell(1, j, title)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DCE6F1")
        cell.border = box
    for i, d in enumerate(excluded, start=2):
        ws.cell(i, 1, d[F_SAMPLE]).number_format = "@"    # 样本编号带前导零，按文本写
        ws.cell(i, 2, d[F_CENTER])
        ws.cell(i, 3, d[F_ITEM])
        ws.cell(i, 4, d[F_STATUS] or None)
        ws.cell(i, 5, EXCLUDE_REASON.get(d[F_STATUS], "状态列为空"))
        for j in range(1, len(head) + 1):
            ws.cell(i, j).border = box
    for col, width in zip("ABCDE", [14, 10, 14, 10, 22]):
        ws.column_dimensions[col].width = width


def build_reports(detail, centers, item_spec, item_order):
    """按报表定义逐份出报表；一份报表里每个中心一个工作表。"""
    from openpyxl import load_workbook

    made = []
    for rep in read_report_def():
        codes = [c.strip() for c in str(rep["中心列表"]).split("|") if c.strip()]
        wb = load_workbook(REPORT_TPL)
        template = wb["报表模板"]
        excluded = [d for d in detail if d[F_STATUS] != STATUS_OK and d[F_CENTER] in codes]
        first = True
        for code in codes:
            if code not in centers:
                raise ValueError(f"中心代码「{code}」在中心字典里找不到")
            name = centers[code]
            rows = aggregate(code, detail, item_spec, item_order)
            ws = template if first else wb.copy_worksheet(template)
            first = False
            fill_sheet(ws, name, rows)
            ws.title = name
            log_line(f"  {norm(rep['报表名'])} / {name}：{len(rows)} 个项目")
        add_excluded_sheet(wb, excluded)
        out_file = OUT_DIR / str(rep["输出文件名"])
        wb.save(out_file)
        wb.close()
        made.append((out_file, len(codes), len(excluded)))
    return made


def verify(out_file):
    """生成后回读断言：不读回来一遍，不敢说填对了。"""
    from openpyxl import load_workbook

    stats = {"ok": 0, "fail": 0}

    def check(label, got, want):
        if got == want:
            stats["ok"] += 1
            print(f"  OK   {label} = {got!r}")
        else:
            stats["fail"] += 1
            print(f"  FAIL {label} = {got!r}，期望 {want!r}")

    wb = load_workbook(out_file)

    def item_map(name):
        ws = wb[name]
        head = head_row(ws, "检测项目")
        total = head_row(ws, "合计")
        rows = {}
        for r in range(head + 1, total):
            item = norm(ws.cell(r, 1).value)
            if item:
                rows[item] = {"行": r, "明细": [ws.cell(r, j).value for j in range(1, 9)]}
        return ws, head, total, rows

    check("工作表清单", wb.sheetnames, ["上海中心", "北京中心", "广州中心", SHEET_EXCLUDED])

    ws, head, total, rows = item_map("上海中心")
    check("上海中心 项目数", len(rows), 8)
    check("上海中心 pH 整行", rows["pH"]["明细"], ["pH", None, 2, 2, 0, 0, 100.0, 6.58])
    check("上海中心 残留溶剂 单位（换算后）", rows["残留溶剂"]["明细"][1], "mg/dL")
    check("上海中心 残留溶剂 均值（12.5×0.1）", rows["残留溶剂"]["明细"][7], 1.25)
    check("上海中心 含量 合格率", rows["含量"]["明细"][6], 50.0)
    check("上海中心 有关物质 均值（0.42+1.86）/2", rows["有关物质"]["明细"][7], 1.14)
    check("上海中心 重金属 均值只吃数值", rows["重金属"]["明细"][7], 2.4)
    check("上海中心 微生物限度 均值按 1 位小数", rows["微生物限度"]["明细"][7], 120.0)
    check("上海中心 炽灼残渣 未判定", rows["炽灼残渣"]["明细"][5], 1)
    check("上海中心 炽灼残渣 合格率空", rows["炽灼残渣"]["明细"][6], None)
    check("上海中心 合计 检测项数", ws.cell(total, 3).value, 13)
    check("上海中心 合计 合格率", ws.cell(total, 7).value, 83.3)
    check("上海中心 第 9~12 行已清空",
          [ws.cell(head + 9, 1).value, ws.cell(head + 12, 1).value], [None, None])

    ws2, head2, total2, rows2 = item_map("北京中心")
    check("北京中心 项目数", len(rows2), 2)
    check("北京中心 第 3 行清值但边框还在",
          [ws2.cell(head2 + 3, 1).value, ws2.cell(head2 + 3, 1).border.left.style], [None, "thin"])
    check("北京中心 合计 合格率", ws2.cell(total2, 7).value, 75.0)

    ws3, head3, total3, rows3 = item_map("广州中心")
    check("广州中心 项目数", len(rows3), 5)
    check("广州中心 含量 项数（含带单位那条）", rows3["含量"]["明细"][2], 4)
    check("广州中心 含量 均值", rows3["含量"]["明细"][7], 97.5)
    check("广州中心 合计 检测项数", ws3.cell(total3, 3).value, 14)
    check("广州中心 合计 合格率", ws3.cell(total3, 7).value, 92.3)

    wsx = wb[SHEET_EXCLUDED]
    ids = [wsx.cell(r, 1).value for r in range(2, wsx.max_row + 1)]
    check("未纳入清单 条数", len(ids), 4)
    check("未纳入清单 前导零样本编号", "0000081" in ids, True)
    check("未纳入清单 编号按文本格式", wsx.cell(2, 1).number_format, "@")
    check("未纳入清单 状态为空的兜底原因",
          wsx.cell(ids.index("0000091") + 2, 5).value, "状态列为空")

    wb.close()
    print(f"  回读断言：{stats['ok']} 项 OK，{stats['fail']} 项 FAIL")
    return stats


def main():
    started = time.time()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    for old in OUT_DIR.iterdir():
        if old.is_file():
            old.unlink()                       # 重跑先清，免得上一版多出来的文件被当这一版

    ensure_deps()
    log_line(f"运行主机：{host_tag()}")
    log_line(f"开始执行：{time.strftime('%Y-%m-%d %H:%M:%S')}  "
             f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）")

    detail = read_detail()
    centers = read_center_dict()
    item_spec, item_order = read_item_spec()
    log_line(f"读入：明细 {len(detail)} 条 / 中心 {len(centers)} 个 / 项目口径 {len(item_order)} 条")
    print(f"读入明细 {len(detail)} 条，中心 {len(centers)} 个，项目口径 {len(item_order)} 条")

    made = build_reports(detail, centers, item_spec, item_order)
    print(f"生成报表 {len(made)} 份：")
    for out_file, sheets, excluded in made:
        print(f"  {out_file.name}（{sheets} 个工作表 + 未纳入清单 {excluded} 条）")

    print("生成后回读校验：")
    stats = verify(made[0][0])
    passed = stats["fail"] == 0
    verdict = "PASS" if passed else "FAIL"
    span = time.time() - started
    log_line(f"执行完成：{time.strftime('%Y-%m-%d %H:%M:%S')}  用时 {span:.1f}s  判定 {verdict}")
    print(f"判定：{verdict}（用时 {span:.1f}s）")

    if "--no-pause" not in sys.argv and sys.stdin and sys.stdin.isatty():
        input("按回车结束……")
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
