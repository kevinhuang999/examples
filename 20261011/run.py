# -*- coding: utf-8 -*-
"""实验数据自动校验 + 超限报警 —— 可运行示例

场景：一批样品检测完，结果从仪器导出后汇总成一张明细表；规格限度写在另一张表里。
一批几十上百条结果要逐条对着限度判「超没超」，人工比对慢、还容易漏；真超了往往
是等日报汇总出来才发现。本示例把明细与限度表读进来，逐条判 合格 / 超限 / 未判定，
超限的单独出一张报警表、未判定的另出一张（写明原因），每一步留痕。

用法：
    python run.py              # 跑完暂停（双击也能用）
    python run.py --no-pause   # 不暂停（脚本 / CI 里跑）
"""
import argparse
import csv
import io
import re
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
TPL_DIR = SRC_DIR / "templates"
SPEC_FILE = TPL_DIR / "规格限度表.csv"
DETAIL_FILE = RAW_DIR / "检测结果明细.csv"
LOG_FILE = SRC_DIR / "run_log.txt"

OUT_NAME = "实验数据校验判定.xlsx"
NO_COL = "样品编号"
ITEM_COL = "检测项目"
RESULT_COL = "结果"
UNIT_COL = "单位"
DATE_COL = "检测日期"

ALIASES = {
    ITEM_COL: ("检测项目", "项目", "检验项目"),
    RESULT_COL: ("结果", "检测结果", "实测值"),
    UNIT_COL: ("单位", "单位名称"),
    NO_COL: ("样品编号", "样品号", "编号"),
    DATE_COL: ("检测日期", "日期", "完成日期"),
}

NO_LIMIT = "不得检出"
ND_TOKENS = {"未检出", "未测出", "nd", "n.d.", "n/a", "na", "—", "<lod", "≤lod", "lod"}
LIMIT_RE = re.compile(r"^([<>≤≥]=?)\s*(-?\d+(?:\.\d+)?)\s*([^\s\d]*)$")
NUM_RE = re.compile(r"^(-?\d+(?:\.\d+)?)\s*([^\s\d]*)$")


def ensure_deps():
    """缺依赖自动装（清华源）；已装则直接返回。"""
    need = []
    for mod, pkg in (("openpyxl", "openpyxl"),):
        try:
            __import__(mod)
        except ImportError:
            need.append(pkg)
    if not need:
        return
    print("缺少依赖：", ", ".join(need), "，正在从清华源安装……")
    cmd = [sys.executable, "-m", "pip", "install", "-i",
           "https://pypi.tuna.tsinghua.edu.cn/simple", *need]
    if subprocess.call(cmd) != 0:
        subprocess.call(cmd + ["--user"])


def norm(name):
    """列名 / 项目名归一：去空白（含全角空格）、全角括号转半角、英文小写。

    明细里的项目名尾巴上多一个全角空格（`炽灼残渣　`）肉眼看不出来，
    直接比对就会判成「限度表里没有这一项」。归一之后再比，这类假挂起就没了。
    """
    s = str(name) if name is not None else ""
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[\s\u00a0\u3000]+", "", s)
    return s.lower()


def to_float(s):
    """空字符串不能当 0——「没有下限」和「下限是 0」是两回事。"""
    s = str(s or "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def load_specs(path):
    """读规格限度表：检测项目 → 判定方式 / 下限 / 上限 / 单位。

    限度是业务资产，写在表里、不写死在代码里——限度改一版只是改这张表，
    不用动代码、不用重新核对脚本。这是校验类脚本能长期用下去的前提。
    """
    with path.open(encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.reader(f) if any(str(c).strip() for c in r)]
    specs = {}
    for r in rows[1:]:
        item = str(r[0]).strip()
        if not item:
            continue
        mode = str(r[1]).strip() if len(r) > 1 else ""
        specs[norm(item)] = {
            "项目": item,
            "方式": mode or "上限",
            "下限": to_float(r[2]) if len(r) > 2 else None,
            "上限": to_float(r[3]) if len(r) > 3 else None,
            "单位": str(r[4]).strip() if len(r) > 4 else "",
        }
    return specs


def spec_text(spec):
    """把限度写成一行给人看的写法（报警表里要能一眼看出按哪条限度判的）。"""
    if spec is None:
        return "—"
    u = spec["单位"]
    if spec["方式"] == NO_LIMIT or (spec["下限"] is None and spec["上限"] is None):
        return NO_LIMIT
    if spec["下限"] is not None and spec["上限"] is not None:
        return f"{spec['下限']:g} ~ {spec['上限']:g} {u}".strip()
    if spec["上限"] is not None:
        return f"≤ {spec['上限']:g} {u}".strip()
    return f"≥ {spec['下限']:g} {u}".strip()


def read_text(path):
    """按 utf-8-sig → gbk → utf-8 依次试；GBK 文件用 utf-8 读会直接抛异常，不是读出乱码。"""
    last = None
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            return path.read_text(encoding=enc), enc
        except UnicodeDecodeError as e:
            last = e
    raise last


def read_matrix(path):
    """把一个导出文件读成二维表（不做任何解释），并返回实际用的编码。"""
    if path.suffix.lower() == ".csv":
        text, enc = read_text(path)
        return list(csv.reader(io.StringIO(text))), enc
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows, "xlsx"


def find_header(rows):
    """找表头行：第一个「非空单元格 >= 2」的行。

    不能写死 header=0——仪器软件导出时会在表头上压几行导出信息；
    也不能靠行号猜，导出格式改一版行号就全错。
    """
    for i, r in enumerate(rows):
        if sum(1 for c in r if str(c or "").strip()) >= 2:
            return i
    return -1


def pick_columns(header):
    """按别名表把要用的列找出来；找不到的留 None。

    列名写死成「结果」的话，遇到写成「检测结果」的模板整张就读空了，
    而且一行错都不报——所以列名一律走别名表，取不到显式留空、由后面判空兜住。
    """
    idx = {}
    for name, aliases in ALIASES.items():
        idx[name] = None
        wanted = {norm(a) for a in aliases}
        for j, h in enumerate(header):
            if norm(h) in wanted:
                idx[name] = j
                break
    return idx


def cell(cells, i):
    """按列号取单元格；列号是 None（那列没找到）时返回空串，不抛异常。"""
    if i is None or i < 0 or i >= len(cells):
        return ""
    return str(cells[i]).strip()


def parse_result(text):
    """把结果单元格拆成 (形态, 数值, 单位)。

    形态只有五种：
      num   —— 一个确定的数（`99.4`、`98.6 g/L`，单位可能写在结果里）
      limit —— 限号值（`<0.01`、`≤0.5`），只给了边界的数
      nd    —— 文字结论（未检出 / ND / —）
      text  —— 认不出的文字（如「见附页」）
      empty —— 空的（待检或漏录）

    为什么要分形态：`<0.01` 与「未检出」不能当 0 算、也不能当 0.01 算——
    它们只是「低于某个界」，直接转成数字再判，超限结论会反过来。
    """
    s = str(text or "").strip()
    if not s:
        return "empty", None, ""
    if s.lower().replace(" ", "") in ND_TOKENS:
        return "nd", None, ""
    m = LIMIT_RE.match(s)
    if m:
        return "limit", float(m.group(2)), m.group(3).strip()
    m = NUM_RE.match(s)
    if m:
        return "num", float(m.group(1)), m.group(2).strip()
    return "text", None, s


def judge_one(item, res_text, unit_col, spec):
    """一条结果的判定：返回 (判定, 方向, 说明)。

    判定只有三种：合格 / 超限 / 未判定。
    「未判定」不等于「合格」——限度缺失、结果空缺、单位对不上的，一律挂起来等人确认：
    悄悄放过是放走了风险，顺手判成超限又会把报警表淹掉，两头都错。
    """
    kind, num, unit_in = parse_result(res_text)
    raw = str(res_text or "").strip()

    if spec is None:
        return "未判定", "", f"规格限度表里没有「{item}」这一项，无法判定"
    if kind == "empty":
        return "未判定", "", "结果列是空的（待检或漏录）"
    if kind == "text":
        return "未判定", "", f"结果不是数值也不是标准的未检出写法（原样保留：{raw}）"

    unit = unit_col or unit_in            # 单位列优先；列空的用结果里拆出来的
    if unit and spec["单位"] and norm(unit) != norm(spec["单位"]):
        return "未判定", "", f"结果单位 {unit} 与限度单位 {spec['单位']} 不一致"

    if kind == "nd":
        if spec["方式"] == NO_LIMIT:
            return "合格", "", "未检出，满足不得检出的要求"
        if spec["上限"] is not None:
            return "合格", "", "未检出，低于上限"
        return "未判定", "", "未检出无法与下限比较"

    if spec["方式"] == NO_LIMIT:
        return "超限", "检出", f"该项要求不得检出，实测 {raw} {unit}".strip()

    if kind == "limit":                   # <0.01 这类只给了边界，只与上限比
        if spec["上限"] is not None and num <= spec["上限"]:
            return "合格", "", f"{raw} 低于上限"
        if spec["上限"] is None:
            return "未判定", "", f"{raw} 无法与下限比较"
        return "超限", "偏高", f"{raw} 高于上限 {spec['上限']:g} {spec['单位']}".strip()

    if spec["下限"] is not None and num < spec["下限"]:
        return "超限", "偏低", f"{raw} 低于下限 {spec['下限']:g} {spec['单位']}".strip()
    if spec["上限"] is not None and num > spec["上限"]:
        return "超限", "偏高", f"{raw} 高于上限 {spec['上限']:g} {spec['单位']}".strip()
    return "合格", "", "在限度内"


def check_all(specs, rows, hdr, cols):
    """逐行判定，返回 (判定明细, 超限报警, 未判定清单)。"""
    detail, alarms, undecided = [], [], []
    for r in rows[hdr + 1:]:
        cells = [str(c or "").strip() for c in r]
        if not any(cells):
            continue
        item = cell(cells, cols.get(ITEM_COL))
        res = cell(cells, cols.get(RESULT_COL))
        unit = cell(cells, cols.get(UNIT_COL))
        spec = specs.get(norm(item))
        verdict, direction, why = judge_one(item, res, unit, spec)
        row = {NO_COL: cell(cells, cols.get(NO_COL)), ITEM_COL: item,
               RESULT_COL: res, UNIT_COL: unit,
               DATE_COL: cell(cells, cols.get(DATE_COL)),
               "适用限度": spec_text(spec), "判定": verdict,
               "方向": direction, "说明": why}
        detail.append(row)
        if verdict == "超限":
            alarms.append(row)
        elif verdict == "未判定":
            undecided.append(row)
    return detail, alarms, undecided


def summarize(detail, alarms, undecided):
    """汇总：合格率的分母是「合格 + 超限」，不含未判定。

    没判定的记录不能进分母——否则一项还没出结果，合格率就被算成 0 或虚高。
    """
    passed = len(detail) - len(alarms) - len(undecided)
    judged = passed + len(alarms)
    rate = round(passed / judged, 4) if judged else 0.0
    return {"明细条数": len(detail), "合格": passed, "超限": len(alarms),
            "未判定": len(undecided), "合格率": rate}


def write_outputs(detail, alarms, undecided):
    """四张表一次落盘：判定明细 / 超限报警 / 未判定清单 / 校验汇总。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / OUT_NAME
    cols = [NO_COL, ITEM_COL, RESULT_COL, UNIT_COL, DATE_COL,
            "适用限度", "判定", "方向", "说明"]
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = Workbook()

    ws = wb.active
    ws.title = "判定明细"
    ws.append(cols)
    for r in detail:
        ws.append([r.get(c, "") for c in cols])
        if r["判定"] == "超限":                       # 报警行整行标红，扫表时一眼看得到
            for j in range(1, len(cols) + 1):
                ws.cell(row=ws.max_row, column=j).fill = red
    for j in range(1, len(cols) + 1):
        ws.cell(row=1, column=j).font = Font(bold=True)
    for i in range(2, ws.max_row + 1):                # 样品编号按文本存，前导零才不会掉
        ws.cell(row=i, column=1).number_format = "@"

    ws2 = wb.create_sheet("超限报警")
    ws2.append(cols[:5] + ["适用限度", "方向", "说明"])
    for r in alarms:
        ws2.append([r.get(c, "") for c in cols[:5]] + [r["适用限度"], r["方向"], r["说明"]])
        for j in range(1, 9):
            ws2.cell(row=ws2.max_row, column=j).fill = red
    ws2.cell(row=1, column=1).font = Font(bold=True)
    for i in range(2, ws2.max_row + 1):
        ws2.cell(row=i, column=1).number_format = "@"

    ws3 = wb.create_sheet("未判定清单")
    ws3.append([NO_COL, ITEM_COL, RESULT_COL, UNIT_COL, "原因"])
    for r in undecided:
        ws3.append([r[NO_COL], r[ITEM_COL], r[RESULT_COL], r[UNIT_COL], r["说明"]])
    for i in range(2, ws3.max_row + 1):
        ws3.cell(row=i, column=1).number_format = "@"

    s = summarize(detail, alarms, undecided)
    ws4 = wb.create_sheet("校验汇总")
    ws4.append(["指标", "数值"])
    for k, v in s.items():
        ws4.append([k, v])
    ws4.cell(row=1, column=1).font = Font(bold=True)

    wb.save(out)
    return out


def verify(detail, alarms, undecided):
    """生成后回读：逐项断言，打印 OK / FAIL，返回总判定。"""
    from openpyxl import load_workbook
    out = OUT_DIR / OUT_NAME
    s = summarize(detail, alarms, undecided)
    checks = [("判定表已生成", out.exists() and out.stat().st_size > 0)]
    checks.append(("明细条数 = 26", s["明细条数"] == 26))
    checks.append(("超限 = 6 条", s["超限"] == 6))
    checks.append(("未判定 = 5 条", s["未判定"] == 5))
    checks.append(("合格 = 15 条", s["合格"] == 15))
    checks.append(("合格率分母不含未判定（15/21）", abs(s["合格率"] - 15 / 21) < 1e-4))
    checks.append(("恰好等于上限判合格（水分 0.5 / 重金属 10）",
                   all(a[ITEM_COL] != "水分" or a[NO_COL] != "0000102" for a in alarms)))
    checks.append(("略超上限判超限（含量 102.3 / 炽灼残渣 0.12）",
                   {(a[NO_COL], a[ITEM_COL]) for a in alarms} >=
                   {("0000103", "含量"), ("0000113", "炽灼残渣")}))
    checks.append(("低于下限判超限且方向为偏低（含量 97.6）",
                   any(a[NO_COL] == "0000102" and a["方向"] == "偏低" for a in alarms)))
    checks.append(("「不得检出」项目数值结果判超限（黄曲霉毒素 0.6）",
                   any(a[NO_COL] == "0000107" for a in alarms)))
    checks.append(("「不得检出」项目未检出判合格",
                   any(d[NO_COL] == "0000106" and d["判定"] == "合格" for d in detail)))
    checks.append(("限号值 <0.01 判合格、不当 0 也不当超限",
                   any(d[NO_COL] == "0000109" and d[ITEM_COL] == "水分"
                       and d["判定"] == "合格" for d in detail)))
    checks.append(("单位不一致判未判定（98.6 g/L 对 %）",
                   any(u[NO_COL] == "0000108" for u in undecided)))
    checks.append(("结果为空判未判定（0000109 / 0000115 各一条）",
                   len([u for u in undecided if not str(u[RESULT_COL]).strip()]) == 2))
    checks.append(("项目不在限度表里判未判定（电导率）",
                   any(u[ITEM_COL] == "电导率" for u in undecided)))
    checks.append(("限度表里有、明细里没有的项目不报警（不溶性微粒无数据）",
                   all(a[ITEM_COL] != "不溶性微粒" for a in alarms)))
    if out.exists():
        wb = load_workbook(out)
        ws = wb["判定明细"]
        checks.append(("样品编号列已设文本格式", ws.cell(row=2, column=1).number_format == "@"))
        checks.append(("报警表条数 = 6", wb["超限报警"].max_row - 1 == 6))
        checks.append(("报警行已标红（判定明细首条报警）",
                       ws.cell(row=[i for i in range(2, ws.max_row + 1)
                                    if ws.cell(row=i, column=7).value == "超限"][0],
                               column=1).fill.fgColor.rgb == "00FFC7CE"))
    ok = True
    for name, passed in checks:
        print(("  OK   " if passed else "  FAIL ") + name)
        ok = ok and bool(passed)
    return ok


def host_info():
    """主机名 + 本机对外 IP，用来分辨这段日志是哪台机器跑的。"""
    name = socket.gethostname()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except OSError:
        try:
            ip = socket.gethostbyname(name)
        except OSError:
            ip = "未知"
    return name, ip


def append_log(text):
    """日志追加不覆盖；每段开头记运行主机，方便分辨机器。"""
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-pause", action="store_true", help="跑完不暂停")
    args = ap.parse_args()

    ensure_deps()
    if not SPEC_FILE.exists() or not DETAIL_FILE.exists():
        print("找不到规格限度表或结果明细：", SPEC_FILE, DETAIL_FILE)
        print("先跑 source/build_fixtures.py 重造模板与数据。")
        return 1
    SRC_DIR.mkdir(parents=True, exist_ok=True)

    t0 = time.time()
    host, ip = host_info()
    append_log(f"\n运行主机：{host}（{ip}）\n开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}\n"
               f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）\n")

    print("=" * 62)
    print("实验数据自动校验 + 超限报警")
    print("=" * 62)
    specs = load_specs(SPEC_FILE)
    print(f"规格限度表：{len(specs)} 项")

    rows, enc = read_matrix(DETAIL_FILE)
    hdr = find_header(rows)
    cols = pick_columns(rows[hdr])
    detail, alarms, undecided = check_all(specs, rows, hdr, cols)
    s = summarize(detail, alarms, undecided)
    print(f"读入明细 {enc} 编码、表头在第 {hdr + 1} 行 → 判定 {s['明细条数']} 条："
          f"合格 {s['合格']} / 超限 {s['超限']} / 未判定 {s['未判定']}"
          f"（合格率 {s['合格率']:.1%}）")
    for a in alarms:
        print(f"  [超限] {a[NO_COL]} {a[ITEM_COL]} = {a[RESULT_COL]} → {a['说明']}")

    out = write_outputs(detail, alarms, undecided)
    print(f"已生成：{out.name}")
    ok = verify(detail, alarms, undecided)

    used = time.time() - t0
    verdict = "PASS" if ok else "FAIL"
    print(f"\n总判定：{verdict}（用时 {used:.1f}s）")
    append_log(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  用时 {used:.1f}s  判定 {verdict}\n")
    if not args.no_pause:
        input("\n按回车键退出……")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
