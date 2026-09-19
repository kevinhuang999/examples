# -*- coding: utf-8 -*-
"""引流长文《临床试验访视时点报告能否自动生成，超窗漏项只能人工一条条对》配套脚本。

按方案规定的访视窗与检项清单，把仪器导出的检验结果，拼成一份份访视时点报告：

    访视计划（一行一个 受试者 x 访视点：计划日期、窗前两天、窗后几天）
  + 检验结果（一行一个 受试者 x 访视点 x 检项，仪器导出）
  + 访视点检项清单（方案规定：每个访视点该做哪几项）
→ 每个到过访的时点出一份报告；谁提前到了、谁晚到了、谁漏做了检查、谁的结果还没出，
  都写进报告和台账；计划表里对不上的编号单列，不混进台账。

核心代码与正文逐字一致，末尾多了自动回读校验，方便确认该标出来的都标出来了。
跨平台：Windows / macOS / Linux 都是一条命令，缺依赖会自动装。
    python run.py
    python run.py --no-pause   # 不暂停（CI 或脚本里用）

打包给别人时，把整个目录拷走即可，路径全部相对定位，无硬编码。
"""
import csv
import shutil
import socket
import subprocess
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

# 控制台编码随环境变，只加容错，不强行改编码。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"
NO_PAUSE = "--no-pause" in sys.argv


def _pause(msg=""):
    """Windows 双击 .py 时留住窗口，避免一闪而过。"""
    if msg:
        print(msg)
    if NO_PAUSE:
        return
    try:
        if sys.platform == "win32" and sys.stdin and sys.stdin.isatty():
            input("按回车退出")
    except Exception:
        pass


def _ensure(*modules_pkgs):
    """缺哪个装哪个。先直接装，失败再退到 --user（无管理员权限的机器）。"""
    for mod, pkg in modules_pkgs:
        try:
            __import__(mod)
            continue
        except ImportError:
            pass
        print(f"缺 {pkg}，正在用清华源安装（解释器: {sys.executable}）...")
        base = [sys.executable, "-m", "pip", "install", "-i", MIRROR, pkg]
        attempts = [base]
        if sys.prefix == sys.base_prefix:      # 不在虚拟环境里，普通安装可能没权限
            attempts.append(base[:4] + ["--user"] + base[4:])
        for cmd in attempts:
            subprocess.run(cmd, check=False)
            try:
                __import__(mod)
                break
            except ImportError:
                continue
        else:
            _pause(f"自动安装失败。手动执行：\n  {sys.executable} -m pip install -i {MIRROR} {pkg}")
            sys.exit(1)


_ensure(("openpyxl", "openpyxl"))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
PLAN_FILE = HERE / "01_raw_data" / "访视计划.csv"                # 一行一个 受试者 x 访视点
RESULT_FILE = HERE / "01_raw_data" / "检验结果_仪器导出.csv"      # 一行一个 检项结果
PANEL_FILE = HERE / "source" / "templates" / "访视点检项清单.csv"  # 方案规定：每个访视点做哪几项
REPORT_DIR = HERE / "02_output" / "访视时点报告"                  # 一份报告一个文件
LEDGER_FILE = HERE / "02_output" / "访视执行台账.xlsx"            # 台账 + 游离记录
LOG_DIR = HERE / "source"                                        # 运行日志（run_log.txt 落在 source/ 下）

PLAN_COLS = ["中心编号", "中心名称", "受试者编号", "访视点", "计划日期", "窗前天", "窗后天"]
RESULT_COLS = ["受试者编号", "访视点", "采样日期", "检项", "结果", "单位", "参考下限", "参考上限"]
PANEL_COLS = ["访视点", "访视序号", "检项"]

LEDGER_COLS = ["中心编号", "中心名称", "受试者编号", "访视点", "计划日期", "窗口起", "窗口止",
               "实际采样日期", "偏离天数", "到访判定", "应做项数", "实做项数", "缺失检项", "结论"]
STRAY_COLS = ["受试者编号", "访视点", "采样日期", "检项", "说明"]
REPORT_HEAD = ["检项", "结果", "单位", "参考范围", "判定"]

# 数据导出日。窗口终点还没到的计划行不能报"未到访"——人家本来就还没来。
DATA_CUTOFF = date(2026, 9, 30)

# 结论词。一次访视可能同时踩两件事，所以用「、」拼起来，不按优先级二选一。
J_OPEN = "未到期"
J_NOVISIT = "未到访"
J_WINDOW = "超窗"
J_MISS = "缺项"
J_NORESULT = "未出结果"
J_DUP = "同项多重结果"
J_OK = "正常"

I_UNKNOWN = "—"

HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
OK_FILL = PatternFill("solid", fgColor="D9EAD3")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")    # 还没到 / 结果本身超范围
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")     # 要人跟的：超窗、缺项、没出结果
GRAY_FILL = PatternFill("solid", fgColor="F2F2F2")

LABEL_FONT = Font(name="微软雅黑", size=10, bold=True)
BODY_FONT = Font(name="微软雅黑", size=10)
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def to_int(text: str) -> int:
    """窗口天数写成空、带单位的，一律当 0 天。"""
    try:
        return int(float(str(text).strip()))
    except Exception:
        return 0


def to_float(text: str):
    """转不出数就返回 None——"<0.5""未检出"这类都不是数，不能当 0 处理。"""
    try:
        return float(str(text).strip())
    except Exception:
        return None


def as_date(text: str):
    """日期串转 date；转不出来返回 None，别拿今天顶上。"""
    try:
        return date.fromisoformat(str(text).strip())
    except Exception:
        return None


def safe_name(text: str) -> str:
    """文件名里不能带 \\ / : * ? " < > |，换掉再拼。"""
    bad = '\\/:*?"<>|'
    return "".join("-" if c in bad else c for c in str(text)).strip() or "未命名"


def read_rows(path: Path, cols, key_col: str):
    """读一张导出表：BOM 去掉、每列去空白、编号空的行丢掉。

    丢的是表尾那两条：合计行（条数填在检项列）和说明行——留着会被当成一条真记录。
    """
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for raw in csv.DictReader(f):
            if not (raw.get(key_col) or "").strip():
                continue
            rows.append({c: (raw.get(c) or "").strip() for c in cols})
    return rows


def read_plan():
    return read_rows(PLAN_FILE, PLAN_COLS, "受试者编号")


def read_panel():
    """访视点 -> (访视序号, [检项...])。报告的先后按访视序号排，不按访视点名字排。"""
    panel = {}
    for r in read_rows(PANEL_FILE, PANEL_COLS, "访视点"):
        _, items = panel.setdefault(r["访视点"], (to_int(r["访视序号"]), []))
        if r["检项"] not in items:
            items.append(r["检项"])
    return panel


def read_results():
    return read_rows(RESULT_FILE, RESULT_COLS, "受试者编号")


def group_visits(results, plan_keys):
    """按 (受试者编号, 访视点) 把结果行归成一次访视。

    - 计划表里查不到的组合单列成游离：既不能丢（不知道现场抄错了哪一条），
      也不能并进台账（并错等于记到别人头上）；
    - 同一次访视里同一个检项出了两条（复测重出），先到的留下，后到的记进「复测」等人工定；
    - 到访日取这一堆里最早的采样日期：复测不算重新到访，不能拿最后一条把超窗洗掉。
    """
    groups, strays = {}, []
    for r in results:
        key = (r["受试者编号"], r["访视点"])
        if key not in plan_keys:
            strays.append(r)
            continue
        g = groups.setdefault(key, {"采样日期": [], "项": {}, "复测": []})
        if r["采样日期"]:
            g["采样日期"].append(r["采样日期"])
        if r["检项"] in g["项"]:
            g["复测"].append(r)
            continue
        g["项"][r["检项"]] = r
    return groups, strays


def window_of(plan):
    """窗口两头 = 计划日往前、往后各让几天。两头分开取，别写成对称的 ±N。"""
    day = as_date(plan["计划日期"])
    return day - timedelta(days=to_int(plan["窗前天"])), day + timedelta(days=to_int(plan["窗后天"]))


def judge_visit(plan, group, panel, items_done):
    """一次访视的判定：到访在不在窗里、方案要做的项漏了哪几项。

    偏离天数带符号：正数是晚到，负数是提前到——提前到也是超窗，只看晚到会漏掉一半。
    """
    start, end = window_of(plan)
    days = sorted(group["采样日期"]) if group else []
    visit_day = as_date(days[0]) if days else None
    should = panel.get(plan["访视点"], (0, []))[1]
    missing = [i for i in should if i not in items_done]

    if visit_day is None:
        verdict = J_OPEN if end > DATA_CUTOFF else J_NOVISIT
        arrive = verdict              # 台账里这两档也写「未到期 / 未到访」，不混成一个词
        drift = None
        missing = []                 # 一次都没来，"漏了哪项"没意义，全列出来只会淹掉台账
    else:
        drift = 0
        if visit_day < start:
            drift = (visit_day - start).days
        elif visit_day > end:
            drift = (visit_day - end).days
        arrive = J_OK if drift == 0 else J_WINDOW
        tags = []
        if drift:
            tags.append(J_WINDOW)
        if missing:
            tags.append(J_MISS)
        if any(not r["结果"] for r in group["项"].values()):
            tags.append(J_NORESULT)
        if group["复测"]:
            tags.append(J_DUP)
        verdict = "、".join(tags) if tags else J_OK

    return {
        "窗口起": start.isoformat(),
        "窗口止": end.isoformat(),
        "实际采样日期": visit_day.isoformat() if visit_day else "",
        "到访判定": arrive,
        "偏离天数": drift,
        "应做项数": len(should),
        "实做项数": len(items_done),
        "缺失检项": "、".join(missing) or I_UNKNOWN,
        "结论": verdict,
    }


def range_text(low: str, high: str) -> str:
    """参考范围怎么印。只给了一头的（如 ≤50）就只印那一头，空的那头写 0 是把范围写错。"""
    lo, hi = str(low).strip(), str(high).strip()
    if lo and hi:
        return f"{lo} ~ {hi}"
    if hi:
        return f"≤ {hi}"
    if lo:
        return f"≥ {lo}"
    return I_UNKNOWN


def judge_item(row) -> str:
    """一个检项的结果判定。

    空结果 = 本次没出结果，不是正常；非数值（"未见异常""阴性"）是文字结果，不判范围；
    是数值才拿去跟参考范围比，且只比给了的那一头。
    """
    text = row["结果"]
    if not text:
        return "未出结果"
    num = to_float(text)
    if num is None:
        return "文字结果"
    low, high = to_float(row["参考下限"]), to_float(row["参考上限"])
    if low is None and high is None:
        return "无参考范围"
    if low is not None and num < low:
        return "超出范围"
    if high is not None and num > high:
        return "超出范围"
    return "在范围内"


def report_items(plan, group, panel):
    """报告明细：先按方案清单的顺序铺，没做的写成"未做"；方案外多做的排最后。

    顺序不能按归档顺序来——报告是拿去对方案的，得跟方案一个次序，少哪一项一眼就看到。
    方案外多做的也不能悄悄丢，那在核查里是"报告与原始记录不符"。
    """
    done = group["项"] if group else {}
    should = panel.get(plan["访视点"], (0, []))[1]
    rows = []
    for item in should:
        r = done.get(item)
        if r is None:
            rows.append([item, "未做", "", "", "未做"])
        else:
            rows.append([item, r["结果"], r["单位"], range_text(r["参考下限"], r["参考上限"]),
                         judge_item(r)])
    for item, r in done.items():
        if item not in should:
            rows.append([item, r["结果"], r["单位"], range_text(r["参考下限"], r["参考上限"]),
                         "方案外项目"])
    return rows


def write_report(path: Path, plan, judged, rows):
    """一份访视报告：抬头 + 明细 + 结论。样式随代码走，不依赖谁先排好的模板。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "访视报告"
    for i, w in enumerate([16, 14, 12, 20, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells("A1:E1")
    ws["A1"] = "临床检验结果报告（访视时点报告）"
    ws["A1"].font = Font(name="微软雅黑", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    head = [
        ("中心", f"{plan['中心编号']} {plan['中心名称']}", "访视点", plan["访视点"]),
        ("受试者编号", plan["受试者编号"], "计划日期", plan["计划日期"]),
        ("允许窗口", f"{judged['窗口起']} ~ {judged['窗口止']}",
         "实际采样日期", judged["实际采样日期"] or I_UNKNOWN),
        ("到访判定", judged["到访判定"], "偏离天数",
         I_UNKNOWN if judged["偏离天数"] is None else str(judged["偏离天数"])),
    ]
    r = 2
    for a, b, c, d in head:
        ws.cell(r, 1, a).font = LABEL_FONT
        ws.cell(r, 2, b).font = BODY_FONT
        ws.cell(r, 4, c).font = LABEL_FONT
        ws.cell(r, 5, d).font = BODY_FONT
        r += 1

    # 结果本身超范围是要写出来的异常值，跟"到访超窗"不是一回事：前者是数据，后者是流程。
    for c, name in enumerate(REPORT_HEAD, 1):
        cell = ws.cell(r, c, name)
        cell.font = LABEL_FONT
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center")
        cell.border = BOX
    r += 1

    item_fill = {"超出范围": WARN_FILL, "未做": BAD_FILL, "未出结果": BAD_FILL}
    for row in rows:
        for c, value in enumerate(row, 1):
            cell = ws.cell(r, c, value)
            cell.font = BODY_FONT
            cell.border = BOX
            if c == 5 and value in item_fill:
                cell.fill = item_fill[value]
        r += 1

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws.cell(r, 1, "结论").font = LABEL_FONT
    cell = ws.cell(r, 2, judged["结论"])
    cell.font = BODY_FONT
    cell.fill = OK_FILL if judged["结论"] == J_OK else BAD_FILL
    wb.save(path)


def write_ledger(rows, strays, path: Path):
    """台账：一行一个访视点，谁该跟一眼就看出来；游离记录另开一页，不跟前两页混。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "访视执行台账"
    for i, w in enumerate([10, 12, 12, 8, 12, 12, 12, 14, 10, 12, 10, 10, 24, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c, name in enumerate(LEDGER_COLS, 1):
        cell = ws.cell(1, c, name)
        cell.font = LABEL_FONT
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center")

    r = 1
    for row in rows:
        r += 1
        if row["结论"] == J_OK:
            fill = None
        elif row["结论"] == J_OPEN:
            fill = GRAY_FILL
        elif row["结论"] == J_NOVISIT:
            fill = WARN_FILL
        else:
            fill = BAD_FILL
        for c, name in enumerate(LEDGER_COLS, 1):
            cell = ws.cell(r, c, row[name])
            cell.font = BODY_FONT
            if fill is not None:
                cell.fill = fill

    ws2 = wb.create_sheet("游离记录")
    for c, name in enumerate(STRAY_COLS, 1):
        cell = ws2.cell(1, c, name)
        cell.font = LABEL_FONT
        cell.fill = HEAD_FILL
    for i, r0 in enumerate(strays, 2):
        ws2.cell(i, 1, r0["受试者编号"]).font = BODY_FONT
        ws2.cell(i, 2, r0["访视点"]).font = BODY_FONT
        ws2.cell(i, 3, r0["采样日期"]).font = BODY_FONT
        ws2.cell(i, 4, r0["检项"]).font = BODY_FONT
        ws2.cell(i, 5, "计划表里没有这个 受试者编号 + 访视点 的组合，多半是编号抄错了（前导零）").font = BODY_FONT
    wb.save(path)


def main():
    plan_rows = read_plan()
    panel = read_panel()
    results = read_results()
    plan_keys = {(p["受试者编号"], p["访视点"]) for p in plan_rows}
    groups, strays = group_visits(results, plan_keys)

    # 顺序按中心、受试者、访视序号排；按访视点名字排 EOS 会跑到 SC 前面。
    order = {v: panel[v][0] for v in panel}
    plan_rows.sort(key=lambda p: (p["中心编号"], p["受试者编号"], order.get(p["访视点"], 99)))

    # 重跑必须先清目录：上一轮多出来的报告留着会被当成这一轮的，跟着一起发出去。
    if REPORT_DIR.exists():
        shutil.rmtree(REPORT_DIR)
    REPORT_DIR.mkdir(parents=True)
    LEDGER_FILE.parent.mkdir(parents=True, exist_ok=True)

    ledger, made, skipped = [], 0, []
    for plan in plan_rows:
        group = groups.get((plan["受试者编号"], plan["访视点"]))
        done = list(group["项"].keys()) if group else []
        judged = judge_visit(plan, group, panel, done)
        ledger.append({**{k: plan[k] for k in PLAN_COLS}, **judged})
        if not judged["实际采样日期"]:
            skipped.append(f"{plan['受试者编号']}/{plan['访视点']}")
            continue
        rows = report_items(plan, group, panel)
        name = f"{safe_name(plan['中心编号'])}_{safe_name(plan['受试者编号'])}_{safe_name(plan['访视点'])}.xlsx"
        write_report(REPORT_DIR / name, plan, judged, rows)
        made += 1

    write_ledger(ledger, strays, LEDGER_FILE)
    print(f"报告 {made} 份 -> 02_output/访视时点报告/")
    print(f"台账 {len(ledger)} 行 -> 02_output/访视执行台账.xlsx")
    print(f"没出报告的（没到访 / 没到期）：{'、'.join(skipped) if skipped else '无'}")
    print(f"游离记录 {len(strays)} 条 -> 台账「游离记录」页签")
    return _verify(made, ledger, strays)


def read_report(path: Path):
    """回读一份报告：抬头拿前 6 行，明细从第 7 行起。"""
    ws = load_workbook(path).active
    head = {}
    for r in range(2, 6):
        head[ws.cell(r, 1).value] = ws.cell(r, 2).value
        head[ws.cell(r, 4).value] = ws.cell(r, 5).value
    rows = []
    for r in range(7, ws.max_row + 1):
        if ws.cell(r, 1).value == "结论":
            head["结论"] = ws.cell(r, 2).value
            break
        rows.append([ws.cell(r, c).value for c in range(1, 6)])
    return head, rows


def read_ledger():
    wb = load_workbook(LEDGER_FILE)
    ws, ws2 = wb["访视执行台账"], wb["游离记录"]
    cols = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        rows.append(dict(zip(cols, r)))
    strays = [r for r in ws2.iter_rows(min_row=2, values_only=True) if r[0]]
    return rows, strays


def _verify(made, ledger, strays):
    """回读产物，把该标出来的逐条对一遍。"""
    rows, strays_back = read_ledger()
    by_key = {(r["受试者编号"], r["访视点"]): r for r in rows}
    r1 = read_report(REPORT_DIR / "S01_00102_SC.xlsx")[0]
    r2 = read_report(REPORT_DIR / "S01_00103_V1.xlsx")
    r3 = read_report(REPORT_DIR / "S02_00201_V1.xlsx")
    r4 = read_report(REPORT_DIR / "S02_00202_V2.xlsx")
    r5 = read_report(REPORT_DIR / "S02_00201_SC.xlsx")
    pk = [x for x in r3[1] if x[0] == "PK 采血"][0]

    checks = [
        ("报告份数 = 到过访的时点数 16", made == 16, f"{made}"),
        ("台账 20 行（含没到访与没到期的）", len(rows) == 20, f"{len(rows)}"),
        ("00102 SC 提前到访也判超窗，偏离 -2", by_key[("00102", "SC")]["偏离天数"] == -2,
         f"{by_key[('00102', 'SC')]['偏离天数']}"),
        ("00102 V2 晚到 3 天", by_key[("00102", "V2")]["偏离天数"] == 3,
         f"{by_key[('00102', 'V2')]['偏离天数']}"),
        ("00103 V1 结论含「缺项」，缺失两项都列出来",
         by_key[("00103", "V1")]["结论"] == J_MISS and by_key[("00103", "V1")]["缺失检项"] == "尿常规、凝血四项",
         f"{by_key[('00103', 'V1')]['结论']} / {by_key[('00103', 'V1')]['缺失检项']}"),
        ("00103 V2 窗口已过又没记录 -> 未到访", by_key[("00103", "V2")]["结论"] == J_NOVISIT,
         f"{by_key[('00103', 'V2')]['结论']}"),
        ("00101 V3 窗口没到 -> 未到期，不算未到访", by_key[("00101", "V3")]["结论"] == J_OPEN,
         f"{by_key[('00101', 'V3')]['结论']}"),
        ("没到访的那行不列缺项，写「—」", by_key[("00103", "V2")]["缺失检项"] == I_UNKNOWN,
         f"「{by_key[('00103', 'V2')]['缺失检项']}」"),
        ("台账里「未到访」与「未到期」是两档，不混成一个词",
         by_key[("00103", "V2")]["到访判定"] == J_NOVISIT and by_key[("00101", "V3")]["到访判定"] == J_OPEN,
         f"{by_key[('00103', 'V2')]['到访判定']} / {by_key[('00101', 'V3')]['到访判定']}"),
        ("报告的到访判定：提前到访那份也写「超窗」", r1["到访判定"] == J_WINDOW, f"{r1['到访判定']}"),
        ("00103 V1 报告里没做的两项各占一行，判定「未做」",
         [x for x in r2[1] if x[0] == "凝血四项"][0][4] == "未做" and len(r2[1]) == 5,
         f"{[x[0] for x in r2[1]]}"),
        ("00201 V1 同项复测 -> 结论含「同项多重结果」", J_DUP in by_key[("00201", "V1")]["结论"],
         f"{by_key[('00201', 'V1')]['结论']}"),
        ("复测取先到那条（18.4），没被后到的 21.7 顶掉", pk[1] == "18.4", f"{pk[1]}"),
        ("00202 V2 生化全套 52 标「超出范围」",
         [x for x in r4[1] if x[0] == "生化全套"][0][4] == "超出范围",
         f"{[x for x in r4[1] if x[0] == '生化全套'][0][4]}"),
        ("方案外多做的「心肌酶」印在最后，判定「方案外项目」",
         r4[1][-1][0] == "心肌酶" and r4[1][-1][4] == "方案外项目", f"{r4[1][-1]}"),
        ("00201 SC 空结果 -> 「未出结果」，不是正常",
         [x for x in r5[1] if x[0] == "尿常规"][0][4] == "未出结果",
         f"{[x for x in r5[1] if x[0] == '尿常规'][0][4]}"),
        ("PK 采血只有上限，参考范围印成「≤ 50」",
         [x for x in r4[1] if x[0] == "PK 采血"][0][3] == "≤ 50",
         f"{[x for x in r4[1] if x[0] == 'PK 采血'][0][3]}"),
        ("前导零编号在文件名里保留 5 位", (REPORT_DIR / "S02_00201_V1.xlsx").exists(), "S02_00201_V1.xlsx"),
        ("抄丢前导零的 2202 单列成游离，没并进台账",
         len(strays_back) == 1 and strays_back[0][0] == "2202" and len(rows) == 20,
         f"{strays_back}"),
        ("仪器导出的表尾合计行与说明行都没被当成检项",
         all(x[0] not in ("合计", "说明：结果为空表示本次未出结果，不是正常") for x in r5[1]),
         f"{[x[0] for x in r5[1]]}"),
    ]
    ok = True
    for name, passed, val in checks:
        print(f"  [{'OK ' if passed else 'FAIL'}] {name}  ->  {val}")
        ok &= bool(passed)

    print("\n总判定:", "PASS" if ok else "FAIL")
    return bool(ok)


class _Tee:
    """同时写控制台和日志文件，万一控制台中文乱码还有日志可看。"""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
            except Exception:
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except Exception:
                pass


def _note_unfinished():
    """上一条执行若被手动关窗口打断，日志里会缺完成行。开跑前补一句说明。"""
    LOG_FILE = LOG_DIR / "run_log.txt"
    if not LOG_FILE.exists():
        return
    tail = ""
    with open(LOG_FILE, encoding="utf-8", errors="replace") as f:
        for ln in f:
            if ln.strip():
                tail = ln.strip()
    if tail and not tail.startswith("执行完成"):
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write("（上一条没有留下完成时间，多半是窗口被手动关闭或中途强制中断）\n")


def _host_info() -> str:
    """主机名 + 本机 IP，用来分辨这段日志是哪台机器跑出来的。"""
    try:
        host = socket.gethostname()
    except Exception:
        host = "unknown"
    try:
        # UDP 不会真的发包，只按路由表取本机对外网卡的地址
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
    except Exception:
        try:
            ip = socket.gethostbyname(host)
        except Exception:
            ip = "未知"
    return f"{host}（{ip}）"


def _run():
    """一次执行 = 追加一段：开始时间 + 过程 + 完成时间。历史日志不覆盖。"""
    LINE = "=" * 60
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    _note_unfinished()
    _log = open(LOG_DIR / "run_log.txt", "a", encoding="utf-8")   # 追加，别用 w
    _log.write(f"\n{LINE}\n")
    _log.write(f"运行主机：{_host_info()}\n")
    _log.write(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}\n")
    _log.write(f"Python   ：{sys.version.split()[0]}  （{Path(sys.executable).name}）\n")
    _log.flush()

    _orig = sys.stdout
    sys.stdout = _Tee(sys.stdout, _log)
    t0, ok, tail = time.time(), False, ""
    try:
        try:
            ok = main()
        except KeyboardInterrupt:
            tail = "\n被手动中断（Ctrl+C）。"
        except Exception as e:
            import traceback
            traceback.print_exc()
            tail = f"\n跑失败了：{type(e).__name__}: {e}"
        else:
            tail = "\n跑完了。结果在 02_output/，运行记录已写入 source/run_log.txt。"
        if tail:
            print(tail)
        # 完成行必须写在暂停之前：双击后关窗口就丢不了
        print(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}"
              f"  用时 {time.time() - t0:.1f}s  判定 {'PASS' if ok else 'FAIL'}")
    finally:
        sys.stdout.flush()
        _log.close()
        sys.stdout = _orig

    _pause()          # 等回车放在日志收尾之后
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    _run()
