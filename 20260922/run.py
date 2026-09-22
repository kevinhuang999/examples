#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""中心实验室 TAT 超时提醒：按样本把已耗时长算准，挑出要超时的，生成催办文本。

一句话：TAT 提醒的难点不在减时间，而在三件事——起点按采集还是按接收、
暂停期间要不要停表、工作日口径跨假日怎么数。这三件事算错，提醒就会吵得没人看。

用法：
    python run.py              # 跑完停住等回车（Windows 双击也行）
    python run.py --no-pause   # 跑完直接退出
"""

import csv
import socket
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime, time as dtime, timedelta
from pathlib import Path

# ----------------------------------------------------------------- ① 配置区（跑之前主要改这里）

ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "01_raw_data"
OUT_DIR = ROOT / "02_output"
LOG_FILE = ROOT / "source" / "run_log.txt"

LEDGER_CSV = RAW_DIR / "样本接收台账.csv"
PAUSE_CSV = RAW_DIR / "计时暂停记录.csv"
CALENDAR_CSV = RAW_DIR / "工作日历.csv"

REPORT_XLSX = OUT_DIR / "TAT监控清单.xlsx"
ALERT_TXT = OUT_DIR / "催办提醒.txt"

# 数据截止时刻 = 台账导出的那一刻。它必须是常量：拿 datetime.now() 去算，
# 今天跑出来的超时清单和明天的对不上，没法复盘，也没法拿去跟中心对账。
DATA_CUTOFF = datetime(2026, 9, 22, 10, 0)

# 快没时间了就提前喊。短时限项目按比例喊不住（4 小时的加急，20% 才 48 分钟），
# 所以取「4 小时」与「时限的 20%」里大的那个。
WARN_MIN_HOURS = 4.0
WARN_RATIO = 0.2

# 「工作日」口径只数工作日里的上班时段。这是本示例唯一需要跟客户确认的假设，见 README。
WORK_START = dtime(9, 0)
WORK_END = dtime(17, 0)

SCOPE_RECEIVE = "接收"
CAL_NATURAL = "自然日"
CAL_WORKDAY = "工作日"

R_OVERDUE = "已超时"
R_WARN = "即将超时"
R_OK = "正常"
R_DELIVERED = "按时交付"
R_LATE = "超时交付"
R_HOLD = "待补数据"

FILL = {R_OVERDUE: "FFC7CE", R_WARN: "FFEB9C", R_LATE: "FFD8B2", R_HOLD: "E6E6E6"}
ORDER = {R_OVERDUE: 0, R_WARN: 1, R_HOLD: 2, R_OK: 3, R_LATE: 4, R_DELIVERED: 5}

ALERT_HEADER = ["样本编号", "中心编号", "检项类别", "起算口径", "计时方式",
                "起点时间", "已耗(小时)", "时限(小时)", "剩余(小时)", "结论"]
ALL_HEADER = ALERT_HEADER + ["报告发出时间", "备注"]
HOLD_HEADER = ["样本编号", "中心编号", "检项类别", "起算口径", "起点时间",
               "时限(小时)", "问题", "说明"]
ALERT_WIDTH = [15, 11, 11, 10, 10, 18, 11, 11, 11, 11]
ALL_WIDTH = ALERT_WIDTH + [18, 34]
HOLD_WIDTH = [15, 11, 11, 10, 18, 11, 13, 44]

# ----------------------------------------------------------------- 装依赖 / 日志


def ensure_deps():
    """缺 openpyxl 就自己装（走清华源）。"""
    try:
        import openpyxl  # noqa: F401
        return
    except ImportError:
        pass
    for args in ([sys.executable, "-m", "pip", "install", "openpyxl",
                  "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"],
                 [sys.executable, "-m", "pip", "install", "--user", "openpyxl",
                  "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"]):
        try:
            subprocess.check_call(args)
            return
        except subprocess.CalledProcessError:
            continue
    raise SystemExit("openpyxl 装不上，请手动执行：pip install openpyxl")


def local_ip():
    """取本机对外网卡地址（不是 127.0.0.1），拿不到就退回主机名解析。"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
        finally:
            s.close()
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "未知"


def log_line(text):
    """日志追加不覆盖：每段开头都先记运行主机，方便分辨是哪台机器跑的。"""
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")
    print(text)


# ----------------------------------------------------------------- ② 读数据（归堆）


def parse_dt(text):
    """把台账里的时间串转成 datetime，认不出来返回 None。

    导出的时间列格式不统一：有的带秒、有的不带，有的还夹了全角空格，
    所以逐个格式试一遍。**千万别在这里兜一个"今天零点"**——
    兜了以后，一条本该挂起待补数据的样本会装成正常，提醒就白做了。
    """
    text = (text or "").replace("\u3000", " ").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def read_ledger():
    """读样本接收台账，一行一个样本。

    用 csv 直接读、不经过 pandas：样本编号是补零到 7 位的编号，
    走 pandas 会默认猜成数字，0000123 变 123，事后 astype 也补不回来。
    表尾的合计行与说明行第一列是空的，靠这一点整行跳过——
    也正因为这样，合计数字千万别写进第一列，写进去就会被当成一条样本。
    """
    rows = []
    with open(LEDGER_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("样本编号") or "").strip():
                continue
            rows.append({k: (v or "").strip() for k, v in row.items()})
    return rows


def read_pauses():
    """读计时暂停记录：一行一段暂停（等复检、等客户确认这种停表时段）。"""
    rows = []
    with open(PAUSE_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("样本编号") or "").strip():
                continue
            rows.append({k: (v or "").strip() for k, v in row.items()})
    return rows


def group_pauses(rows):
    """按样本编号把暂停段归堆 —— 台账一行一个样本，暂停可能一段都没有、也可能好几段，
    两边粒度不同，先把暂停挂到编号上，判的时候逐条查。"""
    bucket = {}
    for r in rows:
        bucket.setdefault(r["样本编号"], []).append(r)
    return bucket


def read_calendar():
    """读工作日历，返回"是工作日"的日期集合。

    日历要单独一张表，不能在代码里写 weekday() < 5 —— 法定假日与调休都不按整周排
    （本示例里 09-17 是假日、09-19 周六却是调休上班），写死周末的人一遇调休就错。
    """
    days = set()
    with open(CALENDAR_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            day = parse_dt((row.get("日期") or "") + " 00:00")
            if day is None:                     # 表尾说明行：日期列是空的
                continue
            if (row.get("是否工作日") or "").strip() in ("是", "Y", "1", "True"):
                days.add(day.date())
    return days


# ----------------------------------------------------------------- ③ 算时长（最容易算错的一段）


def work_span_hours(start, end, workdays):
    """算 [start, end) 里真正落在上班时段的小时数（工作日 09:00–17:00，一天 8 小时）。

    逐格累加，不按天粗扣：调休和假日不按整周排，
    "总时长 - 周末天数 × 24" 遇到跨假期的样本一定错；"按 8 小时乘天数" 遇到半天也错。
    格子取 30 分钟——再细没必要，TAT 的口径本来就在小时级。
    """
    if end <= start:
        return 0.0
    step = timedelta(minutes=30)
    total = 0.0
    cur = start
    while cur < end:
        nxt = min(cur + step, end)
        if cur.date() in workdays and WORK_START <= cur.time() < WORK_END:
            total += (nxt - cur).total_seconds() / 3600
        cur = nxt
    return total


def elapsed_hours(start, end, mode, workdays):
    """按计时口径算 start → end 的已耗时长。

    同一批样本里两种口径是混着的：自然日＝连续时间（夜里、周末照样走）；
    工作日＝只数上班时段。口径这一列不能省，省了就只能挑一列时间去当起点。
    """
    if end <= start:
        return 0.0
    if mode == CAL_WORKDAY:
        return work_span_hours(start, end, workdays)
    return (end - start).total_seconds() / 3600


def overlap_hours(a_start, a_end, b_start, b_end, mode, workdays):
    """两段时间重叠的部分有多长（用来算暂停吃掉了多少小时）。"""
    s, e = max(a_start, b_start), min(a_end, b_end)
    return elapsed_hours(s, e, mode, workdays) if e > s else 0.0


def pause_span(pauses, start, end, mode, workdays, cutoff):
    """把一条样本的暂停段折成"停表时长"，顺手把可疑的暂停记下来。

    两种脏数据都要接住，而且都不能猜：
    - **结束时间为空** = 这段暂停还没结束（还在等复检/等客户回话）→ 截断到数据截止时刻。
      绝不能当成 0 点、更不能当成"没暂停过"，否则一条正在停表的样本会被误报超时；
    - **结束早于开始** = 录反了 → 这一段整段丢弃，宁可少扣，也不要把时长扣成负的。
    """
    stopped, notes = 0.0, []
    for p in pauses:
        p_start = parse_dt(p.get("暂停开始时间"))
        p_end = parse_dt(p.get("暂停结束时间")) or cutoff
        if p_start is None:
            notes.append("暂停记录缺开始时间，已忽略")
            continue
        if p_end < p_start:
            notes.append(f'暂停记录时间倒挂（{p["暂停开始时间"]} → {p["暂停结束时间"]}），已忽略')
            continue
        stopped += overlap_hours(start, end, p_start, p_end, mode, workdays)
    return stopped, notes


def pick_start(rec):
    """取计时起点：按台账里的「起算口径」选列。

    有的研究方案按采集时间起算（冷链路上耗的时间也算），有的按中心接收完成时间算。
    硬拿一列当起点，另一半样本的时长全是错的。口径说按采集算、采集时间却空着，
    就挂起等补，**不许拿接收时间顶上**——顶上就是自己编了一个更宽松的起点。
    """
    if rec["起算口径"] == SCOPE_RECEIVE:
        return parse_dt(rec["接收完成时间"]), "接收完成时间"
    return parse_dt(rec["采集时间"]), "采集时间"


def _hold(out, reason, text):
    """把一条样本钉成"待补数据"：结论、问题类型、说明三处一起写，别只写一处。"""
    out["结论"], out["问题"], out["说明"] = R_HOLD, reason, text
    return out


def judge(rec, pauses, workdays, cutoff, dup):
    """判一条样本：先看数据全不全，再算时长，最后才分级。

    顺序不能换 —— 先分级再补数据，会把"缺时限"的样本算成 0 小时、装成正常。
    """
    out = {"样本编号": rec["样本编号"], "中心编号": rec["中心编号"],
           "检项类别": rec["检项类别"], "起算口径": rec["起算口径"],
           "计时方式": rec["计时方式"], "起点时间": "", "报告发出时间": "",
           "时限": None, "已耗": 0.0, "剩余": 0.0,
           "结论": R_HOLD, "问题": "", "说明": "", "备注": []}
    if dup > 1:
        out["备注"].append(f"台账里同一编号出现 {dup} 条")

    start, where = pick_start(rec)
    if start is None:
        return _hold(out, "起点缺失", f'口径是「{rec["起算口径"]}」，{where}空着，先补时间再盯')

    limit_text = rec["TAT时限(小时)"]
    if not limit_text:
        return _hold(out, "缺时限", "台账里 TAT 时限是空的，先跟方案确认时限再来盯")
    try:
        limit = float(limit_text)
    except ValueError:
        return _hold(out, "时限非数字", f'TAT 时限填的是「{limit_text}」，没法算')
    if limit <= 0:
        return _hold(out, "时限非正数", f"TAT 时限是 {limit:g} 小时，先确认为什么不是正数")

    sampled = parse_dt(rec["采集时间"])
    receive = parse_dt(rec["接收完成时间"])
    report = parse_dt(rec["报告发出时间"])
    if report and report < start:
        return _hold(out, "时间录反", "报告发出时间早于计时起点，先核对这两个时间")
    if sampled and receive and sampled > receive:
        out["备注"].append("采集时间晚于接收完成时间，对一下")

    end = report or cutoff
    used = elapsed_hours(start, end, rec["计时方式"], workdays)
    stopped, notes = pause_span(pauses, start, end, rec["计时方式"], workdays, cutoff)
    out["备注"].extend(notes)
    used = max(used - stopped, 0.0)

    out["起点时间"] = start.strftime("%Y-%m-%d %H:%M")
    out["报告发出时间"] = report.strftime("%Y-%m-%d %H:%M") if report else ""
    out["时限"], out["已耗"] = limit, used
    out["剩余"] = limit - used

    if report:                                  # 报告已经发出：只复盘当时超没超，不再催
        out["结论"] = R_LATE if used > limit else R_DELIVERED
    elif used > limit:                          # 已耗正好等于时限不算超时，算即将超时（剩余 0）
        out["结论"] = R_OVERDUE
    elif out["剩余"] <= max(WARN_MIN_HOURS, limit * WARN_RATIO):
        out["结论"] = R_WARN
    else:
        out["结论"] = R_OK
    return out


# ----------------------------------------------------------------- ④ 输出（写表 + 催办文本）


def as_row(it):
    """把一条判完的记录铺成输出表的一行。待补数据的样本不写时长数字——
    写 0 会被人当成"还没耗时间"，比空着更误导。"""
    hold = it["结论"] == R_HOLD
    return [it["样本编号"], it["中心编号"], it["检项类别"], it["起算口径"], it["计时方式"],
            it["起点时间"],
            "" if hold else round(it["已耗"], 1),
            "" if it["时限"] is None else it["时限"],
            "" if hold else round(it["剩余"], 1),
            it["结论"]]


def write_sheet(ws, header, rows, widths, fixed=None):
    """写一张工作表：表头加粗居中、冻结首行、整行刷底色。

    取色按「结论」列（前两张表的第 10 列）查色表；挂起那页没有结论列，
    整页一个颜色，直接传 fixed 进来。判颜色用 fill.patternType == "solid"，
    别去比 fgColor.rgb 的字符串——不同 openpyxl 版本给的 rgb 前缀（00 / FF）
    不一样，比字符串会时对时错。
    """
    from openpyxl.styles import Alignment, Font, PatternFill

    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append(row)
        color = FILL.get(fixed) if fixed else FILL.get(row[9])
        if color:
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", start_color=color)
    ws.freeze_panes = "A2"
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width


def write_report(items):
    """写监控清单三页：要人动手的、全量的、数据不全要人补的。"""
    from openpyxl import Workbook

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    alert = sorted([i for i in items if i["结论"] in (R_OVERDUE, R_WARN)],
                   key=lambda i: i["剩余"])
    holds = [i for i in items if i["结论"] == R_HOLD]
    everything = sorted(items, key=lambda i: (ORDER[i["结论"]], i["剩余"]))

    ws = wb.active
    ws.title = "提醒名单"
    write_sheet(ws, ALERT_HEADER, [as_row(i) for i in alert], ALERT_WIDTH)
    write_sheet(wb.create_sheet("全量监控"), ALL_HEADER,
                [as_row(i) + [i["报告发出时间"], i["备注文本"]] for i in everything], ALL_WIDTH)
    write_sheet(wb.create_sheet("挂起与异常"), HOLD_HEADER,
                [[i["样本编号"], i["中心编号"], i["检项类别"], i["起算口径"], i["起点时间"],
                  "" if i["时限"] is None else i["时限"], i["问题"], i["说明"]] for i in holds],
                HOLD_WIDTH, fixed=R_HOLD)
    wb.save(REPORT_XLSX)
    return len(alert), len(holds)


def write_alert_text(items):
    """写一段能直接粘进群里的催办文本。

    只列要人动手的（已超时 + 即将超时），按紧急度排；已出报告的不进——
    提醒里混进"已经办完的"，看两次就没人看了。数据不全的单独列一行，
    它不是催检测员，是催补数据的人。
    """
    alert = sorted([i for i in items if i["结论"] in (R_OVERDUE, R_WARN)],
                   key=lambda i: i["剩余"])
    overdue = [i for i in alert if i["结论"] == R_OVERDUE]
    warns = [i for i in alert if i["结论"] == R_WARN]
    holds = [i for i in items if i["结论"] == R_HOLD]

    lines = [f"【TAT 催办】数据截止 {DATA_CUTOFF:%Y-%m-%d %H:%M}，"
             f"共盯 {len(items)} 条，需要处理 {len(alert)} 条"]
    for title, group, tail in (("■ 已超时", overdue, "超"), ("■ 即将超时", warns, "剩")):
        if not group:
            continue
        lines.append("")
        lines.append(f"{title} {len(group)} 条")
        for it in group:
            sid, site, cat = it["样本编号"], it["中心编号"], it["检项类别"]
            hours = -it["剩余"] if it["结论"] == R_OVERDUE else it["剩余"]
            lines.append(f"  {sid}  {site}  {cat}  已耗 {it['已耗']:.1f}h / "
                         f"时限 {it['时限']:g}h  {tail} {hours:.1f}h")
    if holds:
        lines.append("")
        lines.append(f"■ 待补数据 {len(holds)} 条（时限或起点没填对，先补再盯）："
                     + "、".join(i["样本编号"] for i in holds))
    lines.append("")
    lines.append("请对应中心在这条下面回复处理进展；超时单请一并说明原因。")
    ALERT_TXT.parent.mkdir(parents=True, exist_ok=True)
    ALERT_TXT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return len(alert)


# ----------------------------------------------------------------- 回读校验


def verify(items):
    """重新从磁盘把 xlsx 与催办文本读回来，逐项断言——不自证，只看落盘的结果。"""
    from openpyxl import load_workbook

    checks = []

    def ok(name, good, detail=""):
        checks.append((name, bool(good), detail))

    by_id = {}
    for it in items:
        by_id.setdefault(it["样本编号"], it)
    overdue = [i for i in items if i["结论"] == R_OVERDUE]
    warns = [i for i in items if i["结论"] == R_WARN]
    holds = [i for i in items if i["结论"] == R_HOLD]
    done = [i for i in items if i["结论"] in (R_DELIVERED, R_LATE)]

    wb = load_workbook(REPORT_XLSX)
    alert, allws, holdws = wb["提醒名单"], wb["全量监控"], wb["挂起与异常"]

    ok("台账 40 条全都判了", len(items) == 40, f"{len(items)} 条")
    ok("提醒名单条数 = 已超时 + 即将超时",
       alert.max_row - 1 == len(overdue) + len(warns),
       f"{alert.max_row - 1} vs {len(overdue) + len(warns)}")
    ok("已超时 10 条 / 即将超时 4 条", (len(overdue), len(warns)) == (10, 4),
       f"{len(overdue)} / {len(warns)}")
    remains = [alert.cell(row=r, column=9).value for r in range(2, alert.max_row + 1)]
    ok("提醒名单按剩余小时升序", remains == sorted(remains), f"{remains}")
    ok("全量监控覆盖全部样本", allws.max_row - 1 == len(items),
       f"{allws.max_row - 1} vs {len(items)}")
    ok("挂起页条数 = 待补数据条数", holdws.max_row - 1 == len(holds),
       f"{holdws.max_row - 1} vs {len(holds)}")

    alert_ids = {alert.cell(row=r, column=1).value for r in range(2, alert.max_row + 1)}
    ok("已出报告的样本一条都不进提醒", not (alert_ids & {i["样本编号"] for i in done}),
       f'{sorted(alert_ids & {i["样本编号"] for i in done})}')
    ok("待补数据的样本一条都不进提醒", not (alert_ids & {i["样本编号"] for i in holds}),
       f'{sorted(alert_ids & {i["样本编号"] for i in holds})}')
    ok("样本编号保持补零原样",
       "0000123" in {i["样本编号"] for i in items} and "123" not in {i["样本编号"] for i in items})
    ok("已耗小时没有负数", all(i["已耗"] >= 0 for i in items))

    expect = {"0000123": (R_OVERDUE, -1.0), "0000142": (R_WARN, 0.0),
              "0000162": (R_OK, 6.5), "0000163": (R_OVERDUE, -0.5),
              "0000144": (R_LATE, None), "0000170": (R_HOLD, None),
              "0000133": (R_OK, 15.0), "0000130": (R_OVERDUE, -0.5)}
    for sid, (verdict, remain) in expect.items():
        got = by_id.get(sid)
        good = got is not None and got["结论"] == verdict
        detail = f'找不到 {sid}' if got is None else f'{got["结论"]}'
        if good and remain is not None:
            good = abs(got["剩余"] - remain) < 0.05
            detail = f'{got["剩余"]:.2f} vs {remain}'
        ok(f"{sid} 结论 {verdict}", good, detail)
    ok("0000161 扣掉未结束的暂停后已耗 13.0h",
       abs(by_id["0000161"]["已耗"] - 13.0) < 0.05, f'{by_id["0000161"]["已耗"]:.2f}')
    ok("0000163 的倒挂暂停被记进备注", "倒挂" in by_id["0000163"]["备注文本"])
    dup = [i for i in items if i["样本编号"] == "0000180"]
    ok("台账重复编号两条都保留且都标了",
       len(dup) == 2 and all("2 条" in i["备注文本"] for i in dup), f"{len(dup)}")

    painted = True
    for r in range(2, alert.max_row + 1):
        want = FILL.get(alert.cell(row=r, column=10).value)
        fill = alert.cell(row=r, column=1).fill
        if want is None or fill.patternType != "solid" \
                or not str(fill.start_color.rgb).endswith(want):
            painted = False
    ok("提醒名单整行按结论刷色", painted)

    text = ALERT_TXT.read_text(encoding="utf-8")
    missing = sorted(sid for sid in alert_ids if sid not in text)
    stray = sorted(i["样本编号"] for i in items
                   if i["结论"] in (R_OK, R_DELIVERED, R_LATE) and i["样本编号"] in text)
    ok("催办文本覆盖全部要处理的样本", not missing, "、".join(missing))
    ok("催办文本里没有正常与已交付的样本", not stray, "、".join(stray))
    ok("催办文本条目数 = 要处理的条数",
       text.count("已耗 ") == len(alert_ids), f'{text.count("已耗 ")} vs {len(alert_ids)}')

    return checks


# ----------------------------------------------------------------- main


def main():
    t0 = time.time()
    log_line(f"运行主机：{socket.gethostname()}（{local_ip()}）")
    log_line(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
             f"Python {sys.version.split()[0]} / {Path(sys.executable).name}")

    ledger = read_ledger()
    pauses = group_pauses(read_pauses())
    workdays = read_calendar()
    dup = Counter(r["样本编号"] for r in ledger)
    print(f"台账 {len(ledger)} 条，暂停记录 {sum(len(v) for v in pauses.values())} 段，"
          f"工作日历 {len(workdays)} 天，数据截止 {DATA_CUTOFF:%Y-%m-%d %H:%M}")

    items = []
    for rec in ledger:
        item = judge(rec, pauses.get(rec["样本编号"], []), workdays, DATA_CUTOFF, dup[rec["样本编号"]])
        item["备注文本"] = "；".join(item["备注"])
        items.append(item)

    made, held = write_report(items)
    write_alert_text(items)

    checks = verify(items)
    for name, good, detail in checks:
        print(f"  [{'OK ' if good else 'FAIL'}] {name}" + (f"   {detail}" if not good else ""))

    passed = all(c[1] for c in checks)
    print(f"\n要处理 {made} 条（已超时 "
          f"{sum(1 for i in items if i['结论'] == R_OVERDUE)} 条），"
          f"待补数据 {held} 条，回读校验 {sum(c[1] for c in checks)}/{len(checks)} 项通过")
    print(f"结果目录：{OUT_DIR}")

    log_line(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  用时 {time.time() - t0:.1f}s  "
             f"判定 {'PASS' if passed else 'FAIL'}")

    if "--no-pause" not in sys.argv:
        try:
            input("\n按回车键关闭窗口…")
        except EOFError:
            pass


if __name__ == "__main__":
    ensure_deps()
    main()


