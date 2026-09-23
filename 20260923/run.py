#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第三方检测 · 委托单批量生成：把受理登记表里的行归成委托单，一单一号出正式委托单。

一句话：批量出单的难点不在填模板，而在归堆与挂起——哪几行属于同一张委托单、
一张单该不该出、出不了的那单缺什么，这三件事定错，出的单比手工还难对。

用法：
    python run.py              # 跑完停住等回车（Windows 双击也行）
    python run.py --no-pause   # 跑完直接退出
"""

import csv
import re
import socket
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path

# ----------------------------------------------------------------- ① 配置区（跑之前主要改这里）

ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "01_raw_data"
OUT_DIR = ROOT / "02_output"
TPL_DIR = ROOT / "source" / "templates"
LOG_FILE = ROOT / "source" / "run_log.txt"

ENTRY_CSV = RAW_DIR / "委托受理登记.csv"
TPL_XLSX = TPL_DIR / "检测委托单模板.xlsx"
ORDER_DIR = OUT_DIR / "委托单"
LEDGER_XLSX = OUT_DIR / "委托受理台账.xlsx"

# 单号规则：WT + 受理日期 + 当天流水（同一天从 01 起，按登记表里出现的先后排）。
PREFIX = "WT"
MAX_DETAIL = 8                       # 单张委托单能写几行样品，跟模板明细区一致

# 受理日期在登记表里被写成过好几种：斜杠、横杠、点号。都要认。
# 认不出来的**不兜今天**——兜了就是把一张日期没定的委托单出成今天的单，对账时说不清。
DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d")

# 检测项目的分隔符：顿号、逗号（半角全角）、分号、斜杠，客户怎么写都有。
SPLIT_RE = re.compile(r"[、,，;；/]+")

RUSH = "加急"
NORMAL = "常规"

DATE_FMT = "yyyy-mm-dd"              # 写进模板的日期格子要顺手设格式，否则 Excel 显示成序列号
FILL = {RUSH: "FFC7CE"}

ORDER_HEADER = ["委托单号", "受理日期", "客户名称", "联系人", "送样方式", "报告交付方式",
                "加急", "样品数", "检测项目数", "委托单文件"]
ORDER_WIDTH = [18, 12, 26, 10, 12, 14, 8, 9, 11, 24]
DETAIL_HEADER = ["委托单号", "序号", "样品编号", "样品名称", "客户样品编号", "数量",
                 "检测项目数", "检测项目", "依据标准", "备注"]
DETAIL_WIDTH = [18, 6, 20, 20, 16, 8, 10, 30, 20, 26]
HOLD_HEADER = ["受理日期(原样)", "客户名称", "涉及样品", "问题", "说明"]
HOLD_WIDTH = [16, 26, 14, 14, 52]

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


def clean(text):
    """去空白。全角空格（U+3000）也要去——客户名称尾巴上粘一个全角空格，
    不去干净就会跟同名客户分成两张单，人眼还看不出来。"""
    return (text or "").replace("\u3000", " ").strip()


def norm_date(text):
    """把受理日期转成 date，认不出来返回 None（连同原始文字一起挂起）。"""
    text = clean(text)
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def read_entries():
    """读受理登记表，一行一个样品。

    用 csv 直接读、不经过 pandas：客户样品编号是补零的（0000713），
    走 pandas 会被猜成数字 713，事后 astype 也补不回前导零。
    表尾的说明行第一列是空的，靠这一点整行跳过——所以说明文字千万别写进第一列。
    """
    rows = []
    with open(ENTRY_CSV, encoding="utf-8-sig", newline="") as f:
        for raw in csv.DictReader(f):
            if not clean(raw.get("受理日期")):
                continue
            rows.append({k: clean(v) for k, v in raw.items()})
    return rows


def group_orders(rows):
    """把行归堆成委托单。

    分组键是「受理日期 + 客户名称」，两个都要在里面，缺一个就出错：
    - 少了客户名称 → 同一天不同客户的样品并到一张单上；
    - 少了受理日期 → 同一客户隔天再来一次会被并成一张，样品数和项目数全错。
    日期和客户名都先归一化再当键：同一天写成 2026/9/18 与 2026.09.18 是一个人填的，
    带全角空格的客户名和干净的客户名也是同一个客户。
    """
    orders, index = [], {}
    for r in rows:
        day = norm_date(r["受理日期"])
        name = r["客户名称"]
        key = (day, name)
        if key not in index:
            index[key] = {"受理日期": day, "受理日期原文": r["受理日期"], "客户名称": name,
                          "联系人": r["客户联系人"], "送样方式": r["送样方式"],
                          "报告交付方式": r["报告交付方式"], "行": [], "备注": []}
            orders.append(index[key])
        order = index[key]
        order["行"].append(r)
        if r["加急"] and RUSH not in order["备注"]:
            order["备注"].append(f'第 {len(order["行"])} 行标了「{r["加急"]}」，按整单处理')
    return orders


# ----------------------------------------------------------------- ③ 校验 / 编号 / 填数据


def split_items(text):
    """把检测项目按分隔符拆开，去空、保序。

    这一列是客户原样写的，同一份单上「甲醛含量、pH值」和「耐摩擦色牢度/耐水色牢度」
    两种写法都会出现，所以顿号逗号分号斜杠一起当分隔符。
    """
    return [x.strip() for x in SPLIT_RE.split(text) if x.strip()]


def as_qty(text):
    """把数量转成正整数，转不了返回 None。

    只认纯数字。有人把数量写成「2 支」「约 3 件」——**不要用正则把数字抠出来**，
    抠出来脚本是能跑，但客户原件上的口径被你改掉了，对单时两边说的不是一回事。
    这种整单挂起，让前台去跟客户确认单位。
    """
    text = clean(text)
    if not text.isdigit():
        return None
    qty = int(text)
    return qty if qty > 0 else None


def check_order(order):
    """判一张单能不能出，不能出就写清缺什么。

    先判整单级（日期、客户名、样品行数），再逐行判——整单级的问题一出现，
    逐行报出来的缺项都是噪音，前台得先拿到"这张单压根出不了、以及为什么"。
    """
    if order["受理日期"] is None:
        return "日期认不出", f'受理日期填的是「{order["受理日期原文"]}」，先跟业务确认受理日期'
    if not order["客户名称"]:
        return "客户名称缺失", "客户名称是空的，出不了单，先补客户单位全称"
    if len(order["行"]) > MAX_DETAIL:
        return "样品超上限", (f'这一单有 {len(order["行"])} 个样品，超过单张委托单 '
                          f"{MAX_DETAIL} 行明细的上限，先拆成多张单再生成")
    for i, r in enumerate(order["行"], 1):
        if not r["样品名称"]:
            return "样品名称缺失", f"第 {i} 行样品名称是空的"
        if not r["检测项目"]:
            return "检测项目缺失", f'第 {i} 行（{r["样品名称"]}）没填检测项目'
        if not r["依据标准"]:
            return "依据标准缺失", f'第 {i} 行（{r["样品名称"]}）的依据标准是空的，先跟客户确认标准号'
        if as_qty(r["样品数量"]) is None:
            return "数量不可用", (f'第 {i} 行（{r["样品名称"]}）数量填的是'
                              f'「{r["样品数量"]}」，请只填正整数，单位写到备注里')
    return None


def number_orders(orders):
    """给能出的单编号：WT + 受理日期 + 当天流水（01 起），样品再挂两级序号。

    流水按受理日期分段，不是全局连号——委托单号要能一眼看出是哪天受理的，
    跨天连号会让前台找单时先算一遍减法。
    **挂起的单不占号**：占了号又出不了单，台账上就多一排空号，
    业务员回头还得解释"这几张单去哪了"。等资料补齐、真的出单那天再编号。
    """
    seq = {}
    for order in orders:
        if order["问题"]:
            order["委托单号"] = ""
            continue
        day = order["受理日期"]
        seq[day] = seq.get(day, 0) + 1
        order["委托单号"] = f'{PREFIX}{day:%Y%m%d}-{seq[day]:02d}'
        order["加急"] = bool(order["备注"])
        for i, r in enumerate(order["行"], 1):
            r["序号"] = i
            r["样品编号"] = f'{order["委托单号"]}-{i:02d}'
            r["数量"] = as_qty(r["样品数量"])
            r["项目列表"] = split_items(r["检测项目"])
            r["项目数"] = len(r["项目列表"])


LABELS = ("委托单号", "受理日期", "加急", "客户名称", "联系人", "送样方式",
          "报告交付方式", "样品数", "检测项目数", "备注")


def label_rows(ws):
    """按标签在模板里反查"这个字段该写哪一格"。

    模板是质量部备案的，栏目顺序随时可能挪，写死坐标（B2、D2…）改一次版就全漂。
    反查一次拿到表，后面全按标签取值，模板换版也不用改代码。
    """
    found = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            label = clean(str(cell.value or ""))
            if label in LABELS and label not in found:
                found[label] = (cell.row, cell.column)
    return found


def detail_span(ws):
    """算出明细区：表头行的下一行起，连续 MAX_DETAIL 行。

    模板里明细空行**必须画边框**：没内容又没样式的行在 openpyxl 眼里根本不存在，
    ws.max_row 只到表头那一行，按 max_row 找明细区的写法会一行都找不到。
    """
    head = 0
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        texts = [clean(str(c.value or "")) for c in row]
        if "样品名称" in texts and "依据标准" in texts:
            head = row[0].row
            break
    if not head:
        raise SystemExit("模板里找不到明细表头（要有「样品名称」「依据标准」两列），先核对模板版本")
    span = min(ws.max_row - head, MAX_DETAIL)
    if span < MAX_DETAIL:
        raise SystemExit(f"模板明细区只认得 {span} 行（要 {MAX_DETAIL} 行），"
                         "检查一下模板里的空明细行是不是没画边框")
    return head + 1, head + span


def fill_order(order):
    """把一张委托单填进模板副本：单头按标签写，明细按行铺。"""
    from openpyxl import load_workbook

    wb = load_workbook(TPL_XLSX)
    ws = wb.active
    pos = label_rows(ws)

    def put(label, value):
        r, c = pos[label]
        ws.cell(row=r, column=c + 1, value=value)

    put("委托单号", order["委托单号"])
    put("客户名称", order["客户名称"])
    put("联系人", order["联系人"])
    put("送样方式", order["送样方式"])
    put("报告交付方式", order["报告交付方式"])
    put("加急", RUSH if order["加急"] else NORMAL)
    put("样品数", len(order["行"]))
    put("检测项目数", sum(r["项目数"] for r in order["行"]))
    put("备注", "；".join(order["备注"]))

    cell = ws.cell(row=pos["受理日期"][0], column=pos["受理日期"][1] + 1, value=order["受理日期"])
    cell.number_format = DATE_FMT          # 不设格式，Excel 里显示成 46265 这种序列号

    start, end = detail_span(ws)
    for i, r in enumerate(order["行"]):
        row = start + i
        ws.cell(row=row, column=1, value=r["序号"])
        ws.cell(row=row, column=2, value=r["样品名称"])
        code = ws.cell(row=row, column=3, value=r["客户样品编号"] or None)
        code.number_format = "@"            # 客户样品编号按文本写，否则 0000713 会掉前导零
        ws.cell(row=row, column=4, value=r["数量"])
        ws.cell(row=row, column=5, value="、".join(r["项目列表"]))
        ws.cell(row=row, column=6, value=r["依据标准"])
        ws.cell(row=row, column=7, value=r["备注"])
    # 明细没占满的行保持空白：塞 0 或 "无" 会被当成"有个数量为 0 的样品"

    ORDER_DIR.mkdir(parents=True, exist_ok=True)
    path = ORDER_DIR / f'{order["委托单号"]}.xlsx'
    wb.save(path)
    return path.name


# ----------------------------------------------------------------- ④ 输出（出单 + 台账）


def write_sheet(ws, header, rows, widths, key=None):
    """写一张工作表：表头加粗居中、冻结首行、按 key 整行刷底色。"""
    from openpyxl.styles import Alignment, Font, PatternFill

    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append(row)
        color = FILL.get(key(row)) if key else None
        if color:
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", start_color=color)
    ws.freeze_panes = "A2"
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width


def write_ledger(orders):
    """写受理台账三页：能出的单、单里的样品明细、出不来的待补资料。"""
    from openpyxl import Workbook

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    done = [o for o in orders if o["委托单号"]]
    held = [o for o in orders if not o["委托单号"]]

    ws = wb.active
    ws.title = "委托单台账"
    write_sheet(ws, ORDER_HEADER,
                [[o["委托单号"], o["受理日期"], o["客户名称"], o["联系人"], o["送样方式"],
                  o["报告交付方式"], RUSH if o["加急"] else NORMAL, len(o["行"]),
                  sum(r["项目数"] for r in o["行"]), o["文件"]] for o in done],
                ORDER_WIDTH, key=lambda row: row[6])

    rows = []
    for o in done:
        for r in o["行"]:
            rows.append([o["委托单号"], r["序号"], r["样品编号"], r["样品名称"],
                         r["客户样品编号"], r["数量"], r["项目数"],
                         "、".join(r["项目列表"]), r["依据标准"], r["备注"]])
    write_sheet(wb.create_sheet("样品明细"), DETAIL_HEADER, rows, DETAIL_WIDTH)

    write_sheet(wb.create_sheet("待补资料"), HOLD_HEADER,
                [[o["受理日期原文"] or "(空)", o["客户名称"] or "(空)",
                  "、".join(r["样品名称"] or "(空)" for r in o["行"][:3])
                  + ("…" if len(o["行"]) > 3 else ""),
                  o["问题"], o["说明"]] for o in held], HOLD_WIDTH)
    wb.save(LEDGER_XLSX)
    return len(done), len(held)


# ----------------------------------------------------------------- 回读校验


def verify(orders, done, held):
    """重新从磁盘把出的单和台账读回来，逐项断言——不自证，只看落盘的结果。"""
    from openpyxl import load_workbook

    checks = []

    def ok(name, good, detail=""):
        checks.append((name, bool(good), detail))

    by_no = {o["委托单号"]: o for o in done}
    hold_reasons = {o["问题"] for o in held}
    all_detail = [r for o in done for r in o["行"]]
    all_hold_rows = [r for o in held for r in o["行"]]

    ok("登记表 24 行全部读进", len(all_detail) + len(all_hold_rows) == 24,
       f'{len(all_detail) + len(all_hold_rows)} 行')
    ok("出单 4 张 / 挂起 6 单", (len(done), len(held)) == (4, 6),
       f"{len(done)} / {len(held)}")
    ok("出单覆盖 10 行样品 / 挂起 14 行样品",
       (len(all_detail), len(all_hold_rows)) == (10, 14),
       f"{len(all_detail)} / {len(all_hold_rows)}")
    ok("单号按日期分段连号",
       sorted(by_no) == ["WT20260918-01", "WT20260918-02", "WT20260918-03", "WT20260919-01"],
       f"{sorted(by_no)}")
    ok("同客户隔天再来一次是另一张单",
       by_no["WT20260918-01"]["客户名称"] == by_no["WT20260919-01"]["客户名称"]
       and len(by_no["WT20260919-01"]["行"]) == 2)
    ok("同一天同客户的三种日期格式与全角空格归成一张单",
       len(by_no["WT20260918-02"]["行"]) == 3, f'{len(by_no["WT20260918-02"]["行"])} 行')
    ok("样品编号 = 单号 + 两位序号",
       [r["样品编号"] for r in by_no["WT20260918-01"]["行"]]
       == ["WT20260918-01-01", "WT20260918-01-02", "WT20260918-01-03"])
    ok("项目数按分隔符拆开",
       by_no["WT20260918-01"]["行"][1]["项目数"] == 2
       and by_no["WT20260918-03"]["行"][0]["项目数"] == 3,
       f'{by_no["WT20260918-01"]["行"][1]["项目数"]} / {by_no["WT20260918-03"]["行"][0]["项目数"]}')
    ok("加急按整单算（同一单只一行标了加急）",
       by_no["WT20260918-03"]["加急"] and not by_no["WT20260918-01"]["加急"])
    ok("六类挂起原因各一单",
       hold_reasons == {"日期认不出", "客户名称缺失", "样品超上限",
                        "依据标准缺失", "数量不可用", "检测项目缺失"},
       f"{sorted(hold_reasons)}")
    held_note = next((o["说明"] for o in held if o["问题"] == "数量不可用"), "")
    ok("数量写成「2 支」的单挂起且说明带原文", "2 支" in held_note, held_note)
    ok("9 行样品那单挂起并提示拆单",
       "拆成多张" in next(o["说明"] for o in held if o["问题"] == "样品超上限"))

    wb = load_workbook(LEDGER_XLSX)
    ows, dws, hws = wb["委托单台账"], wb["样品明细"], wb["待补资料"]
    ok("台账行数 = 出单数", ows.max_row - 1 == len(done), f"{ows.max_row - 1}")
    ok("明细页行数 = 出单样品数", dws.max_row - 1 == len(all_detail),
       f"{dws.max_row - 1} vs {len(all_detail)}")
    ok("待补资料页行数 = 挂起单数", hws.max_row - 1 == len(held), f"{hws.max_row - 1}")
    ok("台账样品数合计 = 明细页行数",
       sum(ows.cell(row=r, column=8).value for r in range(2, ows.max_row + 1)) == len(all_detail))
    ok("客户样品编号保持补零原样",
       "0000713" in {r["客户样品编号"] for r in all_detail}
       and dws.cell(row=2, column=5).value == "0000713",
       f'{dws.cell(row=2, column=5).value}')
    ok("明细页第 2 行就是 01 号单的第 1 个样品",
       dws.cell(row=2, column=1).value == "WT20260918-01"
       and dws.cell(row=2, column=2).value == 1)

    # 逐张回读出的单，看单头与明细对不对
    head_ok, blank_ok, fmt_ok = True, True, True
    for no, o in by_no.items():
        ws = load_workbook(ORDER_DIR / f"{no}.xlsx").active
        pos = label_rows(ws)
        val = lambda name: ws.cell(row=pos[name][0], column=pos[name][1] + 1).value
        if (val("委托单号") != no or val("客户名称") != o["客户名称"]
                or val("样品数") != len(o["行"])
                or val("检测项目数") != sum(r["项目数"] for r in o["行"])):
            head_ok = False
        start, end = detail_span(ws)
        if ws.cell(row=start, column=2).value != o["行"][0]["样品名称"]:
            head_ok = False
        for i in range(len(o["行"])):
            if ws.cell(row=start + i, column=1).value != i + 1:
                head_ok = False
        for row in range(start + len(o["行"]), end + 1):
            if any(ws.cell(row=row, column=c).value not in (None, "") for c in range(1, 8)):
                blank_ok = False
        if ws.cell(row=start, column=3).number_format != "@" \
                or ws.cell(row=pos["受理日期"][0], column=pos["受理日期"][1] + 1) \
                .number_format != DATE_FMT:
            fmt_ok = False
    ok("每张单的单头（单号/客户/样品数/项目数）与明细一致", head_ok)
    ok("明细没占满的行保持空白，没被填 0 或「无」", blank_ok)
    ok("客户样品编号是文本、受理日期带日期格式", fmt_ok)

    tpl = load_workbook(TPL_XLSX).active
    tstart, tend = detail_span(tpl)
    ok("模板原件没被动过（明细区仍是空的）",
       all(tpl.cell(row=r, column=c).value in (None, "") for r in range(tstart, tend + 1)
           for c in range(1, 8)))

    return checks


# ----------------------------------------------------------------- main


def main():
    t0 = time.time()
    log_line(f"运行主机：{socket.gethostname()}（{local_ip()}）")
    log_line(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
             f"Python {sys.version.split()[0]} / {Path(sys.executable).name}")

    rows = read_entries()
    orders = group_orders(rows)
    for order in orders:
        problem = check_order(order)
        order["问题"], order["说明"] = problem if problem else ("", "")
    number_orders(orders)
    for order in orders:
        order["文件"] = fill_order(order) if order["委托单号"] else ""
    done, held = write_ledger(orders)
    print(f"登记 {len(rows)} 行 → {len(orders)} 张委托单：出单 {done} 张，挂起 {held} 张")

    checks = verify(orders, [o for o in orders if o["委托单号"]],
                    [o for o in orders if not o["委托单号"]])
    for name, good, detail in checks:
        print(f"  [{'OK ' if good else 'FAIL'}] {name}" + (f"   {detail}" if not good else ""))

    passed = all(c[1] for c in checks)
    print(f"\n出单 {done} 张（{sum(len(o['行']) for o in orders if o['委托单号'])} 个样品），"
          f"待补资料 {held} 张，回读校验 {sum(c[1] for c in checks)}/{len(checks)} 项通过")
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
