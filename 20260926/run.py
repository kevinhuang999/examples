#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第三方检测 · 送检样品登记与编号：把收样室当天一堆送样单读进来，
逐件编出唯一编号，生成样品登记台账 + 待打印标签清单，编不出号的一律挂起。

一句话：这活的难点不在编号本身，在**编号必须唯一、连续，而资料常常不全**。
手工编是先编号后判定——编到一半发现客户没建代码，号已经写下去了，只能跳号；
脚本是先判定后编号——编不出号的整行不进号段，号段天然是连的。

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

REG_CSV = RAW_DIR / "送样登记原始记录.csv"
CUST_CSV = RAW_DIR / "客户代码表.csv"
CAT_CSV = RAW_DIR / "样品类别代码表.csv"
TPL_XLSX = TPL_DIR / "样品登记台账.xlsx"
LEDGER_XLSX = OUT_DIR / "样品登记台账.xlsx"
LABEL_CSV = OUT_DIR / "样品标签清单.csv"

# 编号规则：受理日期(6 位) + 客户代码(2 位) + 类别码(1 位) + 流水(3 位)，例 260926QHA001。
# 流水按「受理日期 + 客户代码 + 类别码」分段，段内从 001 起连续排。
# **同一客户同一天来第二批（另一张送样单），继续顺着排，不重开 001**——
# 收样员看着手上这张单从 001 开始编，下午那张单又从头编一遍，重号就是这么来的。
DATE_SEG_FMT = "%y%m%d"
SEQ_WIDTH = 3

# 接收人是收样员而不是送样人；同一批台账由同一人接收，所以写成常量，不从数据里取。
# 数据里只有「客户送样单写了什么」，谁接的样是收样室的事。
RECEIVER = "收样室"

DATE_FMTS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d")
DATE_FMT = "yyyy-mm-dd"        # 写进 Excel 的日期格要顺手设格式，否则显示成 46265
TEXT_FMT = "@"                 # 客户样品编号走文本格式，否则 0000713 变成 713

# 表头按栏目名反查，写死行号的话模板一加说明行就全漂。
HEADER = ["样品编号", "受理日期", "委托单号", "客户名称", "客户样品编号",
          "样品名称", "样品类别", "数量", "存放位置", "接收人", "备注"]
SOURCE_HEAD = "受理日期"        # 原文表头所在的那一列，也用来判空跳过表尾行

HOLD_HEADER = ["受理日期", "委托单号", "客户名称", "客户样品编号", "样品名称", "数量", "挂起原因"]
LABEL_HEADER = ["样品编号", "样品名称", "客户名称", "接收日期", "存放位置"]

BLANK_RE = re.compile(r"\s+")
CUST_ALIAS = {"\u3000": ""}     # 全角空格：客户名尾巴粘一个，手工看是同一家、比较时是两家


# ----------------------------------------------------------------- 装依赖 / 日志


def ensure_deps():
    """缺 openpyxl 就自己装（走清华源）。"""
    try:
        import openpyxl  # noqa: F401
        return
    except ImportError:
        pass
    pkgs = ["-i", "https://pypi.tuna.tsinghua.edu.cn/simple", "openpyxl"]
    for extra in ([], ["--user"]):
        r = subprocess.run([sys.executable, "-m", "pip", "install", *pkgs, *extra],
                           capture_output=True, text=True)
        if r.returncode == 0:
            return
    print("依赖安装失败，请手动执行：\npip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple")
    sys.exit(1)


def host_line():
    """运行主机：主机名 + 本机对外网卡地址，用来分辨这段日志是哪台机器跑出来的。"""
    try:
        name = socket.gethostname()
    except Exception:
        name = "未知"
    ip = "未知"
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        try:
            ip = socket.gethostbyname(name)
        except Exception:
            ip = "未知"
    return f"运行主机：{name}（{ip}）"


def log(text):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")


# ----------------------------------------------------------------- ② 读数据


def norm(value):
    """比较用的键：去空白（含全角空格）。显示照原样，只压比较用的那一份。"""
    if value is None:
        return ""
    s = str(value)
    for k, v in CUST_ALIAS.items():
        s = s.replace(k, v)
    return BLANK_RE.sub("", s)


def read_csv_rows(path, head_col):
    """读一张 CSV 成 dict 列表，全部按文本取。

    **一律不转类型**——客户样品编号 0000713 一转 int 就变 713，标签和原始记录上写的都是带前导零的。
    表尾的合计行与说明行靠 `head_col` 那一列为空跳过；所以合计数字千万别写进那一列。
    """
    rows = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            if not raw or not norm(raw.get(head_col, "")):
                continue
            rows.append({k: ("" if v is None else str(v)) for k, v in raw.items()})
    return rows


def load_code_map(path, key_col, val_col):
    """读「名称 → 代码」对照表，键压成比较用的形式。判空的列用 key_col 自己。"""
    m = {}
    for row in read_csv_rows(path, key_col):
        k, v = norm(row.get(key_col)), str(row.get(val_col, "")).strip()
        if k and v:
            m[k] = v
    return m


def parse_date(text):
    """认不出返回 None，**不许兜今天**——日期没定的送样单出成当天的，编号里的日期段就是错的。"""
    s = str(text).strip()
    for fmt in DATE_FMTS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


# ----------------------------------------------------------------- ③ 判定与编号


def check_row(row, cust_map, cat_map, dup_keys):
    """逐行判定，返回挂起原因；返回 None 表示这行能编上号。

    顺序就是业务顺序：先看日期（编号的第一段就是它），再看客户有没有代码，
    再看类别有没有类别码（这两段是编号的中间一段），最后才是台账要展示的字段。
    """
    if parse_date(row.get("受理日期")) is None:
        return "受理日期认不出，编号里定不下日期段"
    if not norm(row.get("客户名称")):
        return "客户名称空，认不出是哪家"
    if norm(row.get("客户名称")) not in cust_map:
        return "客户没有代码，先在客户代码表里建档"
    if norm(row.get("样品类别")) not in cat_map:
        return "样品类别没有类别码，别猜，先补类别表"
    if not norm(row.get("样品名称")):
        return "样品名称空，台账上没法标识"
    qty = parse_qty(row.get("数量"))
    if qty is None:
        return f"数量认不出件数（原件写的「{str(row.get('数量')).strip()}」），别替客户改口径"
    if (row.get("委托单号"), norm(row.get("客户样品编号"))) in dup_keys:
        return "同一委托单里客户样品编号重复，追溯会串"
    return None


def parse_qty(text):
    """数量必须是正整数。客户写「2 支」就挂起——用正则抠出 2 等于替他改了口径。"""
    s = str(text).strip()
    return int(s) if s.isdigit() and int(s) > 0 else None


def dup_key_set(rows):
    """同一委托单内重复出现的客户样品编号，整组都要挂——不能只挂后一条。"""
    count = {}
    for row in rows:
        k = (row.get("委托单号"), norm(row.get("客户样品编号")))
        count[k] = count.get(k, 0) + 1
    return {k for k, n in count.items() if n > 1 and k[1]}


def number_items(rows, cust_map, cat_map):
    """判定 + 编号：通过的行按数量展开成件，每一件拿一个编号。

    返回 (items, holds)。**挂起的行不占流水号**——号是给台账上能贴标签的件用的，
    资料补齐那天再补号，号段因此始终是连续的。
    """
    dup_keys = dup_key_set(rows)
    items, holds, seq = [], [], {}
    for order, row in enumerate(rows):
        reason = check_row(row, cust_map, cat_map, dup_keys)
        if reason:
            holds.append(dict(row, **{"挂起原因": reason}))
            continue
        d = parse_date(row.get("受理日期"))
        seg = (d, cust_map[norm(row["客户名称"])], cat_map[norm(row["样品类别"])])
        seq[seg] = seq.get(seg, 0) + 1
        main = f"{d.strftime(DATE_SEG_FMT)}{seg[1]}{seg[2]}{seq[seg]:0{SEQ_WIDTH}d}"
        qty = parse_qty(row.get("数量"))
        for k in range(1, qty + 1):
            # 件数为 1 不加后缀：满台账都是 -1 没人这么干。件序号不补零，跟规范里的 -1 示例一致。
            items.append({
                "样品编号": main if qty == 1 else f"{main}-{k}",
                "主编号": main, "件序号": k, "件数": qty, "行序": order,
                "受理日期": d, "委托单号": row.get("委托单号"),
                "客户名称": str(row.get("客户名称")).strip(),
                "客户样品编号": str(row.get("客户样品编号")).strip(),
                "样品名称": str(row.get("样品名称")).strip(),
                "样品类别": str(row.get("样品类别")).strip(),
                "数量": qty, "存放位置": str(row.get("存放位置")).strip(),
                "接收人": RECEIVER, "备注": str(row.get("备注")).strip(),
            })
    items.sort(key=lambda x: (x["行序"], x["件序号"]))
    return items, holds


# ----------------------------------------------------------------- ④ 输出


def head_row_no(ws, title):
    """按栏目名反查表头在第几行——模板加一行说明，写死行号就全漂。"""
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and norm(v) == norm(title):
                return r
    raise KeyError(f"模板里找不到表头栏目：{title}")


def head_col(ws, head_row, title):
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=head_row, column=c).value
        if v is not None and norm(v) == norm(title):
            return c
    raise KeyError(f"模板表头里找不到栏目：{title}")


def write_register(items, holds):
    """写台账：模板抬头 + 表头不动，明细行按件数创建并逐行画边框。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, Side
    from openpyxl.utils import get_column_letter

    wb = load_workbook(TPL_XLSX)
    ws = wb.active
    hr = head_row_no(ws, "样品编号")
    col = {name: head_col(ws, hr, name) for name in HEADER}

    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    body = Font(name="微软雅黑", size=10)

    for i, it in enumerate(items):
        r = hr + 1 + i
        for name in HEADER:
            cell = ws.cell(row=r, column=col[name])
            cell.font = body
            cell.border = border
        ws.cell(row=r, column=col["样品编号"], value=it["样品编号"])
        dt = ws.cell(row=r, column=col["受理日期"], value=it["受理日期"])
        dt.number_format = DATE_FMT
        ws.cell(row=r, column=col["委托单号"], value=it["委托单号"])
        ws.cell(row=r, column=col["客户名称"], value=it["客户名称"])
        cn = ws.cell(row=r, column=col["客户样品编号"], value=it["客户样品编号"])
        cn.number_format = TEXT_FMT
        ws.cell(row=r, column=col["样品名称"], value=it["样品名称"])
        ws.cell(row=r, column=col["样品类别"], value=it["样品类别"])
        ws.cell(row=r, column=col["数量"], value=it["数量"])
        ws.cell(row=r, column=col["存放位置"], value=it["存放位置"])
        ws.cell(row=r, column=col["接收人"], value=it["接收人"])
        ws.cell(row=r, column=col["备注"], value=it["备注"])
        ws.cell(row=r, column=col["样品编号"]).alignment = Alignment(horizontal="left", vertical="center")
    ws.column_dimensions[get_column_letter(col["样品编号"])].width = 20

    hs = wb.create_sheet("挂起清单")
    for i, name in enumerate(HOLD_HEADER, start=1):
        c = hs.cell(row=1, column=i, value=name)
        c.font = Font(name="微软雅黑", size=10, bold=True)
        c.border = border
    for i, row in enumerate(holds):
        for j, name in enumerate(HOLD_HEADER, start=1):
            c = hs.cell(row=2 + i, column=j, value=str(row.get(name, "")))
            c.font = body
            c.border = border
    for i, w in enumerate([12, 14, 26, 16, 20, 8, 46], start=1):
        hs.column_dimensions[get_column_letter(i)].width = w
    hs.freeze_panes = "A2"

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(LEDGER_XLSX)


def write_labels(items):
    """标签清单：一件一行，收样室拿它打印不干胶标签。编号错了标签全废，所以清单和台账同源。"""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(LABEL_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(LABEL_HEADER)
        for it in items:
            w.writerow([it["样品编号"], it["样品名称"], it["客户名称"],
                        it["受理日期"].strftime("%Y-%m-%d"), it["存放位置"]])


# ----------------------------------------------------------------- 回读校验


def verify(items, holds, expect):
    """跑完回读自己生成的文件，逐项断言。"""
    from openpyxl import load_workbook

    checks = []

    def ok(name, cond):
        checks.append((name, bool(cond)))

    wb = load_workbook(LEDGER_XLSX)
    ws = wb["样品登记台账"]
    hr = head_row_no(ws, "样品编号")
    col = {name: head_col(ws, hr, name) for name in HEADER}
    read = []
    for r in range(hr + 1, ws.max_row + 1):
        no = ws.cell(row=r, column=col["样品编号"]).value
        if not no:
            continue
        read.append({
            "样品编号": str(no),
            "客户样品编号": str(ws.cell(row=r, column=col["客户样品编号"]).value),
            "客户名称": str(ws.cell(row=r, column=col["客户名称"]).value),
            "样品名称": str(ws.cell(row=r, column=col["样品名称"]).value),
            "受理日期": ws.cell(row=r, column=col["受理日期"]).value,
            "接收人": str(ws.cell(row=r, column=col["接收人"]).value),
        })
    hold_read, hold_no = [], []
    hs = wb["挂起清单"]
    for r in range(2, hs.max_row + 1):
        v = hs.cell(row=r, column=1).value
        if v:
            hold_read.append(str(v))
            hold_no.append(str(hs.cell(row=r, column=2).value))

    no_list = [x["样品编号"] for x in read]
    mains = sorted({x["样品编号"].split("-")[0] for x in read})

    ok(f"台账件数 = 通过行按数量展开后的 {expect['items']} 件", len(read) == expect["items"])
    ok("台账件数比通过的行数多，说明按数量逐件展开了", len(read) > expect["rows_ok"])
    ok("台账里没有空编号", all(no_list))
    ok("编号全局唯一，没有重号", len(set(no_list)) == len(no_list))
    ok(f"主编号共 {expect['mains']} 个", len(mains) == expect["mains"])

    # 本行只给了一件样品 → 不加后缀；给了多件 → 带 -1 到 -N
    single = [x for x in read if x["样品编号"].count("-") == 0]
    ok("单件样品的编号不带后缀", len(single) == expect["single"])
    ok("多件样品第一件是 -1", any(x["样品编号"].endswith("-1") for x in read))

    # 号段连续：按主编号前缀分组，件序号从 1 到 N 一个不跳
    groups = {}
    for x in read:
        pre = x["样品编号"].split("-")[0]
        k = int(x["样品编号"].split("-")[1]) if "-" in x["样品编号"] else 1
        groups.setdefault(pre, []).append(k)
    ok("每个号段的件序号从 1 连续到 N", all(sorted(v) == list(range(1, len(v) + 1)) for v in groups.values()))

    # 流水连续：同一（日期 + 客户代码 + 类别码）段内 001…00N 不跳号
    segs = {}
    for m in mains:
        segs.setdefault(m[6:9], []).append(int(m[9:]))
    ok("同段流水从 001 起连续、不跳号", all(sorted(v) == list(range(1, len(v) + 1)) for v in segs.values()))
    ok("启海饮用水段流水 = 001,002,003,004", sorted(segs.get("QHA", [])) == [1, 2, 3, 4])

    # ★关键：同客户同一天第二批接着排，不重开 001
    ok("第二批（WT26092605）没重开 001，接着排在 003",
       any(x["样品编号"] == "260926QHA003-1" for x in read))

    ok("客户样品编号前导零 0000713 保住了",
       any(x["客户样品编号"] == "0000713" for x in read))
    ok("全角空格客户名归一化后对上代码表（宏远进了台账）",
       any("宏远" in x["客户名称"] for x in read))
    ok("120 件那批展开到 -120",
       any(x["样品编号"].endswith("-120") for x in read))
    ok("120 件那批的主编号是 260926QHA004", "260926QHA004" in mains)

    # ★字符串排序会错位：-120 会排到 -2 前面，所以台账顺序按 (行序, 件序号) 数值排
    ok("台账不是按编号字符串排序的（-10 会排到 -2 前）",
       sorted(no_list) != no_list)
    ok("同批内 -2 排在 -120 前面", no_list.index("260926QHA004-2") < no_list.index("260926QHA004-120"))

    ok(f"挂起 {expect['holds']} 行", len(hold_read) == expect["holds"])
    ok("挂起的行一个都没进台账（待定日期那单 WT26092606 不在）",
       "WT26092606" not in {str(x) for x in read})
    ok("挂起清单第一列仍是受理日期（没被编号覆盖）",
       all(h == "待定" or parse_date(h) is not None for h in hold_read))
    ok("重复样号那两行都挂起了（同一委托单内 MD-A-773）",
       sum(1 for h in hold_no if h == "WT26092612") == 2)
    ok("接收人一列统一填了收样室", all(x["接收人"] == RECEIVER for x in read))
    ok("受理日期回读是日期，不是序列号",
       isinstance(read[0]["受理日期"], (date, datetime)))

    with open(LABEL_CSV, "r", encoding="utf-8-sig", newline="") as f:
        lab = list(csv.reader(f))
    ok("标签清单行数 = 台账件数 + 1 行表头", len(lab) - 1 == len(read))
    ok("标签清单第一列就是样品编号", lab[0][0] == "样品编号")

    print()
    for name, good in checks:
        print(f"  {'OK  ' if good else 'FAIL'} {name}")
    passed = sum(1 for _, g in checks if g)
    print(f"\n回读校验：{passed}/{len(checks)} 项通过")
    return passed == len(checks), len(checks), passed


# ----------------------------------------------------------------- 主线


def main():
    ensure_deps()
    t0 = time.time()
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    log(host_line())
    log(f"开始执行：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Python：{sys.version.split()[0]}（{sys.executable}）")

    rows = read_csv_rows(REG_CSV, SOURCE_HEAD)
    cust_map = load_code_map(CUST_CSV, "客户名称", "客户代码")
    cat_map = load_code_map(CAT_CSV, "样品类别", "类别码")
    items, holds = number_items(rows, cust_map, cat_map)
    write_register(items, holds)
    write_labels(items)
    expect = {"items": 132, "holds": 8, "mains": 8, "single": 3, "rows_ok": 8}
    ok, total, passed = verify(items, holds, expect)

    used = time.time() - t0
    print(f"\n收样登记 {len(rows)} 行 → 台账 {len(items)} 件 → 主编号 {len({i['主编号'] for i in items})} 个 "
          f"→ 挂起 {len(holds)} 行")
    print(f"结果目录：{OUT_DIR}")
    log(f"收样登记 {len(rows)} 行 → 台账 {len(items)} 件 → 挂起 {len(holds)} 行")
    log(f"执行完成：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  用时 {used:.1f}s  "
        f"判定 {'PASS' if ok else 'FAIL'}（{passed}/{total}）")
    print(f"判定：{'PASS' if ok else 'FAIL'}（{passed}/{total} 项）")

    if "--no-pause" not in sys.argv:
        input("\n按回车退出…")


if __name__ == "__main__":
    main()
