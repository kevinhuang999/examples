#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第三方检测 · 检测报告批量改名分发：按每家客户自己的命名规则给报告改名、按客户拆包，
再出一份符合《检验检测机构报告发放管理要求》的发放登记表，外加一份给客服用的分发清单。

一句话：这活的难点不在改名，在**名字是客户定的、格式各不一样，而发放记录是制度定的、一份都不能少**。
名字改错是客户收到别人的报告；记录漏了是追溯时拿不出证据。

用法：
    python run.py              # 跑完停住等回车（Windows 双击也行）
    python run.py --no-pause   # 跑完直接退出
"""

import csv
import re
import shutil
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
ORIG_DIR = RAW_DIR / "报告原件"

REPORT_CSV = RAW_DIR / "报告台账.csv"
NAMING_CSV = RAW_DIR / "客户命名规则.csv"
RECIPIENT_CSV = RAW_DIR / "客户收件人.csv"
TPL_XLSX = TPL_DIR / "报告发放登记表.xlsx"
LEDGER_XLSX = OUT_DIR / "发放登记台账.xlsx"
MANIFEST_CSV = OUT_DIR / "客户分发清单.csv"
PACK_DIR = OUT_DIR / "报告分发"

# 发放日期：登记表上「发放日期」这一栏填哪天，得是定的。
# **不许用 date.today()**——隔天重跑，同一批报告的发放日期就变了，跟客服手上的邮件时间对不上。
SEND_DAY = date(2026, 9, 25)

# 不对外分发的报告状态。作废报告的电子件流出去是合规事故，宁可挂起也不改名。
BLOCK_STATUS = {"待审核": "报告未出具，不能分发", "已作废": "报告已作废，不对外分发"}

# 文件名里不许出现的字符。产品名、委托单号里都可能带斜杠，不换掉连文件都建不出来。
ILLEGAL_RE = re.compile(r'[\\/:*?"<>|\r\n\t]+')
LEFTOVER_RE = re.compile(r"\{[^}]*\}")          # 拼完还剩花括号 = 模板里有不认识的字段

DATE_FMTS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d")
DATE_FMT = "yyyy-mm-dd"                         # 写进 Excel 的日期格要顺手设格式，否则显示成序列号

LEDGER_HEADER = ["报告编号", "产品名称", "报告数量（份）", "领取单位", "领取人", "联系电话",
                 "发放方式", "发放日期", "文件名", "经办人", "备注"]
LEDGER_WIDTH = [16, 22, 12, 26, 10, 14, 10, 12, 40, 10, 26]
HOLD_HEADER = ["报告编号", "客户名称", "问题", "说明"]
HOLD_WIDTH = [16, 26, 20, 52]
ORPHAN_HEADER = ["原件文件", "大小(KB)", "说明"]
ORPHAN_WIDTH = [20, 12, 52]


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


# ----------------------------------------------------------------- ② 归一化与清洗


def clean(text):
    """去空白。全角空格（U+3000）也要去——客户名是从 Excel 里粘过来的，
    尾巴上带一个全角空格，不去干净就会跟同名客户分成两家，收件人查不到、整家报告卡住。"""
    return (text or "").replace("\u3000", " ").strip()


def cust_key(name):
    """客户名的比较键：再去掉内部空白。「上海启海化工 」和「上海启海化工」是同一家。"""
    return re.sub(r"\s+", "", clean(name))


def parse_date(text):
    """认四种日期写法，认不出返回 None（不兜今天）。"""
    s = clean(text)
    for fmt in DATE_FMTS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def read_csv_rows(path):
    """读 CSV。第一列为空的行一律跳过——表尾的合计行、说明行第一列都是空的，
    说明文字要是写进第一列，就会被当成一份报告混进分发清单。"""
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    header = [clean(x) for x in rows[0]]
    out = []
    for raw in rows[1:]:
        if not raw or not clean(raw[0]):
            continue
        out.append({header[i]: clean(raw[i]) if i < len(raw) else "" for i in range(len(header))})
    return out


# ----------------------------------------------------------------- ③ 读数据（归堆）


def read_reports():
    """读报告台账。报告编号、委托单号全程当字符串用，不转 int——
    编号里有前导零的话，`int()` 一转就没了，改出来的文件名跟原件编号对不上。"""
    reports = []
    for row in read_csv_rows(REPORT_CSV):
        reports.append({
            "报告编号": row["报告编号"],
            "委托单号": row["委托单号"],
            "客户名称": row["客户名称"],
            "产品名称": row["产品名称"],
            "报告份数": row["报告份数"],
            "报告日期": parse_date(row["报告日期"]),
            "状态": row["状态"],
        })
    return reports


def read_naming():
    """读客户命名规则，按客户名的比较键收成字典。"""
    rules = {}
    for row in read_csv_rows(NAMING_CSV):
        rules[cust_key(row["客户名称"])] = {
            "客户名称": row["客户名称"],
            "写客户": row["文件名里客户写法"],
            "模板": row["命名模板"],
            "日期写法": row["日期写法"],
        }
    return rules


def read_recipients():
    """读客户收件人。一家客户可能有好几个收件人（客户本人 + 客户指定的甲方），
    所以这里收成「客户 → 收件人列表」。"""
    recs = {}
    for row in read_csv_rows(RECIPIENT_CSV):
        recs.setdefault(cust_key(row["客户名称"]), []).append({
            "收件人": row["收件人"],
            "电话": row["联系电话"],
            "邮箱": row["邮箱"],
            "发放方式": row["发放方式"],
        })
    return recs


def scan_originals():
    """扫报告原件目录，用文件名（去扩展名）当报告编号。

    文件名一律按字符串处理，**不做 int 转换**——老系统留下的编号是 `0000122`，
    转成整数就变成 `122`，跟台账永远对不上，人还看不出是哪儿错了。
    """
    found = {}
    for path in sorted(ORIG_DIR.glob("*.pdf")):
        found[path.stem] = path
    return found


# ----------------------------------------------------------------- ④ 分发判定与改名


def check_report(rep, naming, recipients, originals):
    """判定顺序就是业务顺序：报告没出 → 不说名字的事；客户没登记收件人 → 不判原件；
    原件不在 → 不拼名字。

    顺序反了挂起理由就会张冠李戴：一份还没审核的报告被标成「原件缺失」，
    客服拿着错理由去补资料，补完还是发不出去。
    """
    if rep["状态"] in BLOCK_STATUS:
        return "挂起", rep["状态"], BLOCK_STATUS[rep["状态"]]
    rule = naming.get(cust_key(rep["客户名称"]))
    if rule is None:
        return "挂起", "客户命名规则缺失", "台账里的客户在命名规则表里查不到，不知道按谁的规矩改名"
    if not recipients.get(cust_key(rep["客户名称"])):
        return "挂起", "未登记收件人", "客户收件人表里没有这一家，报告不知道发给谁"
    if clean(rep["报告编号"]) not in originals:
        return "挂起", "原件缺失", "报告原件目录里找不到这份报告，改不了名"
    if rep["报告日期"] is None:
        return "挂起", "报告日期认不出", "日期列写法不在支持列表里，不能用它拼文件名"
    return "正常", "", ""


def make_filename(rep, rule):
    """按客户给的模板拼文件名：先换占位符，再换非法字符，最后统一加 .pdf。

    拼完必须回头查一遍还有没有没换掉的花括号——客户把字段名写错一个字
    （`{产品名}`），不检查就会把一个带花括号的文件名发给客户。
    """
    name = rule["模板"]
    for key in ("报告编号", "委托单号", "产品名称"):
        name = name.replace("{%s}" % key, clean(rep[key]))
    name = name.replace("{客户}", clean(rule["写客户"]) or clean(rep["客户名称"]))
    name = name.replace("{报告日期}", rep["报告日期"].strftime(rule["日期写法"]))
    if LEFTOVER_RE.search(name):
        raise ValueError(f"模板里有不认识的占位符：{LEFTOVER_RE.findall(name)}")
    return ILLEGAL_RE.sub("-", name).strip() + ".pdf"


def plan(reports, naming, recipients, originals):
    """先把能发的挑出来，再回头查撞名。撞名要两份都挂起——
    只挂后一份的话，前一份照样发出去，两家客户拿到的报告就混了。"""
    items, held, keys = [], [], {}
    for rep in reports:
        kind, problem, note = check_report(rep, naming, recipients, originals)
        if kind == "挂起":
            held.append({"报告编号": rep["报告编号"], "客户名称": rep["客户名称"],
                         "问题": problem, "说明": note})
            continue
        rule = naming[cust_key(rep["客户名称"])]
        try:
            name = make_filename(rep, rule)
        except ValueError as exc:
            held.append({"报告编号": rep["报告编号"], "客户名称": rep["客户名称"],
                         "问题": "文件名拼不出来", "说明": str(exc)})
            continue
        item = {"rep": rep, "rule": rule, "文件名": name,
                "收件人": recipients[cust_key(rep["客户名称"])],
                "键": (cust_key(rep["客户名称"]), name.lower())}
        items.append(item)
        keys.setdefault(item["键"], []).append(item)

    dup = {k for k, group in keys.items() if len(group) > 1}
    if dup:
        kept = []
        for item in items:
            if item["键"] in dup:
                held.append({"报告编号": item["rep"]["报告编号"], "客户名称": item["rep"]["客户名称"],
                             "问题": "命名撞名",
                             "说明": f"同一家客户下另有报告算出同名文件 {item['文件名']}，"
                                     "两份都挂起，先跟客户确认改名规则"})
            else:
                kept.append(item)
        items = kept
    return items, held


def pack(items, originals):
    """按客户建目录、把原件**复制**成分发件。

    复制不是移动：原件是归档凭证，动它等于动了可追溯性；分发件只是副本，
    目录每次重跑先清空，免得上一版留下的文件被当成这一版发出去。
    """
    if PACK_DIR.exists():
        shutil.rmtree(PACK_DIR)
    rows = []
    for item in items:
        cust_dir = PACK_DIR / ILLEGAL_RE.sub("-", clean(item["rep"]["客户名称"]))
        cust_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(originals[clean(item["rep"]["报告编号"])], cust_dir / item["文件名"])
        rows.append(item)
    return rows


# ----------------------------------------------------------------- ⑤ 输出


def ledger_rows(items):
    """发放登记表按「报告 × 收件人」一行——制度里的登记表记的是**每一次发放动作**，
    一份报告发给两个人就是两次发放。

    份数按发放方式给：纸质按委托约定的份数（台账里的报告份数），电子只记 1 份。
    电子件多记份数会让台账上的数量虚高，对不上客户实际拿到的纸。
    """
    rows = []
    for item in items:
        rep = item["rep"]
        for rec in item["收件人"]:
            paper = rec["发放方式"] == "纸质"
            rows.append([rep["报告编号"], rep["产品名称"], rep["报告份数"] if paper else "1",
                         rep["客户名称"], rec["收件人"], rec["电话"], rec["发放方式"],
                         SEND_DAY, item["文件名"], "",
                         f"纸质件按 {rep['报告份数']} 份寄出" if paper else "电子件加密发送"])
    return rows


def head_row(ws, label):
    """表头行按栏目名反查：模板上下加标题行、改列顺序都不影响定位，写死行号必漂。"""
    for row in ws.iter_rows():
        for cell in row:
            if clean(str(cell.value or "")) == label:
                return cell.row
    raise SystemExit(f"模板里找不到「{label}」这一栏")


def head_col(ws, row, label):
    for cell in ws[row]:
        if clean(str(cell.value or "")) == label:
            return cell.column
    raise SystemExit(f"表头行里找不到「{label}」这一栏")


def fill_ledger(rows):
    """把登记行写进模板：模板不动表头与样式，只从表头下面第一行开始填。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    thin = Side(style="thin", color="8C8C8C")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = load_workbook(TPL_XLSX)
    ws = wb[wb.sheetnames[0]]
    head = head_row(ws, "报告编号")
    date_col = head_col(ws, head, "发放日期")
    start = head + 1
    for i, row in enumerate(rows):
        for col, value in enumerate(row, start=1):
            cell = ws.cell(row=start + i, column=col, value=value)
            cell.border = border
            cell.font = Font(name="等线", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=False)
        ws.cell(row=start + i, column=date_col).number_format = DATE_FMT
    return head, start, wb


def write_sheet(wb, title, header, rows, widths, first=False):
    from openpyxl.styles import Alignment, Font

    ws = wb[wb.sheetnames[0]] if first else wb.create_sheet(title)
    ws.title = title
    for col, name in enumerate(header, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.font = Font(name="等线", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = widths[col - 1]
    for i, row in enumerate(rows, start=2):
        for col, value in enumerate(row, start=1):
            ws.cell(row=i, column=col, value=value).font = Font(name="等线", size=10)
    return ws


def write_manifest(items):
    """给客服一份分发清单：哪个收件人、要发哪几份、邮箱是哪个，照着发就行。"""
    groups = {}
    for item in items:
        for rec in item["收件人"]:
            key = (item["rep"]["客户名称"], rec["收件人"], rec["发放方式"], rec["邮箱"])
            groups.setdefault(key, []).append(item["文件名"])
    with open(MANIFEST_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["客户名称", "收件人", "电子邮箱", "发放方式", "报告数", "分发文件", "建议邮件主题"])
        for (cust, who, way, mail), files in groups.items():
            w.writerow([cust, who, mail, way, len(files), "；".join(sorted(files)),
                        f"【检测报告】{cust} 报告已出具（{len(files)} 份）"])
    return groups


# ----------------------------------------------------------------- 回读校验


def verify(reports, items, held, orphans):
    from openpyxl import load_workbook

    checks = []

    def ok(name, good, detail=""):
        checks.append((name, bool(good), detail))

    ok("台账 11 行全读到（表尾说明行没被当成报告）", len(reports) == 11, f"{len(reports)}")
    ok("分发件 4 份，都成功复制到分发目录",
       len(items) == 4 and all((PACK_DIR / ILLEGAL_RE.sub("-", clean(i["rep"]["客户名称"]))
                                / i["文件名"]).exists() for i in items),
       f"{len(items)}")
    ok("挂起 7 份", len(held) == 7, f"{len(held)}")
    ok("五种挂起理由齐全（命名撞名2/未登记收件人2/原件缺失1/待审核1/已作废1）",
       sorted(r["问题"] for r in held) == sorted(["命名撞名", "命名撞名", "未登记收件人",
                                                  "未登记收件人", "原件缺失", "待审核", "已作废"]),
       f'{[r["问题"] for r in held]}')

    filed = {i["rep"]["报告编号"] for i in items}
    ok("已作废的 WX260918010 没进分发件", "WX260918010" not in filed)
    ok("待审核的 WX260918009 没进分发件", "WX260918009" not in filed)
    ok("原件缺失的 QH260918011 没进分发件", "QH260918011" not in filed)
    ok("撞名的 JD260917005 / JD260917006 两份都挂起",
       {"JD260917005", "JD260917006"} & filed == set()
       and {r["报告编号"] for r in held if r["问题"] == "命名撞名"} == {"JD260917005", "JD260917006"})
    ok("原件目录里的已作废件仍在原处（只复制、不动原件）",
       (ORIG_DIR / "WX260918010.pdf").exists() and len(list(ORIG_DIR.glob("*.pdf"))) == 10,
       f"{len(list(ORIG_DIR.glob('*.pdf')))}")
    ok("分发目录里没有出现作废件文件名",
       all("WX260918010" not in p.name for p in PACK_DIR.rglob("*.pdf")))

    names = {i["rep"]["报告编号"]: i["文件名"] for i in items}
    ok("宏远：日期写成 20260915、产品名里的斜杠换成横杠",
       names["WX260915001"] == "WX260915001_铝合金-塑钢型材_20260915.pdf"
       and "/" not in names["WX260915001"], names.get("WX260915001", ""))
    ok("启海：文件名里用客户的简称「启海化工」、日期写成 2026-09-16",
       names["QH260916003"] == "启海化工-WT26091620-QH260916003.pdf", names.get("QH260916003", ""))
    ok("台账里 2026.09.15 这种写法也认得出，文件名按客户要的格式写",
       names["WX260915002"] == "WX260915002_棉麻混纺面料_20260915.pdf", names.get("WX260915002", ""))
    ok("金鼎两份报告算出的文件名确实撞在一起（挂起理由不是瞎写的）",
       len({r["说明"] for r in held if r["问题"] == "命名撞名"}) == 1
       and "建筑用钢化玻璃_2026.09.17.pdf" in
       [r["说明"] for r in held if r["问题"] == "命名撞名"][0])

    wb = load_workbook(LEDGER_XLSX)
    ws = wb["发放登记"]
    head = head_row(ws, "报告编号")
    ws2, ws3 = wb["待处理"], wb["无对应台账原件"]
    ok("发放登记表数据行 = 报告×收件人 = 6 行", ws.max_row - head == 6, f"{ws.max_row - head}")
    ok("待处理页 7 行", ws2.max_row - 1 == 7, f"{ws2.max_row - 1}")
    ok("无对应台账原件页 1 行", ws3.max_row - 1 == 1, f"{ws3.max_row - 1}")
    ok("游离原件的编号保留前导零 0000122（没被当数字吃掉）",
       ws3.cell(row=2, column=1).value == "0000122.pdf", f'{ws3.cell(row=2, column=1).value}')

    ledger = [[ws.cell(row=r, column=c).value for c in range(1, len(LEDGER_HEADER) + 1)]
              for r in range(head + 1, ws.max_row + 1)]
    ok("登记表按「报告×收件人」展开：启海 003 有两行（陆志远 / 李慧）",
       len([r for r in ledger if r[0] == "QH260916003"]) == 2)
    paper = [r for r in ledger if r[6] == "纸质"]
    ok("纸质收件人的份数按台账（3 份与 1 份），电子一律记 1 份",
       sorted(r[2] for r in paper) == ["1", "3"]
       and all(r[2] == "1" for r in ledger if r[6] == "电子"),
       f"{[r[2] for r in ledger]}")
    # 写进去的是 date，读回来是 datetime —— 直接比 date 必挂，得先取 .date()
    ok("登记表的发放日期 = 常量 SEND_DAY，且是日期格式",
       all(getattr(r[7], "date", lambda: r[7])() == SEND_DAY for r in ledger)
       and ws.cell(row=head + 1, column=head_col(ws, head, "发放日期")).number_format == DATE_FMT,
       f'{ledger[0][7]!r}')
    ok("登记表的领取单位与领取人跟收件人表对得上",
       {(r[3], r[4]) for r in ledger} == {("浙江宏远纺织有限公司", "周敏"),
                                          ("上海启海化工有限公司", "陆志远"),
                                          ("上海启海化工有限公司", "李慧")})
    ok("收件人表里客户名带全角空格，仍归到同一家（启海 4 行 = 2 报告 × 2 收件人）",
       len([r for r in ledger if r[3] == "上海启海化工有限公司"]) == 4,
       f"{len([r for r in ledger if r[3] == '上海启海化工有限公司'])}")

    with open(MANIFEST_CSV, "r", encoding="utf-8-sig", newline="") as f:
        manifest = list(csv.reader(f))
    ok("分发清单 3 行（宏远 1 个收件人 + 启海 2 个），按收件人归堆",
       len(manifest) - 1 == 3, f"{len(manifest) - 1}")
    ok("分发清单里没有登记不上收件人的金鼎、新洲",
       all("金鼎" not in row[0] and "新洲" not in row[0] for row in manifest[1:]))

    tpl = load_workbook(TPL_XLSX)["发放登记"]
    tpl_head = head_row(tpl, "报告编号")
    ok("模板原件没被动过（表头下面还是空的）",
       all(tpl.cell(row=r, column=c).value in (None, "") for r in range(tpl_head + 1, tpl.max_row + 2)
           for c in range(1, len(LEDGER_HEADER) + 1)))

    return checks


# ----------------------------------------------------------------- main


def main():
    t0 = time.time()
    log_line(f"运行主机：{socket.gethostname()}（{local_ip()}）")
    log_line(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
             f"Python {sys.version.split()[0]} / {Path(sys.executable).name}")

    reports = read_reports()
    naming = read_naming()
    recipients = read_recipients()
    originals = scan_originals()

    items, held = plan(reports, naming, recipients, originals)
    items = pack(items, originals)
    used = {clean(i["rep"]["报告编号"]) for i in items} | {clean(r["报告编号"]) for r in held}
    orphans = [{"文件": f"{no}.pdf", "大小": f"{p.stat().st_size / 1024:.1f}",
                "说明": "台账里查不到这个编号，不丢弃也不分发，等业务确认"}
               for no, p in originals.items() if no not in used]

    rows = ledger_rows(items)
    wb = fill_ledger(rows)[2]
    write_sheet(wb, "待处理", HOLD_HEADER,
                [[r["报告编号"], r["客户名称"], r["问题"], r["说明"]] for r in held], HOLD_WIDTH)
    write_sheet(wb, "无对应台账原件", ORPHAN_HEADER,
                [[o["文件"], o["大小"], o["说明"]] for o in orphans], ORPHAN_WIDTH)
    wb.save(LEDGER_XLSX)
    write_manifest(items)

    print(f"台账 {len(reports)} 行 → 分发件 {len(items)} 份 → 登记 {len(rows)} 行 → "
          f"挂起 {len(held)} 份 → 无对应台账原件 {len(orphans)} 个")

    checks = verify(reports, items, held, orphans)
    for name, good, detail in checks:
        print(f"  [{'OK ' if good else 'FAIL'}] {name}" + (f"   {detail}" if not good else ""))

    passed = all(c[1] for c in checks)
    print(f"\n回读校验 {sum(c[1] for c in checks)}/{len(checks)} 项通过，判定 "
          f"{'PASS' if passed else 'FAIL'}")
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
