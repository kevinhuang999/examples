# -*- coding: utf-8 -*-
"""样品编号批量生成：读历史编号台账 → 按规则表给本批样品续编编号 → 出编号清单与台账 Excel。

用法：
    python run.py            # 跑完暂停等回车（Windows 双击也能用）
    python run.py --no-pause # 跑完直接退出
"""
import csv
import datetime as dt
import io
import re
import socket
import sys
import time
from pathlib import Path


def ensure_deps():
    """缺依赖自动从清华源装，不挑解释器。"""
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl",
                               "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"])


ensure_deps()

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

# ---------------- 路径与字段常量区 ----------------
BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
TPL_DIR = SRC_DIR / "templates"
SAMPLES_FILE = RAW_DIR / "本批样品清单.csv"
LEDGER_FILE = RAW_DIR / "编号台账_已用编号.csv"
RULE_FILE = TPL_DIR / "编号规则表.csv"
LOG_FILE = SRC_DIR / "run_log.txt"

OUT_NAME = "样品编号清单.xlsx"

ID_COL = "编号"
NAME_COL = "样品名称"
TYPE_COL = "检测类型"
SRC_COL = "样品来源"
DATE_COL = "送样日期"
MADE_COL = "编号生成日"

# 编号形状：前缀-日期段(年月)-流水号，如 WS-202610-0004
ID_RE = re.compile(r"^([A-Za-z]+)-(\d{6})-(\d+)$")

ALIASES = {
    ID_COL: ("编号", "样品编号", "样品号"),
    NAME_COL: ("样品名称", "样品名"),
    TYPE_COL: ("检测类型", "检测类别", "检测项目"),
    SRC_COL: ("样品来源", "来源", "委托方"),
    DATE_COL: ("送样日期", "收样日期", "送检日期", "采样日期"),
    MADE_COL: ("编号生成日", "生成日期"),
}

FILL_PENDING = PatternFill("solid", fgColor="FFE699")
FILL_OVERFLOW = PatternFill("solid", fgColor="FFC7CE")


def now_host_ip():
    """运行主机与对外网卡 IP（用来分辨这台记录是哪台机器跑的）。"""
    host = socket.gethostname()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except OSError:
        try:
            ip = socket.gethostbyname(host)
        except OSError:
            ip = "未知"
    return host, ip


def log(msg):
    """追加写日志，不覆盖历史段。"""
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(msg + "\n")


def read_text(path):
    """读文本并返回实际用的编码：系统导出的 CSV 常在 UTF-8 和 GBK 之间摇摆。"""
    for enc in ("utf-8-sig", "gbk"):
        try:
            return path.read_text(encoding=enc), enc
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("csv", b"", 0, 1, f"{path.name}: 编码既不是 utf-8 也不是 gbk")


def find_header(rows):
    """找表头行：第一个「至少 2 个单元格命中已知字段名」的行。

    不能只看「非空单元格 >= 2」——导出信息那行正好也是两格（台账名 + 导出时间），
    会被当成表头，整张表从第一行就开始读错。按字段名认表头才认得住。
    """
    known = {n for names in ALIASES.values() for n in names}
    for i, r in enumerate(rows):
        if sum(1 for c in r if str(c or "").strip() in known) >= 2:
            return i
    return -1


def cell(cells, idx):
    if idx is None or idx >= len(cells):
        return ""
    return cells[idx]


def pick_columns(header):
    """按别名表把表头映射到标准字段；对不上的字段显式返回 None，由判空兜住。"""
    cols = {}
    norm_header = [str(h or "").strip() for h in header]
    for std, names in ALIASES.items():
        cols[std] = None
        for i, h in enumerate(norm_header):
            if h in names:
                cols[std] = i
                break
    return cols


def to_int(s):
    s = str(s or "").strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def parse_date(s):
    """把日期单元格转成 date；空返回 None；认不出的写法返回 None（交给待确认，不猜）。"""
    s = str(s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def load_rules(path):
    """读编号规则表：检测类型 → (前缀, 流水位数)。

    编号规则是管理口径，写在表里不写死在代码里——新开一种检测类型、
    或者流水位数要从 3 位改 4 位，加一行表就行，不用动脚本。
    """
    text, _ = read_text(path)
    rows = [r for r in csv.reader(io.StringIO(text)) if any(str(c).strip() for c in r)]
    rules = {}
    for r in rows[1:]:
        t = str(r[0]).strip()
        prefix = str(r[1]).strip().upper() if len(r) > 1 else ""
        width = to_int(r[2]) if len(r) > 2 else None
        if t and prefix and width:
            rules[t] = (prefix, width)
    if not rules:
        raise ValueError(f"{path.name}: 规则表是空的，没法编编号")
    return rules


def parse_id(s):
    """拆编号为 (前缀, 日期段, 流水号)；认不出的返回 None，交给待确认，不猜。"""
    m = ID_RE.match(str(s or "").strip())
    if not m:
        return None
    return m.group(1).upper(), m.group(2), int(m.group(3))


def load_ledger(path):
    """读历史编号台账，返回 (台账明细, 每组最大流水号, 历史重号清单)。

    流水号一律 int() 之后再比大小。台账在 Excel 里过过手，「9」和「0010」的排序
    跟人想的不一样；直接按字符串取 max 会取到 "9"，下一号又编成 0010，
    跟台账里已经有的 0010 撞个正着——这是重号最常见的来路。
    """
    text, enc = read_text(path)
    rows = list(csv.reader(io.StringIO(text)))
    hdr = find_header(rows)
    if hdr < 0:
        raise ValueError(f"{path.name}: 找不到表头行")
    cols = pick_columns(rows[hdr])
    if cols.get(ID_COL) is None:
        raise ValueError(f"{path.name}: 表头里找不到「编号」列")

    ledger, seq_max, dups, seen = [], {}, [], set()
    for r in rows[hdr + 1:]:
        cells = [str(c or "").strip() for c in r]
        if not any(cells):
            continue
        sid = cell(cells, cols.get(ID_COL))
        if not sid:
            continue
        row = {
            ID_COL: sid,
            NAME_COL: cell(cells, cols.get(NAME_COL)),
            TYPE_COL: cell(cells, cols.get(TYPE_COL)),
            MADE_COL: cell(cells, cols.get(MADE_COL)),
        }
        ledger.append(row)
        if sid in seen:
            dups.append(sid)
        seen.add(sid)
        parts = parse_id(sid)
        if parts is None:
            continue
        prefix, month, seq = parts
        key = (prefix, month)
        if seq > seq_max.get(key, 0):
            seq_max[key] = seq
    return ledger, seq_max, dups, enc


def load_samples(path):
    """读本批样品清单，按文件顺序返回，每行带一个原始行号（待确认清单要指回原表）。"""
    text, enc = read_text(path)
    rows = list(csv.reader(io.StringIO(text)))
    hdr = find_header(rows)
    if hdr < 0:
        raise ValueError(f"{path.name}: 找不到表头行")
    cols = pick_columns(rows[hdr])
    missing = [k for k in (NAME_COL, TYPE_COL, DATE_COL) if cols.get(k) is None]
    if missing:
        raise ValueError(f"{path.name}: 样品清单缺字段：{missing}")

    samples = []
    for i, r in enumerate(rows[hdr + 1:], start=hdr + 2):
        cells = [str(c or "").strip() for c in r]
        if not any(cells):
            continue
        samples.append({
            "行号": i,
            NAME_COL: cell(cells, cols.get(NAME_COL)),
            TYPE_COL: cell(cells, cols.get(TYPE_COL)),
            SRC_COL: cell(cells, cols.get(SRC_COL)),
            DATE_COL: parse_date(cell(cells, cols.get(DATE_COL))),
        })
    return samples, enc


def make_ids(samples, rules, seq_max, today):
    """给本批样品续编编号；编不了的四类进待确认，一条都不静默丢掉。

    日期段取样品自己的送样月份：补录上个月的样品要回到上个月那一段接着编，
    跟当月混在一起编，等于给自己埋重号。
    """
    new_rows, pending = [], []
    for s in samples:
        if not s[NAME_COL]:
            pending.append((s, "样品名称为空，编号没法挂靠"))
            continue
        if s[TYPE_COL] not in rules:
            pending.append((s, f"检测类型「{s[TYPE_COL]}」不在编号规则表里"))
            continue
        if s[DATE_COL] is None:
            pending.append((s, "送样日期为空或认不出，取不到编号日期段"))
            continue
        prefix, width = rules[s[TYPE_COL]]
        month = s[DATE_COL].strftime("%Y%m")
        key = (prefix, month)
        seq = seq_max.get(key, 0) + 1
        seq_max[key] = seq
        new_rows.append({
            ID_COL: f"{prefix}-{month}-{seq:0{width}d}",
            NAME_COL: s[NAME_COL],
            TYPE_COL: s[TYPE_COL],
            SRC_COL: s[SRC_COL],
            DATE_COL: s[DATE_COL].isoformat(),
            MADE_COL: today.isoformat(),
            "溢出": seq > 10 ** width - 1,
        })
    return new_rows, pending


def summarize(new_rows):
    """按「前4位（前缀）＋日期段」归堆计数，给汇总表用。"""
    counts = {}
    for r in new_rows:
        key = r[ID_COL].rsplit("-", 1)[0]
        counts[key] = counts.get(key, 0) + 1
    return counts


def write_outputs(new_rows, ledger, pending, dups, counts):
    """一次落三张表：新编号清单、编号台账（全量）、待确认清单。"""
    cols = [ID_COL, NAME_COL, TYPE_COL, SRC_COL, DATE_COL]
    wb = Workbook()

    ws = wb.active
    ws.title = "新编号清单"
    ws.append(cols)
    for r in new_rows:
        ws.append([r.get(c, "") for c in cols])
        if r["溢出"]:
            for j in range(1, len(cols) + 1):
                ws.cell(row=ws.max_row, column=j).fill = FILL_OVERFLOW
    for j in range(1, len(cols) + 1):
        ws.cell(row=1, column=j).font = Font(bold=True)
    for i in range(2, ws.max_row + 1):
        ws.cell(row=i, column=1).number_format = "@"    # 编号按文本存，前导零才不掉

    ws2 = wb.create_sheet("编号台账")
    ws2.append([ID_COL, NAME_COL, TYPE_COL, MADE_COL])
    for i, r in enumerate(ledger + [{ID_COL: x[ID_COL], NAME_COL: x[NAME_COL],
                                     TYPE_COL: x[TYPE_COL], MADE_COL: x[MADE_COL]}
                                    for x in new_rows], start=2):
        ws2.append([r.get(c, "") for c in (ID_COL, NAME_COL, TYPE_COL, MADE_COL)])
        ws2.cell(row=i, column=1).number_format = "@"   # 编号按文本存，前导零才不掉
    for j in range(1, 5):
        ws2.cell(row=1, column=j).font = Font(bold=True)

    ws3 = wb.create_sheet("待确认")
    ws3.append(["样品清单行号", NAME_COL, TYPE_COL, "问题"])
    for s, why in pending:
        ws3.append([s["行号"], s[NAME_COL], s[TYPE_COL], why])
    for sid in dups:
        ws3.append(["", "", "", f"历史台账里已有重号：{sid}，编号前先人工核一遍"])
    for j in range(1, 5):
        ws3.cell(row=1, column=j).font = Font(bold=True)
    for i in range(2, ws3.max_row + 1):
        if str(ws3.cell(row=i, column=4).value or "").strip():
            ws3.cell(row=i, column=4).fill = FILL_PENDING

    ws4 = wb.create_sheet("编号汇总")
    ws4.append(["生成日期", new_rows[0][MADE_COL] if new_rows else ""])
    for key in sorted(counts):
        ws4.append([key, counts[key]])
    ws4.append(["本批合计", len(new_rows)])
    ws4.append(["历史台账条数", len(ledger)])
    ws4.cell(row=1, column=1).font = Font(bold=True)

    out = OUT_DIR / OUT_NAME
    wb.save(out)
    return out


def verify(out, new_rows, ledger, pending, dups, counts):
    """生成后回读断言：数据故意造了边界值（丢前导零、位数溢出、缺字段），跑通不等于对。

    最后两条是真回读：把刚存的 Excel 打开一次，确认编号是按文本存的——
    内存里对没用，文件里存成数字，前导零在客户打开那一刻就掉了。
    """
    new_ids = [r[ID_COL] for r in new_rows]
    old_ids = {r[ID_COL] for r in ledger}
    by_prefix = {}
    for r in new_rows:
        by_prefix.setdefault(r[ID_COL].rsplit("-", 1)[0], []).append(r[ID_COL])

    def first(key):
        return sorted(by_prefix.get(key, []))[:1]

    oks = []
    oks.append(("本批编号 16 条，与可编号样品数一致", len(new_ids) == 16))
    oks.append(("编号之间没有重号", len(new_ids) == len(set(new_ids))))
    oks.append(("新编号与历史台账无交集", not (set(new_ids) & old_ids)))
    oks.append(("水质当月从 0004 起编（上月 0007 不参与）",
                first("WS-202610") == ["WS-202610-0004"]))
    oks.append(("补录上月的样品编回上月段 WS-202609-0008",
                "WS-202609-0008" in set(new_ids)))
    oks.append(("台账里丢前导零的 TR-202610-9 按 9 认，本批从 0011 起（按字符比会撞 0010）",
                first("TR-202610") == ["TR-202610-0011"]))
    oks.append(("位数溢出不重号：FCM 编到 1000 仍唯一",
                "FCM-202610-1000" in set(new_ids)))
    oks.append(("待确认 4 条，每条都带了原因", len(pending) == 4 and all(w for _, w in pending)))
    oks.append(("历史台账重号 WS-202610-0003 被报出", "WS-202610-0003" in dups))

    wb = load_workbook(out)
    first_cell = wb["新编号清单"].cell(row=2, column=1)
    oks.append(("回读：编号存成文本且前导零在位",
                first_cell.number_format == "@" and first_cell.value == new_ids[0]))
    ledger_col = [wb["编号台账"].cell(row=i, column=1).value
                  for i in range(2, 2 + len(ledger) + len(new_rows))]
    oks.append(("回读：台账里 12 条历史 + 16 条新编号，全部按文本读回",
                len(ledger_col) == len(ledger) + len(new_rows)
                and all(isinstance(v, str) for v in ledger_col)))
    ok = all(ok for _, ok in oks)
    for name, passed in oks:
        print(f"  [{'OK' if passed else 'FAIL'}] {name}")
    return ok


def main():
    t0 = time.time()
    host, ip = now_host_ip()
    log(f"")
    log(f"运行主机：{host}（{ip}）")
    log(f"开始执行：{dt.datetime.now().isoformat(timespec='seconds')}  "
        f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）")

    OUT_DIR.mkdir(exist_ok=True)
    today = dt.date.today()
    rules = load_rules(RULE_FILE)
    ledger, seq_max, dups, led_enc = load_ledger(LEDGER_FILE)
    samples, smp_enc = load_samples(SAMPLES_FILE)
    print(f"规则表 {len(rules)} 种检测类型；历史台账 {led_enc} 编码 {len(ledger)} 条；"
          f"本批样品 {smp_enc} 编码 {len(samples)} 条")

    new_rows, pending = make_ids(samples, rules, seq_max, today)
    counts = summarize(new_rows)
    print(f"本批编出 {len(new_rows)} 个编号，待确认 {len(pending)} 条：")
    for key in sorted(counts):
        ids = [r[ID_COL] for r in new_rows if r[ID_COL].rsplit("-", 1)[0] == key]
        print(f"  {key}  {counts[key]} 个：{ids[0]} … {ids[-1]}")

    over = [r[ID_COL] for r in new_rows if r["溢出"]]
    if over:
        print(f"  [注意] 流水号已超出规则位数，编号变长但没重号：{over}")
    for sid in dups:
        print(f"  [注意] 历史台账里已有重号：{sid}，建议先人工核一遍")

    out = write_outputs(new_rows, ledger, pending, dups, counts)
    print(f"已生成：{out.name}（新编号清单 / 编号台账 / 待确认 / 编号汇总 四张表）")
    ok = verify(out, new_rows, ledger, pending, dups, counts)
    print("总判定：" + ("PASS" if ok else "FAIL"))

    log(f"执行完成：{dt.datetime.now().isoformat(timespec='seconds')}  "
        f"用时 {time.time() - t0:.1f}s  判定 {'PASS' if ok else 'FAIL'}")

    if "--no-pause" not in sys.argv:
        input("\n按回车退出……")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
