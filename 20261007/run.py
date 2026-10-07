# -*- coding: utf-8 -*-
"""仪器导出的检测明细 → 按列规则清洗、格式转换（每处改动留痕）。

跑法：
    python run.py              # 双击也行
    python run.py --no-pause   # 跑完不等回车

做什么：
    读 01_raw_data/ 下的仪器导出文件，按 RULES 里"一列一组规则"清洗：去零宽字符、去空白、
    全角转半角、项目名与单位别名归一、日期归一，再把结果列拆成数值与状态。
    每改一处就往 02_output/清洗记录.csv 记一条「原值 → 新值 → 用了哪条规则」；认不出的行
    整行进 02_output/挂起清单.csv，不猜、不改。清洗后数据落 02_output/清洗后数据.csv。

工程外壳（正文里不逐行讲）：只用标准库，无需装包；跑完自动回读断言打印 PASS/FAIL；
日志追加写 source/run_log.txt，每段开头记运行主机（主机名 + 本机 IP）。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import re
import socket
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "01_raw_data"
OUT_DIR = ROOT / "02_output"
SRC_DIR = ROOT / "source"
LOG_PATH = SRC_DIR / "run_log.txt"
CLEAN_PATH = OUT_DIR / "清洗后数据.csv"
CHANGE_PATH = OUT_DIR / "清洗记录.csv"
PENDING_PATH = OUT_DIR / "挂起清单.csv"
LEDGER_PATH = OUT_DIR / "清洗台账.csv"

# ------------------------------------------------- ① 配置：格式表 + 清洗规则表

# 一种导出格式一条记录。后缀相同的两种格式也能分开——靠「必需列」认，不靠文件名。
FORMATS = [
    {
        "名称": "仪器导出_逗号分隔",
        "后缀": (".csv",),
        "编码": ("utf-8-sig", "gbk"),
        "分隔符": ",",
        "必需列": ("样品编号", "客户名称", "检测项目", "结果", "单位", "检测日期"),
        "判空列": "检测项目",
        "列映射": {
            "样品编号": "样品编号",
            "客户名称": "客户名称",
            "检测项目": "检测项目",
            "结果": "结果",
            "单位": "单位",
            "检测日期": "检测日期",
        },
    },
    {
        "名称": "仪器导出_制表符分隔",
        "后缀": (".txt",),
        "编码": ("gbk", "utf-8-sig"),
        "分隔符": "\t",
        "必需列": ("样品编号", "客户名称", "检测项目", "结果", "单位", "检测日期"),
        "判空列": "检测项目",
        "列映射": {
            "样品编号": "样品编号",
            "客户名称": "客户名称",
            "检测项目": "检测项目",
            "结果": "结果",
            "单位": "单位",
            "检测日期": "检测日期",
        },
    },
]

# 清洗规则：一列一条，动作按数组顺序执行。加字段只加一行，不动清洗逻辑。
# 这一层只改"写法"，不改"含义"——所以没有任何推断性动作（换算、补值、去重）在表里。
RULES = [
    {"列": "样品编号", "动作": ["去零宽", "去空白", "全角转半角"], "必填": True},
    {"列": "客户名称", "动作": ["去零宽", "去空白", "全角转半角"], "必填": True},
    {"列": "检测项目", "动作": ["去零宽", "去空白", "全角转半角", "别名归一"], "必填": True,
     "映射": {"含量": "含量(%)", "含量%": "含量(%)", "有关物质": "有关物质(%)"}},
    {"列": "结果", "动作": ["去零宽", "去空白", "全角转半角"]},
    {"列": "单位", "动作": ["去零宽", "去空白", "全角转半角", "别名归一"],
     "映射": {"mg/l": "mg/L", "MG/L": "mg/L", "ug/ml": "μg/mL"}},
    {"列": "检测日期", "动作": ["去零宽", "去空白", "全角转半角", "日期归一"], "必填": True},
]

# 结果列能当"结论"用的文字只有这几个。表外的文字一律挂起，不许当结论塞进数据。
DECISIONS = {"未检出", "ND"}

STATUS_NUM, STATUS_LIMIT, STATUS_TEXT, STATUS_EMPTY = "数值", "低于下限", "文字结论", "空"

# 数值 + 可选单位：12.3mg/L、<0.01、98.50 都能拆。带符号的一律按"低于下限"处理。
NUM_UNIT = re.compile(r"^(?P<sign>[<>≤≥]?=?)(?P<num>-?\d+(?:\.\d+)?)(?P<unit>[^\d]*)$")

CLEAN_COLS = ["样品编号", "客户名称", "检测项目", "结果", "结果数值", "结果状态",
              "单位", "检测日期", "来源文件", "行号"]
CHANGE_COLS = ["来源文件", "行号", "列", "动作", "原值", "新值"]
PENDING_COLS = ["来源文件", "行号", "挂起原因"]
LEDGER_COLS = ["来源文件", "读出行数", "输出行数", "留痕条数", "挂起条数", "判定"]

# ---------------------------------------------------------------- ② 读文件

def read_text(path: Path, encodings) -> str:
    """按候选编码依次试。GBK 文件用 utf-8 读会直接抛 UnicodeDecodeError，不是乱码。"""
    last = None
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError as exc:
            last = exc
    raise last


def pick_format(path: Path):
    """先看后缀，再用表头特征列确认——同样后缀、表头不同的两种格式就靠这一步分开。"""
    lines = None
    for fmt in FORMATS:
        if path.suffix.lower() not in fmt["后缀"]:
            continue
        if lines is None:
            lines = read_text(path, fmt["编码"]).splitlines()
        for raw in lines[:10]:
            cells = [c.strip() for c in raw.split(fmt["分隔符"])]
            if all(col in cells for col in fmt["必需列"]):
                return fmt, lines
    return None, None


def read_export(path: Path, fmt, lines) -> list[dict]:
    """切出表头行之后的数据行，按判空列丢掉表尾说明行。"""
    header_at = None
    for i, raw in enumerate(lines):
        cells = [c.strip() for c in raw.split(fmt["分隔符"])]
        if all(col in cells for col in fmt["必需列"]):
            header_at = i
            break
    if header_at is None:
        return []
    rows = list(csv.reader(lines[header_at:], delimiter=fmt["分隔符"]))
    header = [c.strip() for c in rows[0]]
    out = []
    for line_no, cells in enumerate(rows[1:], start=header_at + 2):
        rec = {h: (cells[i] if i < len(cells) else "") for i, h in enumerate(header)}
        if not rec.get(fmt["判空列"], "").strip():
            continue                      # 表尾说明行 / 空行：判空列一空就到头
        item = {dst: rec.get(src, "") for dst, src in fmt["列映射"].items()}
        item["_来源文件"] = path.name
        item["_行号"] = line_no
        out.append(item)
    return out


# ---------------------------------------------------------------- ③ 清洗

def drop_zero_width(value: str) -> str:
    """去掉零宽字符——从系统或网页里复制出来的值很常见，肉眼完全看不见。"""
    return re.sub(r"[\u200b\u200c\u200d\ufeff]", "", str(value))


def strip_ws(value: str) -> str:
    """去掉所有空白，含全角空格 U+3000 与不间断空格 U+00A0（strip() 都不管）。"""
    return re.sub(r"[\s\u00a0\u3000]+", "", str(value))


def to_halfwidth(value: str) -> str:
    """全角转半角：只动 0xFF01-0xFF5E（！到～）与全角空格，其他一个字符都不碰。

    不图省事用 unicodedata.normalize("NFKC")：它顺手会把 µ(U+00B5) 换成 μ(U+03BC)、
    把 Ⅲ 换成 III，单位名一旦被悄悄改掉，两批数据就合不到一起了。
    """
    out = []
    for ch in str(value):
        code = ord(ch)
        if code == 0x3000:
            out.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out)


def parse_date(raw: str):
    """2026-10-06 / 2026/10/06 / 2026.10.6 / 20261006 都认；认不出返回 None，不兜今天。"""
    text = str(raw).strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, pattern).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


CELL_ACTIONS = {"去零宽": drop_zero_width, "去空白": strip_ws, "全角转半角": to_halfwidth}


def clean_cell(value, rule: dict, changes: list, line_no) -> tuple[str, str]:
    """按规则里的动作顺序清洗一格，每改一处记一条留痕。返回 (清洗后的值, 挂起原因)。"""
    cur = str(value if value is not None else "")
    col = rule["列"]
    for action in rule["动作"]:
        if action == "别名归一":
            new = rule.get("映射", {}).get(cur, cur)      # 映射表里没有的一个字都不动
            if new != cur:
                changes.append((line_no, col, action, cur, new))
                cur = new
        elif action == "日期归一":
            new = parse_date(cur)
            if new is None:
                return cur, f"检测日期认不出：{cur}"
            if new != cur:
                changes.append((line_no, col, action, cur, new))
                cur = new
        else:
            new = CELL_ACTIONS[action](cur)
            if new != cur:
                changes.append((line_no, col, action, cur, new))
            cur = new
    return cur, ""


def split_result(text: str):
    """结果原值 → (数值, 单位, 状态, 挂起原因)。四态：数值 / 低于下限 / 文字结论 / 空。"""
    if not text:
        return None, "", STATUS_EMPTY, ""
    if "," in text:
        return None, "", "", "结果含逗号，分不清千分位还是分隔符"
    m = NUM_UNIT.match(text)
    if m:
        unit = m.group("unit").strip()
        if m.group("sign"):
            return None, unit, STATUS_LIMIT, ""
        return float(m.group("num")), unit, STATUS_NUM, ""
    if text in DECISIONS:
        return None, "", STATUS_TEXT, ""
    return None, "", "", f"结果认不出，且不在受控词表内：{text}"


def clean_row(row: dict) -> tuple[dict | None, list, str]:
    """清洗一行：先过规则表，再拆结果列。任何一列认不出就整行挂起——不改，也不记留痕。"""
    line_no = row.get("_行号")
    changes: list = []
    out = {"来源文件": row.get("_来源文件", ""), "行号": line_no}
    for rule in RULES:
        col = rule["列"]
        value, reason = clean_cell(row.get(col, ""), rule, changes, line_no)
        if reason:
            return None, changes, reason
        if rule.get("必填") and not value:
            return None, changes, f"{col}为空"
        out[col] = value
    value, unit, status, reason = split_result(out["结果"])
    if reason:
        return None, changes, reason
    out["结果数值"], out["结果状态"] = value, status
    out["单位"] = unit or out["单位"]      # 结果里带了单位就用它，没带才用单位列
    return out, changes, ""


# ---------------------------------------------------------------- ④ 输出

def run_all():
    """扫 01_raw_data/ 全部文件：读 → 清洗 → 归堆。返回 (明细, 留痕, 挂起, 台账)。"""
    clean_rows, changes, pending, ledger = [], [], [], []
    for path in sorted(RAW_DIR.iterdir()):
        if not path.is_file():
            continue
        fmt, lines = pick_format(path)
        if fmt is None:                       # 认不出的文件照样登记，不安静跳过
            pending.append({"来源文件": path.name, "行号": "", "挂起原因": "文件格式认不出"})
            ledger.append({"来源文件": path.name, "读出行数": 0, "输出行数": 0,
                           "留痕条数": 0, "挂起条数": 1, "判定": "挂起"})
            continue
        read_n = out_n = chg_n = pend_n = 0
        for row in read_export(path, fmt, lines):
            read_n += 1
            cleaned, cell_changes, reason = clean_row(row)
            if cleaned is None:
                pend_n += 1
                pending.append({"来源文件": path.name, "行号": row.get("_行号"),
                                "挂起原因": reason})
                continue                      # 整行作废：前面已改好的格子也一起作废
            out_n += 1
            clean_rows.append(cleaned)
            for line_no, col, action, old, new in cell_changes:
                changes.append({"来源文件": path.name, "行号": line_no, "列": col,
                                "动作": action, "原值": old, "新值": new})
            chg_n += len(cell_changes)
        ledger.append({"来源文件": path.name, "读出行数": read_n, "输出行数": out_n,
                       "留痕条数": chg_n, "挂起条数": pend_n,
                       "判定": "PASS" if out_n else "挂起"})
    return clean_rows, changes, pending, ledger


def write_csv(path: Path, header: list, rows: list) -> None:
    """一律 utf-8-sig：Excel 双击打开不乱码。"""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        for row in rows:
            w.writerow(["" if row.get(c) is None else row.get(c) for c in header])


def digest(rows: list) -> str:
    """把明细摊成一行行文本算指纹，用来证明同一份输入跑两遍结果一字不差。"""
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=CLEAN_COLS, extrasaction="ignore", lineterminator="\n")
    for row in rows:
        w.writerow({c: ("" if row.get(c) is None else row.get(c)) for c in CLEAN_COLS})
    return hashlib.sha1(buf.getvalue().encode("utf-8")).hexdigest()


# ---------------------------------------------------------------- 回读校验

def verify(clean_rows, changes, pending, ledger) -> list:
    checks = []

    def add(name, ok, detail=""):
        checks.append((name, bool(ok), detail))

    add("清洗后明细 14 行", len(clean_rows) == 14, f"实际 {len(clean_rows)}")
    add("留痕 17 条", len(changes) == 17, f"实际 {len(changes)}")
    add("挂起 5 条", len(pending) == 5, f"实际 {len(pending)}")
    add("台账 5 行（一个文件一行）", len(ledger) == 5, f"实际 {len(ledger)}")

    items = [r["检测项目"] for r in clean_rows]
    add("项目别名归一：含量(%) 共 3 行", items.count("含量(%)") == 3, f"实际 {items.count('含量(%)')}")
    add("项目别名归一：有关物质(%) 共 2 行", items.count("有关物质(%)") == 2)
    add("明细里不再有裸写法的项目名", "含量" not in items and "有关物质" not in items)

    ids = [r["样品编号"] for r in clean_rows]
    add("前导零编号保住了（0000081）", "0000081" in ids)
    add("编号没被剥成 81", "81" not in ids and "86" not in ids)
    add("零宽字符被去掉（0000086）", "0000086" in ids)

    r98 = [r for r in clean_rows if r["结果"] == "98.50"]
    add("全角数字 ９８.５０ 转成 98.50（两行同值，数值都是 98.5）",
        len(r98) == 2 and all(r["结果数值"] == 98.5 for r in r98), f"实际 {len(r98)} 行")
    half = [r for r in clean_rows if r["结果"] == "0.025"]
    add("半全角混写的 0.0２５ 转成 0.025（数值=0.025）",
        len(half) == 1 and half[0]["结果数值"] == 0.025)
    add("结果里混的单位被拆出来（12.3 mg/L → 值 12.3、单位 mg/L）",
        any(r["结果"] == "12.3mg/L" and r["结果数值"] == 12.3 and r["单位"] == "mg/L"
            for r in clean_rows))
    add("单位别名归一：mg/l 与 MG/L 都成了 mg/L",
        any(r["单位"] == "mg/L" for r in clean_rows)
        and not any(r["单位"] in ("mg/l", "MG/L") for r in clean_rows))

    dates = {r["检测日期"] for r in clean_rows}
    add("斜杠、点号、全角横线三种日期都归一成 2026-10-06", "2026-10-06" in dates)
    add("日期列只剩 YYYY-MM-DD 一种写法",
        all(re.fullmatch(r"\d{4}-\d{2}-\d{2}", d) for d in dates), f"实际 {sorted(dates)}")

    statuses = {r["结果状态"] for r in clean_rows}
    add("四态都出现：数值 / 低于下限 / 文字结论 / 空",
        statuses == {STATUS_NUM, STATUS_LIMIT, STATUS_TEXT, STATUS_EMPTY}, f"实际 {sorted(statuses)}")
    add("<0.01 数值列为空、状态=低于下限",
        any(r["结果状态"] == STATUS_LIMIT and r["结果数值"] is None for r in clean_rows))
    add("文字结论不进数值列",
        all(r["结果数值"] is None for r in clean_rows if r["结果状态"] == STATUS_TEXT))
    add("原值列保留输入的位数（98.50 没被修成 98.5）",
        any(r["结果"] == "98.50" for r in clean_rows))

    reasons = {p["挂起原因"] for p in pending}
    add("五类挂起原因各一条",
        len(reasons) == 5 and "文件格式认不出" in reasons
        and "结果含逗号，分不清千分位还是分隔符" in reasons,
        f"实际 {sorted(reasons)}")
    add("挂起行没进明细（3,250 / 待检 / 待定 那三行）",
        not any(r["结果"] in ("3,250", "待检") for r in clean_rows)
        and not any(r["检测日期"] == "待定" for r in clean_rows))
    add("空客户名称那行也挂起了", any("客户名称为空" in p["挂起原因"] for p in pending))

    acts = {a: sum(1 for c in changes if c["动作"] == a) for a in
            ("去零宽", "去空白", "全角转半角", "别名归一", "日期归一")}
    add("留痕按动作分布：去零宽 1 / 去空白 5 / 全角转半角 4 / 别名归一 5 / 日期归一 2",
        acts == {"去零宽": 1, "去空白": 5, "全角转半角": 4, "别名归一": 5, "日期归一": 2},
        f"实际 {acts}")
    add("留痕每条都写得清原值和新值",
        all(c["原值"] != c["新值"] and c["列"] and c["行号"] for c in changes))
    add("同一份输入跑两遍，明细指纹一致（清洗是纯函数）",
        digest(clean_rows) == digest(run_all()[0]))

    for path in (CLEAN_PATH, CHANGE_PATH, PENDING_PATH, LEDGER_PATH):
        add(f"{path.name} 已落盘且非空", path.exists() and path.stat().st_size > 0)
    return checks


# ---------------------------------------------------------------- 日志与入口

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def local_ip() -> str:
    """取本机对外网卡地址，用来分辨是哪台机器跑的（不发包，只为让系统选出口网卡）。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(("8.8.8.8", 80))
        return sock.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "未知"
    finally:
        sock.close()


def log(line: str) -> None:
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main() -> int:
    started = time.time()
    parser = argparse.ArgumentParser()
    parser.add_argument("--no-pause", action="store_true")
    args = parser.parse_args()

    if not RAW_DIR.exists():
        print(f"找不到输入目录：{RAW_DIR}")
        return 2
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    SRC_DIR.mkdir(parents=True, exist_ok=True)

    log(f"运行主机：{socket.gethostname()}（{local_ip()}）")
    log(f"开始执行：{now_str()}  Python {sys.version.split()[0]}  解释器 {Path(sys.executable).name}")

    clean_rows, changes, pending, ledger = run_all()
    write_csv(CLEAN_PATH, CLEAN_COLS, clean_rows)
    write_csv(CHANGE_PATH, CHANGE_COLS, changes)
    write_csv(PENDING_PATH, PENDING_COLS, pending)
    write_csv(LEDGER_PATH, LEDGER_COLS, ledger)

    print(f"输入：{RAW_DIR.name}/ 共 {len(ledger)} 个文件")
    for row in ledger:
        print(f"  [{row['判定']}] {row['来源文件']}：读出 {row['读出行数']} 行 → "
              f"输出 {row['输出行数']} 行、留痕 {row['留痕条数']} 条、挂起 {row['挂起条数']} 条")
    print(f"\n清洗后数据 {len(clean_rows)} 行 / 清洗记录 {len(changes)} 条 / 挂起 {len(pending)} 条\n")

    checks = verify(clean_rows, changes, pending, ledger)
    for name, ok, detail in checks:
        print(f"[{'OK ' if ok else 'NG '}] {name}" + (f"  → {detail}" if detail and not ok else ""))
    bad = [c for c in checks if not c[1]]
    verdict = "PASS" if not bad else "FAIL"
    print(f"\n回读校验：{len(checks) - len(bad)}/{len(checks)} 项通过  总判定 {verdict}")

    log(f"执行完成：{now_str()}  用时 {time.time() - started:.1f}s  判定 {verdict}")
    if not args.no_pause:
        input("按回车结束……")
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
