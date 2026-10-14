# -*- coding: utf-8 -*-
"""检验记录批量电子化：按台账把受控模板复制成每份电子检验记录，回填仪器导出结果并逐项判定，
   另出记录完整性与编号连续性核查清单。

跑法：
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

from openpyxl import Workbook, load_workbook          # noqa: E402
from openpyxl.styles import Font, PatternFill         # noqa: E402

# ---------------- 路径与字段常量区 ----------------
BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
TPL_DIR = SRC_DIR / "templates"
RECORD_FILE = RAW_DIR / "检验记录台账.csv"
RESULT_FILE = RAW_DIR / "仪器导出_结果明细.csv"
ITEM_FILE = TPL_DIR / "检项清单.csv"
TPL_FILE = TPL_DIR / "检验记录模板.xlsx"
LOG_FILE = SRC_DIR / "run_log.txt"

BOOK_NAME = "检验记录册.xlsx"
CHECK_NAME = "记录核查清单.xlsx"
TPL_SHEET = "模板"

REC_NO = "记录编号"
PRODUCT_COL = "品种"
LOT_COL = "批号"
ANALYST_COL = "检验人"
REVIEWER_COL = "复核人"
DATE_COL = "检验日期"
ITEM_COL = "检验项目"
VALUE_COL = "结果"
UNIT_COL = "单位"
SRC_TYPE_COL = "数据来源"
LOW_COL = "下限"
HIGH_COL = "上限"
NOTE_COL = "说明"
SPEC_COL = "标准规定"
JUDGE_COL = "单项判定"

MACHINE = "仪器"      # 仪器出的结果，脚本负责填
MANUAL = "人工"       # 检查项，结果留空等检验员填

# 判定只用这五个值：判得了的、判不了的、以及填不了的两种
OK, OVER, BY_METHOD, TO_FILL, NO_DATA = "符合", "超限", "见方法", "待填", "缺结果"

DETAIL_HEAD = ("序号", ITEM_COL, SPEC_COL, "检验结果", UNIT_COL, JUDGE_COL)
HEAD_LABELS = (REC_NO, PRODUCT_COL, DATE_COL, LOT_COL, ANALYST_COL, REVIEWER_COL,
               "检验人签名", "复核人签名")

# 记录编号形状：前缀-年月-流水号，如 QC-202610-0001
REC_RE = re.compile(r"^([A-Za-z]+)-(\d{6})-(\d+)$")
NUM_RE = re.compile(r"^-?\d+(\.\d+)?$")
LIMIT_RE = re.compile(r"^([<>≤≥])\s*(-?\d+(?:\.\d+)?)$")

ALIASES = {
    REC_NO: ("记录编号", "记录号"),
    PRODUCT_COL: ("品种", "品名", "产品名称"),
    LOT_COL: ("批号", "生产批号", "批号/编号"),
    ANALYST_COL: ("检验人", "检验员", "检测人"),
    REVIEWER_COL: ("复核人", "复核", "审核人"),
    DATE_COL: ("检验日期", "检测日期", "日期"),
    ITEM_COL: ("检验项目", "检测项目", "项目"),
    VALUE_COL: ("结果", "检验结果", "检测结果"),
    UNIT_COL: ("单位", "计量单位"),
    SRC_TYPE_COL: ("数据来源", "来源"),
    LOW_COL: ("下限", "标准下限"),
    HIGH_COL: ("上限", "标准上限"),
    NOTE_COL: ("说明", "备注"),
}

FILL_OVER = PatternFill("solid", fgColor="FFC7CE")
FILL_TODO = PatternFill("solid", fgColor="FFE699")
FILL_HEAD = PatternFill("solid", fgColor="DDEBF7")


def now_host_ip():
    """运行主机与对外网卡 IP（用来分辨这段记录是哪台机器跑的）。"""
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

    仪器结果文件导出时表头前面压着两行导出信息，按「非空单元格 >= 2」认表头
    会认到那两行上去（导出名 + 导出时间正好也是两格），整张表从第一行就开始读错。
    按字段名认表头才认得住。
    """
    known = {n for names in ALIASES.values() for n in names}
    for i, r in enumerate(rows):
        if sum(1 for c in r if str(c or "").strip() in known) >= 2:
            return i
    return -1


def read_rows(path):
    """读 CSV，跳过空行，并定位表头；返回 (表头, 数据行, 实际编码)。"""
    text, enc = read_text(path)
    rows = [r for r in csv.reader(io.StringIO(text)) if any(str(c).strip() for c in r)]
    hi = find_header(rows)
    if hi < 0:
        raise ValueError(f"{path.name}: 找不到表头行")
    header = [str(c).strip() for c in rows[hi]]
    body = [[str(c).strip() for c in r] for r in rows[hi + 1:]]
    return header, body, enc


def cell(cells, idx):
    if idx is None or idx >= len(cells):
        return ""
    return cells[idx]


def pick_columns(header, keys):
    """按别名表把列名定下来，返回「字段 → 列下标」；缺列时值为 None，调用处判空。"""
    idx = {}
    for key in keys:
        idx[key] = None
        for name in ALIASES.get(key, (key,)):
            if name in header:
                idx[key] = header.index(name)
                break
    return idx


def to_num(s):
    """把「结果」里的纯数值取出来；取不到返回 None。

    限号值（`<0.05`）、文字结论（`未检出`）、空值都走 None 这条路——
    它们不是「0」，也不是合格，后面会分开处理。
    """
    t = str(s or "").strip().replace("%", "").replace("％", "")
    return float(t) if NUM_RE.match(t) else None


def parse_limit(s):
    """限号值 `<0.05` / `>100`：返回 (符号, 数)；不是限号值返回 None。"""
    m = LIMIT_RE.match(str(s or "").strip())
    return (m.group(1), float(m.group(2))) if m else None


def load_items(path):
    """读检项清单：品种 → 该品种该做哪些项、限度、该项的数据来源。

    「该检哪些项」是品种属性，写在表里不写死在代码里——新增品种、改限度，
    加一行表就行，不用动脚本。限度原文连着数值一起留着，记录上照样显示 93.0%。
    """
    header, body, _ = read_rows(path)
    idx = pick_columns(header, (PRODUCT_COL, ITEM_COL, SRC_TYPE_COL,
                                LOW_COL, HIGH_COL, UNIT_COL, NOTE_COL))
    items = {}
    for r in body:
        product = cell(r, idx[PRODUCT_COL])
        name = cell(r, idx[ITEM_COL])
        if not product or not name:
            continue
        items.setdefault(product, []).append({
            ITEM_COL: name,
            SRC_TYPE_COL: cell(r, idx[SRC_TYPE_COL]) or MACHINE,
            "low_t": cell(r, idx[LOW_COL]),
            "high_t": cell(r, idx[HIGH_COL]),
            LOW_COL: to_num(cell(r, idx[LOW_COL])),
            HIGH_COL: to_num(cell(r, idx[HIGH_COL])),
            UNIT_COL: cell(r, idx[UNIT_COL]),
            NOTE_COL: cell(r, idx[NOTE_COL]),
        })
    if not items:
        raise ValueError(f"{path.name}: 检项清单是空的，没法生成记录")
    return items


def spec_text(it):
    """记录里「标准规定」一列的文字：范围 / ≥下限 / ≤上限 / 人工项直接写方法规定。"""
    lo, hi = it["low_t"], it["high_t"]
    if lo and hi:
        return f"{lo}~{hi}"
    if lo:
        return f"≥{lo}"
    if hi:
        return f"≤{hi}"
    return it[NOTE_COL] or "按方法规定"


def judge_result(value, it):
    """按结果值的形态判定，返回 (判定, 说明)。

    只有纯数值才去跟限度比；限号值按报告值的量级判，量级和限度谁大说不准就不硬判；
    文字结论原样照抄、判「见方法」——把「未检出」当成合格是自欺。
    """
    low, high = it[LOW_COL], it[HIGH_COL]
    if value is None or str(value).strip() == "":
        return NO_DATA, "仪器导出里这项是空的，先核是没跑还是没导出来"
    v = to_num(value)
    if v is not None:
        if low is not None and v < low:
            return OVER, f"低于下限 {it['low_t']}"
        if high is not None and v > high:
            return OVER, f"高于上限 {it['high_t']}"
        return OK, ""
    lim = parse_limit(value)
    if lim:
        sym, n = lim
        if sym in ("<", "≤"):
            if high is None:
                return OK, f"报告值 {value}，上限未规定"
            if n <= high:
                return OK, ""
            return BY_METHOD, f"报告值 {value} 与上限 {high} 谁大说不准，要看方法"
        if low is None:
            return OK, f"报告值 {value}，下限未规定"
        if n >= low:
            return OK, ""
        return BY_METHOD, f"报告值 {value} 与下限 {low} 谁大说不准，要看方法"
    return BY_METHOD, f"文字结论「{value}」，按方法规定判"


def norm_lot(lot):
    """批号归一：去掉前导零再比。

    仪器导出的文件被 Excel 打开过一遭，`0B1207` 会被吃成 `B1207`；
    不归一会整批配不上，判定成「仪器结果缺失」——那是假报警，比漏报还费人。
    """
    return str(lot or "").strip().lstrip("0").upper() or "0"


def load_results(path):
    """读仪器导出的结果明细：按 (归一批号, 检验项目) 归堆，返回 (归堆表, 重复项, 编码)。"""
    header, body, enc = read_rows(path)
    idx = pick_columns(header, (LOT_COL, ITEM_COL, VALUE_COL, UNIT_COL))
    grouped, dups = {}, []
    for r in body:
        lot, item = cell(r, idx[LOT_COL]), cell(r, idx[ITEM_COL])
        if not lot or not item:
            continue
        key = (norm_lot(lot), item)
        row = {LOT_COL: lot, ITEM_COL: item, VALUE_COL: cell(r, idx[VALUE_COL]),
               UNIT_COL: cell(r, idx[UNIT_COL])}
        if key in grouped:
            old = grouped[key]
            if str(old[VALUE_COL]).strip() == str(row[VALUE_COL]).strip():
                dups.append((lot, item, "同批同项重复导出，按一条算"))
            else:
                dups.append((lot, item,
                             f"同批同项两条结果不一致（{old[VALUE_COL]} / {row[VALUE_COL]}），"
                             f"已按第一条填，发出前必须核"))
            continue
        grouped[key] = row
    return grouped, dups, enc


def load_records(path):
    """读检验记录台账：本批要出哪些记录。"""
    header, body, enc = read_rows(path)
    idx = pick_columns(header, (REC_NO, PRODUCT_COL, LOT_COL,
                                ANALYST_COL, REVIEWER_COL, DATE_COL))
    recs = []
    for r in body:
        recs.append({k: cell(r, idx[k]) for k in
                     (REC_NO, PRODUCT_COL, LOT_COL, ANALYST_COL, REVIEWER_COL, DATE_COL)})
    return recs, enc


def make_pages(records, items, grouped, dups):
    """把台账、检项清单、仪器结果三份数据拼成「每份记录该有哪些行」。

    这里是整件事的中枢：明细行以「检项清单」为准，不是以仪器结果为准——
    以结果为准的话，漏做、漏导出的项目会整行消失，记录看上去是满的，其实是缺的。
    """
    dup_note = {(norm_lot(l), i): w for l, i, w in dups}
    pages, held, extras, notes = [], [], [], []
    noted = set()
    for rec in records:
        product = rec[PRODUCT_COL]
        if product not in items:
            held.append((rec, f"品种「{product}」不在检项清单里，不知道该检哪些项，不出记录页"))
            continue
        lot_key = norm_lot(rec[LOT_COL])
        rows, problems = [], []
        for it in items[product]:
            res = grouped.get((lot_key, it[ITEM_COL]))
            if it[SRC_TYPE_COL] == MANUAL:
                rows.append({ITEM_COL: it[ITEM_COL], SPEC_COL: spec_text(it),
                             VALUE_COL: "", UNIT_COL: it[UNIT_COL],
                             JUDGE_COL: TO_FILL, NOTE_COL: "检查项，等检验员填写"})
                continue
            if res is None:
                rows.append({ITEM_COL: it[ITEM_COL], SPEC_COL: spec_text(it),
                             VALUE_COL: "", UNIT_COL: it[UNIT_COL],
                             JUDGE_COL: NO_DATA, NOTE_COL: "仪器导出里没有这一项"})
                problems.append(f"{it[ITEM_COL]}：仪器结果缺失")
                continue
            if res[LOT_COL] != rec[LOT_COL] and (rec[LOT_COL], res[LOT_COL]) not in noted:
                noted.add((rec[LOT_COL], res[LOT_COL]))
                notes.append(("批号写法不一致",
                              f"台账写 {rec[LOT_COL]}、仪器文件写 {res[LOT_COL]}，"
                              f"已按去掉前导零匹配，建议两边统一"))
            judge, why = judge_result(res[VALUE_COL], it)
            if (lot_key, it[ITEM_COL]) in dup_note:
                why = (why + "；" if why else "") + dup_note[(lot_key, it[ITEM_COL])]
            rows.append({ITEM_COL: it[ITEM_COL], SPEC_COL: spec_text(it),
                         VALUE_COL: res[VALUE_COL], UNIT_COL: it[UNIT_COL] or res[UNIT_COL],
                         JUDGE_COL: judge, NOTE_COL: why})
            if judge == OVER:
                problems.append(f"{it[ITEM_COL]}：{why}")
            elif judge == NO_DATA:
                problems.append(f"{it[ITEM_COL]}：仪器结果是空的")
        for (lot, item), res in grouped.items():
            if lot == lot_key and item not in [i[ITEM_COL] for i in items[product]]:
                extras.append((rec[REC_NO], rec[LOT_COL], item, res[VALUE_COL]))
        page = dict(rec)
        page["rows"] = rows
        page["problems"] = problems
        page["miss"] = sum(1 for r in rows if r[JUDGE_COL] == NO_DATA)
        page["over"] = sum(1 for r in rows if r[JUDGE_COL] == OVER)
        page["todo"] = sum(1 for r in rows if r[JUDGE_COL] == TO_FILL)
        pages.append(page)
    return pages, held, extras, notes


def check_recnos(records):
    """记录编号连续性：同一「前缀-年月」段内要连号。

    受控记录本断号意味着有一页不知去向；重号更麻烦——两份记录挂同一个号，
    出了事谁也说不清哪份是哪份。这两类都不能静默过去。
    """
    seen, groups = {}, {}
    for rec in records:
        no = rec[REC_NO]
        seen[no] = seen.get(no, 0) + 1
        m = REC_RE.match(no or "")
        if m:
            groups.setdefault(f"{m.group(1)}-{m.group(2)}", []).append(int(m.group(3)))
    dup_nos = sorted(no for no, c in seen.items() if c > 1)
    gaps = [f"{seg}-{n:04d}" for seg, seqs in groups.items()
            for n in range(min(seqs), max(seqs) + 1) if n not in seqs]
    return gaps, dup_nos


def find_label_cells(ws, labels):
    """按标签文字找右边的值格——模板挪了位置也不会填错地方。"""
    found = {}
    for row in ws.iter_rows():
        for c in row:
            t = str(c.value or "").strip().rstrip("：:")
            if t in labels and t not in found:
                found[t] = (c.row, c.column + 1)
    return found


def find_detail_header(ws):
    """明细表头行＝第一列写着「序号」的那行；明细数据从它下面一行开始。"""
    for row in ws.iter_rows(min_col=1, max_col=1):
        if str(row[0].value or "").strip() == "序号":
            return row[0].row
    raise ValueError("模板里找不到明细表头（第一列写着「序号」的那一行）")


def copy_sheet(wb, tpl, title):
    """复制模板页，并把打印区域补回来。

    实测：copy_worksheet 会带走单元格、样式、列宽和页面方向，但**不带打印区域**
    （打印区域是挂在工作表上的定义名，复制时丢了）。模板排好的「一页一份记录」
    全靠它，不补的话打印出来会跟着内容乱跑。
    """
    ws = wb.copy_worksheet(tpl)
    ws.title = title
    if tpl.print_area:
        ws.print_area = str(tpl.print_area).split("!", 1)[-1]
    return ws


def fill_page(ws, page):
    """往复制出来的记录页里写值：表头信息区按标签找格子，明细区从表头下面一行往下填。"""
    labels = find_label_cells(ws, HEAD_LABELS)
    for key in (REC_NO, PRODUCT_COL, DATE_COL, LOT_COL, ANALYST_COL, REVIEWER_COL):
        r, c = labels[key]
        ws.cell(row=r, column=c).value = page[key]
        ws.cell(row=r, column=c).number_format = "@"    # 批号按文本写，前导零才留得住
    sig_analyst = labels["检验人签名"]
    sig_reviewer = labels["复核人签名"]
    ws.cell(*sig_analyst).value = page[ANALYST_COL] or "（未签）"
    ws.cell(*sig_reviewer).value = page[REVIEWER_COL] or "（未签）"

    head_row = find_detail_header(ws)
    cols = {str(ws.cell(row=head_row, column=c).value or "").strip(): c for c in range(1, 8)}
    first = head_row + 1
    last = first
    while str(ws.cell(row=last, column=1).value or "").strip() == "" and last < head_row + 30:
        last += 1
    if len(page["rows"]) > last - first:
        raise ValueError(f"{page[REC_NO]}：明细 {len(page['rows'])} 行，"
                         f"模板只预置了 {last - first} 行，请先在模板里加行"
                         f"（insert_rows 会把签名区格式顶乱，别在脚本里插）")
    for i, row in enumerate(page["rows"]):
        rr = first + i
        ws.cell(row=rr, column=cols["序号"]).value = i + 1
        ws.cell(row=rr, column=cols[ITEM_COL]).value = row[ITEM_COL]
        ws.cell(row=rr, column=cols[SPEC_COL]).value = row[SPEC_COL]
        ws.cell(row=rr, column=cols["检验结果"]).value = row[VALUE_COL]
        ws.cell(row=rr, column=cols[UNIT_COL]).value = row[UNIT_COL]
        ws.cell(row=rr, column=cols[JUDGE_COL]).value = row[JUDGE_COL]
        if row[JUDGE_COL] == OVER:
            ws.cell(row=rr, column=cols[JUDGE_COL]).fill = FILL_OVER
        elif row[JUDGE_COL] in (NO_DATA, TO_FILL):
            ws.cell(row=rr, column=cols[JUDGE_COL]).fill = FILL_TODO


def write_book(pages):
    """复制模板生成记录册：一页一份记录，sheet 名就是记录编号。

    sheet 名有两道坎：不能重复（重号的记录得加后缀），长度也不能超 31 字。
    """
    wb = load_workbook(TPL_FILE)
    tpl = wb[TPL_SHEET]
    used = {}
    for page in pages:
        title = page[REC_NO]
        used[title] = used.get(title, 0) + 1
        if used[title] > 1:
            title = f"{title}({used[title]})"
        fill_page(copy_sheet(wb, tpl, title[:31]), page)
    del wb[TPL_SHEET]
    out = OUT_DIR / BOOK_NAME
    wb.save(out)
    return out


def write_check(pages, held, gaps, dup_nos, extras, notes, dups):
    """出核查清单：四张表——记录一览 / 缺项明细 / 超限明细 / 编号与来源异常。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "记录一览"
    cols = [REC_NO, PRODUCT_COL, LOT_COL, ANALYST_COL, REVIEWER_COL, DATE_COL,
            "明细项数", "缺结果", "超限", "待填", "结论"]
    ws.append(cols)
    for page in pages:
        issues = []
        if page["miss"]:
            issues.append(f"{page['miss']} 项缺结果")
        if page["over"]:
            issues.append(f"{page['over']} 项超限")
        if not page[ANALYST_COL]:
            issues.append("检验人未填")
        if not page[REVIEWER_COL]:
            issues.append("复核人未签")
        ws.append([page[REC_NO], page[PRODUCT_COL], page[LOT_COL], page[ANALYST_COL],
                   page[REVIEWER_COL], page[DATE_COL], len(page["rows"]),
                   page["miss"], page["over"], page["todo"],
                   "完整" if not issues else "待处理（" + "；".join(issues) + "）"])
        if issues:
            for j in range(1, len(cols) + 1):
                ws.cell(row=ws.max_row, column=j).fill = FILL_TODO
    for j in range(1, len(cols) + 1):
        ws.cell(row=1, column=j).font = Font(bold=True)
        ws.cell(row=1, column=j).fill = FILL_HEAD

    ws2 = wb.create_sheet("缺项明细")
    ws2.append([REC_NO, LOT_COL, ITEM_COL, "说明"])
    for page in pages:
        for row in page["rows"]:
            if row[JUDGE_COL] == NO_DATA:
                ws2.append([page[REC_NO], page[LOT_COL], row[ITEM_COL], row[NOTE_COL]])
    for j in range(1, 5):
        ws2.cell(row=1, column=j).font = Font(bold=True)

    ws3 = wb.create_sheet("超限明细")
    ws3.append([REC_NO, LOT_COL, ITEM_COL, SPEC_COL, "检验结果", UNIT_COL, "说明"])
    for page in pages:
        for row in page["rows"]:
            if row[JUDGE_COL] == OVER:
                ws3.append([page[REC_NO], page[LOT_COL], row[ITEM_COL], row[SPEC_COL],
                            row[VALUE_COL], row[UNIT_COL], row[NOTE_COL]])
    for j in range(1, 8):
        ws3.cell(row=1, column=j).font = Font(bold=True)

    ws4 = wb.create_sheet("编号与来源异常")
    ws4.append(["类型", "内容"])
    for gap in gaps:
        ws4.append(["记录编号断号", f"{gap}：这一段里没有这份记录，先查那页去哪了"])
    for no in dup_nos:
        ws4.append(["记录编号重号", f"{no}：台账里出现多次，两份记录挂了同一个号，必须先核归属"])
    for lot, item, why in dups:
        ws4.append(["重复/不一致结果", f"{lot} / {item}：{why}"])
    for rec_no, lot, item, value in extras:
        ws4.append(["清单外的仪器结果", f"{rec_no}（{lot}）/ {item} = {value}："
                                      f"检项清单里没这一项，没往记录里塞，先问清该不该检"])
    for kind, msg in notes:
        ws4.append([kind, msg])
    for rec, why in held:
        ws4.append(["记录未生成", f"{rec[REC_NO]}（{rec[PRODUCT_COL]} / {rec[LOT_COL]}）：{why}"])
    for j in range(1, 3):
        ws4.cell(row=1, column=j).font = Font(bold=True)
    for i in range(2, ws4.max_row + 1):
        ws4.cell(row=i, column=1).fill = FILL_TODO

    out = OUT_DIR / CHECK_NAME
    wb.save(out)
    return out


def verify(out_book, out_check, pages, held, gaps, dup_nos, extras, notes, dups):
    """生成后回读断言：三份数据都故意造了边界值，跑通不等于对。

    最后几条是真回读：把刚存的记录册打开一次，确认批号是按文本存的、
    模板的标题与打印区域还在——内存里对没用，文件里丢了就是丢了。
    """
    all_rows = [r for p in pages for r in p["rows"]]
    miss_pages = [p[REC_NO] for p in pages if p["miss"]]
    over_pages = [p[REC_NO] for p in pages if p["over"]]
    b0231A = next(p for p in pages if p[LOT_COL] == "0231A")
    b0231C = next(p for p in pages if p[LOT_COL] == "0231C")
    b1207 = next(p for p in pages if p[LOT_COL] == "0B1207")
    b0231B = next(p for p in pages if p[LOT_COL] == "0231B")

    def row_of(page, item):
        return next(r for r in page["rows"] if r[ITEM_COL] == item)

    oks = []
    oks.append(("生成记录页 10 张（台账 11 行，1 行品种不在检项清单挂起）",
                len(pages) == 10 and len(held) == 1))
    oks.append(("批号前导零被仪器侧吃掉仍能配上：0B1207 的含量测定填了 95.8",
                row_of(b1207, "含量测定")[VALUE_COL] == "95.8"
                and row_of(b1207, "含量测定")[JUDGE_COL] == OK))
    oks.append(("同值重复导出只算一条：0231C 溶出度填一次、判符合",
                row_of(b0231C, "溶出度")[VALUE_COL] == "88"
                and row_of(b0231C, "溶出度")[JUDGE_COL] == OK))
    oks.append(("限号值 <0.05 与上限 0.5 比：判符合，不当缺项",
                row_of(b0231B, "水分")[JUDGE_COL] == OK))
    oks.append(("文字结论「未检出」原样填入并判「见方法」，不计超限",
                row_of(b1207, "有关物质")[VALUE_COL] == "未检出"
                and row_of(b1207, "有关物质")[JUDGE_COL] == BY_METHOD))
    oks.append(("超限 4 项，落在 2 份记录上", sum(p["over"] for p in pages) == 4
                and sorted(over_pages) == ["QC-202610-0002", "QC-202610-0003"]))
    oks.append(("缺结果 6 项，落在 4 份记录上（含 0231F 整批 3 项）",
                sum(p["miss"] for p in pages) == 6 and len(miss_pages) == 4))
    oks.append(("人工项「性状」留空判「待填」，10 页各 1 项",
                b0231A["todo"] == 1 and row_of(b0231A, "性状")[JUDGE_COL] == TO_FILL
                and sum(1 for r in all_rows if r[JUDGE_COL] == TO_FILL) == 10))
    oks.append(("漏签 3 份：0002 与 0005 缺复核人、0007 缺检验人",
                sorted(p[REC_NO] for p in pages
                       if not p[ANALYST_COL] or not p[REVIEWER_COL]) ==
                ["QC-202610-0002", "QC-202610-0005", "QC-202610-0007"]))
    oks.append(("编号断号 0004、重号 0003 都报出来",
                gaps == ["QC-202610-0004"] and dup_nos == ["QC-202610-0003"]))
    oks.append(("清单外的仪器结果 1 条报出，且没塞进记录",
                len(extras) == 1 and extras[0][2] == "干燥失重"
                and all(r[ITEM_COL] != "干燥失重" for r in all_rows)))
    oks.append(("重复/不一致结果各报 1 条",
                len(dups) == 2 and any("不一致" in w for _, _, w in dups)))
    oks.append(("批号写法不一致记了 1 条", len(notes) == 1))

    wb = load_workbook(out_book)
    ws = wb["QC-202610-0001"]
    lot_cell = ws.cell(*find_label_cells(ws, (LOT_COL,))[LOT_COL])
    oks.append(("回读：批号按文本存且前导零在位",
                lot_cell.number_format == "@" and lot_cell.value == "0231A"))
    oks.append(("回读：模板标题还在，打印区域补回来了",
                str(ws["A1"].value) == "检验记录表" and bool(ws.print_area)))
    oks.append(("回读：记录册 10 页，重号那页带 (2) 后缀",
                len(wb.sheetnames) == 10 and "QC-202610-0003(2)" in wb.sheetnames))
    wbc = load_workbook(out_check)
    oks.append(("回读：核查清单四张表都在",
                set(wbc.sheetnames) == {"记录一览", "缺项明细", "超限明细", "编号与来源异常"}))
    ok = all(passed for _, passed in oks)
    for name, passed in oks:
        print(f"  [{'OK' if passed else 'FAIL'}] {name}")
    return ok


def main():
    t0 = time.time()
    host, ip = now_host_ip()
    log("")
    log(f"运行主机：{host}（{ip}）")
    log(f"开始执行：{dt.datetime.now().isoformat(timespec='seconds')}  "
        f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）")

    OUT_DIR.mkdir(exist_ok=True)
    records, rec_enc = load_records(RECORD_FILE)
    items = load_items(ITEM_FILE)
    grouped, dups, res_enc = load_results(RESULT_FILE)
    print(f"台账 {rec_enc} 编码 {len(records)} 份记录；检项清单 {len(items)} 个品种；"
          f"仪器结果 {res_enc} 编码归堆 {len(grouped)} 条（重复/不一致 {len(dups)} 条）")

    pages, held, extras, notes = make_pages(records, items, grouped, dups)
    gaps, dup_nos = check_recnos(records)
    print(f"生成记录页 {len(pages)} 张；缺结果 {sum(p['miss'] for p in pages)} 项"
          f"（{sum(1 for p in pages if p['miss'])} 份）、超限 {sum(p['over'] for p in pages)} 项"
          f"（{sum(1 for p in pages if p['over'])} 份）")
    for page in pages:
        if page["problems"]:
            print(f"  [注意] {page[REC_NO]}（{page[LOT_COL]}）：{'；'.join(page['problems'])}")
    for gap in gaps:
        print(f"  [注意] 记录编号断号：{gap}")
    for no in dup_nos:
        print(f"  [注意] 记录编号重号：{no}")
    for rec, why in held:
        print(f"  [挂起] {rec[REC_NO]}：{why}")

    out_book = write_book(pages)
    out_check = write_check(pages, held, gaps, dup_nos, extras, notes, dups)
    print(f"已生成：{out_book.name}（{len(pages)} 页记录）+ {out_check.name}（四张核查表）")
    ok = verify(out_book, out_check, pages, held, gaps, dup_nos, extras, notes, dups)
    print("总判定：" + ("PASS" if ok else "FAIL"))

    log(f"执行完成：{dt.datetime.now().isoformat(timespec='seconds')}  "
        f"用时 {time.time() - t0:.1f}s  判定 {'PASS' if ok else 'FAIL'}")

    if "--no-pause" not in sys.argv:
        input("\n按回车退出……")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
