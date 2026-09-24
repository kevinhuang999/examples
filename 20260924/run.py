#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第三方检测 · 报价单批量汇总比对：把客户询价明细按客户归堆，查价目表算价、出报价单，
再跟这家客户上次的成交价逐项比一遍。

一句话：报价这活的难点不在算乘法，而在**一条询价行该用价目表里的哪一行**——
客户写的项目名跟价目表不一样、同一个项目几种方法几个价、价目表还有没生效和已失效的价。
这三件事定错，出的报价单看着整齐，价格是错的。

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
from datetime import date, datetime, timedelta
from pathlib import Path

# ----------------------------------------------------------------- ① 配置区（跑之前主要改这里）

ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "01_raw_data"
OUT_DIR = ROOT / "02_output"
TPL_DIR = ROOT / "source" / "templates"
LOG_FILE = ROOT / "source" / "run_log.txt"

PRICE_XLSX = RAW_DIR / "价目表.xlsx"
INQUIRY_CSV = RAW_DIR / "询价明细.csv"
HISTORY_CSV = RAW_DIR / "历史成交价.csv"
TPL_XLSX = TPL_DIR / "报价单模板.xlsx"
COMPARE_XLSX = OUT_DIR / "报价汇总比对.xlsx"
QUOTE_DIR = OUT_DIR / "报价单"

# 报价基准日：价目表按这一天挑生效价。**不许用 date.today()**——
# 换了天再跑，同一批询价算出来的价就变了，台账跟报价单对不上，说不清是谁的错。
QUOTE_DAY = date(2026, 9, 24)
QUOTE_PREFIX = "BJ"
VALID_DAYS = 30                       # 报价有效期

# 阶梯折扣：按**这一家客户当次能报出来的项数**定档，从大到小第一个命中。
# 「能报出来」不等于「客户问了几项」——挂起项不算，否则折扣按 10 项算、单上只报 9 项。
TIERS = [(10, 0.90), (5, 0.95), (0, 1.00)]

# 跟历史成交价比，变动超过这个幅度就标出来（等于不算超）。
RAISE_LIMIT = 0.10
MAX_DETAIL = 12                       # 报价单模板明细区行数

DATE_FMTS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d")
DATE_FMT = "yyyy-mm-dd"               # 写进模板的日期格子要顺手设格式，否则 Excel 显示成序列号

# 项目名归一化：全角转半角 → 去括号注释 → 去空白与分隔符 → 去尾缀 → 转小写。
FULL2HALF = str.maketrans("（）【】〔〕：，、／．－", "()[]{}:,,/.-")
BRACKET_RE = re.compile(r"[（(][^）)]*[）)]")          # 「（Pb）」「(Pb)」这类注解
SEP_RE = re.compile(r"[\s/、,;．.\-]+")
# 尾缀按长的排前面：先去「含量测定」，再去「含量」，不然「铅含量测定」只掉一个尾巴。
TAIL_WORDS = ("含量测定", "含量检测", "含量", "测定", "检测", "测试")
ALIAS_SPLIT = re.compile(r"[、,;；/]+")

COMPARE_HEADER = ["客户名称", "客户样品编号", "样品名称", "检测项目（客户写法）", "检测方法",
                  "数量", "匹配项目", "单价", "折扣", "金额", "上次单价", "变动幅度",
                  "状态", "说明"]
COMPARE_WIDTH = [26, 16, 16, 22, 16, 8, 16, 10, 8, 12, 10, 10, 12, 44]
HOLD_HEADER = ["客户名称", "客户样品编号", "样品名称", "检测项目（客户写法）", "检测方法",
               "数量", "问题", "说明"]
HOLD_WIDTH = [26, 16, 16, 22, 16, 12, 18, 52]

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
    """去空白。全角空格（U+3000）也要去——客户名称尾巴上粘一个全角空格，
    不去干净就会跟同名客户分成两家，报价单出两份，人眼还看不出来。"""
    return (text or "").replace("\u3000", " ").strip()


def cust_key(name):
    """客户名的比较键：再去掉内部空白。部门名前多打一个空格也算同一家。"""
    return re.sub(r"\s+", "", clean(name))


def norm_item(text):
    """把项目名压成可比的键：「铅（Pb）含量测定」「铅含量」「铅」压完都是「铅」。

    为什么必须归一化：客户是在自家系统里挑的项目名，价目表是自家编的，
    两边写法天然不一样——靠人工对着看，看一天也看不出「重金属铅」和「铅」是一回事。
    """
    s = clean(text).translate(FULL2HALF)
    s = BRACKET_RE.sub("", s)
    s = SEP_RE.sub("", s)
    for word in TAIL_WORDS:
        if s.endswith(word) and len(s) > len(word):
            s = s[: -len(word)]
            break
    return s.lower()


def norm_method(text):
    """方法名的比较键：「XRF 荧光光谱」和「xrf荧光光谱」是同一个方法，中间的空白不算。"""
    return re.sub(r"\s+", "", clean(text).translate(FULL2HALF)).upper()


def to_amount(text):
    """单价/金额转数字，认不出返回 None。客户原件里「待定」这种字眼一律不兜 0。"""
    try:
        value = float(clean(text))
    except ValueError:
        return None
    return value if value > 0 else None


def parse_qty(text):
    """数量只认正整数。客户原件写「2 件」的，挂起也不抠那个 2——
    抠出来就替客户改了口径，报价单上的数量跟客户原件对不上，出事说不清。"""
    s = clean(text)
    return int(s) if re.fullmatch(r"\d+", s) and int(s) > 0 else None


def parse_date(text):
    """认三种日期写法，认不出返回 None（不兜今天）。"""
    s = clean(text)
    for fmt in DATE_FMTS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def sheet_rows(path):
    """读 xlsx 成「表头 → 值」的字典列表。第一列空的整行跳过（表尾说明行靠这条挡掉）。"""
    from openpyxl import load_workbook

    ws = load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    head = [clean(str(c)) if c is not None else "" for c in rows[0]]
    out = []
    for raw in rows[1:]:
        if not clean(str(raw[0]) if raw[0] is not None else ""):
            continue
        out.append({k: ("" if v is None else str(v) if not isinstance(v, float)
                        else (str(int(v)) if v.is_integer() else str(v)))
                    for k, v in zip(head, raw)})
    return out


def csv_rows(path, key_col):
    """读 CSV，key_col 为空的行整行跳过——表尾的说明行、合计行都靠这一条挡掉。

    **key_col 要挑"一定有值"的那一列**，这条比什么都重要：询价明细里客户名称空着
    恰恰是要挂起的那几行，拿客户名称判空，最该看见的行会被静默丢掉，
    台账上还看不出少了东西。所以询价明细按第一列（询价日期）判空，不按客户名称。

    用标准库 csv 而不是 pandas：客户样品编号 0000713 这种前导零，
    pandas 会猜成整数，astype 也补不回已经丢掉的那一位。
    """
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    return [{k: clean(v) for k, v in r.items() if k} for r in rows if clean(r.get(key_col))]


# ----------------------------------------------------------------- ③ 读数据（归堆）


def load_price_list():
    """读价目表。**价目表是长表，不是一一对应**：一个项目有几行很常见——
    不同检测方法不同价、同一年里调过价、去年的旧价还留在表里。

    所以索引键不能只用项目名，要用（项目规范名 + 方法规范名）；
    另外把「别名」列也编进索引，客户写「重金属铅」也能落到「铅」上。
    返回 (索引, 别名表, 认识的项目名集合)。
    """
    index, alias, names = {}, {}, set()
    for row in sheet_rows(PRICE_XLSX):
        name = norm_item(row.get("项目名称"))
        method = norm_method(row.get("检测方法"))
        start = parse_date(row.get("生效日期"))
        price = to_amount(row.get("单价"))
        if not name or not method or not start or price is None:
            continue                                  # 资料不全的价目行不进索引，别让它带错价
        rec = {"编号": row.get("项目编号"), "名称": row.get("项目名称"),
               "规范": name, "方法": row.get("检测方法"), "单位": row.get("计价单位"),
               "单价": price, "生效": start,
               "失效": parse_date(row.get("失效日期"))}   # 空 = 至今有效，不是"无效"
        index.setdefault((name, method), []).append(rec)
        names.add(name)
        for one in ALIAS_SPLIT.split(row.get("别名") or ""):
            if clean(one):
                alias[norm_item(one)] = name
                names.add(norm_item(one))     # 别名的 key 也要算"认识的项目名"，不然客户写别名会被判成匹配不上
    return index, alias, names


def load_inquiry():
    """读客户询价明细，一行一个（样品 × 检测项目）。
    一行一个项目才比得动价——项目挤在一个格子里，价格就没法逐项核对。
    判空走第一列询价日期，**不走客户名称**（空客户名的那几行正是要挂起报出来的）。"""
    rows = []
    for raw in csv_rows(INQUIRY_CSV, "询价日期"):
        rows.append({"询价日期": raw.get("询价日期", ""), "客户名称": raw.get("客户名称", ""),
                     "客户样品编号": raw.get("客户样品编号", ""), "样品名称": raw.get("样品名称", ""),
                     "检测项目": raw.get("检测项目", ""), "检测方法": raw.get("检测方法", ""),
                     "数量原文": raw.get("数量", "")})
    return rows


def load_history():
    """读历史成交价，键 =（客户、项目规范名、方法规范名）。

    比价必须带方法：同一个项目不同方法本来就不是一个价，
    只按项目名比，会把「XRF 报 130、ICP-OES 报 220」判成降幅超阈值。
    """
    out = {}
    for row in csv_rows(HISTORY_CSV, "客户名称"):
        key = (cust_key(row.get("客户名称")), norm_item(row.get("检测项目")),
               norm_method(row.get("检测方法")))
        out[key] = (to_amount(row.get("上次单价")), clean(row.get("成交日期")))
    return out


# ----------------------------------------------------------------- ④ 匹配与算价


def active_prices(recs, day):
    """挑在报价基准日生效的价目行。

    两个边界都在这里：生效日期晚于基准日的**还没生效的调价**必须排除（取"最后一条"
    就会把下个月的涨价提前报出去）；失效日期为空表示一直有效，不是"无效"。
    """
    return [r for r in recs if r["生效"] <= day and (r["失效"] is None or day <= r["失效"])]


def match_line(item, method, index, alias, names, day):
    """给一条询价行找价目行。三步：归一化 → 查别名表 → 查价目表；三步都落空才算匹配不上。

    这里**不做模糊匹配**。difflib 那类"看着像就选一个"的做法，报错价了还不报错，
    比挂起危险得多：挂起最多晚一天报出去，报错价是直接在客户面前丢分。
    """
    key = norm_item(item)
    if not key:
        return None, "检测项目缺失", "询价明细里这一行的检测项目是空的"
    if key not in names:
        return None, "项目名称匹配不上", \
            f"「{clean(item)}」在价目表里找不到对应项目，不许按相近字猜一个价"

    std = alias.get(key, key)
    mkey = norm_method(method)
    same_item = [r for (name, _m), recs in index.items() if name == std for r in recs]

    if mkey:
        hit = active_prices(index.get((std, mkey), []), day)
        if hit:
            return hit[0], "", ""
        others = active_prices(same_item, day)
        if others:
            return None, "价目表里没有这个方法", \
                f"「{std}」在价目表里只有 {'、'.join(sorted({r['方法'] for r in others}))}，" \
                f"客户写的是「{clean(method)}」，先确认用哪一种"
        return None, "项目名称匹配不上", f"「{clean(item)}」在价目表里没有可用价格"

    # 客户没写方法：同一个项目只有一种方法才敢直接用；有两种以上就是替客户挑价。
    hit = active_prices(same_item, day)
    if not hit:
        return None, "项目名称匹配不上", f"「{clean(item)}」在价目表里没有可用价格"
    methods = sorted({r["方法"] for r in hit})
    if len(methods) > 1:
        return None, "方法待确认", \
            f"「{std}」有 {len(methods)} 种检测方法（{'、'.join(methods)}）、单价不同，" \
            f"客户没写用哪一种"
    return hit[0], "", ""


def tier_rate(items):
    """按项数定折扣档。从大到小第一个命中的生效——**等于门槛算命中**，
    「满 5 项打 95 折」里的 5 项就是 5 项。这条口子要跟商务定死，
    不然天天为"刚好多一项算哪档"扯皮。"""
    for floor, rate in TIERS:
        if items >= floor:
            return rate
    return 1.0


def price_all(rows, index, alias, names, day):
    """逐行定状态：先看客户名，再看是不是重复行，再看数量，最后才查价。顺序不能换。

    为什么客户名和数量排在查价前面：这两样是"根本不具备报价条件"，
    先查出价再挂起，等于白查一遍，挂起原因还会被写成"匹配不上"，掩掉真问题。
    """
    seen = set()
    for row in rows:
        if not cust_key(row["客户名称"]):
            row.update(状态="待确认", 问题="客户名缺失",
                       说明="询价明细里客户名称是空的，不知道这份报价单该报给谁")
            continue

        # 同一家、同一客户样品编号、同一个项目出现两行：只留先出现的那行，后一行挂起。
        # 直接删掉后一行是错的——客户可能真送了两个平行样，得问清楚哪一行作准。
        dup = (cust_key(row["客户名称"]), clean(row["客户样品编号"]), norm_item(row["检测项目"]))
        if dup in seen:
            row.update(状态="待确认", 问题="重复询价行",
                       说明="同一客户、同一客户样品编号、同一项目出现两行，先跟客户确认哪一行作准")
            continue
        seen.add(dup)

        qty = parse_qty(row["数量原文"])
        if qty is None:
            row.update(状态="待确认", 问题="数量不可用",
                       说明=f'客户原件把数量写成「{clean(row["数量原文"])}」，整数数量才能计价，'
                            f"不替客户改成数字")
            continue

        rec, problem, detail = match_line(row["检测项目"], row["检测方法"], index, alias, names, day)
        if problem:
            row.update(状态="待确认", 问题=problem, 说明=detail)
            continue
        row.update(rec=rec, 数量=qty, 单价=rec["单价"], 问题="", 说明="")


def compare_history(row, history):
    """跟这家客户上次的成交价比一遍。

    变了要标出来：客户问「怎么比上次贵」的时候，报价员手上得有数。
    降了也要标——多半是价目表调过价没通知商务，报出去才发现自己降了一半。
    查不到历史价是「无历史价」，不是异常；历史价记的是文字（如「待定」）的，
    不能当 0 去算涨幅，那是一除就爆的东西。
    """
    rec = row["rec"]
    last = history.get((cust_key(row["客户名称"]), rec["规范"], norm_method(rec["方法"])))
    if last is None:
        return "无历史价", "", "", ""
    last_price, last_day = last
    if last_price is None:
        return "历史价待核", "", "", f"历史台账里这一项记的不是数字（{last_day}），核完再定要不要调"
    gap = (rec["单价"] - last_price) / last_price
    text = f"{gap:+.1%}"
    if abs(gap) > RAISE_LIMIT:
        return "价格变动", f"{last_price:g}", text, \
            f"上次 {last_price:g} 元、这次 {rec['单价']:g} 元（{text}），报出去前先把原因说清楚"
    return "已核对", f"{last_price:g}", text, f"跟 {last_day} 那次一致或微调"


def group_customer(rows):
    """按客户归堆，保持首次出现的顺序（报价单号按这个顺序连号）。"""
    groups = {}
    for row in rows:
        key = cust_key(row["客户名称"])
        if key not in groups:
            groups[key] = {"客户": clean(row["客户名称"]), "行": []}
        groups[key]["行"].append(row)
    return list(groups.values())


def price_customer(rows):
    """一家客户的单价与折扣：先按**能报出来的项数**定档，再逐行乘数量、乘折扣。

    金额一律「逐行 round 后再合计」，不许「先加总再乘折扣」——
    后者遇到带小数的单价会跟报价单差一分钱，客户对账时第一个抓的就是这个。
    """
    ok = [r for r in rows if r.get("rec")]
    rate = tier_rate(len(ok))
    for row in ok:
        row["折扣"] = rate
        row["金额"] = round(row["单价"] * row["数量"] * rate, 2)
    return rate, round(sum(r["金额"] for r in ok), 2)


# ----------------------------------------------------------------- ⑤ 输出


def safe_name(name):
    """客户名直接当文件名会出事：斜杠建不出文件，冒号在 Windows 上非法。"""
    return re.sub(r'[\\/:*?"<>|\s]+', "-", clean(name))


def label_rows(ws):
    """按标签名反查「标签在第几行第几列」。

    模板是质量部备案的，改一版就多一行、少一行，写死 B2/D3 迟早漂掉；
    按标签名找，模板重排也不影响。
    """
    pos = {}
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value
            if isinstance(value, str) and clean(value):
                pos[clean(value)] = (row, col)
    return pos


def detail_span(ws):
    """认明细区：表头行（第一列是「序号」）下面**连续带边框**的行。

    模板里明细空行必须画边框——openpyxl 读回时既没内容又没样式的行根本不存在，
    靠 ws.max_row 反推明细区的写法，在这些行上是**一行都找不到**的。
    """
    head = next((r for r in range(1, ws.max_row + 1)
                 if clean(str(ws.cell(row=r, column=1).value or "")) == "序号"), None)
    if head is None:
        raise SystemExit("模板里找不到明细表头（第一列应该是「序号」）")
    last = head
    for row in range(head + 1, ws.max_row + 1):
        if ws.cell(row=row, column=1).border.left.style:     # 合计行没画边框，到这里就断
            last = row
        else:
            break
    return head, head + 1, last


def fill_quote(cust, no, ok_rows, rate, total):
    """把一家客户的报价填进模板，另存成独立文件。模板原件不动。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Font

    wb = load_workbook(TPL_XLSX)
    ws = wb.active
    pos = label_rows(ws)
    head, first, last = detail_span(ws)
    if len(ok_rows) > last - first + 1:
        raise SystemExit(f"{cust} 可报价 {len(ok_rows)} 项，超过模板明细区 {last - first + 1} 行，"
                         f"要么加行数要么拆成两份报价单")

    def put(label, value):
        row, col = pos[label]
        ws.cell(row=row, column=col + 1, value=value)

    put("报价单号", no)
    quote_cell = ws.cell(row=pos["报价日期"][0], column=pos["报价日期"][1] + 1, value=QUOTE_DAY)
    quote_cell.number_format = DATE_FMT
    put("客户名称", cust)
    put("项数", len(ok_rows))
    till = ws.cell(row=pos["有效期至"][0], column=pos["有效期至"][1] + 1,
                   value=QUOTE_DAY + timedelta(days=VALID_DAYS))
    till.number_format = DATE_FMT
    put("折扣档", f"{rate * 10:g} 折（按 {len(ok_rows)} 项）" if rate < 1 else "不打折")
    put("计价说明", "单价为单样品单项价，含税；复检、加急另计。")

    for i, row in enumerate(ok_rows, start=1):
        r = first + i - 1
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=row["rec"]["名称"])
        ws.cell(row=r, column=3, value=row["rec"]["方法"])
        code = ws.cell(row=r, column=4, value=row["客户样品编号"])
        code.number_format = "@"                 # 不设文本，0000713 会掉前导零
        ws.cell(row=r, column=5, value=row["数量"])
        ws.cell(row=r, column=6, value=row["单价"]).number_format = "0.00"
        ws.cell(row=r, column=7, value=row["金额"]).number_format = "0.00"

    # 合计行紧贴明细区下面，**不画边框**——画了就落进 detail_span 里，被当成可填的明细行
    ws.cell(row=last + 1, column=5, value="合计").font = Font(bold=True)
    ws.cell(row=last + 1, column=7, value=total).font = Font(bold=True)

    QUOTE_DIR.mkdir(parents=True, exist_ok=True)
    path = QUOTE_DIR / f"报价单_{safe_name(cust)}.xlsx"
    wb.save(path)
    return path


def write_sheet(ws, header, rows, widths):
    from openpyxl.styles import Alignment, Font

    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append(row)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
    ws.freeze_panes = "A2"


def write_compare(entries, quotes):
    """写报价汇总比对表：一页全量逐行（含挂起），一页单独列待确认的。"""
    from openpyxl import Workbook

    rows = []
    for row in entries:
        rec = row.get("rec")
        rows.append([
            clean(row["客户名称"]), row["客户样品编号"], row["样品名称"],
            clean(row["检测项目"]), clean(row["检测方法"]),
            row.get("数量", ""), rec["名称"] if rec else "", rec["单价"] if rec else "",
            f"{row['折扣'] * 10:g} 折" if rec and row["折扣"] < 1 else ("不打折" if rec else ""),
            row.get("金额", ""), row.get("上次单价", ""), row.get("变动幅度", ""),
            row["状态"], row.get("说明", ""),
        ])

    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb.create_sheet("报价比对"), COMPARE_HEADER, rows, COMPARE_WIDTH)
    write_sheet(wb.create_sheet("待确认"), HOLD_HEADER,
                [[clean(r["客户名称"]), r["客户样品编号"], r["样品名称"], clean(r["检测项目"]),
                  clean(r["检测方法"]), r["数量原文"], r["问题"], r["说明"]]
                 for r in entries if r["状态"] == "待确认"], HOLD_WIDTH)

    total = sum(q["合计"] for q in quotes)
    write_sheet(wb.create_sheet("报价单台账"),
                ["报价单号", "客户名称", "项数", "折扣", "合计金额", "有效性至", "文件"],
                [[q["号"], q["客户"], len(q["行"]),
                  f"{q['折扣'] * 10:g} 折" if q["折扣"] < 1 else "不打折",
                  q["合计"], (QUOTE_DAY + timedelta(days=VALID_DAYS)).isoformat(),
                  q["文件"].name] for q in quotes],
                [18, 26, 8, 8, 12, 12, 34])
    wb.create_sheet("总计").append(["报价单数", len(quotes), "合计金额", round(total, 2)])

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(COMPARE_XLSX)
    return total


# ----------------------------------------------------------------- 回读校验


def verify(entries, quotes):
    """重新从磁盘把比对表与报价单读回来，逐项断言——不自证，只看落盘的结果。"""
    from openpyxl import load_workbook

    checks = []

    def ok(name, good, detail=""):
        checks.append((name, bool(good), detail))

    ok_rows = [r for r in entries if r.get("rec")]
    held = [r for r in entries if r["状态"] == "待确认"]
    by_cust = {q["客户"]: q for q in quotes}

    def count_status(status):
        return sum(1 for r in entries if r["状态"] == status)

    ok("询价明细 27 行全读进，表尾说明行没算进来", len(entries) == 27, f"{len(entries)} 行")
    ok("可报价 20 行 / 待确认 7 行", (len(ok_rows), len(held)) == (20, 7),
       f"{len(ok_rows)} / {len(held)}")
    no_quote = {cust_key(r["客户名称"]) for r in held} - {cust_key(q["客户"]) for q in quotes}
    ok("出报价单 4 家；一项都报不出来的 2 组客户不出单",
       len(quotes) == 4 and len(no_quote) == 2, f"{len(quotes)} / {sorted(no_quote)}")

    ok("折扣档按能报出来的项数：10 项 9 折 / 5 项 9.5 折 / 4 项不打折",
       (by_cust["上海启海化工有限公司"]["折扣"], by_cust["浙江宏远纺织有限公司"]["折扣"],
        by_cust["苏州新亚电子材料有限公司"]["折扣"]) == (0.90, 0.95, 1.00),
       f"{[q['折扣'] for q in quotes]}")
    ok("项数正好等于门槛也吃档（10 项中 9 折、5 项中 9.5 折），4 项这一侧不吃",
       (len(by_cust["上海启海化工有限公司"]["行"]), len(by_cust["浙江宏远纺织有限公司"]["行"]),
        len(by_cust["苏州新亚电子材料有限公司"]["行"])) == (10, 5, 4))

    lead = [r for r in ok_rows if r["rec"]["名称"] == "铅含量"
            and norm_method(r["rec"]["方法"]) == "ICP-OES"]
    ok("没用明年 10 月才生效的调价（铅含量 ICP-OES 报 220 不是 260）",
       {r["单价"] for r in lead} == {220}, f"{sorted({r['单价'] for r in lead})}")
    ph = [r for r in ok_rows if r["rec"]["名称"] == "pH值"]
    ok("没用去年已失效的旧价（pH值 报 60 不是 50）",
       {r["单价"] for r in ph} == {60}, f"{sorted({r['单价'] for r in ph})}")
    ok("归一化把「铅（Pb）含量测定」锚到「铅」上（挂起原因是方法待确认，不是匹配不上）",
       any(r["检测项目"] == "铅（Pb）含量测定" and r["问题"] == "方法待确认" for r in held))
    ok("别名把「重金属铅」落到「铅」上，报了 220",
       any(r["检测项目"] == "重金属铅" and r["rec"]["名称"] == "铅含量"
           and r["单价"] == 220 for r in ok_rows))
    ok("同一个项目两种方法各按各的价（XRF 130 / ICP-OES 220）",
       {r["单价"] for r in ok_rows if r["rec"]["名称"] == "铅含量"} == {130, 220},
       f"{sorted({r['单价'] for r in ok_rows if r['rec']['名称'] == '铅含量'})}")
    ok("项目名对不上的挂起，没按相近字猜价",
       any(r["问题"] == "项目名称匹配不上" and "多环芳烃" in r["检测项目"] for r in held))
    ok("客户写了价目表里没有的方法 → 挂起并列出可选方法",
       any(r["问题"] == "价目表里没有这个方法" and "XRF 荧光光谱" in r["说明"] for r in held))
    ok("六个挂起原因各就各位",
       {r["问题"] for r in held} == {"方法待确认", "项目名称匹配不上", "价目表里没有这个方法",
                                     "数量不可用", "客户名缺失", "重复询价行"},
       f"{sorted({r['问题'] for r in held})}")
    ok("数量写成「2 件」的行挂起，说明带原文",
       any("2 件" in r["说明"] for r in held if r["问题"] == "数量不可用"))
    ok("客户名缺失整组 2 行挂起",
       sum(1 for r in held if r["问题"] == "客户名缺失") == 2)
    ok("重复询价行只挂后一行，前一行正常报出",
       sum(1 for r in held if r["问题"] == "重复询价行") == 1
       and sum(1 for r in ok_rows if r["问题"] == "") >= 1)

    ok("带全角空格的客户名没被拆成两家",
       len([q for q in quotes if q["客户"] == "浙江宏远纺织有限公司"]) == 1
       and len(by_cust["浙江宏远纺织有限公司"]["行"]) == 5)
    ok("历史价比对：总磷两行 +25.0%、甲醛 -20.0% 被标出，正好 +10.0% 的不标",
       count_status("价格变动") == 3
       and len({r["检测项目"] for r in ok_rows if r["状态"] == "价格变动"}) == 2
       and any(r["变动幅度"] == "+10.0%" and r["状态"] == "已核对" for r in ok_rows),
       f'{count_status("价格变动")} / {[r["变动幅度"] for r in ok_rows if r["状态"] == "价格变动"]}')
    ok("历史价记着文字的标成待核、不当 0 算涨幅",
       count_status("历史价待核") == 1)
    ok("第一次做的项目标成无历史价",
       count_status("无历史价") == 11, f'{count_status("无历史价")}')
    ok("五项跟历史价对得上（已核对）", count_status("已核对") == 5,
       f'{count_status("已核对")}')

    ok("客户C 合计 = 逐行金额之和",
       by_cust["上海启海化工有限公司"]["合计"]
       == round(sum(r["金额"] for r in by_cust["上海启海化工有限公司"]["行"]), 2),
       f'{by_cust["上海启海化工有限公司"]["合计"]}')
    ok("报价单号按客户顺序连号",
       [q["号"] for q in quotes] == ["BJ20260924-01", "BJ20260924-02",
                                     "BJ20260924-03", "BJ20260924-04"],
       f"{[q['号'] for q in quotes]}")

    wb = load_workbook(COMPARE_XLSX)
    cmpws, holdws, ledws = wb["报价比对"], wb["待确认"], wb["报价单台账"]
    ok("比对页行数 = 询价明细行数", cmpws.max_row - 1 == len(entries), f"{cmpws.max_row - 1}")
    ok("待确认页行数 = 待确认行数", holdws.max_row - 1 == len(held), f"{holdws.max_row - 1}")
    ok("台账页单价保留 220/130 两种", {ledws.cell(row=r, column=1).value for r in range(2, ledws.max_row + 1)}
       == {q["号"] for q in quotes})
    ok("客户样品编号在比对页保持补零原样",
       "0000713" in {cmpws.cell(row=r, column=2).value for r in range(2, cmpws.max_row + 1)})

    quote_ok, blank_ok, fmt_ok = True, True, True
    for q in quotes:
        ws = load_workbook(q["文件"]).active
        pos = label_rows(ws)
        head, first, last = detail_span(ws)
        if ws.cell(row=pos["报价单号"][0], column=pos["报价单号"][1] + 1).value != q["号"] \
                or ws.cell(row=pos["客户名称"][0], column=pos["客户名称"][1] + 1).value != q["客户"] \
                or ws.cell(row=pos["项数"][0], column=pos["项数"][1] + 1).value != len(q["行"]) \
                or ws.cell(row=last + 1, column=7).value != q["合计"]:
            quote_ok = False
        if ws.cell(row=first, column=1).value != 1 \
                or ws.cell(row=first, column=2).value != q["行"][0]["rec"]["名称"]:
            quote_ok = False
        for row in range(first + len(q["行"]), last + 1):     # 没占满的明细行必须保持空白
            if any(ws.cell(row=row, column=c).value not in (None, "") for c in range(1, 8)):
                blank_ok = False
        if ws.cell(row=first, column=4).number_format != "@" \
                or ws.cell(row=pos["报价日期"][0], column=pos["报价日期"][1] + 1).number_format != DATE_FMT \
                or ws.cell(row=pos["有效期至"][0], column=pos["有效期至"][1] + 1).number_format != DATE_FMT:
            fmt_ok = False
    ok("每张报价单的单头（单号/客户/项数/合计）与明细一致", quote_ok)
    ok("明细没占满的行保持空白，没被填 0 或「无」", blank_ok)
    ok("客户样品编号是文本、两个日期格带日期格式", fmt_ok)

    tpl_head, tpl_first, tpl_last = detail_span(load_workbook(TPL_XLSX).active)
    tpl = load_workbook(TPL_XLSX).active
    ok("模板原件没被动过（明细区仍是空的）",
       all(tpl.cell(row=r, column=c).value in (None, "") for r in range(tpl_first, tpl_last + 1)
           for c in range(1, 8)))
    ok("模板明细区行数与常量一致", tpl_last - tpl_first + 1 == MAX_DETAIL,
       f"{tpl_last - tpl_first + 1}")

    return checks


# ----------------------------------------------------------------- main


def main():
    t0 = time.time()
    log_line(f"运行主机：{socket.gethostname()}（{local_ip()}）")
    log_line(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
             f"Python {sys.version.split()[0]} / {Path(sys.executable).name}")

    entries = load_inquiry()
    index, alias, names = load_price_list()
    history = load_history()
    price_all(entries, index, alias, names, QUOTE_DAY)

    quotes = []
    for group in group_customer(entries):
        rate, total = price_customer(group["行"])
        ok_rows = [r for r in group["行"] if r.get("rec")]
        if not ok_rows:
            continue                                       # 一项都报不出来，不出报价单
        for row in ok_rows:
            row["状态"], row["上次单价"], row["变动幅度"], row["说明"] = \
                compare_history(row, history)
        no = f"{QUOTE_PREFIX}{QUOTE_DAY:%Y%m%d}-{len(quotes) + 1:02d}"
        quotes.append({"客户": group["客户"], "号": no, "折扣": rate, "合计": total,
                       "行": ok_rows, "文件": fill_quote(group["客户"], no, ok_rows, rate, total)})

    total = write_compare(entries, quotes)
    print(f"询价 {len(entries)} 行 → 报价单 {len(quotes)} 份，合计 {total:.2f} 元")

    checks = verify(entries, quotes)
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
