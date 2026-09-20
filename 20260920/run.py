# -*- coding: utf-8 -*-
"""引流长文《中心实验室报告模板多中心格式不统一，一份份手工重排实在做不动》配套脚本。

一份数据、三套模板，把仪器导出的检验结果按**每个中心自己的模板**填成报告：

    仪器导出（一行一个 检项结果）
  + 中心模板配置（这个中心用哪个模板、什么单位制、缺项怎么写）
  + 各中心自己的 xlsx 模板（字段位置、字段集合、标签写法都不一样）
→ 一个样本一份报告，填进对应中心的模板里，格式和图章位置一个不动。

核心代码与正文逐字一致，末尾多了自动回读校验，方便确认每一份都填对了。
跨平台：Windows / macOS / Linux 都是一条命令，缺依赖会自动装。
    python run.py
    python run.py --no-pause   # 不暂停（CI 或脚本里用）

打包给别人时，把整个目录拷走即可，路径全部相对定位，无硬编码。
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
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
RESULT_FILE = HERE / "01_raw_data" / "检验结果_仪器导出.csv"    # 一行一个 检项结果
CONFIG_FILE = HERE / "01_raw_data" / "中心模板配置.csv"          # 一行一个中心
TPL_DIR = HERE / "source" / "templates"                          # 各中心自己的模板
OUT_DIR = HERE / "02_output"                                     # 报告 + 台账
LOG_DIR = HERE / "source"                                        # 运行日志

RESULT_COLS = ["样本编号", "受试者编号", "中心编号", "访视点", "检项",
               "结果", "单位", "参考下限", "参考上限", "检测日期"]
CONFIG_COLS = ["中心编号", "中心名称", "模板文件", "单位制", "缺项写法", "报告编号前缀"]

LEDGER_COLS = ["中心编号", "中心名称", "样本编号", "状态", "检项数", "字段命中", "换算项数",
               "模板文件", "说明"]
STRAY_COLS = ["中心编号", "样本编号", "受试者编号", "检项", "结果", "说明"]

# 模板里的标签写法千奇百怪，这里给出「标准字段名 → 模板上可能怎么写」的对照，
# 归一化之后按这张表反查；新增一个中心只要在配置与模板里出现同一套标签，不用改代码。
FIELD_ALIASES = {
    "样本编号": ["样本编号", "样本ID"],
    "受试者编号": ["受试者编号", "受试者ID"],
    "中心编号": ["中心编号", "中心代码"],
    "中心名称": ["中心名称", "中心"],
    "访视点": ["访视点", "访视"],
    "检测日期": ["检测日期", "采样日期"],
    "报告编号": ["报告编号"],
}
ALIAS_TO_STD = {alias: std for std, aliases in FIELD_ALIASES.items() for alias in aliases}

# 国际单位 → 常规单位。换算系数挂在这里、不挂在数据上：同一份结果，
# A 中心要 mmol/L、B 中心要 mg/dL，是模板的属性，不是数据的属性。
UNIT_CONVERSIONS = {
    "葡萄糖": ("mmol/L", "mg/dL", 18.0182),
    "总胆固醇": ("mmol/L", "mg/dL", 38.67),
    "甘油三酯": ("mmol/L", "mg/dL", 88.57),
}

R_MAKE = "出报告"
R_HOLD = "挂起"
MAX_DETAIL_ROWS = 200               # 模板明细区最多认多少行，防止一张空表被认成几万行


def to_float(text):
    """转不出数就返回 None——"未检出""<0.5"这类都不是数，不能当 0 处理。"""
    try:
        return float(str(text).strip())
    except Exception:
        return None


def as_date(text):
    """日期串转 date；转不出来返回 None，别拿今天顶上。"""
    try:
        return date.fromisoformat(str(text).strip())
    except Exception:
        return None


def d_of(value):
    """回读的日期是 datetime（Excel 里还跟着区域设置走），比之前先统一。"""
    return value.date() if isinstance(value, datetime) else value


def safe_name(text: str) -> str:
    """文件名里不能带 \\ / : * ? " < > |，换掉再拼。"""
    bad = '\\/:*?"<>|'
    return "".join("-" if c in bad else c for c in str(text)).strip() or "未命名"


def norm_label(text) -> str:
    """把模板标签压成能比的样子：去掉括号里的中英文注释、冒号（含全角）和空白。

    「样本编号：」「样本编号\u3000：」「样本编号 (Sample ID)：」三种写法压完都是「样本编号」——
    不先压一遍，同一个字段换个模板就找不着了。
    """
    s = str(text if text is not None else "")
    s = re.sub(r"[（(].*?[)）]", "", s)
    s = re.sub(r"[\s\u3000：:]+", "", s)
    return s


def match_field(text):
    return ALIAS_TO_STD.get(norm_label(text))


def read_rows(path: Path, cols, key_col: str):
    """读一张导出表：BOM 去掉、每列去空白、判空列空的行丢掉。

    丢的是表尾那两条：合计行（条数填在检项列）和说明行——留着会被当成一条真记录。
    """
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for raw in csv.DictReader(f):
            if not (raw.get(key_col) or "").strip():
                continue
            rows.append({c: (raw.get(c) or "").strip() for c in cols})
    return rows


def read_results():
    return read_rows(RESULT_FILE, RESULT_COLS, "样本编号")


def read_config():
    """中心编号 → 这个中心用哪个模板、什么单位制、缺项怎么写。"""
    return {r["中心编号"]: r for r in read_rows(CONFIG_FILE, CONFIG_COLS, "中心编号")}


def group_samples(results, known_centers):
    """按 (中心编号, 样本编号) 归堆。

    配置表里没有这个中心时**不猜模板**，单列出来等人工处理：三个模板的字段和单位都不一样，
    猜错了出的是错报告，比不出更麻烦。
    """
    groups, strays = {}, []
    for r in results:
        if r["中心编号"] not in known_centers:
            strays.append(r)
            continue
        groups.setdefault((r["中心编号"], r["样本编号"]), []).append(r)
    return groups, strays


def target_cell(ws, label_cell):
    """值格 = 标签右边一格；右边那格若落在合并区里，就用标签自己的坐标。"""
    for rng in ws.merged_cells.ranges:
        if label_cell.coordinate in rng:
            return label_cell
    return ws.cell(row=label_cell.row, column=label_cell.column + 1)


def find_fields(ws):
    """扫模板，找出「标准字段名 → 值格坐标」。

    不写死坐标：三个中心的行号各不相同，谁在模板上面插一行，写死的那份脚本就整体错位，
    而且错的是"值填进隔壁标签的格子"这种看着像对的错。
    约定只有一条：值写在标签右边一格。
    """
    spots = {}
    for row in ws.iter_rows():
        for cell in row:
            std = match_field(cell.value)
            if not std or std in spots:
                continue
            spots[std] = target_cell(ws, cell).coordinate
    return spots


def find_detail_block(ws):
    """从模板里读出明细区：表头在第几行第几列、明细从哪行起、模板一共留了几行。

    明细区 = 表头下面连续的空行，碰到第一行有内容就停。所以模板里**必须把明细行
    （哪怕空着）的边框画好**，否则 openpyxl 读回来认不出这些行存在，脚本会以为一行都填不了。
    """
    head_row = head_col = None
    for row in ws.iter_rows():
        for cell in row:
            if norm_label(cell.value) == "检项":
                head_row, head_col = cell.row, cell.column
                break
        if head_row:
            break
    if not head_row:
        return None, None, []
    rows, r = [], head_row + 1
    while r <= ws.max_row and len(rows) < MAX_DETAIL_ROWS:
        blank = all(norm_label(ws.cell(row=r, column=head_col + i).value) == "" for i in range(4))
        if not blank:
            break
        rows.append(r)
        r += 1
    return head_row, head_col, rows


def range_text(low: str, high: str) -> str:
    """参考范围按两头拼；只给一头就只写一头——空的那头补个 0，等于把范围写错。"""
    lo, hi = to_float(low), to_float(high)
    if lo is None and hi is None:
        return ""
    if lo is None:
        return f"≤ {high.strip()}"
    if hi is None:
        return f"≥ {low.strip()}"
    return f"{low.strip()}~{high.strip()}"


def convert_result(item, text, unit, unit_system):
    """把结果按中心要的单位制处理好，返回 (写入值, 写入单位, 要人工看的说明)。

    三条规矩：
    - 国际单位的中心 → 原样写；
    - 常规单位的中心 + 检项配了系数 + 数据单位正是系数表里的源单位 → 乘系数，
      **同时把单位列一起改掉**——只改数值不改单位，数字看着对、单位是错的，审核根本看不出来；
    - 数据单位已经是目标单位，或跟系数表对不上 → 一律不动它，只报出来。
      差距是 18 倍这个量级，这一步宁可停下也不能猜。
    """
    text = (text or "").strip()
    if not text:
        return None, "", ""                    # 缺项，交给缺项写法那一步
    value = to_float(text) if to_float(text) is not None else text
    conv = UNIT_CONVERSIONS.get(item)
    if unit_system != "常规单位" or not conv:
        return value, unit, ""
    src, dst, factor = conv
    if unit != src:
        return value, unit, f"单位 {unit} 与换算来源 {src} 不一致，按原值写入"
    if not isinstance(value, float):
        return value, unit, "结果不是数值，未换算"
    return round(value * factor, 1), dst, ""


def derive(field, cfg, sample_id):
    """模板要、但数据里没有的字段——只有「报告编号」拼得出来，其余返回 None 等人工补。"""
    if field == "报告编号":
        return f"{cfg['报告编号前缀']}-RPT-{sample_id}"
    return None


def fill_one(cfg, sample_id, rows):
    """把这个样本的结果填进它自己中心的模板，另存成一份报告。

    返回 (台账记录, 报告路径或 None)：要么出报告，要么挂起并写清为什么。
    """
    rec = {"中心编号": cfg["中心编号"], "中心名称": cfg["中心名称"], "样本编号": sample_id,
           "状态": "", "检项数": len(rows), "字段命中": "", "换算项数": 0,
           "模板文件": cfg["模板文件"], "说明": ""}
    tpl = TPL_DIR / cfg["模板文件"]
    if not tpl.exists():
        rec["状态"], rec["说明"] = R_HOLD, f"模板文件不存在：{cfg['模板文件']}"
        return rec, None

    wb = load_workbook(tpl)
    ws = wb.active
    spots = find_fields(ws)
    head_row, head_col, detail_rows = find_detail_block(ws)
    if not detail_rows:
        rec["状态"], rec["说明"] = R_HOLD, "模板里找不到明细表头「检项」，明细区定位不了"
        return rec, None

    notes, todo, converted = [], [], 0
    for r in rows:
        value, unit, note = convert_result(r["检项"], r["结果"], r["单位"], cfg["单位制"])
        if unit and unit != r["单位"]:
            converted += 1
        if note:
            notes.append(f"{r['检项']}：{note}")
        if value is None:
            value = cfg["缺项写法"] or None          # 各中心对"没做"的写法不一样：留空 / 未做 / —
        todo.append([r["检项"], value, unit, range_text(r["参考下限"], r["参考上限"])])
    if len(todo) > len(detail_rows):
        rec["状态"] = R_HOLD
        rec["说明"] = f"模板明细区只有 {len(detail_rows)} 行，本样本 {len(todo)} 项"
        return rec, None

    head = rows[0]
    ctx = {"样本编号": sample_id, "受试者编号": head["受试者编号"], "中心编号": cfg["中心编号"],
           "中心名称": cfg["中心名称"], "访视点": head["访视点"],
           "检测日期": as_date(head["检测日期"])}
    written = 0
    for std, coord in spots.items():
        if std in ctx:
            value = ctx[std]
            if value == "":
                value = None
            if isinstance(value, date) and ws[coord].number_format == "General":
                ws[coord].number_format = "yyyy-mm-dd"   # 没设日期格式的格子会显示成 45625
            ws[coord] = value
            written += 1
        else:
            value = derive(std, cfg, sample_id)
            if value is None:
                notes.append(f"模板要「{std}」但数据里没有")
            else:
                ws[coord] = value
                written += 1
    # 数据里有、这个模板没有的字段：不硬塞。中心自己的模板长什么样，就按什么样出。
    extra = [std for std in ctx if std not in spots]
    if extra:
        notes.append("模板无此字段：" + "、".join(extra))

    for i, item in enumerate(todo):
        r = detail_rows[i]
        for j, value in enumerate(item):
            ws.cell(row=r, column=head_col + j, value=value)

    out_dir = OUT_DIR / safe_name(cfg["中心名称"])
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"样本{safe_name(sample_id)}_报告.xlsx"
    wb.save(out)

    rec["状态"] = R_MAKE
    rec["字段命中"] = f"{written}/{len(spots)}"
    rec["换算项数"] = converted
    rec["说明"] = "；".join(notes)
    return rec, out


def write_ledger(ledger, strays):
    """台账：一个样本一行（挂起的也在），加上"配置表里没有这个中心"的那几行。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "填充台账"
    ws.append(LEDGER_COLS)
    for r in ledger:
        ws.append([r[c] for c in LEDGER_COLS])
    ws2 = wb.create_sheet("待人工处理")
    ws2.append(STRAY_COLS)
    for r in strays:
        ws2.append([r["中心编号"], r["样本编号"], r["受试者编号"], r["检项"], r["结果"],
                    "中心编号不在中心模板配置表里，不猜模板"])
    for sheet, cols in ((ws, LEDGER_COLS), (ws2, STRAY_COLS)):
        for i, name in enumerate(cols, 1):
            sheet.column_dimensions[get_column_letter(i)].width = max(12, len(name) * 2 + 6)
    path = OUT_DIR / "填充台账.xlsx"
    wb.save(path)
    return path


def main():
    # 重跑先把上次的输出清干净：旧文件留着，就得一份份对哪些是新出的
    if OUT_DIR.exists():
        shutil.rmtree(OUT_DIR)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    config = read_config()
    groups, strays = group_samples(read_results(), set(config))
    ledger, made = [], []
    for key in sorted(groups):
        center, sample = key
        rec, out = fill_one(config[center], sample, groups[key])
        ledger.append(rec)
        if out:
            made.append(out)
        flag = "OK  " if rec["状态"] == R_MAKE else "挂起"
        print(f"{flag} {center} / {sample}  检项 {rec['检项数']}  字段 {rec['字段命中'] or '-'}"
              f"  换算 {rec['换算项数']}" + (f"  {rec['说明']}" if rec["说明"] else ""))

    write_ledger(ledger, strays)
    print(f"\n出报告 {len(made)} 份，挂起 {len(ledger) - len(made)} 份，"
          f"配置表里没有的中心 {len(strays)} 条，明细见 02_output/填充台账.xlsx")
    return _verify(made, ledger, strays)


def read_report(path: Path):
    """把已生成的报告读回来：表头键值对 + 明细。校验用，跟填的时候是两套读法。"""
    ws = load_workbook(path).active
    kv = {std: ws[coord].value for std, coord in find_fields(ws).items()}
    head_row, head_col, _ = find_detail_block(ws)
    items = {}
    for r in range(head_row + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=head_col).value
        if not isinstance(name, str) or not name.strip():
            continue
        items[name] = (ws.cell(row=r, column=head_col + 1).value,
                       ws.cell(row=r, column=head_col + 2).value)
    return kv, items


def read_ledger():
    wb = load_workbook(OUT_DIR / "填充台账.xlsx")
    ws = wb["填充台账"]
    rows = list(ws.iter_rows(values_only=True))
    ledger = [dict(zip(rows[0], r)) for r in rows[1:]]
    ws2 = wb["待人工处理"]
    return ledger, list(ws2.iter_rows(values_only=True))[1:]


def _verify(made, ledger, strays):
    """回读每一份报告，逐项确认：编号没变、单位换对了、缺项按各中心的写法写了。"""
    ok = True

    def check(label, got, want):
        nonlocal ok
        good = got == want
        ok = ok and good
        print(f"{'OK  ' if good else 'FAIL'} {label}：{got!r}" + ("" if good else f"（期望 {want!r}）"))

    print(f"\n共生成 {len(made)} 份报告")
    kv, items = read_report(OUT_DIR / "华东中心" / "样本0091_报告.xlsx")
    check("A01/0091 样本编号（前导零）", kv.get("样本编号"), "0091")
    check("A01/0091 访视点", kv.get("访视点"), "V2")
    check("A01/0091 检测日期", d_of(kv.get("检测日期")), date(2026, 9, 8))
    check("A01/0091 葡萄糖（国际单位，原样）", items.get("葡萄糖"), (5.4, "mmol/L"))
    check("A01/0091 总胆红素（缺项→留空）", items.get("总胆红素", (None,))[0], None)

    kv, items = read_report(OUT_DIR / "North Site" / "样本B-1007_报告.xlsx")
    check("B02/B-1007 中心名称（模板专属字段）", kv.get("中心名称"), "North Site")
    check("B02/B-1007 葡萄糖（换算+改单位）", items.get("葡萄糖"), (91.9, "mg/dL"))
    check("B02/B-1007 甘油三酯（换算+改单位）", items.get("甘油三酯"), (159.4, "mg/dL"))
    check("B02/B-1007 钾（没配系数→原样）", items.get("钾"), (4.2, "mmol/L"))
    _, items = read_report(OUT_DIR / "North Site" / "样本B-1008_报告.xlsx")
    check("B02/B-1008 葡萄糖（单位已是 mg/dL→不换算）", items.get("葡萄糖"), (108, "mg/dL"))
    check("B02/B-1008 总胆红素（缺项→未做）", items.get("总胆红素", (None,))[0], "未做")

    kv, items = read_report(OUT_DIR / "华南中心" / "样本0093_报告.xlsx")
    check("C03/0093 报告编号（数据里没有→按规则拼）", kv.get("报告编号"), "C-RPT-0093")
    check("C03/0093 样本编号（标签夹全角空格照样认）", kv.get("样本编号"), "0093")
    check("C03/0093 总胆红素（缺项→—）", items.get("总胆红素", (None,))[0], "—")

    led = {r["样本编号"]: r for r in ledger}
    check("台账行数（表尾两行未计入）", len(ledger), 6)
    check("A01/0092 状态", led["0092"]["状态"], R_HOLD)
    check("A01/0092 未生成报告", (OUT_DIR / "华东中心" / "样本0092_报告.xlsx").exists(), False)
    check("A01/0092 挂起原因写清了", "明细区只有 12 行" in led["0092"]["说明"], True)
    check("B02/B-1008 单位对不上已报出", "不一致" in led["B-1008"]["说明"], True)
    check("B02/B-1007 模板无「访视点」已报出", "访视点" in led["B-1007"]["说明"], True)
    check("配置表里没有的中心（不猜模板）", len(strays), 2)
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
