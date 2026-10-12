# -*- coding: utf-8 -*-
"""试剂效期盘点：读试剂台账（系统导出）→ 算实际到期日与剩余天数 → 按阈值表定状态 → 出盘点 Excel。

用法：
    python run.py            # 跑完暂停等回车（Windows 双击也能用）
    python run.py --no-pause # 跑完直接退出
"""
import csv
import datetime as dt
import io
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

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ---------------- 路径与字段常量区 ----------------
BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
TPL_DIR = SRC_DIR / "templates"
LEDGER_FILE = RAW_DIR / "试剂台账_系统导出.csv"
THRESH_FILE = TPL_DIR / "状态阈值表.csv"
LOG_FILE = SRC_DIR / "run_log.txt"

OUT_NAME = "试剂效期盘点.xlsx"

NAME_COL = "试剂名称"
SPEC_COL = "规格"
LOT_COL = "批号"
QTY_COL = "库存数量"
UNIT_COL = "单位"
LOC_COL = "存放位置"
EXP_COL = "有效期至"
OPEN_COL = "开封日期"
OPEN_DAYS_COL = "开封后效期天数"

ALIASES = {
    NAME_COL: ("试剂名称", "名称", "试剂"),
    SPEC_COL: ("规格", "规格型号"),
    LOT_COL: ("批号", "批次"),
    QTY_COL: ("库存数量", "数量", "余量"),
    UNIT_COL: ("单位", "数量单位"),
    LOC_COL: ("存放位置", "位置", "储存位置"),
    EXP_COL: ("有效期至", "有效期", "失效日期"),
    OPEN_COL: ("开封日期", "开封时间"),
    OPEN_DAYS_COL: ("开封后效期天数", "开封效期天数"),
}

# 已过期 / 紧急 / 临期 / 正常 四档的底色（明细里整行标，扫一眼就知道先处理谁）
FILL_EXPIRED = PatternFill("solid", fgColor="FFC7CE")
FILL_URGENT = PatternFill("solid", fgColor="FFE699")


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
    """读文本并返回实际用的编码：仪器/系统导出的 CSV 常在 UTF-8 和 GBK 之间摇摆。"""
    for enc in ("utf-8-sig", "gbk"):
        try:
            return path.read_text(encoding=enc), enc
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("csv", b"", 0, 1, f"{path.name}: 编码既不是 utf-8 也不是 gbk")


def find_header(rows):
    """找表头行：第一个「非空单元格 >= 2」的行。

    系统导出常在表头前压几行导出信息（系统名、导出时间、过滤条件），
    写死 header=0 会把这些当成表头；靠行号猜则导出格式一改就全错。
    """
    for i, r in enumerate(rows):
        if sum(1 for c in r if str(c or "").strip()) >= 2:
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


def parse_date(s):
    """把日期单元格转成 date；空返回 None；认不出的写法返回 None（交给未判定，不猜）。"""
    s = str(s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def to_int(s):
    s = str(s or "").strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def load_thresholds(path):
    """读状态阈值表：状态 → 剩余天数上限（含）。

    阈值是管理口径，写在表里不写死在代码里——「紧急」改「8 天以内」只是改一行表。
    按顺序找第一个命中的档，所以表里行的顺序就是判定优先级。
    """
    text, _ = read_text(path)
    rows = [r for r in csv.reader(io.StringIO(text)) if any(str(c).strip() for c in r)]
    rules = []
    for r in rows[1:]:
        name = str(r[0]).strip()
        limit = to_int(r[1]) if len(r) > 1 else None
        if name and limit is not None:
            rules.append((name, limit))
    if not rules:
        raise ValueError(f"{path.name}: 阈值表是空的，没法判状态")
    return rules


def actual_expiry(exp, opened, open_days):
    """算实际到期日 = min(有效期至, 开封日期 + 开封后效期天数)。

    返回 (到期日, 错误说明)。到期日为 None 时说明写在 reason 里：
    试剂开封后往往只能再用几十天，比瓶身印的效期短得多——只看有效期至会漏掉
    「开封了却放了半年」的那批，这是效期盘点最容易漏的地方。
    """
    if exp is None:
        return None, "有效期至为空"
    if opened is not None and open_days is None:
        return None, "已开封但没记开封后效期天数"
    if opened is not None and open_days is not None:
        if opened > exp:
            return None, "开封日期晚于有效期至，台账有误"
        return min(exp, opened + dt.timedelta(days=open_days)), ""
    return exp, ""


def judge_status(days, rules):
    """剩余天数 → 状态：按阈值表顺序找第一个「剩余天数 ≤ 上限」的档。"""
    for name, limit in rules:
        if days <= limit:
            return name
    return rules[-1][0]


def check_all(rows, hdr, cols, rules, today):
    """逐行算实际到期日与状态，返回 (明细, 到期提醒, 未判定清单)。"""
    detail, alerts, undecided = [], [], []
    for r in rows[hdr + 1:]:
        cells = [str(c or "").strip() for c in r]
        if not any(cells):
            continue
        exp = parse_date(cell(cells, cols.get(EXP_COL)))
        opened = parse_date(cell(cells, cols.get(OPEN_COL)))
        open_days = to_int(cell(cells, cols.get(OPEN_DAYS_COL)))
        expiry, why = actual_expiry(exp, opened, open_days)
        row = {
            NAME_COL: cell(cells, cols.get(NAME_COL)),
            SPEC_COL: cell(cells, cols.get(SPEC_COL)),
            LOT_COL: cell(cells, cols.get(LOT_COL)),
            QTY_COL: cell(cells, cols.get(QTY_COL)),
            UNIT_COL: cell(cells, cols.get(UNIT_COL)),
            LOC_COL: cell(cells, cols.get(LOC_COL)),
            EXP_COL: cell(cells, cols.get(EXP_COL)),
            "实际到期日": expiry.isoformat() if expiry else "",
            "剩余天数": (expiry - today).days if expiry else "",
            "状态": "",
            "说明": why,
        }
        if expiry is None:
            row["状态"] = "未判定"
            undecided.append(row)
        else:
            row["状态"] = judge_status(row["剩余天数"], rules)
            if row["状态"] != "正常":
                alerts.append(row)
        detail.append(row)
    return detail, alerts, undecided


def summarize(detail):
    """按状态归堆计数，给汇总表用。"""
    counts = {}
    for r in detail:
        counts[r["状态"]] = counts.get(r["状态"], 0) + 1
    return counts


def write_outputs(detail, alerts, undecided, counts, today):
    """一次落三张表：盘点明细（红/黄标）、到期提醒（过期+紧急+临期）、盘点汇总。"""
    cols = [NAME_COL, SPEC_COL, LOT_COL, QTY_COL, UNIT_COL, LOC_COL,
            EXP_COL, "实际到期日", "剩余天数", "状态", "说明"]
    wb = Workbook()

    ws = wb.active
    ws.title = "盘点明细"
    ws.append(cols)
    for r in detail:
        ws.append([r.get(c, "") for c in cols])
        if r["状态"] == "已过期":
            for j in range(1, len(cols) + 1):
                ws.cell(row=ws.max_row, column=j).fill = FILL_EXPIRED
        elif r["状态"] == "紧急":
            for j in range(1, len(cols) + 1):
                ws.cell(row=ws.max_row, column=j).fill = FILL_URGENT
    for j in range(1, len(cols) + 1):
        ws.cell(row=1, column=j).font = Font(bold=True)
    lot_idx = cols.index(LOT_COL) + 1
    for i in range(2, ws.max_row + 1):
        ws.cell(row=i, column=lot_idx).number_format = "@"   # 批号按文本存，前导零才不掉

    ws2 = wb.create_sheet("到期提醒")
    ws2.append(cols)
    for r in sorted(alerts, key=lambda x: x["剩余天数"]):
        ws2.append([r.get(c, "") for c in cols])
    for j in range(1, len(cols) + 1):
        ws2.cell(row=1, column=j).font = Font(bold=True)

    ws3 = wb.create_sheet("盘点汇总")
    ws3.append(["盘点日期", today.isoformat()])
    ws3.append(["台账条数", len(detail)])
    for name, _ in THRESH_ORDER:
        ws3.append([name, counts.get(name, 0)])
    ws3.append(["未判定", counts.get("未判定", 0)])
    ws3.append(["合计", sum(counts.values())])
    ws3.cell(row=1, column=1).font = Font(bold=True)

    out = OUT_DIR / OUT_NAME
    wb.save(out)
    return out


def verify(detail, alerts, undecided, counts):
    """生成后回读断言：数据是故意造了边界值的，正常数据跑通不代表对。"""
    oks = []
    oks.append(("明细条数 = 提醒 + 未判定 + 正常",
                len(alerts) + len(undecided) + counts.get("正常", 0) == len(detail)))
    oks.append(("汇总合计 = 明细条数", sum(counts.values()) == len(detail)))
    expired_lots = {r[LOT_COL] for r in detail if r["状态"] == "已过期"}
    oks.append(("已知过期批号 000210 判成了已过期", "000210" in expired_lots))
    oks.append(("每条明细都有状态", all(r["状态"] for r in detail)))
    oks.append(("未判定的每条都带了原因", all(r["说明"] for r in undecided)))
    oks.append(("开了封没记效期的进了未判定", any("开封" in r["说明"] for r in undecided)))
    ok = all(ok for _, ok in oks)
    for name, passed in oks:
        print(f"  [{'OK' if passed else 'FAIL'}] {name}")
    return ok


THRESH_ORDER = [("已过期", None), ("紧急", None), ("临期", None), ("正常", None)]


def main():
    t0 = time.time()
    host, ip = now_host_ip()
    log(f"")
    log(f"运行主机：{host}（{ip}）")
    log(f"开始执行：{dt.datetime.now().isoformat(timespec='seconds')}  "
        f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）")

    OUT_DIR.mkdir(exist_ok=True)
    today = dt.date.today()
    rules = load_thresholds(THRESH_FILE)
    text, enc = read_text(LEDGER_FILE)
    rows = list(csv.reader(io.StringIO(text)))
    hdr = find_header(rows)
    if hdr < 0:
        raise ValueError(f"{LEDGER_FILE.name}: 找不到表头行")
    cols = pick_columns(rows[hdr])
    missing = [k for k, v in cols.items() if v is None]
    if missing:
        raise ValueError(f"台账表头缺字段：{missing}")

    detail, alerts, undecided = check_all(rows, hdr, cols, rules, today)
    counts = summarize(detail) | {"未判定": len(undecided)}
    s = {name: counts.get(name, 0) for name, _ in THRESH_ORDER}
    print(f"读入台账 {enc} 编码、表头在第 {hdr + 1} 行 → 共 {len(detail)} 条："
          f"已过期 {s['已过期']} / 紧急 {s['紧急']} / 临期 {s['临期']} / "
          f"正常 {s['正常']} / 未判定 {counts.get('未判定', 0)}")
    for r in sorted(alerts, key=lambda x: x["剩余天数"])[:10]:
        print(f"  [{r['状态']}] {r[NAME_COL]} 批号 {r[LOT_COL]} "
              f"剩 {r['剩余天数']} 天（{r['实际到期日']}）")

    out = write_outputs(detail, alerts, undecided, counts, today)
    print(f"已生成：{out.name}（盘点明细 / 到期提醒 / 盘点汇总 三张表）")
    ok = verify(detail, alerts, undecided, counts)
    print("总判定：" + ("PASS" if ok else "FAIL"))

    log(f"执行完成：{dt.datetime.now().isoformat(timespec='seconds')}  "
        f"用时 {time.time() - t0:.1f}s  判定 {'PASS' if ok else 'FAIL'}")

    if "--no-pause" not in sys.argv:
        input("\n按回车退出……")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
