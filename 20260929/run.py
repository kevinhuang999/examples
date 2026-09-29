# -*- coding: utf-8 -*-
"""多客户检测数据汇总报表自动生成。

读 01_raw_data/ 里的三张表（检测结果明细 / 客户报表口径表 / 判定标准表），
按每家客户各自的报表口径出一页「检测数据汇总报表」，汇成一份工作簿，
再单独导一份不合格项明细。跑完自动回读校验，打印 PASS/FAIL。

用法：python run.py      （Windows 双击 run.py 也行；加 --no-pause 跳过结尾回车）
"""

import csv
import socket
import subprocess
import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
TPL_FILE = SRC_DIR / "templates" / "客户检测数据汇总报表模板.xlsx"
LOG_FILE = SRC_DIR / "run_log.txt"

RESULT_FILE = RAW_DIR / "检测结果明细.csv"
SPEC_FILE = RAW_DIR / "客户报表口径表.csv"
RULE_FILE = RAW_DIR / "判定标准表.csv"
REPORT_FILE = OUT_DIR / "客户检测数据汇总报表.xlsx"
BAD_FILE = OUT_DIR / "不合格项明细.csv"

PERIOD = "2026-09"          # 统计期间
CLOSE_DATE = "2026-09-30"   # 数据截止日：这天之后出的报告不算本期

# 结果明细的列名（跟着仪器导出的表头写，换表头就换这里）
COL_ORDER = "委托单号"
COL_CLIENT = "客户名称"
COL_SAMPLE = "样品编号"
COL_CATEGORY = "检测类别"
COL_ITEM = "检测项目"
COL_VALUE = "结果值"
COL_UNIT = "单位"
COL_STATUS = "报告状态"
COL_REPORT = "报告日期"

STATUS_DONE = "已出报告"
STATUS_VOID = "已作废"

INDICATORS = ["送检样品数", "检测项数", "已判定项数", "合格项数",
              "不合格项数", "未完成项数", "待确认项数", "合格率"]


def ensure_openpyxl():
    """缺依赖自动装（走清华源），装不上再退到 --user。"""
    try:
        import openpyxl  # noqa: F401
        return
    except ImportError:
        print("缺少 openpyxl，正在用清华源安装 …")
    mirror = ["-i", "https://pypi.tuna.tsinghua.edu.cn/simple"]
    for extra in ([], ["--user"]):
        args = [sys.executable, "-m", "pip", "install", "openpyxl"] + extra + mirror
        if subprocess.call(args) == 0:
            return
    raise SystemExit("openpyxl 装不上，请手动执行：pip install openpyxl")


def machine_tag():
    """运行主机标识：主机名 + 本机对外网卡 IP，用来分辨是哪台机器跑的。"""
    host = socket.gethostname()
    ip = "未知"
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
    except OSError:
        try:
            ip = socket.gethostbyname(host)
        except OSError:
            ip = "未知"
    return f"{host}（{ip}）"


def log_open():
    """开一段追加日志：先记运行主机，再记开始时间与解释器。"""
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    fh = open(LOG_FILE, "a", encoding="utf-8")
    fh.write(f"\n运行主机：{machine_tag()}\n")
    fh.write(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
             f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）\n")
    return fh


def norm(text):
    """比较用的归一化：去掉所有空白（含全角空格），统一成字符串。

    客户名尾巴粘一个空格或全角空格（U+3000），不归一化就会把同一家客户拆成两三家。
    """
    return "".join(str(text or "").split()).replace("\u3000", "")


def clean_name(text):
    """显示用客户名：只去掉首尾空白，中间的字保留原样。"""
    return str(text or "").strip()


def read_table(path, key_col):
    """按 key_col 判空读一张 CSV：该列为空的行（表尾合计/说明行）一律跳过。

    判空列必须一张表传一个——拿结果明细的「委托单号」去读口径表，
    那张表根本没有这一列，每行都会被当成空行丢掉，整张表读空还不报错。
    """
    if not path.exists():
        raise FileNotFoundError(f"找不到数据文件：{path.name}")
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        raw_rows = list(csv.DictReader(fh))
    rows = []
    for row in raw_rows:
        if norm(row.get(key_col)) == "":
            continue
        rows.append({k: str(v or "").strip() for k, v in row.items()})
    if not rows:
        raise ValueError(f"{path.name} 按「{key_col}」判空后一行都没读到，先看列名对不对")
    return rows


def parse_value(text):
    """解析结果值：返回 (符号, 数值)；符号是「=」「<」「>」，读不出数值时返回原文。

    仪器导出的结果值不只有纯数字：有「<0.01」这种报告值、有「未检出」这种文字结论，
    也有整格空的（结果没传回来）。直接 float() 会在这些行上崩。
    """
    raw = norm(text)
    if raw == "":
        return None, None
    head, tail = raw[0], raw[1:]
    if head in "<＜":
        try:
            return "<", float(tail)
        except ValueError:
            return raw, None
    if head in ">＞":
        try:
            return ">", float(tail)
        except ValueError:
            return raw, None
    try:
        return "=", float(raw)
    except ValueError:
        return raw, None


def judge(value_text, rule):
    """按判定标准判一项结果：返回「合格」「不合格」，判不了返回 None。

    判定方式不是只有上下限——只卡上限的（限量值）、只有下限的、纯文字的（微生物结论），
    所以这里按 rule 里的方式分派，不写死一个比较符号。
    """
    way = norm(rule.get("判定方式"))
    low = norm(rule.get("下限"))
    high = norm(rule.get("上限"))
    expect = norm(rule.get("合格结论文字"))
    sign, num = parse_value(value_text)
    if sign is None:
        return None
    if way == "文字":
        return "合格" if sign == expect else "不合格"
    if sign in ("<", ">"):
        # 报告值只给了边界：<L 落在上限以内算合格，>L 落在下限以上算合格，跨界的说不清
        if sign == "<":
            return "合格" if high and num <= float(high) else None
        return "合格" if low and num >= float(low) else None
    if num is None:
        return None
    if way == "上限":
        return "合格" if high and num <= float(high) else "不合格"
    if way == "下限":
        return "合格" if low and num >= float(low) else "不合格"
    if way == "区间":
        ok = True
        if low:
            ok = ok and num >= float(low)
        if high:
            ok = ok and num <= float(high)
        return "合格" if ok else "不合格"
    return None


def rule_text(rule):
    """把判定标准写成报表上的一句话（不合格明细里要印出来）。"""
    way = norm(rule.get("判定方式"))
    if way == "文字":
        return f"结论须为「{norm(rule.get('合格结论文字'))}」"
    low, high = norm(rule.get("下限")), norm(rule.get("上限"))
    if way == "上限":
        return f"≤ {high}"
    if way == "下限":
        return f"≥ {low}"
    return f"{low} ~ {high}"


def pick_latest(rows):
    """同一样品同一项目可能有好几条结果（复检重出）：剔掉作废的，再取报告日期最新的一条。

    作废件不能参与「取最新」——它是被撤回的结论，日期再新也不算数。
    两条并列为最新（同一天出两份报告）说明数据有问题，返回 None 交给人工。
    """
    live = [r for r in rows if norm(r.get(COL_STATUS)) != STATUS_VOID]
    if not live:
        return None, "只剩被作废的结果"
    latest = max(live, key=lambda r: norm(r.get(COL_REPORT)) or "0000-00-00")
    same = [r for r in live if norm(r.get(COL_REPORT)) == norm(latest.get(COL_REPORT))]
    if len(same) > 1:
        return None, "同一报告日期有两条结果，需要人工确认"
    return latest, ""


def group_by_client(rows):
    """按客户归堆，客户名先归一化再当键，显示名另存干净的。"""
    buckets = {}
    for row in rows:
        key = norm(row.get(COL_CLIENT))
        if key == "":
            continue
        buckets.setdefault(key, []).append(row)
    return buckets


def sample_set(rows):
    """这一堆记录里出现过多少份样品（作废件不算）。"""
    return {norm(r.get(COL_SAMPLE)) for r in rows if norm(r.get(COL_STATUS)) != STATUS_VOID}


def summarize_client(rows, spec, rules):
    """一家客户的汇总：按「样品 + 检测项目」去重，逐项判定，再按客户口径分组统计。

    去重这步不能省——复检重出的样品在明细里留下两三条同项目的结果，
    直接汇总等于把这些项算了两遍；取哪一条、作废的怎么处理都在 pick_latest() 里。
    """
    items = {}
    for row in rows:
        items.setdefault((norm(row.get(COL_SAMPLE)), norm(row.get(COL_ITEM))), []).append(row)
    live_items = {k: v for k, v in items.items()
                  if any(norm(r.get(COL_STATUS)) != STATUS_VOID for r in v)}

    stat = {name: 0 for name in INDICATORS}
    stat["送检样品数"] = len(sample_set(rows))
    stat["检测项数"] = len(live_items)
    detail, bad, pending = {}, [], []
    dim_col = COL_CATEGORY if norm(spec["分组维度"]) == "检测类别" else COL_ORDER
    for (sample, item), group in live_items.items():
        bucket = detail.setdefault(norm(group[0].get(dim_col)),
                                   {"检测项数": 0, "合格": 0, "不合格": 0})
        bucket["检测项数"] += 1
        latest, note = pick_latest(group)
        if latest is None:
            stat["待确认项数"] += 1
            pending.append((sample, item, note))
            continue
        if norm(latest.get(COL_STATUS)) != STATUS_DONE:
            # 报告还没出：不计入合格率分母，也不算丢——单独计一个数
            stat["未完成项数"] += 1
            continue
        rule = rules.get((norm(latest.get(COL_CATEGORY)), norm(latest.get(COL_ITEM))))
        if rule is None:
            stat["待确认项数"] += 1
            pending.append((sample, item, "判定标准表里没有这一项"))
            continue
        verdict = judge(latest.get(COL_VALUE), rule)
        if verdict is None:
            stat["待确认项数"] += 1
            pending.append((sample, item, "结果值读不出，判不了"))
            continue
        stat["已判定项数"] += 1
        stat["合格项数" if verdict == "合格" else "不合格项数"] += 1
        bucket["合格" if verdict == "合格" else "不合格"] += 1
        if verdict == "不合格":
            bad.append({"客户名称": clean_name(latest.get(COL_CLIENT)),
                        "委托单号": latest.get(COL_ORDER),
                        "样品编号": latest.get(COL_SAMPLE),
                        "检测类别": latest.get(COL_CATEGORY),
                        "检测项目": latest.get(COL_ITEM),
                        "结果值": latest.get(COL_VALUE),
                        "单位": latest.get(COL_UNIT),
                        "判定标准": rule_text(rule),
                        "报告日期": latest.get(COL_REPORT)})
    return {"stat": stat, "detail": detail, "bad": bad, "pending": pending}


def label_row(ws, text):
    """按 A 列的标签名反查行号——模板挪一行也不会填错地方。"""
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, 1).value) == norm(text):
            return r
    return None


def section_row(ws, prefix):
    """找「一、」「二、」「三、」这种小节标题所在行。"""
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(r, 1).value).startswith(prefix):
            return r
    return None


def rate_text(qualified, judged):
    """合格率：分母是已判定项数，一条都没判出来时给「—」，不给 0%。"""
    return f"{qualified / judged:.1%}" if judged else "—"


def write_client_sheet(ws, client, spec, summary):
    """把一家客户的汇总结果填进按模板复制出来的这一页。"""
    ws.cell(2, 2).value = client
    ws.cell(3, 2).value = PERIOD
    stat = summary["stat"]
    for name in INDICATORS:
        row = label_row(ws, name)
        if name == "合格率":
            ws.cell(row, 2).value = rate_text(stat["合格项数"], stat["已判定项数"])
        else:
            ws.cell(row, 2).value = stat[name]

    head = section_row(ws, "二、") + 1
    ws.cell(head, 1).value = spec["表头组名"]
    for i, (gkey, cell) in enumerate(summary["detail"].items()):
        r = head + 1 + i
        ws.cell(r, 1).value = gkey
        ws.cell(r, 2).value = cell["检测项数"]
        ws.cell(r, 3).value = cell["合格"]
        ws.cell(r, 4).value = cell["不合格"]
        ws.cell(r, 5).value = rate_text(cell["合格"], cell["合格"] + cell["不合格"])

    bad_head = section_row(ws, "三、") + 1
    if spec["含不合格明细"] == "是":
        for i, item in enumerate(summary["bad"]):
            r = bad_head + 1 + i
            ws.cell(r, 1).value = item["样品编号"]
            ws.cell(r, 2).value = item["检测项目"]
            ws.cell(r, 3).value = f"{item['结果值']} {item['单位']}".strip()
            ws.cell(r, 4).value = item["判定标准"]
            ws.cell(r, 5).value = item["报告日期"]
    else:
        ws.cell(bad_head + 1, 1).value = "按客户口径，本报表只给统计数，不列不合格明细。"


def fill_summary(ws, rows):
    """把客户一览写进「汇总」页。"""
    for i, row in enumerate(rows):
        for j, value in enumerate(row):
            ws.cell(2 + i, 1 + j).value = value


def write_bad_csv(rows):
    cols = ["客户名称", "委托单号", "样品编号", "检测类别", "检测项目",
            "结果值", "单位", "判定标准", "报告日期"]
    with open(BAD_FILE, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def add(checks, name, got, want):
    ok = got == want
    checks.append(("OK   " if ok else "FAIL ") + f"{name}：{got!r}"
                  + ("" if ok else f" ≠ {want!r}"))
    return ok


def verify(load_workbook):
    """回读生成的报表，逐项断言。正常数据跑通不代表对，边界值全对才算。"""
    checks = []
    wb = load_workbook(REPORT_FILE)
    add(checks, "工作簿页签", wb.sheetnames,
        ["汇总", "华泰新材料有限公司", "同安环境检测有限公司", "恒基食品科技有限公司"])

    want = {
        "华泰新材料有限公司": {"送检样品数": 4, "检测项数": 7, "已判定项数": 5, "合格项数": 4,
                               "不合格项数": 1, "未完成项数": 2, "待确认项数": 0,
                               "合格率": "80.0%"},
        "同安环境检测有限公司": {"送检样品数": 4, "检测项数": 5, "已判定项数": 4, "合格项数": 3,
                                 "不合格项数": 1, "未完成项数": 0, "待确认项数": 1,
                                 "合格率": "75.0%"},
        "恒基食品科技有限公司": {"送检样品数": 4, "检测项数": 5, "已判定项数": 4, "合格项数": 2,
                                 "不合格项数": 2, "未完成项数": 0, "待确认项数": 1,
                                 "合格率": "50.0%"},
    }
    for name, values in want.items():
        ws = wb[name]
        add(checks, f"{name}·客户名", ws.cell(2, 2).value, name)
        for key, value in values.items():
            add(checks, f"{name}·{key}", ws.cell(label_row(ws, key), 2).value, value)

    ht = wb["华泰新材料有限公司"]
    add(checks, "华泰·分组表头", ht.cell(16, 1).value, "检测类别")
    add(checks, "华泰·水质行", [ht.cell(17, c).value for c in range(1, 6)],
        ["水质", 5, 3, 0, "100.0%"])
    add(checks, "华泰·土壤行", [ht.cell(18, c).value for c in range(1, 6)],
        ["土壤", 2, 1, 1, "50.0%"])
    add(checks, "华泰·不合格明细", [ht.cell(30, c).value for c in range(1, 3)], ["0000103", "镉"])

    ta = wb["同安环境检测有限公司"]
    add(checks, "同安·分组表头", ta.cell(16, 1).value, "委托单号")
    add(checks, "同安·委托单1", [ta.cell(17, c).value for c in range(1, 6)],
        ["WT2026-0000704", 3, 2, 0, "100.0%"])
    add(checks, "同安·委托单2", [ta.cell(18, c).value for c in range(1, 6)],
        ["WT2026-0000705", 2, 1, 1, "50.0%"])
    add(checks, "同安·不留不合格明细",
        str(ta.cell(30, 1).value).startswith("按客户口径"), True)

    hj = wb["恒基食品科技有限公司"]
    add(checks, "恒基·分组表头", hj.cell(16, 1).value, "检测项目类别")
    add(checks, "恒基·食品行", [hj.cell(17, c).value for c in range(1, 6)],
        ["食品", 5, 2, 2, "50.0%"])
    add(checks, "恒基·不合格明细两条", [hj.cell(30, 2).value, hj.cell(31, 2).value], ["菌落总数", "大肠菌群"])

    idx = wb["汇总"]
    add(checks, "汇总·正常客户1", [idx.cell(2, c).value for c in range(1, 8)],
        ["华泰新材料有限公司", 4, 7, 4, 1, "80.0%", "正常"])
    add(checks, "汇总·挂起客户", [idx.cell(5, c).value for c in range(1, 3)],
        ["信远电子科技有限公司", 1])
    add(checks, "汇总·挂起原因", idx.cell(5, 7).value, "未进客户报表口径表，先补口径再出报表")

    with open(BAD_FILE, "r", encoding="utf-8-sig", newline="") as fh:
        bad_rows = list(csv.DictReader(fh))
    add(checks, "不合格明细 CSV 条数", len(bad_rows), 4)
    add(checks, "不合格明细 CSV 首条样品", bad_rows[0]["样品编号"], "0000103")

    failed = [c for c in checks if c.startswith("FAIL")]
    return not failed, checks


def main():
    ensure_openpyxl()
    from openpyxl import load_workbook

    started = datetime.now()
    log = log_open()
    verdict = "FAIL"
    try:
        results = read_table(RESULT_FILE, COL_ORDER)
        specs = {norm(r["客户名称"]): r for r in read_table(SPEC_FILE, "客户名称")}
        rules = {(norm(r["检测类别"]), norm(r["检测项目"])): r
                 for r in read_table(RULE_FILE, "检测项目")}

        OUT_DIR.mkdir(parents=True, exist_ok=True)
        wb = load_workbook(TPL_FILE)
        tpl = wb["报表"]
        summary_rows, bad_all, pending_clients = [], [], []
        for key, rows in group_by_client(results).items():
            display = clean_name(rows[0].get(COL_CLIENT))
            spec = specs.get(key)
            if spec is None:
                # 客户没进口径表：不猜口径、不出这一页，只在一览表里挂起
                pending_clients.append(display)
                summary_rows.append([display, len(sample_set(rows)), "—", "—", "—", "—",
                                     "未进客户报表口径表，先补口径再出报表"])
                continue
            summary = summarize_client(rows, spec, rules)
            sheet = wb.copy_worksheet(tpl)
            sheet.title = display[:31]
            write_client_sheet(sheet, display, spec, summary)
            stat = summary["stat"]
            summary_rows.append([display, stat["送检样品数"], stat["检测项数"], stat["合格项数"],
                                 stat["不合格项数"],
                                 rate_text(stat["合格项数"], stat["已判定项数"]), "正常"])
            bad_all.extend(summary["bad"])

        del wb["报表"]
        fill_summary(wb["汇总"], summary_rows)
        wb.save(REPORT_FILE)
        write_bad_csv(bad_all)

        ok, checks = verify(load_workbook)
        for line in checks:
            print(line)
        verdict = "PASS" if ok else "FAIL"
        log.write(f"结果：客户 {len(summary_rows)} 家，出报表 "
                  f"{len(summary_rows) - len(pending_clients)} 家，挂起 {len(pending_clients)} 家，"
                  f"不合格项 {len(bad_all)} 条\n")
    except Exception as exc:
        verdict = "FAIL"
        print(f"执行中断：{type(exc).__name__}: {exc}")
        log.write(f"执行中断：{type(exc).__name__}: {exc}\n")

    elapsed = (datetime.now() - started).total_seconds()
    # 完成行写在暂停之前——双击跑完直接关窗口会杀进程，日志就只剩半截
    log.write(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  用时 {elapsed:.1f}s  判定 {verdict}\n")
    log.close()
    print(f"\n执行完成  用时 {elapsed:.1f}s  判定 {verdict}")
    print(f"报表：{REPORT_FILE.relative_to(BASE)}")
    print(f"日志：{LOG_FILE.relative_to(BASE)}")

    if "--no-pause" not in sys.argv:
        try:
            input("\n按回车退出 …")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
