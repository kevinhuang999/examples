#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把存量纸质「样本接收单」批量转成线上接收单，并跟设备扫码导出的明细逐单核对。

一句话：纸面上才有的字段集中补录一次，明细里本来就有的字段脚本自己带出来，
再把纸面手填的数字和明细实际条数对一遍，对不上的单挂起不放过。

用法：
    python run.py              # 跑完停住等回车（Windows 双击也行）
    python run.py --no-pause   # 跑完直接退出
"""

import csv
import socket
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

# ----------------------------------------------------------------- ① 配置区（跑之前主要改这里）

ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "01_raw_data"
OUT_DIR = ROOT / "02_output"
TPL_DIR = ROOT / "source" / "templates"
FORM_DIR = OUT_DIR / "线上接收单"
LOG_FILE = ROOT / "source" / "run_log.txt"

PAPER_CSV = RAW_DIR / "纸质接收单_人工补录.csv"
DETAIL_CSV = RAW_DIR / "扫码导出_样本明细.csv"
TEMPLATE = TPL_DIR / "接收单模板.xlsx"

LEDGER_CSV = OUT_DIR / "接收单台账.csv"
DIFF_CSV = OUT_DIR / "差异清单.csv"

# 发运中心 → 线上单号里的三段码（中心多了就往这张表里加一行）
CENTER_CODE = {"上海中心": "SHA", "北京中心": "BJS", "广州中心": "CAN"}

# 只有纸面上才有的字段，逐栏搬进线上单；扫码文件里没有，也没法算
# （纸质单上登记的箱数/管数不搬：那两个数字要跟明细比，比完写进「纸面对账」一行）
PAPER_TO_ONLINE = {
    "发运中心": "发运中心",
    "到达日期": "到达日期",
    "到达时间": "到达时间",
    "接收人": "接收人",
}

ID_PREFIX = "RCV"
TEMP_EMPTY = "未记录"
STATUS_READY = "已生成"
STATUS_HOLD = "挂起"

LEDGER_HEADER = ["线上单号", "运单号", "发运中心", "到达日期", "到达时间",
                 "实际箱数", "实际管数", "纸质箱数", "纸质管数",
                 "核对结论", "差异项数", "接收人", "确认状态", "生成时间"]
DIFF_HEADER = ["运单号", "线上单号", "问题类型", "说明"]


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
    """日志追加不覆盖：每次都从"运行主机"开头，方便分辨是哪台机器跑的。"""
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")
    print(text)


# ----------------------------------------------------------------- ② 读数据（归堆）

def read_paper_forms():
    """读「人工补录表」——一行一张纸质单，字段是从纸面上抄下来的。

    用 csv 直接读，不经过 pandas：运单号这类带前导零的编号，
    走 pandas 会默认猜成 int，0000712 变 712，事后 astype 也补不回来。
    真要上 pandas，必须 read_csv(dtype=str)。
    """
    forms = []
    with open(PAPER_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("运单号") or "").strip():
                continue        # 表尾说明行：第一列留空，这里被跳过
            forms.append({k: (v or "").strip() for k, v in row.items()})
    return forms


def read_details():
    """读「扫码导出明细」——一管一行，一趟货几十上百行。"""
    rows = []
    with open(DETAIL_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("运单号") or "").strip():
                continue        # 表尾合计行：同样靠第一列留空跳过
            rows.append({k: (v or "").strip() for k, v in row.items()})
    return rows


def group_details(rows):
    """按运单号归堆——纸面是一单一行，明细是一管一行，两边粒度不同，先各归各的。"""
    bucket = {}
    for r in rows:
        bucket.setdefault(r["运单号"], []).append(r)
    return bucket


# ----------------------------------------------------------------- ③ 核对与编号

def check_one(form, rows):
    """先判能不能出单，再逐项比纸面数字与明细实际条数。

    返回 (状态, 汇总, 问题列表)。问题列表里每条是 (类型, 说明)。
    """
    waybill = form["运单号"]
    issues = []

    if not rows:
        return STATUS_HOLD, None, [("明细缺失", "扫码导出里查不到这个运单号，先确认货有没有到")]
    if not form.get("到达日期"):
        return STATUS_HOLD, None, [("日期缺失", "纸质单上到达日期没填，编不出线上单号")]
    if not form.get("接收人"):
        return STATUS_HOLD, None, [("接收人缺失", "没有接收人就没人确认，单子悬空")]

    summary = {
        "实际箱数": len({r["箱号"] for r in rows}),
        "实际管数": len(rows),
        "样本类型分布": Counter(r["样本类型"] or "未填" for r in rows),
    }

    for col, name in (("纸质登记箱数", "箱数"), ("纸质登记管数", "管数")):
        paper = form.get(col, "").strip()
        actual = summary[f"实际{name}"]
        if paper and paper.isdigit() and int(paper) != actual:
            issues.append((f"{name}不符", f"纸质单写 {paper}，明细实际 {actual}"))

    if not form.get("运输温度上限") and not form.get("运输温度下限"):
        issues.append(("温度未记录", "纸质单上温度栏空着，线上单标未记录，回头补"))

    blank = summary["样本类型分布"].get("未填", 0)
    if blank:
        issues.append(("样本类型缺失", f"{blank} 管明细里样本类型是空的"))

    return STATUS_READY, summary, issues


def assign_ids(forms):
    """给能出单的编线上单号：RCV-<中心代码>-<到达日期>-<当天该中心的第几单>。

    只在同一天、同一中心内排序编流水，号短、能一眼看出是哪天的哪一批。
    """
    grouped = {}
    for form in forms:
        if form["状态"] != STATUS_READY:
            continue
        code = CENTER_CODE.get(form["发运中心"], "")
        if not code:
            form["状态"] = STATUS_HOLD
            form["问题"].append(("中心代码缺失", f"代码表里没有「{form['发运中心']}」"))
            continue
        grouped.setdefault((form["到达日期"], code), []).append(form)

    for (date, code), group in grouped.items():
        group.sort(key=lambda f: f["运单号"])
        for i, form in enumerate(group, 1):
            form["线上单号"] = f"{ID_PREFIX}-{code}-{date.replace('-', '')}-{i:02d}"


# ----------------------------------------------------------------- ④ 输出

def label_rows(ws):
    """把 A 列的文字读成 {栏目名: 行号}——按名字反查位置，不写死坐标。"""
    out = {}
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        if isinstance(cell.value, str) and cell.value.strip():
            out[cell.value.strip()] = cell.row
    return out


def paper_vs_actual(form):
    """纸面登记的数字与明细实际条数并排写出来——接收人签确认前一眼就能看到差在哪。"""
    pb, pc = form.get("纸质登记箱数", ""), form.get("纸质登记管数", "")
    ab, ac = form["汇总"]["实际箱数"], form["汇总"]["实际管数"]
    same = (not pb or int(pb) == ab) and (not pc or int(pc) == ac)
    text = f"纸质 {pb or '—'} 箱/{pc or '—'} 管  →  实际 {ab} 箱/{ac} 管"
    return text + ("，数量一致" if same else "，数量不符")


def fill_form(form, template, target):
    """按栏目名往模板的空值格（B 列）填，原版式一律不动。"""
    from openpyxl import load_workbook

    wb = load_workbook(template)
    ws = wb.active
    rows = label_rows(ws)

    lo, hi = form.get("运输温度下限", ""), form.get("运输温度上限", "")
    temp = f"{lo} ~ {hi} ℃" if (lo or hi) else TEMP_EMPTY
    types = " / ".join(f"{k} {v}" for k, v in form["汇总"]["样本类型分布"].items())

    values = {
        "线上单号": form["线上单号"],
        "运单号": form["运单号"],
        "箱数": form["汇总"]["实际箱数"],
        "管数": form["汇总"]["实际管数"],
        "运输温度": temp,
        "样本类型分布": types,
        "纸面对账": paper_vs_actual(form),
        "核对结论": "一致" if not form["问题"] else "有差异",
        "确认状态": "待确认",
        "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    for paper_col, online_col in PAPER_TO_ONLINE.items():
        values[online_col] = form.get(paper_col, "")

    for label, value in values.items():
        if label not in rows:
            raise KeyError(f"模板里找不到栏目「{label}」，模板被改过？")
        ws.cell(row=rows[label], column=2, value=value)

    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target)


def write_outputs(forms):
    """写台账与差异清单。台账一行一单（挂起的也在），差异清单只记有问题的。"""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    with open(LEDGER_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(LEDGER_HEADER)
        for form in forms:
            summary = form["汇总"] or {}
            s = form["状态"]
            w.writerow([
                form["线上单号"], form["运单号"], form["发运中心"],
                form["到达日期"], form["到达时间"],
                summary.get("实际箱数", ""), summary.get("实际管数", ""),
                form["纸质登记箱数"], form["纸质登记管数"],
                s if s == STATUS_HOLD else ("一致" if not form["问题"] else "有差异"),
                len(form["问题"]), form["接收人"], "—" if s == STATUS_HOLD else "待确认",
                "" if s == STATUS_HOLD else stamp,
            ])

    with open(DIFF_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(DIFF_HEADER)
        for form in forms:
            for kind, note in form["问题"]:
                w.writerow([form["运单号"], form["线上单号"], kind, note])


def write_forms(forms):
    """挂起的单一份都不出——宁可少几份，也不能把对不上的数据灌进线上台账。"""
    if FORM_DIR.exists():
        for old in FORM_DIR.glob("*.xlsx"):
            old.unlink()
    count = 0
    for form in forms:
        if form["状态"] != STATUS_READY:
            continue
        fill_form(form, TEMPLATE, FORM_DIR / f"接收单_{form['运单号']}.xlsx")
        count += 1
    return count


# ----------------------------------------------------------------- 回读校验

def verify(forms):
    from openpyxl import load_workbook

    checks = []

    def ok(name, cond, detail=""):
        checks.append((name, bool(cond), detail))

    ledger = list(csv.DictReader(open(LEDGER_CSV, encoding="utf-8-sig")))
    summary_rows = {r["运单号"]: r for r in ledger}
    diffs = list(csv.DictReader(open(DIFF_CSV, encoding="utf-8-sig")))

    ok("台账行数 = 纸质单张数", len(ledger) == len(forms), f"{len(ledger)} vs {len(forms)}")
    ok("运单号前导零未丢", all(len(r["运单号"]) == 7 for r in ledger),
       str(sorted({len(r["运单号"]) for r in ledger})))
    ok("台账顺序与源表一致", [r["运单号"] for r in ledger] == [f["运单号"] for f in forms])

    ready = [f for f in forms if f["状态"] == STATUS_READY]
    hold = [f for f in forms if f["状态"] == STATUS_HOLD]
    files = sorted(p.name for p in FORM_DIR.glob("*.xlsx"))
    ok("出单份数 = 非挂起单数", len(files) == len(ready), f"{len(files)} vs {len(ready)}")
    ok("挂起单未出单", all(f'接收单_{f["运单号"]}.xlsx' not in files for f in hold))
    ok("全部单都已给号或缺号", all(bool(f["线上单号"]) == (f["状态"] == STATUS_READY) for f in forms))

    ids = [f["线上单号"] for f in ready]
    ok("线上单号无重号", len(set(ids)) == len(ids))

    for form in ready:
        ws = load_workbook(FORM_DIR / f'接收单_{form["运单号"]}.xlsx').active
        row = label_rows(ws)
        got = {"线上单号": ws.cell(row=row["线上单号"], column=2).value,
               "运单号": ws.cell(row=row["运单号"], column=2).value,
               "管数": ws.cell(row=row["管数"], column=2).value,
               "核对结论": ws.cell(row=row["核对结论"], column=2).value,
               "接收人": ws.cell(row=row["接收人"], column=2).value}
        expect = {"线上单号": form["线上单号"], "运单号": form["运单号"],
                  "管数": form["汇总"]["实际管数"], "接收人": form["接收人"],
                  "核对结论": "一致" if not form["问题"] else "有差异"}
        for key in expect:
            ok(f'{form["运单号"]} 线上单.{key}', got[key] == expect[key], f"{got[key]!r} vs {expect[key]!r}")

    ok("台账管数与明细条数一致",
       all(int(summary_rows[f["运单号"]]["实际管数"]) == len(f["明细"]) for f in ready))
    ok("有差异的单全都进了差异清单",
       all(any(d["运单号"] == f["运单号"] for d in diffs) for f in forms if f["问题"]))
    ok("一致的单一条都不在差异清单",
       not any(d["运单号"] in {f["运单号"] for f in forms if not f["问题"]} for d in diffs))
    ok("差异清单条数 = 问题总数", len(diffs) == sum(len(f["问题"]) for f in forms),
       f'{len(diffs)} vs {sum(len(f["问题"]) for f in forms)}')
    ok("差异清单每条都指得出单号",
       all(d["运单号"] in summary_rows for d in diffs)
       and all((d["线上单号"] or summary_rows[d["运单号"]]["核对结论"] == STATUS_HOLD) for d in diffs))

    return checks


# ----------------------------------------------------------------- main

def main():
    t0 = time.time()
    log_line(f"运行主机：{socket.gethostname()}（{local_ip()}）")
    log_line(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
             f"Python {sys.version.split()[0]} / {Path(sys.executable).name}")

    forms = read_paper_forms()
    details = group_details(read_details())
    print(f"纸质单 {len(forms)} 张，明细 {sum(len(v) for v in details.values())} 管，"
          f"覆盖 {len(details)} 个运单号")

    for form in forms:
        form["明细"] = details.get(form["运单号"], [])
        form["线上单号"] = ""
        form["状态"], form["汇总"], form["问题"] = check_one(form, form["明细"])

    assign_ids(forms)
    made = write_forms(forms)
    write_outputs(forms)

    checks = verify(forms)
    for name, good, detail in checks:
        print(f"  [{'OK ' if good else 'FAIL'}] {name}" + (f"   {detail}" if not good else ""))

    passed = all(c[1] for c in checks)
    hold = [f["运单号"] for f in forms if f["状态"] == STATUS_HOLD]
    print(f"\n出单 {made} 份，挂起 {len(hold)} 张（{'、'.join(hold) or '无'}），"
          f"回读校验 {sum(c[1] for c in checks)}/{len(checks)} 项通过")
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
